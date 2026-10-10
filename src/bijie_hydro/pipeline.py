from __future__ import annotations

import json
import hashlib
import re
import uuid
from dataclasses import asdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

from . import __version__
from .archive import ArchiveStore
from .config import Settings
from .errors import CollectorError, DecodeError, QualityError
from .excel_writer import validate_template, write_workbook
from .filtering import matches_province, normalize_province_names
from .font_decoder import DynamicFontDecoder
from .http_client import HttpClient, discover_endpoint, parse_api_payload, validate_source_url
from .locking import collection_lock
from .models import BatchLog, CheckResult, RawRecord, ReservoirRecord
from .normalize import clean_text, parse_measurement, parse_observation_time
from .quality import check_historical_delta, check_historical_time, check_source_structure, check_target_records, has_blocking_failure
from .state import StateStore, record_key, snapshot_fingerprint


CHINA_TZ = ZoneInfo("Asia/Shanghai")


def _batch_id(now: datetime) -> str:
    return f"{now:%Y%m%dT%H%M%S}_{uuid.uuid4().hex[:8]}"


def _base_url(url: str) -> str:
    match = re.match(r"^(https?://[^/]+)", url)
    if not match:
        raise QualityError(f"非法来源URL: {url}")
    return match.group(1)


class CollectionPipeline:
    def __init__(self, settings: Settings):
        self.settings = settings
        source = settings.section("source")
        self.client = HttpClient(
            timeout=int(source.get("timeout_seconds", 30)),
            max_attempts=int(source.get("max_attempts", 3)),
            retry_base_seconds=float(source.get("retry_base_seconds", 2.0)),
            user_agent=str(source.get("user_agent", "GuizhouHydroCollector/0.2")),
        )

    def run(self, input_json: Path | None = None) -> BatchLog:
        state_path = self.settings.resolve_path(self.settings.section("state")["file"])
        with collection_lock(state_path.with_suffix(".lock")):
            self.client.total_attempts = 0
            return self._run_locked(input_json)

    def _run_locked(self, input_json: Path | None = None) -> BatchLog:
        started_at = datetime.now(CHINA_TZ)
        batch_id = _batch_id(started_at)
        paths = self.settings.section("paths")
        source = self.settings.section("source")
        quality = self.settings.section("quality")
        output = self.settings.section("output")
        state_config = self.settings.section("state")
        filter_config = self.settings.section("filter")
        font_config = self.settings.section("font_decoder")
        archive = ArchiveStore(self.settings.resolve_path(paths["archive_root"]), batch_id)
        log = BatchLog(
            batch_id=batch_id,
            status="运行中",
            started_at=started_at,
            source_page=str(source["page_url"]),
            version=__version__,
            report_date=started_at.strftime("%Y%m%d"),
        )

        try:
            template = self.settings.resolve_path(paths["template"])
            template_info = validate_template(template)
            archive.save_json("template_validation.json", template_info)
            province_names = normalize_province_names(filter_config.get("province_names", []))
            archive.save_json("region_filter.json", {
                "mode": "province",
                "province_names": sorted(province_names),
            })
            archive.save_json("provenance.json", {
                "batch_id": batch_id, "mode": "offline_replay" if input_json else "online",
                "input_json": str(input_json) if input_json else None,
                "collected_at": started_at.isoformat(), "timezone": "Asia/Shanghai",
                "template_sha256": hashlib.sha256(template.read_bytes()).hexdigest(),
                "filter_mode": "province", "province_names": sorted(province_names), "processor_version": __version__,
                "short_time_year_policy": "优先官网完整年份；仅无年份时以采集年份推定，跨年回退并受时效检查约束",
            })

            endpoint_override = source.get("endpoint_override")
            if input_json is not None:
                endpoint = str(endpoint_override or source["page_url"])
            else:
                page_response = self.client.request(str(source["page_url"]))
                log.http_status = page_response.status
                archive.save_bytes("source_page.html", page_response.body)
                html = page_response.body.decode("utf-8-sig", errors="strict")
                if "大型水库" not in html:
                    raise QualityError("来源页面不再包含大型水库标识，需要复核页面结构")
                endpoint = endpoint_override or discover_endpoint(str(source["page_url"]), html)
            endpoint = validate_source_url(endpoint, str(source["page_url"]))
            log.endpoint_url = endpoint

            if input_json is not None:
                response_body = input_json.read_bytes()
                archive.save_bytes("offline_response.json", response_body)
            else:
                api_response = self.client.request(endpoint, method="POST", body=b"", referer=str(source["page_url"]))
                log.http_status = api_response.status
                response_body = api_response.body
            archive.save_bytes("raw_response.json", response_body)
            source_records, payload = parse_api_payload(response_body)
            log.national_count = len(source_records)
            log.checks.extend(
                check_source_structure(
                    source_records,
                    int(quality.get("min_national_records", 100)),
                    int(quality.get("max_national_records", 2000)),
                )
            )
            if has_blocking_failure(log.checks):
                raise QualityError("接口字段结构或全国记录数检查失败")

            decoder = DynamicFontDecoder(
                client=self.client,
                base_url=_base_url(str(source["page_url"])),
                cache_dir=self.settings.resolve_path(paths["font_cache"]),
                reference_fonts=[self.settings.resolve_path(value) if not Path(value).is_absolute() else Path(value) for value in font_config.get("reference_fonts", [])],
                offline=input_json is not None,
            )
            required_decode_fields = tuple(font_config.get("required_fields", []))
            raw_records: list[RawRecord] = []
            for index, values in enumerate(source_records, start=1):
                raw = RawRecord(index=index, values=values)
                for field in required_decode_fields:
                    try:
                        raw.decoded[field] = decoder.decode(values.get(field))
                    except DecodeError as exc:
                        raw.decode_errors.append(f"{field}:{exc}")
                raw.decoded["tm"] = values.get("tm")
                raw_records.append(raw)
            decode_error_count = sum(bool(item.decode_errors) for item in raw_records)
            critical_decode_errors = []
            for item in raw_records:
                if not item.decode_errors:
                    continue
                area = clean_text(item.decoded.get("addvnm"))
                area_failed = any(error.startswith("addvnm:") for error in item.decode_errors)
                if area_failed or matches_province(area, province_names):
                    critical_decode_errors.append(item.index)
            for font_id, body in decoder.used_fonts.items():
                archive.save_bytes(f"font_{font_id}.ttf", body)
            log.checks.append(
                CheckResult(
                    "动态字体解码",
                    "正常" if decode_error_count == 0 else ("异常" if critical_decode_errors else "警告"),
                    f"全国解码失败记录数={decode_error_count}; 目标相关失败序号={critical_decode_errors[:20]}",
                    bool(critical_decode_errors),
                )
            )
            archive.save_json(
                "decoded_records.json",
                [{"index": item.index, "decoded": item.decoded, "errors": item.decode_errors} for item in raw_records],
            )
            if critical_decode_errors and font_config.get("fail_on_unresolved", True):
                raise DecodeError(f"目标相关动态字体解码失败，记录序号={critical_decode_errors[:20]}")

            target_records: list[ReservoirRecord] = []
            for raw in raw_records:
                if not matches_province(raw.decoded.get("addvnm"), province_names):
                    raw.filter_status = "非贵州省记录"
                    continue
                reservoir_name = clean_text(raw.decoded.get("stnm"))
                raw.filter_status = "贵州省记录"
                observation_time = parse_observation_time(raw.decoded.get("tm"), started_at)
                water_level, water_missing = parse_measurement(raw.decoded.get("rz"))
                daily_change, change_missing = parse_measurement(raw.decoded.get("rzRange"))
                for field, reason in (("库水位", water_missing), ("日变幅", change_missing)):
                    if reason and reason.startswith(("非数值", "非有限")):
                        log.checks.append(CheckResult("数值有效性", "异常", f"{reservoir_name or '未命名水库'}/{field}:{reason}", True))
                if water_missing:
                    raw.missing_reasons.append(f"库水位:{water_missing}")
                if change_missing:
                    raw.missing_reasons.append(f"日变幅:{change_missing}")
                administrative_area = clean_text(raw.decoded.get("addvnm"))
                target_records.append(
                    ReservoirRecord(
                        basin=clean_text(raw.decoded.get("bsnm")),
                        administrative_area=administrative_area,
                        river_name=clean_text(raw.decoded.get("rvnm")),
                        reservoir_name=reservoir_name,
                        observation_time=observation_time,
                        water_level=water_level,
                        daily_change=daily_change,
                        id_no=clean_text(raw.decoded.get("idNo")),
                        source_index=raw.index,
                    )
                )

            original_count = len(target_records)
            target_records = self._deduplicate(target_records)
            log.checks.append(CheckResult("快照去重", "正常", f"候选={original_count}; 保留每库最新观测={len(target_records)}"))
            log.target_count = len(target_records)
            archive.save_json("parsed_records.json", [asdict(item) for item in raw_records])
            archive.save_json("selected_records.json", [asdict(item) for item in target_records])
            log.checks.extend(
                check_target_records(
                    target_records,
                    started_at,
                    int(quality.get("max_observation_age_hours", 72)),
                    int(quality.get("max_future_minutes", 30)),
                    Decimal(str(quality.get("max_level_change_meters", 20))),
                    bool(quality.get("require_at_least_one_target", True)),
                )
            )
            if not target_records and quality.get("require_at_least_one_target", True):
                log.status = "无匹配数据"
                log.finished_at = datetime.now(CHINA_TZ)
                # Upstream success is distinct from an unreachable page or invalid payload.
                return log
            state_store = StateStore(self.settings.resolve_path(state_config["file"]))
            previous_state = state_store.load()
            log.checks.append(check_historical_time(target_records, previous_state))
            previous_keys = set(previous_state.get("records", {}))
            current_keys = {record_key(item) for item in target_records}
            lost = previous_keys - current_keys
            log.checks.append(CheckResult("历史覆盖变化", "需复核" if lost else "正常", f"较上次缺少站点={sorted(lost)}", bool(lost)))
            log.checks.append(
                check_historical_delta(
                    target_records,
                    previous_state,
                    Decimal(str(quality.get("max_water_level_delta_meters", 20))),
                )
            )
            archive.save_json("quality_checks.json", [asdict(check) for check in log.checks])
            if has_blocking_failure(log.checks):
                raise QualityError("目标数据质量检查失败")

            fingerprint = snapshot_fingerprint(target_records)
            if bool(state_config.get("skip_duplicate_snapshot", True)) and fingerprint == previous_state.get("fingerprint"):
                log.status = "重复跳过"
                log.output_count = 0
                log.finished_at = datetime.now(CHINA_TZ)
                return log

            report_date = started_at.strftime("%Y%m%d")
            log.report_date = report_date
            output_root = self.settings.resolve_path(paths["output_root"])
            file_name = str(output["file_name"]).format(report_date=report_date, batch_id=batch_id)
            if Path(file_name).name != file_name or not file_name.endswith(".xlsx"):
                raise QualityError("输出文件名必须是单个xlsx文件名，不能包含路径")
            destination = output_root / report_date / file_name
            log.status = "成功"
            log.finished_at = datetime.now(CHINA_TZ)
            log.output_file = str(destination)
            log.output_count = len(target_records)
            log.attempts = self.client.total_attempts
            write_workbook(
                template=template,
                destination=destination,
                records=target_records,
                raw_records=raw_records,
                log=log,
                source_url=endpoint,
                collected_at=started_at,
                date_format=str(output.get("date_format", "yyyy-mm-dd hh:mm")),
                number_format=str(output.get("number_format", "0.####")),
                raw_sheet_name=str(output.get("raw_sheet", "原始数据")),
                log_sheet_name=str(output.get("log_sheet", "采集日志")),
            )
            state_store.save(batch_id, fingerprint, target_records)
            return log
        except Exception as exc:
            log.status = "待复核" if any(check.status == "需复核" for check in log.checks) else "失败"
            if log.output_file and not Path(log.output_file).is_file():
                log.output_file = None
                log.output_count = 0
            log.errors.append(f"{type(exc).__name__}: {exc}")
            log.finished_at = datetime.now(CHINA_TZ)
            if isinstance(exc, CollectorError):
                raise
            raise CollectorError(str(exc)) from exc
        finally:
            log.attempts = self.client.total_attempts
            archive.save_json("quality_checks.json", [asdict(check) for check in log.checks])
            archive.save_json("batch_log.json", log.to_dict())
            archive.finalize_manifest()

    @staticmethod
    def _deduplicate(records: list[ReservoirRecord]) -> list[ReservoirRecord]:
        chosen: dict[tuple[str, datetime], ReservoirRecord] = {}
        for record in records:
            key = (record_key(record), record.observation_time)
            previous = chosen.get(key)
            if previous and any(getattr(previous, field) != getattr(record, field) for field in ("basin", "administrative_area", "river_name", "reservoir_name", "water_level", "daily_change")):
                raise QualityError(f"同站同观测时间记录冲突: {record.reservoir_name}")
            chosen.setdefault(key, record)
        latest: dict[str, ReservoirRecord] = {}
        for record in chosen.values():
            key = record_key(record)
            old = latest.get(key)
            if old is None or record.observation_time > old.observation_time:
                latest[key] = record
        return sorted(latest.values(), key=lambda item: (item.administrative_area, item.reservoir_name, item.id_no))

