from __future__ import annotations

from datetime import datetime, timedelta
from decimal import Decimal

from .models import CheckResult, ReservoirRecord


REQUIRED_SOURCE_FIELDS = {"bsnm", "addvnm", "rvnm", "stnm", "tm", "rz", "rzRange", "idNo"}


def check_source_structure(records: list[dict], minimum: int, maximum: int) -> list[CheckResult]:
    checks: list[CheckResult] = []
    count = len(records)
    checks.append(
        CheckResult(
            "全国记录数",
            "正常" if minimum <= count <= maximum else "异常",
            f"记录数={count}, 允许范围={minimum}..{maximum}",
            not minimum <= count <= maximum,
        )
    )
    missing = sorted({field for item in records for field in REQUIRED_SOURCE_FIELDS - set(item)})
    checks.append(
        CheckResult(
            "字段结构",
            "正常" if not missing else "异常",
            "必需字段完整" if not missing else f"缺少字段: {','.join(missing)}",
            bool(missing),
        )
    )
    wrong_types = []
    text_fields = {"bsnm", "addvnm", "rvnm", "stnm", "tm"}
    for index, item in enumerate(records, 1):
        for field in REQUIRED_SOURCE_FIELDS:
            value = item.get(field)
            allowed = (str, type(None)) if field in text_fields else (str, int, float, type(None))
            if not isinstance(value, allowed) or isinstance(value, bool):
                wrong_types.append(f"{index}:{field}")
    checks.append(CheckResult("字段类型", "异常" if wrong_types else "正常", f"非法类型={wrong_types[:20]}", bool(wrong_types)))
    return checks


def check_target_records(
    records: list[ReservoirRecord],
    collected_at: datetime,
    max_age_hours: int,
    max_future_minutes: int,
    max_level_change: Decimal,
    require_one: bool,
) -> list[CheckResult]:
    checks: list[CheckResult] = []
    checks.append(
        CheckResult(
            "目标记录数",
            "正常" if records or not require_one else "无数据",
            f"毕节大型水库记录数={len(records)}",
            require_one and not records,
        )
    )
    keys = [(item.id_no or item.reservoir_name, item.observation_time) for item in records]
    duplicate_count = len(keys) - len(set(keys))
    checks.append(CheckResult("重复记录", "正常" if duplicate_count == 0 else "异常", f"重复数={duplicate_count}", duplicate_count > 0))
    too_old = [item.reservoir_name for item in records if collected_at - item.observation_time > timedelta(hours=max_age_hours)]
    future = [item.reservoir_name for item in records if item.observation_time - collected_at > timedelta(minutes=max_future_minutes)]
    checks.append(CheckResult("观测时间时效", "正常" if not too_old and not future else "异常", f"超时={too_old}; 未来时间={future}", bool(too_old or future)))
    extreme = [item.reservoir_name for item in records if item.daily_change is not None and abs(item.daily_change) > max_level_change]
    checks.append(CheckResult("日变幅范围", "正常" if not extreme else "需复核", f"超过阈值的水库={extreme}", bool(extreme)))
    missing_text = [item.reservoir_name for item in records if not item.basin or not item.river_name or not item.administrative_area]
    missing_level = [item.reservoir_name for item in records if item.water_level is None]
    missing_change = [item.reservoir_name for item in records if item.daily_change is None]
    checks.append(CheckResult("缺失字段", "警告" if missing_text or missing_level or missing_change else "正常",
                              f"流域/河名/区划缺失={missing_text}; 水位缺测={missing_level}; 日变幅缺测={missing_change}", False))
    return checks


def check_historical_delta(records: list[ReservoirRecord], previous: dict, maximum: Decimal) -> CheckResult:
    previous_records = previous.get("records", {}) if isinstance(previous, dict) else {}
    extreme: list[str] = []
    for record in records:
        key = record.id_no or record.reservoir_name
        old = previous_records.get(key)
        if not isinstance(old, dict) or record.water_level is None or old.get("water_level") in (None, ""):
            continue
        try:
            difference = abs(record.water_level - Decimal(str(old["water_level"])))
        except Exception:
            continue
        if difference > maximum:
            extreme.append(f"{record.reservoir_name}:{difference}")
    return CheckResult(
        "历史水位差异",
        "正常" if not extreme else "需复核",
        "无超过阈值记录" if not extreme else f"超过{maximum}米: {','.join(extreme)}",
        bool(extreme),
    )


def check_historical_time(records: list[ReservoirRecord], previous: dict) -> CheckResult:
    rollbacks = []
    for record in records:
        old = previous.get("records", {}).get(record.id_no or record.reservoir_name, {})
        if old.get("observation_time"):
            try:
                old_time = datetime.fromisoformat(old["observation_time"])
                if record.observation_time < old_time:
                    rollbacks.append(record.reservoir_name)
            except (ValueError, TypeError):
                return CheckResult("历史观测时间", "异常", "历史状态时间字段损坏", True)
    return CheckResult("历史观测时间", "需复核" if rollbacks else "正常", f"观测时间回退={rollbacks}", bool(rollbacks))


def has_blocking_failure(checks: list[CheckResult]) -> bool:
    return any(check.status in {"异常", "需复核"} for check in checks)

