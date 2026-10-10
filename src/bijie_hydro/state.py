from __future__ import annotations

import hashlib
import json
import os
import re
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .models import ReservoirRecord
from .errors import QualityError


def canonical_measurement(value):
    if value is None:
        return None
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def snapshot_fingerprint(records: list[ReservoirRecord]) -> str:
    normalized = [
        {
            "key": record.id_no or record.reservoir_name,
            "name": record.reservoir_name,
            "basin": record.basin,
            "administrative_area": record.administrative_area,
            "river_name": record.river_name,
            "time": record.observation_time.isoformat(),
            "water_level": canonical_measurement(record.water_level),
            "daily_change": canonical_measurement(record.daily_change),
        }
        for record in sorted(records, key=lambda item: (item.id_no or item.reservoir_name, item.observation_time))
    ]
    value = json.dumps(normalized, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


class StateStore:
    def __init__(self, path: Path):
        self.path = path

    def load(self) -> dict:
        if not self.path.is_file():
            return {}
        try:
            value = json.loads(self.path.read_text(encoding="utf-8"))
            if not isinstance(value, dict) or not isinstance(value.get("records"), dict) or not value.get("fingerprint"):
                raise QualityError("历史状态结构异常，需要人工检查，禁止静默重置")
            if not isinstance(value["fingerprint"], str) or not re.fullmatch(r"[0-9a-f]{64}", value["fingerprint"]):
                raise QualityError("历史快照指纹损坏")
            try:
                for key, item in value["records"].items():
                    if not isinstance(key, str) or not isinstance(item, dict):
                        raise ValueError("站点状态必须为对象")
                    if datetime.fromisoformat(item["observation_time"]).tzinfo is None:
                        raise ValueError("历史时间缺少时区")
                    for field in ("water_level", "daily_change"):
                        if item.get(field) is not None and not Decimal(str(item[field])).is_finite():
                            raise ValueError("历史数字无效")
            except (ValueError, TypeError, KeyError, InvalidOperation) as exc:
                raise QualityError(f"历史站点状态损坏: {exc}") from exc
            return value
        except (OSError, json.JSONDecodeError) as exc:
            raise QualityError(f"无法读取历史状态，需要人工检查: {exc}") from exc

    def save(self, batch_id: str, fingerprint: str, records: list[ReservoirRecord]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        data = {
            "batch_id": batch_id,
            "fingerprint": fingerprint,
            "records": {
                (record.id_no or record.reservoir_name): {
                    "reservoir_name": record.reservoir_name,
                    "observation_time": record.observation_time.isoformat(),
                    "water_level": str(record.water_level) if record.water_level is not None else None,
                    "daily_change": str(record.daily_change) if record.daily_change is not None else None,
                }
                for record in records
            },
        }
        temporary = self.path.with_suffix(self.path.suffix + ".tmp")
        try:
            temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(temporary, self.path)
        finally:
            temporary.unlink(missing_ok=True)
