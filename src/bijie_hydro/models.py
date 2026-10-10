from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import datetime
from decimal import Decimal
from typing import Any


@dataclass(slots=True)
class RawRecord:
    index: int
    values: dict[str, Any]
    decoded: dict[str, Any] = field(default_factory=dict)
    decode_errors: list[str] = field(default_factory=list)
    filter_status: str = "未筛选"
    missing_reasons: list[str] = field(default_factory=list)


@dataclass(slots=True, frozen=True)
class ReservoirRecord:
    basin: str
    administrative_area: str
    river_name: str
    reservoir_name: str
    observation_time: datetime
    water_level: Decimal | None
    daily_change: Decimal | None
    id_no: str
    source_index: int


@dataclass(slots=True)
class CheckResult:
    name: str
    status: str
    message: str
    requires_review: bool = False


@dataclass(slots=True)
class BatchLog:
    batch_id: str
    status: str
    started_at: datetime
    finished_at: datetime | None = None
    source_page: str = ""
    endpoint_url: str = ""
    http_status: int | None = None
    attempts: int = 0
    national_count: int = 0
    target_count: int = 0
    output_count: int = 0
    report_date: str = ""
    checks: list[CheckResult] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    output_file: str | None = None
    version: str = "0.3.0"

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["started_at"] = self.started_at.isoformat()
        data["finished_at"] = self.finished_at.isoformat() if self.finished_at else None
        return data

