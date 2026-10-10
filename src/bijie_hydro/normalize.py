from __future__ import annotations

import math
import re
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any
from zoneinfo import ZoneInfo

from .errors import QualityError


CHINA_TZ = ZoneInfo("Asia/Shanghai")
MISSING_TEXT = {"", "--", "—", "-", "null", "none", "nan"}


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def parse_observation_time(value: Any, collected_at: datetime) -> datetime:
    text = clean_text(value)
    short_time = bool(re.fullmatch(r"\d{2}-\d{2}\d{2}:\d{2}", text))
    if short_time:
        text = f"{collected_at.astimezone(CHINA_TZ).year}-{text}"
    formats = ("%Y-%m-%d%H:%M:%S", "%Y-%m-%d%H:%M")
    parsed: datetime | None = None
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)
            if short_time:
                if parsed - collected_at.astimezone(CHINA_TZ).replace(tzinfo=None) > timedelta(days=30):
                    parsed = parsed.replace(year=parsed.year - 1)
            break
        except ValueError:
            continue
    if parsed is None:
        raise QualityError(f"无法解析观测时间: {value!r}")
    return parsed.replace(tzinfo=CHINA_TZ)


def parse_measurement(value: Any) -> tuple[Decimal | None, str | None]:
    if value is None:
        return None, "空值"
    text = clean_text(value).lower()
    if text in MISSING_TEXT:
        return None, "缺测标记"
    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError):
        return None, f"非数值:{value}"
    if not number.is_finite() or not math.isfinite(float(number)) or abs(number) > Decimal("10000000"):
        return None, f"非有限数值:{value}"
    if number == 0:
        return None, "接口零值按缺测处理"
    return number, None

