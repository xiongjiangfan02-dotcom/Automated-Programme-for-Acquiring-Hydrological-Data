from __future__ import annotations

from collections.abc import Iterable

from .errors import ConfigurationError
from .normalize import clean_text


def normalize_province_names(values: Iterable[object]) -> frozenset[str]:
    names = frozenset(clean_text(value) for value in values if clean_text(value))
    if not names:
        raise ConfigurationError("必须配置至少一个贵州省名称")
    return names


def matches_province(administrative_area: object, province_names: Iterable[object]) -> bool:
    """Match a province value without arbitrary substring matching."""
    area = clean_text(administrative_area)
    if not area:
        return False
    names = normalize_province_names(province_names)
    if area in names:
        return True
    # Detailed administrative values such as 贵州省毕节市 remain in scope.
    return any(name.endswith("省") and area.startswith(name) for name in names)

