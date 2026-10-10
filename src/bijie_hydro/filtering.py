from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import ConfigurationError, QualityError
from .normalize import clean_text


@dataclass(slots=True, frozen=True)
class AllowlistEntry:
    canonical_name: str
    aliases: frozenset[str]
    id_nos: frozenset[str]
    province: str
    city: str
    administrative_aliases: frozenset[str] = frozenset()


class ReservoirAllowlist:
    def __init__(self, entries: list[AllowlistEntry], version: str):
        self.entries = entries
        self.version = version
        self.by_id = {id_no: entry for entry in entries for id_no in entry.id_nos}
        self.by_name = {name: entry for entry in entries for name in entry.aliases}
        if len(self.by_id) != sum(len(entry.id_nos) for entry in entries) or len(self.by_name) != sum(len(entry.aliases) for entry in entries):
            raise ConfigurationError("白名单存在重复idNo或别名，无法唯一识别水库")

    @classmethod
    def load(cls, path: Path) -> "ReservoirAllowlist":
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            raise ConfigurationError(f"无法读取水库白名单: {exc}") from exc
        entries = []
        for item in data.get("reservoirs", []):
            if not item.get("enabled", True):
                continue
            if item.get("evidence_status") != "已核实" or not item.get("source_url"):
                raise ConfigurationError("启用的水库必须有已核实名录证据和source_url")
            canonical = clean_text(item.get("canonical_name"))
            if not canonical or clean_text(item.get("province")) != "贵州省" or clean_text(item.get("city")) != "毕节市":
                raise ConfigurationError("白名单仅允许有名称的贵州省毕节市水库")
            aliases = {clean_text(value) for value in item.get("aliases", []) if clean_text(value)}
            aliases.add(canonical)
            entries.append(
                AllowlistEntry(
                    canonical_name=canonical,
                    aliases=frozenset(aliases),
                    id_nos=frozenset(clean_text(value) for value in item.get("id_nos", []) if clean_text(value)),
                    province=clean_text(item.get("province")),
                    city=clean_text(item.get("city")),
                    administrative_aliases=frozenset(clean_text(value) for value in item.get("administrative_aliases", [])),
                )
            )
        if not entries:
            raise ConfigurationError("水库白名单为空")
        return cls(entries, str(data.get("version", "unknown")))

    def match(self, record: dict[str, Any], province_names: set[str], city_names: set[str]) -> AllowlistEntry | None:
        id_no = clean_text(record.get("idNo"))
        name = clean_text(record.get("stnm"))
        administrative_area = clean_text(record.get("addvnm"))
        if id_no and id_no in self.by_id:
            entry = self.by_id[id_no]
            if name not in entry.aliases or (administrative_area and not self._area_matches(entry, administrative_area, province_names, city_names)):
                raise QualityError(f"已绑定idNo的名称或行政区划冲突: {id_no}")
            return entry
        entry = self.by_name.get(name)
        if entry is None:
            return None
        if not administrative_area:
            return None
        if not self._area_matches(entry, administrative_area, province_names, city_names):
            return None
        return entry

    @staticmethod
    def _area_matches(entry, value, province_names, city_names):
        # Match controlled complete strings, never an arbitrary substring.
        allowed = {clean_text(name) for name in province_names | city_names}
        allowed.update(clean_text(p + c) for p in province_names for c in city_names)
        allowed.update(entry.administrative_aliases)
        return value in allowed

