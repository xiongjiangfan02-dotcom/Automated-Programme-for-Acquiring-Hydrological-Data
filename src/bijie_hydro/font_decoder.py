from __future__ import annotations

import hashlib
import io
import re
import unicodedata
from pathlib import Path
from typing import Any, Iterable

from .errors import DecodeError
from .http_client import HttpClient


_TAG_PATTERN = re.compile(r"#(?P<font>[A-Za-z0-9_]+)otltag(?P<text>.*?)#FontTag", re.DOTALL)


def extract_font_ids(value: Any) -> set[str]:
    if not isinstance(value, str):
        return set()
    return {match.group("font") for match in _TAG_PATTERN.finditer(value)}


def _glyph_signature(glyph_set: Any, glyph_name: str, units_per_em: int = 1000) -> str:
    try:
        from fontTools.pens.recordingPen import DecomposingRecordingPen
    except ImportError as exc:
        raise DecodeError("缺少fonttools依赖，无法解码官网动态字体") from exc
    pen = DecomposingRecordingPen(glyph_set)
    glyph_set[glyph_name].draw(pen)
    normalized = []
    for command, points in pen.value:
        rounded = tuple(None if point is None else tuple(round(float(value) / units_per_em, 6) for value in point) for point in points)
        normalized.append((command, rounded))
    return hashlib.sha256(repr(normalized).encode("utf-8")).hexdigest()


def _load_font(source: str | Path | bytes):
    try:
        from fontTools.ttLib import TTCollection, TTFont
    except ImportError as exc:
        raise DecodeError("缺少fonttools依赖，无法解码官网动态字体") from exc
    if isinstance(source, bytes):
        return [TTFont(io.BytesIO(source), lazy=False)]
    path = Path(source)
    if not path.is_file():
        return []
    if path.suffix.lower() == ".ttc":
        return list(TTCollection(path, lazy=False).fonts)
    return [TTFont(path, lazy=False)]


class ReferenceGlyphIndex:
    def __init__(self, font_paths: Iterable[str | Path]):
        self.by_signature: dict[str, set[str]] = {}
        loaded = 0
        for path in font_paths:
            for font in _load_font(path):
                loaded += 1
                self._add_font(font)
                font.close()
        if loaded == 0:
            raise DecodeError("没有可用的参考中文字体，请检查font_decoder.reference_fonts")

    def _add_font(self, font: Any) -> None:
        cmap = font.getBestCmap() or {}
        glyph_set = font.getGlyphSet()
        for codepoint, glyph_name in cmap.items():
            if codepoint < 32:
                continue
            try:
                signature = _glyph_signature(glyph_set, glyph_name, font["head"].unitsPerEm)
            except Exception:
                continue
            self.by_signature.setdefault(signature, set()).add(chr(codepoint))

    def resolve(self, signature: str) -> str | None:
        candidates = self.by_signature.get(signature, set())
        return next(iter(candidates)) if len(candidates) == 1 else None


class DynamicFontDecoder:
    def __init__(self, client: HttpClient, base_url: str, cache_dir: Path, reference_fonts: Iterable[str | Path], offline: bool = False):
        self.client = client
        self.base_url = base_url.rstrip("/")
        self.cache_dir = cache_dir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.reference_fonts = list(reference_fonts)
        self.references: ReferenceGlyphIndex | None = None
        self._mappings: dict[str, dict[str, str]] = {}
        self._failures: dict[str, str] = {}
        self.offline = offline
        self.used_fonts: dict[str, bytes] = {}

    def _references(self) -> ReferenceGlyphIndex:
        if self.references is None:
            self.references = ReferenceGlyphIndex(self.reference_fonts)
        return self.references

    def _font_bytes(self, font_id: str) -> bytes:
        cached = self.cache_dir / f"{font_id}.ttf"
        if cached.is_file() and cached.stat().st_size > 0:
            return cached.read_bytes()
        if self.offline:
            raise DecodeError(f"离线重放缺少缓存字体: {font_id}")
        response = self.client.request(f"{self.base_url}/ttf/{font_id}.ttf")
        if response.status != 200 or not response.body:
            raise DecodeError(f"动态字体下载失败: {font_id}")
        cached.write_bytes(response.body)
        return response.body

    def _mapping(self, font_id: str) -> dict[str, str]:
        if font_id in self._failures:
            raise DecodeError(self._failures[font_id])
        try:
            return self._build_mapping(font_id)
        except Exception as exc:
            self._failures[font_id] = f"字体{font_id}无法解码: {exc}"
            raise DecodeError(self._failures[font_id]) from exc

    def _build_mapping(self, font_id: str) -> dict[str, str]:
        if font_id in self._mappings:
            return self._mappings[font_id]
        font_bytes = self._font_bytes(font_id)
        self.used_fonts[font_id] = font_bytes
        fonts = _load_font(font_bytes)
        if not fonts:
            raise DecodeError(f"无法加载动态字体: {font_id}")
        font = fonts[0]
        cmap = font.getBestCmap() or {}
        glyph_set = font.getGlyphSet()
        mapping: dict[str, str] = {}
        for codepoint, glyph_name in cmap.items():
            resolved = self._references().resolve(_glyph_signature(glyph_set, glyph_name, font["head"].unitsPerEm))
            if resolved:
                mapping[chr(codepoint)] = resolved
        font.close()
        self._mappings[font_id] = mapping
        return mapping

    def decode(self, value: Any) -> Any:
        if isinstance(value, str) and "\ufffd" in value:
            raise DecodeError("响应含Unicode替换字符，疑似编码损坏")
        if not isinstance(value, str) or "otltag" not in value:
            return value

        unresolved: list[str] = []

        def replace(match: re.Match[str]) -> str:
            font_id = match.group("font")
            text = match.group("text")
            mapping = self._mapping(font_id)
            output: list[str] = []
            for character in text:
                if character in mapping:
                    output.append(mapping[character])
                elif character.isascii() or character.isspace() or unicodedata.category(character).startswith(("P", "S")):
                    output.append(character)
                else:
                    unresolved.append(f"U+{ord(character):04X}")
                    output.append(character)
            return "".join(output)

        decoded = _TAG_PATTERN.sub(replace, value)
        if "otltag" in decoded or "#FontTag" in decoded or unresolved:
            detail = ",".join(sorted(set(unresolved))[:20])
            raise DecodeError(f"动态字体存在未解析字符: {detail or decoded[:80]}")
        return decoded

