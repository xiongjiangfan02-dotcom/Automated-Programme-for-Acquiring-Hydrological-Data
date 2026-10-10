from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .errors import ConfigurationError


@dataclass(slots=True, frozen=True)
class Settings:
    data: dict[str, Any]
    config_path: Path

    def section(self, name: str) -> dict[str, Any]:
        value = self.data.get(name)
        if not isinstance(value, dict):
            raise ConfigurationError(f"缺少配置节: {name}")
        return value

    def resolve_path(self, value: str) -> Path:
        path = Path(value).expanduser()
        if path.is_absolute():
            return path
        return (self.config_path.parent.parent / path).resolve()


def load_settings(path: str | Path) -> Settings:
    config_path = Path(path).expanduser().resolve()
    if not config_path.is_file():
        raise ConfigurationError(f"配置文件不存在: {config_path}")
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ConfigurationError(f"无法读取配置文件: {exc}") from exc
    settings = Settings(data=data, config_path=config_path)
    for required in ("source", "paths", "font_decoder", "filter", "quality", "state", "output", "upload"):
        settings.section(required)
    if settings.section("upload").get("enabled") is not False:
        raise ConfigurationError("MVP阶段禁止启用上传，upload.enabled 必须为 false")
    source = settings.section("source")
    if not 1 <= int(source.get("timeout_seconds", 30)) <= 60 or not 1 <= int(source.get("max_attempts", 3)) <= 3:
        raise ConfigurationError("超时应为1..60秒，单次请求尝试次数应为1..3")
    if not 0 <= float(source.get("retry_base_seconds", 2)) <= 10:
        raise ConfigurationError("退避基数应为0..10秒")
    from urllib.parse import urlsplit
    page = urlsplit(str(source.get("page_url", "")))
    if page.scheme not in {"http", "https"} or not page.hostname or page.username or page.password:
        raise ConfigurationError("来源页面URL无效")
    output = settings.section("output")
    if output.get("date_format") != "yyyy-mm-dd hh:mm" or output.get("number_format") != "0.####":
        raise ConfigurationError("必须使用约定的日期和最多四位小数格式")
    if output.get("raw_sheet") != "原始数据" or output.get("log_sheet") != "采集日志":
        raise ConfigurationError("审计工作表名称必须为原始数据和采集日志")
    if settings.section("font_decoder").get("fail_on_unresolved") is not True:
        raise ConfigurationError("不得关闭字体解码失败阻断")
    required = {"bsnm", "addvnm", "rvnm", "stnm", "idNo", "rz", "rzRange"}
    if not required.issubset(settings.section("font_decoder").get("required_fields", [])):
        raise ConfigurationError("动态字体配置缺少必要解码字段")
    filter_config = settings.section("filter")
    if filter_config.get("match_mode") != "province":
        raise ConfigurationError("筛选模式必须为province")
    if any(key in filter_config for key in ("allowlist_file", "require_allowlist_match", "city_names")):
        raise ConfigurationError("贵州省级采集不得配置城市或水库白名单")
    schedule = data.get("schedule", {})
    if not isinstance(schedule, dict):
        raise ConfigurationError("schedule配置必须是对象")
    daily_time = str(schedule.get("daily_time", "09:00"))
    if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", daily_time):
        raise ConfigurationError("schedule.daily_time必须使用24小时制HH:mm")
    daily_count = schedule.get("daily_count", 1)
    if type(daily_count) is not int or not 1 <= daily_count <= 24:
        raise ConfigurationError("schedule.daily_count必须为1..24之间的整数")
    task_name = str(schedule.get("task_name", "贵州省大型水库实时水情采集"))
    if not task_name.strip() or any(character in task_name for character in "\\/*?[]"):
        raise ConfigurationError("schedule.task_name不能为空或包含路径分隔符、通配符")
    if not isinstance(schedule.get("popup_enabled", True), bool):
        raise ConfigurationError("schedule.popup_enabled必须为布尔值")
    try:
        popup_timeout = int(schedule.get("popup_timeout_seconds", 15))
    except (TypeError, ValueError) as exc:
        raise ConfigurationError("schedule.popup_timeout_seconds必须为整数") from exc
    if not 1 <= popup_timeout <= 120:
        raise ConfigurationError("schedule.popup_timeout_seconds必须在1..120秒之间")
    from .filtering import normalize_province_names
    normalize_province_names(filter_config.get("province_names", []))
    return settings

