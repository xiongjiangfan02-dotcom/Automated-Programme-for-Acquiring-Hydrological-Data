from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .config import load_settings
from .errors import AccessRestrictedError, CollectorError
from .excel_writer import validate_template
from .pipeline import CollectionPipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="贵州毕节市大型水库实时水情采集模块")
    parser.add_argument("--config", default="config/default.json", help="配置文件路径")
    subparsers = parser.add_subparsers(dest="command", required=True)
    validate_parser = subparsers.add_parser("validate-template", help="只读验证Excel模板")
    validate_parser.add_argument("--template", help="覆盖配置中的模板路径")
    collect_parser = subparsers.add_parser("collect", help="执行一次采集")
    collect_parser.add_argument("--input-json", type=Path, help="使用已保存的接口JSON离线重放")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        settings = load_settings(args.config)
        if args.command == "validate-template":
            template = Path(args.template).resolve() if args.template else settings.resolve_path(settings.section("paths")["template"])
            print(json.dumps(validate_template(template), ensure_ascii=False, indent=2))
            return 0
        log = CollectionPipeline(settings).run(args.input_json)
        print(json.dumps(log.to_dict(), ensure_ascii=False, indent=2))
        return 4 if log.status == "无匹配数据" else 0
    except AccessRestrictedError as exc:
        print(f"访问限制，停止自动重试: {exc}", file=sys.stderr)
        return 5
    except CollectorError as exc:
        print(f"采集失败: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"文件或网络错误: {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())

