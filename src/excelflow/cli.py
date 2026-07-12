import argparse
import sys
from pathlib import Path

from .service import ExtractionService
from .template import create_template


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 驱动的 Pandas 数据抽取工具")
    sub = parser.add_subparsers(dest="command", required=True)
    template = sub.add_parser("template"); template.add_argument("path", nargs="?", default="extraction_plan.xlsx")
    validate = sub.add_parser("validate"); validate.add_argument("path")
    preview = sub.add_parser("preview"); preview.add_argument("path"); preview.add_argument("task_id")
    run = sub.add_parser("run"); run.add_argument("path"); run.add_argument("task_id"); run.add_argument("source_excel"); run.add_argument("output_format", choices=["csv", "jsonl", "xlsx"]); run.add_argument("output_path")
    args, service = parser.parse_args(), ExtractionService()
    try:
        if args.command == "template": create_template(Path(args.path)); print(f"已生成: {args.path}")
        elif args.command == "validate":
            result = service.validate(Path(args.path))
            for error in result.errors: print(f"错误: {error}")
            if result.ok: print("校验通过")
            return 0 if result.ok else 1
        elif args.command == "preview": print(service.preview(Path(args.path), args.task_id))
        else:
            count, output = service.run(Path(args.path), args.task_id, Path(args.source_excel), args.output_format, Path(args.output_path)); print(f"抽取完成: {count} 行 -> {output}")
    except Exception as exc: print(f"失败: {exc}", file=sys.stderr); return 1
    return 0
