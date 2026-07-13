import argparse
import sys
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path

from .service import ExtractionService
from .template import create_template


def package_version() -> str:
    try:
        return version("excelflow")
    except PackageNotFoundError:
        return "0.0.0+unknown"


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="excelflow", description="Excel 驱动的 Pandas 数据抽取工具", allow_abbrev=False
    )
    parser.add_argument(
        "-V", "--version", action="version", version=f"excelflow {package_version()}"
    )
    sub = parser.add_subparsers(dest="command", required=True)
    template = sub.add_parser("template", help="生成抽取计划模板", allow_abbrev=False)
    template.add_argument("-o", "--output", default="extraction_plan.xlsx", help="模板输出路径")
    validate = sub.add_parser("validate", help="校验抽取计划", allow_abbrev=False)
    validate.add_argument("-p", "--plan", required=True, help="抽取计划 Excel")
    preview = sub.add_parser("preview", help="预览 Pandas 执行计划", allow_abbrev=False)
    preview.add_argument("-p", "--plan", required=True, help="抽取计划 Excel")
    preview.add_argument("-t", "--task", required=True, help="任务ID")
    run = sub.add_parser("run", help="执行数据抽取", allow_abbrev=False)
    run.add_argument("-p", "--plan", required=True, help="抽取计划 Excel")
    run.add_argument("-t", "--task", required=True, help="任务ID")
    run.add_argument("-s", "--source", required=True, help="源数据 Excel")
    run.add_argument(
        "-f", "--format", required=True, choices=["csv", "jsonl", "xlsx"], help="输出格式"
    )
    run.add_argument("-o", "--output", required=True, help="输出路径")
    args, service = parser.parse_args(), ExtractionService()
    try:
        if args.command == "template":
            create_template(Path(args.output))
            print(f"已生成: {args.output}")
        elif args.command == "validate":
            result = service.validate(Path(args.plan))
            for error in result.errors:
                print(f"错误: {error}", file=sys.stderr)
            if result.ok:
                print("校验通过")
            return 0 if result.ok else 1
        elif args.command == "preview":
            print(service.preview(Path(args.plan), args.task))
        else:
            count, output = service.run(
                Path(args.plan), args.task, Path(args.source), args.format, Path(args.output)
            )
            print(f"抽取完成: {count} 行 -> {output}")
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1
    return 0
