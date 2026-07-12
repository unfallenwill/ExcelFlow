from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sqlite3
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation


PLAN_HEADERS = [
    "任务ID", "启用", "数据源ID", "源对象", "抽取模式", "增量字段", "开始值", "结束值",
    "过滤条件", "输出格式", "输出路径", "批次大小", "负责人", "备注",
]
SOURCE_HEADERS = ["数据源ID", "类型", "连接地址", "端口", "数据库", "Schema", "用户名", "密码环境变量", "扩展参数"]
FIELD_HEADERS = ["任务ID", "源字段", "目标字段", "目标类型", "转换表达式", "字段顺序", "备注"]

IDENTIFIER = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*(\.[A-Za-z_][A-Za-z0-9_$]*)*$")
SAFE_FIELD = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*$")


@dataclass
class ValidationResult:
    errors: list[str]
    warnings: list[str]

    @property
    def ok(self) -> bool:
        return not self.errors


def _style_sheet(ws, widths: list[int]) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width


def create_template(path: Path) -> None:
    wb = Workbook()
    plan = wb.active
    plan.title = "抽取计划"
    plan.append(PLAN_HEADERS)
    plan.append(["demo_orders", "否", "local_demo", "orders", "增量", "updated_at", "${START_TIME}", "${END_TIME}", "status = 'paid'", "csv", "./output/orders.csv", 10000, "数据组", "示例任务，启用前请修改"])
    _style_sheet(plan, [18, 9, 16, 22, 12, 18, 20, 20, 28, 12, 30, 12, 14, 28])

    source = wb.create_sheet("数据源")
    source.append(SOURCE_HEADERS)
    source.append(["local_demo", "sqlite", "./data/demo.db", "", "", "", "", "", "{}"])
    source.append(["excel_demo", "excel", "./data/source.xlsx", "", "", "", "", "", '{"header_row": 1}'])
    source.append(["json_demo", "json", "./data/source.json", "", "", "", "", "", '{"encoding": "utf-8"}'])
    _style_sheet(source, [16, 12, 28, 10, 18, 14, 16, 20, 28])

    fields = wb.create_sheet("字段映射")
    fields.append(FIELD_HEADERS)
    fields.append(["demo_orders", "id", "order_id", "integer", "", 1, ""])
    fields.append(["demo_orders", "amount", "amount", "decimal", "", 2, ""])
    fields.append(["demo_orders", "updated_at", "updated_at", "datetime", "", 3, ""])
    _style_sheet(fields, [18, 20, 20, 16, 30, 12, 28])

    guide = wb.create_sheet("填写说明")
    guide.append(["项目", "说明"])
    rows = [
        ("使用流程", "填写数据源和任务 → validate 校验 → preview 预览 SQL → run 执行抽取"),
        ("启用", "是/否；只有“是”的任务允许执行"),
        ("源对象", "SQLite: 表/视图名；Excel: 工作表名；JSON: 根对象中的数组键，根为数组时可填 data"),
        ("抽取模式", "全量或增量；增量任务必须填写增量字段"),
        ("开始值/结束值", "增量区间为 [开始值, 结束值)；支持 ${ENV_NAME} 环境变量占位符"),
        ("过滤条件", "可选的 SQL 条件，不要填写 WHERE；内容会原样拼接，执行前务必 preview"),
        ("字段映射", "不配置时抽取 *；转换表达式是源端 SQL 表达式，目标字段是输出列名"),
        ("密码", "只填写环境变量名，例如 PD_PROD_PASSWORD，禁止把密码明文放进 Excel"),
        ("当前连接器", "sqlite、excel、json"),
        ("输出格式", "csv、jsonl 或 xlsx；相对路径以当前工作目录为基准"),
    ]
    for row in rows:
        guide.append(row)
    _style_sheet(guide, [20, 90])
    guide.column_dimensions["B"].width = 100

    for ws in (plan,):
        enabled = DataValidation(type="list", formula1='"是,否"')
        mode = DataValidation(type="list", formula1='"全量,增量"')
        fmt = DataValidation(type="list", formula1='"csv,jsonl,xlsx"')
        ws.add_data_validation(enabled); enabled.add("B2:B1000")
        ws.add_data_validation(mode); mode.add("E2:E1000")
        ws.add_data_validation(fmt); fmt.add("J2:J1000")
    source_type = DataValidation(type="list", formula1='"sqlite,excel,json"')
    source.add_data_validation(source_type); source_type.add("B2:B1000")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _records(ws) -> list[dict[str, Any]]:
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in ws[1]]
    result = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(value is not None and str(value).strip() for value in row):
            continue
        result.append({headers[i]: value for i, value in enumerate(row)})
    return result


def read_spec(path: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    required = {"抽取计划", "数据源", "字段映射"}
    missing = required - set(wb.sheetnames)
    if missing:
        raise ValueError(f"缺少工作表: {', '.join(sorted(missing))}")
    return _records(wb["抽取计划"]), _records(wb["数据源"]), _records(wb["字段映射"])


def validate(path: Path) -> ValidationResult:
    errors: list[str] = []
    warnings: list[str] = []
    try:
        plans, sources, fields = read_spec(path)
    except Exception as exc:
        return ValidationResult([str(exc)], [])

    source_ids = [str(x.get("数据源ID") or "").strip() for x in sources]
    if len(source_ids) != len(set(source_ids)):
        errors.append("数据源ID存在重复")
    task_ids: list[str] = []
    for row_no, plan in enumerate(plans, 2):
        task_id = str(plan.get("任务ID") or "").strip()
        task_ids.append(task_id)
        prefix = f"抽取计划第{row_no}行"
        if not task_id:
            errors.append(f"{prefix}: 任务ID不能为空")
        if str(plan.get("数据源ID") or "").strip() not in source_ids:
            errors.append(f"{prefix}: 数据源ID不存在")
        obj = str(plan.get("源对象") or "").strip()
        source_id = str(plan.get("数据源ID") or "").strip()
        plan_source = next((x for x in sources if str(x.get("数据源ID") or "").strip() == source_id), None)
        source_type = str((plan_source or {}).get("类型") or "").lower()
        if not obj:
            errors.append(f"{prefix}: 源对象不能为空")
        elif source_type == "sqlite" and not IDENTIFIER.fullmatch(obj):
            errors.append(f"{prefix}: SQLite 源对象只能是表/视图标识符，可带 schema")
        mode = str(plan.get("抽取模式") or "").strip()
        if mode not in {"全量", "增量"}:
            errors.append(f"{prefix}: 抽取模式必须为全量或增量")
        if mode == "增量":
            inc = str(plan.get("增量字段") or "").strip()
            if not SAFE_FIELD.fullmatch(inc):
                errors.append(f"{prefix}: 增量字段不能为空且必须是简单字段名")
            if plan.get("开始值") in (None, "") or plan.get("结束值") in (None, ""):
                errors.append(f"{prefix}: 增量抽取必须同时填写开始值和结束值")
        if str(plan.get("输出格式") or "").lower() not in {"csv", "jsonl", "xlsx"}:
            errors.append(f"{prefix}: 输出格式必须为 csv、jsonl 或 xlsx")
        if not str(plan.get("输出路径") or "").strip():
            errors.append(f"{prefix}: 输出路径不能为空")
        if str(plan.get("启用") or "").strip() not in {"是", "否"}:
            errors.append(f"{prefix}: 启用必须为是或否")
    if len(task_ids) != len(set(task_ids)):
        errors.append("任务ID存在重复")
    for row_no, field in enumerate(fields, 2):
        if str(field.get("任务ID") or "").strip() not in task_ids:
            errors.append(f"字段映射第{row_no}行: 任务ID不存在")
        src = str(field.get("源字段") or "").strip()
        expr = str(field.get("转换表达式") or "").strip()
        target = str(field.get("目标字段") or "").strip()
        if not expr and not SAFE_FIELD.fullmatch(src):
            errors.append(f"字段映射第{row_no}行: 源字段必须是简单字段名，或填写转换表达式")
        if not SAFE_FIELD.fullmatch(target):
            errors.append(f"字段映射第{row_no}行: 目标字段格式不合法")
    for source in sources:
        source_type = str(source.get("类型") or "").lower()
        if source_type not in {"sqlite", "excel", "json"}:
            errors.append(f"数据源 {source.get('数据源ID')}: 类型必须为 sqlite、excel 或 json")
        if not str(source.get("连接地址") or "").strip():
            errors.append(f"数据源 {source.get('数据源ID')}: 连接地址不能为空")
        secret = str(source.get("密码环境变量") or "").strip()
        if secret and secret not in os.environ:
            warnings.append(f"环境变量 {secret} 当前未设置")
    return ValidationResult(errors, warnings)


ENV_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")


def _resolve(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    match = ENV_PATTERN.fullmatch(value.strip())
    if match:
        name = match.group(1)
        if name not in os.environ:
            raise ValueError(f"环境变量 {name} 未设置")
        return os.environ[name]
    return value


def build_query(plan: dict[str, Any], fields: list[dict[str, Any]], table_name: str | None = None) -> tuple[str, list[Any]]:
    selected = [x for x in fields if str(x.get("任务ID") or "").strip() == str(plan["任务ID"]).strip()]
    selected.sort(key=lambda x: int(x.get("字段顺序") or 999999))
    if selected:
        columns = []
        for item in selected:
            expression = str(item.get("转换表达式") or item.get("源字段")).strip()
            columns.append(f'{expression} AS "{item["目标字段"]}"')
        select_clause = ", ".join(columns)
    else:
        select_clause = "*"
    sql = f'SELECT {select_clause} FROM {table_name or plan["源对象"]}'
    clauses: list[str] = []
    params: list[Any] = []
    if str(plan.get("抽取模式") or "").strip() == "增量":
        field = plan["增量字段"]
        clauses.append(f"{field} >= ? AND {field} < ?")
        params.extend([_resolve(plan["开始值"]), _resolve(plan["结束值"])])
    custom_filter = str(plan.get("过滤条件") or "").strip()
    if custom_filter:
        clauses.append(f"({custom_filter})")
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    return sql, params


def _find_task(path: Path, task_id: str):
    plans, sources, fields = read_spec(path)
    plan = next((x for x in plans if str(x.get("任务ID")) == task_id), None)
    if not plan:
        raise ValueError(f"任务不存在: {task_id}")
    source = next((x for x in sources if x.get("数据源ID") == plan.get("数据源ID")), None)
    return plan, source, fields


def _write_rows(output: Path, fmt: str, columns: list[str], rows: Iterable[tuple[Any, ...]]) -> int:
    output.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    if fmt == "csv":
        with output.open("w", newline="", encoding="utf-8-sig") as file:
            writer = csv.writer(file); writer.writerow(columns)
            for row in rows: writer.writerow(row); count += 1
    elif fmt == "jsonl":
        with output.open("w", encoding="utf-8") as file:
            for row in rows:
                file.write(json.dumps(dict(zip(columns, row)), ensure_ascii=False, default=str) + "\n"); count += 1
    else:
        wb = Workbook(write_only=True); ws = wb.create_sheet("data"); ws.append(columns)
        for row in rows: ws.append(list(row)); count += 1
        wb.save(output)
    return count


def _options(source: dict[str, Any]) -> dict[str, Any]:
    raw = str(source.get("扩展参数") or "{}").strip()
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"数据源 {source.get('数据源ID')} 的扩展参数不是合法 JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ValueError("扩展参数必须是 JSON 对象")
    return value


def _normalize_value(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bytes)):
        return value
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat(sep=" ")
    return json.dumps(value, ensure_ascii=False, default=str)


def _load_virtual_source(conn: sqlite3.Connection, source: dict[str, Any], object_name: str) -> None:
    """Load an Excel sheet or JSON array into an in-memory SQLite table."""
    source_type = str(source["类型"]).lower()
    path = Path(str(source["连接地址"]))
    options = _options(source)
    records: list[dict[str, Any]] = []
    if source_type == "excel":
        workbook = load_workbook(path, data_only=True, read_only=True)
        if object_name not in workbook.sheetnames:
            raise ValueError(f"Excel 中不存在工作表: {object_name}")
        ws = workbook[object_name]
        header_row = int(options.get("header_row", 1))
        rows = ws.iter_rows(min_row=header_row, values_only=True)
        try:
            headers = [str(x).strip() if x is not None else "" for x in next(rows)]
        except StopIteration:
            raise ValueError(f"Excel 工作表为空: {object_name}")
        if not all(SAFE_FIELD.fullmatch(x) for x in headers) or len(headers) != len(set(headers)):
            raise ValueError("Excel 表头必须是非空、唯一的简单字段名")
        records = [dict(zip(headers, row)) for row in rows if any(x is not None for x in row)]
    elif source_type == "json":
        encoding = str(options.get("encoding", "utf-8"))
        with path.open(encoding=encoding) as file:
            payload = json.load(file)
        data = payload if isinstance(payload, list) else payload.get(object_name) if isinstance(payload, dict) else None
        if not isinstance(data, list):
            raise ValueError(f"JSON 数据必须是数组，或根对象的 {object_name} 必须是数组")
        if not all(isinstance(item, dict) for item in data):
            raise ValueError("JSON 数组的每个元素必须是对象")
        records = data
    if not records:
        raise ValueError(f"源对象没有数据: {object_name}")
    columns: list[str] = []
    for record in records:
        for key in record:
            if not SAFE_FIELD.fullmatch(str(key)):
                raise ValueError(f"源字段名不合法: {key}")
            if key not in columns:
                columns.append(key)
    quoted = ", ".join(f'"{column}"' for column in columns)
    conn.execute(f'CREATE TABLE "__source_data" ({quoted})')
    placeholders = ", ".join("?" for _ in columns)
    conn.executemany(
        f'INSERT INTO "__source_data" ({quoted}) VALUES ({placeholders})',
        ([ _normalize_value(record.get(column)) for column in columns] for record in records),
    )


def run_task(path: Path, task_id: str) -> tuple[int, Path]:
    result = validate(path)
    if not result.ok:
        raise ValueError("Excel 校验失败:\n" + "\n".join(result.errors))
    plan, source, fields = _find_task(path, task_id)
    if str(plan.get("启用")) != "是":
        raise ValueError(f"任务 {task_id} 未启用")
    source_type = str(source.get("类型") or "").lower()
    sql, params = build_query(plan, fields, '"__source_data"' if source_type in {"excel", "json"} else None)
    database = Path(str(source.get("连接地址") or "")) if source_type == "sqlite" else Path(":memory:")
    batch_size = int(plan.get("批次大小") or 10000)
    output = Path(str(plan["输出路径"]))
    with sqlite3.connect(database) as conn:
        if source_type in {"excel", "json"}:
            _load_virtual_source(conn, source, str(plan["源对象"]))
        cursor = conn.execute(sql, params)
        columns = [item[0] for item in cursor.description]
        def rows():
            while batch := cursor.fetchmany(batch_size):
                yield from batch
        count = _write_rows(output, str(plan["输出格式"]).lower(), columns, rows())
    return count, output


def main() -> int:
    parser = argparse.ArgumentParser(description="Excel 驱动的数据抽取工具")
    sub = parser.add_subparsers(dest="command", required=True)
    p_template = sub.add_parser("template", help="生成 Excel 模板")
    p_template.add_argument("path", nargs="?", default="extraction_plan.xlsx")
    p_validate = sub.add_parser("validate", help="校验 Excel 声明")
    p_validate.add_argument("path")
    p_preview = sub.add_parser("preview", help="预览任务 SQL")
    p_preview.add_argument("path"); p_preview.add_argument("task_id")
    p_run = sub.add_parser("run", help="执行一个已启用任务")
    p_run.add_argument("path"); p_run.add_argument("task_id")
    args = parser.parse_args()
    try:
        if args.command == "template":
            create_template(Path(args.path)); print(f"已生成: {args.path}")
        elif args.command == "validate":
            result = validate(Path(args.path))
            for item in result.errors: print(f"错误: {item}")
            for item in result.warnings: print(f"警告: {item}")
            if result.ok: print("校验通过")
            return 0 if result.ok else 1
        elif args.command == "preview":
            plan, _, fields = _find_task(Path(args.path), args.task_id)
            sql, params = build_query(plan, fields)
            print(sql); print("参数:", params)
        elif args.command == "run":
            count, output = run_task(Path(args.path), args.task_id)
            print(f"抽取完成: {count} 行 -> {output}")
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
