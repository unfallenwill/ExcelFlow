import re
from typing import Any

from .schema import QUALIFIED_FIELD, SAFE_FIELD, ExtractionSpec, ValidationResult

# 任务ID 在多任务模式下被拼成输出文件名，禁止路径分隔符、Windows 非法文件名字符与控制字符。
_FORBIDDEN_TASK_ID_CHAR = re.compile(r"[\\/:*?\"<>|\x00-\x1f]")


class SpecValidator:
    operators = {
        "=",
        "!=",
        ">",
        ">=",
        "<",
        "<=",
        "IN",
        "NOT IN",
        "BETWEEN",
        "LIKE",
        "NOT LIKE",
        "IS NULL",
        "IS NOT NULL",
    }
    aggregate_functions = {
        "count",
        "count_all",
        "count_distinct",
        "sum",
        "avg",
        "min",
        "max",
        "first",
        "last",
        "concat_agg",
    }
    target_types = {"integer", "decimal", "string", "datetime"}

    def validate(self, spec: ExtractionSpec) -> ValidationResult:
        errors: list[str] = []
        task_ids = [str(x.get("任务ID") or "").strip() for x in spec.plans]
        if any(not x for x in task_ids):
            errors.append("任务ID不能为空")
        if len(task_ids) != len(set(task_ids)):
            errors.append("任务ID存在重复")
        for task_id in task_ids:
            # 禁止路径分隔符、Windows 非法文件名字符、控制字符与点段，避免目录穿越和运行时 OS 错误。
            if task_id in {".", ".."} or _FORBIDDEN_TASK_ID_CHAR.search(task_id):
                errors.append(f"任务ID不能包含路径分隔符或文件名非法字符: {task_id}")
        aliases: dict[str, set[str]] = {task: set() for task in task_ids}

        for plan in spec.plans:
            task = str(plan.get("任务ID") or "").strip()
            if str(plan.get("启用") or "") not in {"是", "否"}:
                errors.append(f"任务 {task}: 启用必须为是或否")

        for row, obj in enumerate(spec.objects, 2):
            task, alias = (
                str(obj.get("任务ID") or "").strip(),
                str(obj.get("对象别名") or "").strip(),
            )
            if task not in aliases:
                errors.append(f"数据对象第{row}行: 任务ID不存在")
                continue
            if not SAFE_FIELD.fullmatch(alias):
                errors.append(f"数据对象第{row}行: 对象别名不合法")
            if alias in aliases[task]:
                errors.append(f"数据对象第{row}行: 对象别名重复")
            aliases[task].add(alias)
            if not str(obj.get("Sheet名称") or "").strip():
                errors.append(f"数据对象第{row}行: Sheet名称不能为空")
            try:
                if int(obj.get("表头行") or 1) < 1:
                    raise ValueError
            except (TypeError, ValueError):
                errors.append(f"数据对象第{row}行: 表头行必须是正整数")
            if str(obj.get("是否主表") or "") not in {"是", "否"}:
                errors.append(f"数据对象第{row}行: 是否主表必须为是或否")

        for task in task_ids:
            objects = spec.for_task(spec.objects, task)
            primaries = [x for x in objects if str(x.get("是否主表")) == "是"]
            if len(primaries) != 1:
                errors.append(f"任务 {task}: 必须且只能配置一个主表")
            joined = {str(primaries[0]["对象别名"])} if len(primaries) == 1 else set()
            join_groups: dict[int, list[dict[str, Any]]] = {}
            for join in spec.for_task(spec.joins, task):
                try:
                    join_groups.setdefault(int(join["关联顺序"]), []).append(join)
                except (TypeError, ValueError):
                    errors.append(f"任务 {task}: 关联顺序必须是正整数")
            for order in sorted(join_groups):
                rows = join_groups[order]
                types = {str(x.get("关联类型") or "").upper() for x in rows}
                rights = {str(x.get("右侧对象") or "") for x in rows}
                if len(types) != 1 or not types <= {"INNER JOIN", "LEFT JOIN"} or len(rights) != 1:
                    errors.append(f"任务 {task} 关联顺序 {order}: 关联类型或右侧对象不一致")
                    continue
                right = next(iter(rights))
                for join in rows:
                    left_field, right_field = (
                        str(join.get("左侧字段") or ""),
                        str(join.get("右侧字段") or ""),
                    )
                    if (
                        not QUALIFIED_FIELD.fullmatch(left_field)
                        or left_field.split(".")[0] not in joined
                    ):
                        errors.append(f"任务 {task} 关联顺序 {order}: 左侧字段无效")
                    if (
                        not QUALIFIED_FIELD.fullmatch(right_field)
                        or right_field.split(".")[0] != right
                    ):
                        errors.append(f"任务 {task} 关联顺序 {order}: 右侧字段无效")
                if right in joined or right not in aliases[task]:
                    errors.append(f"任务 {task} 关联顺序 {order}: 右侧对象无效或重复")
                joined.add(right)
            if joined != aliases[task]:
                unjoined = ", ".join(sorted(aliases[task] - joined))
                errors.append(f"任务 {task}: 存在未关联的数据对象: {unjoined}")

        for collection, label in ((spec.fields, "字段映射"), (spec.filters, "过滤条件")):
            for row, item in enumerate(collection, 2):
                task = str(item.get("任务ID") or "").strip()
                if task not in aliases:
                    errors.append(f"{label}第{row}行: 任务ID不存在")
                    continue
                field = str(item.get("源字段" if label == "字段映射" else "字段") or "").strip()
                expression = str(item.get("转换表达式") or "").strip()
                if not expression and (
                    not QUALIFIED_FIELD.fullmatch(field) or field.split(".")[0] not in aliases[task]
                ):
                    errors.append(f"{label}第{row}行: 字段引用无效")
                if label == "字段映射":
                    if not SAFE_FIELD.fullmatch(str(item.get("目标字段") or "")):
                        errors.append(f"字段映射第{row}行: 目标字段不合法")
                    if str(item.get("目标类型") or "").strip().lower() not in self.target_types:
                        errors.append(f"字段映射第{row}行: 目标类型必须从下拉框选择")
                if label == "过滤条件":
                    operator = str(item.get("运算符") or "").upper()
                    if operator not in self.operators:
                        errors.append(f"过滤条件第{row}行: 不支持的运算符")
                    if operator not in {"IS NULL", "IS NOT NULL"} and item.get("值1") in (None, ""):
                        errors.append(f"过滤条件第{row}行: 值1不能为空")
                    if operator == "BETWEEN" and item.get("值2") in (None, ""):
                        errors.append(f"过滤条件第{row}行: 值2不能为空")

        for collection, label in ((spec.groups, "分组字段"), (spec.aggregations, "聚合规则")):
            for row, item in enumerate(collection, 2):
                if str(item.get("任务ID") or "").strip() not in aliases:
                    errors.append(f"{label}第{row}行: 任务ID不存在")

        for task in task_ids:
            groups, aggregations = (
                spec.for_task(spec.groups, task),
                spec.for_task(spec.aggregations, task),
            )
            fields = spec.for_task(spec.fields, task)
            if groups and not aggregations:
                errors.append(f"任务 {task}: 配置分组字段时必须至少配置一条聚合规则")
            if aggregations and fields:
                errors.append(f"任务 {task}: 聚合任务不能同时配置字段映射")
            targets: list[str] = []
            group_sources: list[str] = []
            for label, rows, order_field in (
                ("分组字段", groups, "分组顺序"),
                ("聚合规则", aggregations, "聚合顺序"),
            ):
                orders: list[int] = []
                for row, item in enumerate(rows, 2):
                    source, target = (
                        str(item.get("源字段") or "").strip(),
                        str(item.get("目标字段") or "").strip(),
                    )
                    function = str(item.get("聚合函数") or "").strip().lower()
                    if label == "聚合规则" and function not in self.aggregate_functions:
                        errors.append(f"聚合规则第{row}行: 不支持的聚合函数")
                    if label == "聚合规则" and function == "count_all" and source:
                        errors.append(f"聚合规则第{row}行: count_all 不应填写源字段")
                    if (label == "分组字段" or function != "count_all") and (
                        not QUALIFIED_FIELD.fullmatch(source)
                        or source.split(".")[0] not in aliases[task]
                    ):
                        errors.append(f"{label}第{row}行: 源字段引用无效")
                    if label == "分组字段":
                        group_sources.append(source)
                    if not SAFE_FIELD.fullmatch(target):
                        errors.append(f"{label}第{row}行: 目标字段不合法")
                    targets.append(target)
                    if str(item.get("目标类型") or "").strip().lower() not in self.target_types:
                        errors.append(f"{label}第{row}行: 目标类型必须从下拉框选择")
                    try:
                        raw_order = item.get(order_field)
                        if raw_order is None:
                            raise ValueError
                        order = int(raw_order)
                        if order < 1:
                            raise ValueError
                        orders.append(order)
                    except (TypeError, ValueError):
                        errors.append(f"{label}第{row}行: {order_field}必须是正整数")
                if len(orders) != len(set(orders)):
                    errors.append(f"任务 {task}: {order_field}不能重复")
            if len(targets) != len(set(targets)):
                errors.append(f"任务 {task}: 分组或聚合目标字段存在重复")
            if len(group_sources) != len(set(group_sources)):
                errors.append(f"任务 {task}: 分组源字段不能重复")

        return ValidationResult(errors=errors)
