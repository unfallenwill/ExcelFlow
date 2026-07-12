from .schema import ExtractionSpec, QUALIFIED_FIELD, SAFE_FIELD, ValidationResult


class SpecValidator:
    operators = {"=", "!=", ">", ">=", "<", "<=", "IN", "NOT IN", "BETWEEN", "LIKE", "NOT LIKE", "IS NULL", "IS NOT NULL"}

    def validate(self, spec: ExtractionSpec) -> ValidationResult:
        errors: list[str] = []
        task_ids = [str(x.get("任务ID") or "").strip() for x in spec.plans]
        if any(not x for x in task_ids): errors.append("任务ID不能为空")
        if len(task_ids) != len(set(task_ids)): errors.append("任务ID存在重复")
        aliases: dict[str, set[str]] = {task: set() for task in task_ids}

        for plan in spec.plans:
            task = str(plan.get("任务ID") or "").strip()
            if str(plan.get("启用") or "") not in {"是", "否"}: errors.append(f"任务 {task}: 启用必须为是或否")

        for row, obj in enumerate(spec.objects, 2):
            task, alias = str(obj.get("任务ID") or "").strip(), str(obj.get("对象别名") or "").strip()
            if task not in aliases: errors.append(f"数据对象第{row}行: 任务ID不存在"); continue
            if not SAFE_FIELD.fullmatch(alias): errors.append(f"数据对象第{row}行: 对象别名不合法")
            if alias in aliases[task]: errors.append(f"数据对象第{row}行: 对象别名重复")
            aliases[task].add(alias)
            if not str(obj.get("Sheet名称") or "").strip(): errors.append(f"数据对象第{row}行: Sheet名称不能为空")
            try:
                if int(obj.get("表头行") or 1) < 1: raise ValueError
            except (TypeError, ValueError): errors.append(f"数据对象第{row}行: 表头行必须是正整数")
            if str(obj.get("是否主表") or "") not in {"是", "否"}: errors.append(f"数据对象第{row}行: 是否主表必须为是或否")

        for task in task_ids:
            objects = spec.for_task(spec.objects, task)
            primaries = [x for x in objects if str(x.get("是否主表")) == "是"]
            if len(primaries) != 1: errors.append(f"任务 {task}: 必须且只能配置一个主表")
            joined = {str(primaries[0]["对象别名"])} if len(primaries) == 1 else set()
            groups: dict[int, list[dict]] = {}
            for join in spec.for_task(spec.joins, task):
                try: groups.setdefault(int(join["关联顺序"]), []).append(join)
                except (TypeError, ValueError): errors.append(f"任务 {task}: 关联顺序必须是正整数")
            for order in sorted(groups):
                rows = groups[order]
                types = {str(x.get("关联类型") or "").upper() for x in rows}
                rights = {str(x.get("右侧对象") or "") for x in rows}
                if len(types) != 1 or not types <= {"INNER JOIN", "LEFT JOIN"} or len(rights) != 1:
                    errors.append(f"任务 {task} 关联顺序 {order}: 关联类型或右侧对象不一致"); continue
                right = next(iter(rights))
                for join in rows:
                    left_field, right_field = str(join.get("左侧字段") or ""), str(join.get("右侧字段") or "")
                    if not QUALIFIED_FIELD.fullmatch(left_field) or left_field.split(".")[0] not in joined: errors.append(f"任务 {task} 关联顺序 {order}: 左侧字段无效")
                    if not QUALIFIED_FIELD.fullmatch(right_field) or right_field.split(".")[0] != right: errors.append(f"任务 {task} 关联顺序 {order}: 右侧字段无效")
                if right in joined or right not in aliases[task]: errors.append(f"任务 {task} 关联顺序 {order}: 右侧对象无效或重复")
                joined.add(right)
            if joined != aliases[task]: errors.append(f"任务 {task}: 存在未关联的数据对象: {', '.join(sorted(aliases[task] - joined))}")

        for collection, label in ((spec.fields, "字段映射"), (spec.filters, "过滤条件")):
            for row, item in enumerate(collection, 2):
                task = str(item.get("任务ID") or "").strip()
                if task not in aliases: errors.append(f"{label}第{row}行: 任务ID不存在"); continue
                field = str(item.get("源字段" if label == "字段映射" else "字段") or "").strip()
                expression = str(item.get("转换表达式") or "").strip()
                if not expression and (not QUALIFIED_FIELD.fullmatch(field) or field.split(".")[0] not in aliases[task]): errors.append(f"{label}第{row}行: 字段引用无效")
                if label == "字段映射" and not SAFE_FIELD.fullmatch(str(item.get("目标字段") or "")): errors.append(f"字段映射第{row}行: 目标字段不合法")
                if label == "过滤条件":
                    operator = str(item.get("运算符") or "").upper()
                    if operator not in self.operators: errors.append(f"过滤条件第{row}行: 不支持的运算符")
                    if operator not in {"IS NULL", "IS NOT NULL"} and item.get("值1") in (None, ""): errors.append(f"过滤条件第{row}行: 值1不能为空")
                    if operator == "BETWEEN" and item.get("值2") in (None, ""): errors.append(f"过滤条件第{row}行: 值2不能为空")

        return ValidationResult(errors=errors)
