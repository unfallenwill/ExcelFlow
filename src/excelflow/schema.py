from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

PLAN_HEADERS = ["任务ID", "启用", "备注"]
OBJECT_HEADERS = ["任务ID", "Sheet名称", "对象别名", "表头行", "是否主表", "备注"]
JOIN_HEADERS = ["任务ID", "关联顺序", "关联类型", "左侧字段", "右侧对象", "右侧字段", "备注"]
FIELD_HEADERS = ["任务ID", "源字段", "目标字段", "目标类型", "转换表达式", "字段顺序", "备注"]
FILTER_HEADERS = ["任务ID", "条件组", "条件序号", "字段", "运算符", "值1", "值2", "备注"]
GROUP_HEADERS = ["任务ID", "源字段", "目标字段", "目标类型", "分组顺序", "备注"]
AGGREGATION_HEADERS = ["任务ID", "源字段", "聚合函数", "目标字段", "目标类型", "分隔符", "聚合顺序", "备注"]
SAFE_FIELD = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*$")
QUALIFIED_FIELD = re.compile(r"^[A-Za-z_][A-Za-z0-9_$]*\.[A-Za-z_][A-Za-z0-9_$]*$")


@dataclass(frozen=True)
class ExtractionSpec:
    plans: list[dict[str, Any]] = field(default_factory=list)
    objects: list[dict[str, Any]] = field(default_factory=list)
    joins: list[dict[str, Any]] = field(default_factory=list)
    fields: list[dict[str, Any]] = field(default_factory=list)
    filters: list[dict[str, Any]] = field(default_factory=list)
    groups: list[dict[str, Any]] = field(default_factory=list)
    aggregations: list[dict[str, Any]] = field(default_factory=list)

    def task(self, task_id: str) -> dict[str, Any]:
        task = next((x for x in self.plans if str(x.get("任务ID")) == task_id), None)
        if task is None:
            raise ValueError(f"任务不存在: {task_id}")
        return task

    def for_task(self, records: list[dict[str, Any]], task_id: str) -> list[dict[str, Any]]:
        return [x for x in records if str(x.get("任务ID") or "").strip() == task_id]


@dataclass(frozen=True)
class ValidationResult:
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors
