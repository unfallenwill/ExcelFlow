from __future__ import annotations

import os
import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

import pandas as pd

from .expression import SafeExpressionEvaluator
from .schema import ExtractionSpec, SAFE_FIELD

ENV_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")


class ExtractionEngine(ABC):
    @abstractmethod
    def execute(self, spec: ExtractionSpec, task_id: str, source: Path) -> pd.DataFrame: ...


class PandasExtractionEngine(ExtractionEngine):
    def __init__(self, evaluator: SafeExpressionEvaluator | None = None):
        self.evaluator = evaluator or SafeExpressionEvaluator()

    @staticmethod
    def _resolve(value: Any) -> Any:
        if isinstance(value, str) and (match := ENV_PATTERN.fullmatch(value.strip())):
            if match.group(1) not in os.environ: raise ValueError(f"环境变量 {match.group(1)} 未设置")
            return os.environ[match.group(1)]
        return value

    def _load(self, spec: ExtractionSpec, task_id: str, source: Path) -> dict[str, pd.DataFrame]:
        frames = {}
        with pd.ExcelFile(source) as excel:
            for obj in spec.for_task(spec.objects, task_id):
                sheet, alias = str(obj["Sheet名称"]), str(obj["对象别名"])
                if sheet not in excel.sheet_names: raise ValueError(f"Excel 中不存在工作表: {sheet}")
                frame = pd.read_excel(excel, sheet_name=sheet, header=int(obj.get("表头行") or 1) - 1)
                columns = [str(x).strip() for x in frame.columns]
                if not all(SAFE_FIELD.fullmatch(x) for x in columns) or len(columns) != len(set(columns)): raise ValueError(f"工作表 {sheet} 的表头不合法")
                frame.columns = [f"{alias}.{x}" for x in columns]
                frames[alias] = frame
        return frames

    def _join(self, spec: ExtractionSpec, task_id: str, frames: dict[str, pd.DataFrame]) -> pd.DataFrame:
        primary = next(x for x in spec.for_task(spec.objects, task_id) if str(x["是否主表"]) == "是")
        result = frames[str(primary["对象别名"])]
        groups: dict[int, list[dict]] = {}
        for join in spec.for_task(spec.joins, task_id): groups.setdefault(int(join["关联顺序"]), []).append(join)
        for order in sorted(groups):
            rows, right = groups[order], str(groups[order][0]["右侧对象"])
            result = result.merge(frames[right], how="left" if str(rows[0]["关联类型"]).upper() == "LEFT JOIN" else "inner",
                                  left_on=[str(x["左侧字段"]) for x in rows], right_on=[str(x["右侧字段"]) for x in rows], sort=False)
        return result

    def _coerce(self, series: pd.Series, value: Any):
        value = self._resolve(value)
        if pd.api.types.is_datetime64_any_dtype(series): return pd.to_datetime(value)
        if pd.api.types.is_numeric_dtype(series) and isinstance(value, str):
            try: return float(value) if "." in value else int(value)
            except ValueError: pass
        return value

    def _condition(self, frame: pd.DataFrame, item: dict) -> pd.Series:
        series, operator = frame[str(item["字段"])], str(item["运算符"]).upper()
        if operator == "IS NULL": return series.isna()
        if operator == "IS NOT NULL": return series.notna()
        if operator in {"IN", "NOT IN"}:
            mask = series.isin([self._coerce(series, x.strip()) for x in str(item["值1"]).split(",") if x.strip()])
            return ~mask if operator == "NOT IN" else mask
        first = self._coerce(series, item["值1"])
        if operator == "BETWEEN": return series.between(first, self._coerce(series, item["值2"]), inclusive="both")
        if operator in {"LIKE", "NOT LIKE"}:
            pattern = "".join(".*" if char == "%" else "." if char == "_" else re.escape(char) for char in str(first))
            mask = series.astype("string").str.fullmatch(pattern, na=False)
            return ~mask if operator == "NOT LIKE" else mask
        return {"=": series.eq, "!=": series.ne, ">": series.gt, ">=": series.ge, "<": series.lt, "<=": series.le}[operator](first).fillna(False)

    def _filter(self, frame: pd.DataFrame, spec: ExtractionSpec, task_id: str) -> pd.DataFrame:
        groups: dict[int, list[dict]] = {}
        for item in spec.for_task(spec.filters, task_id): groups.setdefault(int(item["条件组"]), []).append(item)
        if groups:
            total = pd.Series(False, index=frame.index)
            for group_id in sorted(groups):
                mask = pd.Series(True, index=frame.index)
                for item in sorted(groups[group_id], key=lambda x: int(x["条件序号"])): mask &= self._condition(frame, item)
                total |= mask
            frame = frame.loc[total]
        return frame

    def _select(self, frame: pd.DataFrame, spec: ExtractionSpec, task_id: str) -> pd.DataFrame:
        fields = sorted(spec.for_task(spec.fields, task_id), key=lambda x: int(x.get("字段顺序") or 999999))
        if not fields: return frame.reset_index(drop=True)
        output = pd.DataFrame(index=frame.index)
        for item in fields:
            expression = str(item.get("转换表达式") or "").strip()
            output[str(item["目标字段"])] = self.evaluator.evaluate(expression, frame) if expression else frame[str(item["源字段"])]
        return output.reset_index(drop=True)

    def execute(self, spec: ExtractionSpec, task_id: str, source: Path) -> pd.DataFrame:
        plan = spec.task(task_id)
        frame = self._join(spec, task_id, self._load(spec, task_id, source))
        if str(plan["抽取模式"]) == "增量":
            series = frame[str(plan["增量字段"])]
            start, end = self._coerce(series, plan["开始值"]), self._coerce(series, plan["结束值"])
            frame = frame.loc[series.ge(start) & series.lt(end)]
        return self._select(self._filter(frame, spec, task_id), spec, task_id)
