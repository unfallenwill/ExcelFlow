from __future__ import annotations

import re
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, cast

import pandas as pd

from .expression import SafeExpressionEvaluator
from .schema import SAFE_FIELD, ExtractionSpec


class ExtractionEngine(ABC):
    @abstractmethod
    def execute(self, spec: ExtractionSpec, task_id: str, source: Path) -> pd.DataFrame: ...


class PandasExtractionEngine(ExtractionEngine):
    def __init__(self, evaluator: SafeExpressionEvaluator | None = None):
        self.evaluator = evaluator or SafeExpressionEvaluator()

    def _load(self, spec: ExtractionSpec, task_id: str, source: Path) -> dict[str, pd.DataFrame]:
        frames = {}
        with pd.ExcelFile(source) as excel:
            for obj in spec.for_task(spec.objects, task_id):
                sheet, alias = str(obj["Sheet名称"]), str(obj["对象别名"])
                if sheet not in excel.sheet_names:
                    raise ValueError(f"Excel 中不存在工作表: {sheet}")
                frame = pd.read_excel(
                    excel, sheet_name=sheet, header=int(obj.get("表头行") or 1) - 1
                )
                columns = [str(x).strip() for x in frame.columns]
                if not all(SAFE_FIELD.fullmatch(x) for x in columns) or len(columns) != len(
                    set(columns)
                ):
                    raise ValueError(f"工作表 {sheet} 的表头不合法")
                frame.columns = [f"{alias}.{x}" for x in columns]
                frames[alias] = frame
        return frames

    def _join(
        self, spec: ExtractionSpec, task_id: str, frames: dict[str, pd.DataFrame]
    ) -> pd.DataFrame:
        primary = next(
            x for x in spec.for_task(spec.objects, task_id) if str(x["是否主表"]) == "是"
        )
        result = frames[str(primary["对象别名"])]
        groups: dict[int, list[dict[str, Any]]] = {}
        for join in spec.for_task(spec.joins, task_id):
            groups.setdefault(int(join["关联顺序"]), []).append(join)
        for order in sorted(groups):
            rows, right = groups[order], str(groups[order][0]["右侧对象"])
            result = result.merge(
                frames[right],
                how="left" if str(rows[0]["关联类型"]).upper() == "LEFT JOIN" else "inner",
                left_on=[str(x["左侧字段"]) for x in rows],
                right_on=[str(x["右侧字段"]) for x in rows],
                sort=False,
            )
        return result

    def _coerce(self, series: pd.Series, value: Any) -> Any:
        if pd.api.types.is_datetime64_any_dtype(series):
            return pd.to_datetime(value)
        if pd.api.types.is_numeric_dtype(series) and isinstance(value, str):
            try:
                return float(value) if "." in value else int(value)
            except ValueError:
                pass
        return value

    def _condition(self, frame: pd.DataFrame, item: dict) -> pd.Series:
        series, operator = cast(pd.Series, frame[str(item["字段"])]), str(item["运算符"]).upper()
        if operator == "IS NULL":
            return series.isna()
        if operator == "IS NOT NULL":
            return series.notna()
        if operator in {"IN", "NOT IN"}:
            mask = series.isin(
                [self._coerce(series, x.strip()) for x in str(item["值1"]).split(",") if x.strip()]
            )
            return ~mask if operator == "NOT IN" else mask
        first = self._coerce(series, item["值1"])
        if operator == "BETWEEN":
            return series.between(first, self._coerce(series, item["值2"]), inclusive="both")
        if operator in {"LIKE", "NOT LIKE"}:
            pattern = "".join(
                ".*" if char == "%" else "." if char == "_" else re.escape(char)
                for char in str(first)
            )
            mask = series.astype("string").str.fullmatch(pattern, na=False)
            return ~mask if operator == "NOT LIKE" else mask
        return {
            "=": series.eq,
            "!=": series.ne,
            ">": series.gt,
            ">=": series.ge,
            "<": series.lt,
            "<=": series.le,
        }[operator](first).fillna(False)

    def _filter(self, frame: pd.DataFrame, spec: ExtractionSpec, task_id: str) -> pd.DataFrame:
        groups: dict[int, list[dict[str, Any]]] = {}
        for item in spec.for_task(spec.filters, task_id):
            groups.setdefault(int(item["条件组"]), []).append(item)
        if groups:
            total = pd.Series(False, index=frame.index)
            for group_id in sorted(groups):
                mask = pd.Series(True, index=frame.index)
                for item in sorted(groups[group_id], key=lambda x: int(x["条件序号"])):
                    mask &= self._condition(frame, item)
                total |= mask
            frame = frame.loc[total]
        return frame

    def _select(self, frame: pd.DataFrame, spec: ExtractionSpec, task_id: str) -> pd.DataFrame:
        fields = sorted(
            spec.for_task(spec.fields, task_id), key=lambda x: int(x.get("字段顺序") or 999999)
        )
        if not fields:
            return frame.reset_index(drop=True)
        output = pd.DataFrame(index=frame.index)
        for item in fields:
            expression = str(item.get("转换表达式") or "").strip()
            value = (
                self.evaluator.evaluate(expression, frame)
                if expression
                else frame[str(item["源字段"])]
            )
            if not isinstance(value, pd.Series):
                value = pd.Series(value, index=frame.index)
            output[str(item["目标字段"])] = self._convert(value, str(item.get("目标类型") or ""))
        return output.reset_index(drop=True)

    @staticmethod
    def _convert(value: Any, target_type: str) -> Any:
        target_type = target_type.strip().lower()
        if target_type in {"integer", "int"}:
            return cast(pd.Series, pd.to_numeric(value, errors="raise")).astype("Int64")
        if target_type in {"decimal", "float", "number"}:
            return cast(pd.Series, pd.to_numeric(value, errors="raise")).astype("Float64")
        if target_type == "string":
            return value.astype("string")
        if target_type in {"datetime", "date"}:
            return pd.to_datetime(value, errors="raise")
        return value

    def _aggregate(self, frame: pd.DataFrame, spec: ExtractionSpec, task_id: str) -> pd.DataFrame:
        groups = sorted(spec.for_task(spec.groups, task_id), key=lambda x: int(x["分组顺序"]))
        rules = sorted(spec.for_task(spec.aggregations, task_id), key=lambda x: int(x["聚合顺序"]))
        keys = [str(item["源字段"]) for item in groups]
        grouped = frame.groupby(keys, dropna=False, sort=False) if keys else None
        columns: dict[str, Any] = {}
        if keys:
            assert grouped is not None
            base = cast(pd.Series, grouped.size()).reset_index(name="__size__")
            for item in groups:
                columns[str(item["目标字段"])] = self._convert(
                    cast(pd.Series, base[str(item["源字段"])]), str(item["目标类型"])
                )
        else:
            base = pd.DataFrame({"__size__": [len(frame)]})

        for item in rules:
            source, function = str(item.get("源字段") or ""), str(item["聚合函数"]).lower()
            if keys:
                assert grouped is not None
                if function == "count_all":
                    value = cast(pd.Series, grouped.size()).reset_index(drop=True)
                else:
                    series_group = grouped[source]
                    if function == "count":
                        value = cast(pd.Series, series_group.count()).reset_index(drop=True)
                    elif function == "count_distinct":
                        value = cast(pd.Series, series_group.nunique(dropna=True)).reset_index(
                            drop=True
                        )
                    elif function == "sum":
                        value = cast(
                            pd.Series,
                            series_group.agg(
                                lambda x: cast(pd.Series, pd.to_numeric(x, errors="raise")).sum(
                                    min_count=1
                                )
                            ),
                        ).reset_index(drop=True)
                    elif function == "avg":
                        value = cast(pd.Series, series_group.mean()).reset_index(drop=True)
                    elif function in {"min", "max", "first", "last"}:
                        value = cast(pd.Series, getattr(series_group, function)()).reset_index(
                            drop=True
                        )
                    else:
                        separator = str(
                            item.get("分隔符") if item.get("分隔符") is not None else ","
                        )
                        value = cast(
                            pd.Series,
                            series_group.agg(
                                lambda x, sep=separator: sep.join(
                                    cast(pd.Series, x).dropna().astype(str)
                                )
                            ),
                        ).reset_index(drop=True)
            else:
                if function == "count_all":
                    raw: Any = len(frame)
                else:
                    series = cast(pd.Series, frame[source])
                    if function == "count":
                        raw = series.count()
                    elif function == "count_distinct":
                        raw = series.nunique(dropna=True)
                    elif function == "sum":
                        raw = cast(pd.Series, pd.to_numeric(series, errors="raise")).sum(
                            min_count=1
                        )
                    elif function == "avg":
                        raw = series.mean()
                    elif function in {"min", "max"}:
                        raw = getattr(series, function)()
                    elif function in {"first", "last"}:
                        raw = (
                            series.dropna().iloc[0 if function == "first" else -1]
                            if cast(bool, series.notna().any())
                            else pd.NA
                        )
                    else:
                        separator = str(
                            item.get("分隔符") if item.get("分隔符") is not None else ","
                        )
                        raw = separator.join(series.dropna().astype(str))
                value = pd.Series([raw])
            columns[str(item["目标字段"])] = self._convert(value, str(item["目标类型"]))
        return pd.DataFrame(columns)

    def execute(self, spec: ExtractionSpec, task_id: str, source: Path) -> pd.DataFrame:
        frame = self._filter(
            self._join(spec, task_id, self._load(spec, task_id, source)), spec, task_id
        )
        return (
            self._aggregate(frame, spec, task_id)
            if spec.for_task(spec.aggregations, task_id)
            else self._select(frame, spec, task_id)
        )
