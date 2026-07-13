import ast
from functools import reduce
from typing import Any, Literal

import pandas as pd


class SafeExpressionEvaluator:
    """Whitelist-based AST interpreter for vectorized derived-column expressions."""

    _date_units: dict[str, Literal["D", "h", "m"]] = {
        "day": "D",
        "days": "D",
        "hour": "h",
        "hours": "h",
        "minute": "m",
        "minutes": "m",
    }

    def evaluate(self, expression: str, frame: pd.DataFrame) -> Any:
        try:
            return self._node(ast.parse(expression, mode="eval").body, frame)
        except SyntaxError as exc:
            raise ValueError("转换表达式语法错误") from exc

    @staticmethod
    def _series(value: Any) -> bool:
        return isinstance(value, pd.Series)

    @staticmethod
    def _strings(value: Any) -> Any:
        if isinstance(value, pd.Series):
            return value.astype("string")
        return pd.NA if pd.isna(value) else str(value)

    @staticmethod
    def _dates(value: Any) -> Any:
        return pd.to_datetime(value, errors="raise")

    def _node(self, node: ast.AST, frame: pd.DataFrame) -> Any:
        if isinstance(node, ast.Constant):
            return node.value
        if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
            field = f"{node.value.id}.{node.attr}"
            if field not in frame:
                raise ValueError(f"衍生列表达式引用了不存在的字段: {field}")
            return frame[field]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub, ast.Not)):
            value = self._node(node.operand, frame)
            if isinstance(node.op, ast.UAdd):
                return +value
            if isinstance(node.op, ast.USub):
                return -value
            return ~value if self._series(value) else not value
        if isinstance(node, ast.BinOp) and isinstance(
            node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod)
        ):
            left, right = self._node(node.left, frame), self._node(node.right, frame)
            operations = {
                ast.Add: lambda: left + right,
                ast.Sub: lambda: left - right,
                ast.Mult: lambda: left * right,
                ast.Div: lambda: left / right,
                ast.Mod: lambda: left % right,
            }
            return operations[type(node.op)]()
        if isinstance(node, ast.Compare):
            left = self._node(node.left, frame)
            masks = []
            operations = {
                ast.Eq: lambda a, b: a == b,
                ast.NotEq: lambda a, b: a != b,
                ast.Gt: lambda a, b: a > b,
                ast.GtE: lambda a, b: a >= b,
                ast.Lt: lambda a, b: a < b,
                ast.LtE: lambda a, b: a <= b,
            }
            for operator, comparator in zip(node.ops, node.comparators, strict=True):
                right = self._node(comparator, frame)
                if type(operator) not in operations:
                    break
                masks.append(operations[type(operator)](left, right))
                left = right
            else:
                return reduce(lambda a, b: a & b, masks)
        if isinstance(node, ast.BoolOp) and isinstance(node.op, (ast.And, ast.Or)):
            values = [self._node(value, frame) for value in node.values]
            operation = (
                (lambda a, b: a & b) if isinstance(node.op, ast.And) else (lambda a, b: a | b)
            )
            return reduce(operation, values)
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
            args = [self._node(value, frame) for value in node.args]
            return self._call(node.func.id.lower(), args)
        raise ValueError("转换表达式包含不受支持的语法")

    def _call(self, name: str, args: list[Any]) -> Any:
        if name == "coalesce" and len(args) == 2:
            return (
                args[0].fillna(args[1])
                if self._series(args[0])
                else args[0]
                if pd.notna(args[0])
                else args[1]
            )
        if name == "abs" and len(args) == 1:
            return abs(args[0])
        if name == "round" and len(args) in {1, 2}:
            digits = int(args[1]) if len(args) == 2 else 0
            return args[0].round(digits) if self._series(args[0]) else round(args[0], digits)
        if name == "clip" and len(args) == 3 and self._series(args[0]):
            return args[0].clip(args[1], args[2])
        if name in {"upper", "lower", "trim", "length"} and len(args) == 1:
            value = self._strings(args[0])
            if self._series(value):
                return {
                    "upper": value.str.upper,
                    "lower": value.str.lower,
                    "trim": value.str.strip,
                    "length": value.str.len,
                }[name]()
            return {
                "upper": value.upper,
                "lower": value.lower,
                "trim": value.strip,
                "length": value.__len__,
            }[name]()
        if name == "replace" and len(args) == 3:
            value = self._strings(args[0])
            old, new = str(args[1]), str(args[2])
            return (
                value.str.replace(old, new, regex=False)
                if self._series(value)
                else value.replace(old, new)
            )
        if name == "substring" and len(args) in {2, 3}:
            value, start = self._strings(args[0]), int(args[1])
            stop = start + int(args[2]) if len(args) == 3 else None
            return value.str.slice(start, stop) if self._series(value) else value[start:stop]
        if name in {"contains", "startswith", "endswith"} and len(args) == 2:
            value, pattern = self._strings(args[0]), str(args[1])
            if self._series(value):
                return (
                    getattr(value.str, name)(pattern, na=False, regex=False)
                    if name == "contains"
                    else getattr(value.str, name)(pattern, na=False)
                )
            return pattern in value if name == "contains" else getattr(value, name)(pattern)
        if name == "concat" and len(args) >= 1:
            values = [self._strings(value) for value in args]
            return reduce(lambda a, b: a + b, values)
        if name == "concat_ws" and len(args) >= 2:
            separator, values = str(args[0]), [self._strings(value) for value in args[1:]]
            if any(self._series(value) for value in values):
                index = next(value.index for value in values if self._series(value))
                series = [
                    value if self._series(value) else pd.Series(value, index=index, dtype="string")
                    for value in values
                ]
                return reduce(lambda left, right: left.str.cat(right, sep=separator), series)
            return pd.NA if any(pd.isna(value) for value in values) else separator.join(values)
        if name == "to_number" and len(args) == 1:
            return pd.to_numeric(args[0], errors="raise")
        if name == "to_string" and len(args) == 1:
            return self._strings(args[0])
        if name == "to_date" and len(args) == 1:
            return self._dates(args[0])
        if name == "dateformat" and len(args) == 2:
            value = self._dates(args[0])
            return (
                value.dt.strftime(str(args[1]))
                if self._series(value)
                else value.strftime(str(args[1]))
            )
        if name in {"year", "month", "day"} and len(args) == 1:
            value = self._dates(args[0])
            return getattr(value.dt, name) if self._series(value) else getattr(value, name)
        if name == "date_add" and len(args) == 3 and str(args[2]).lower() in self._date_units:
            return self._dates(args[0]) + pd.to_timedelta(
                args[1], unit=self._date_units[str(args[2]).lower()]
            )
        if name == "date_diff" and len(args) == 3 and str(args[2]).lower() in self._date_units:
            delta = self._dates(args[0]) - self._dates(args[1])
            unit = self._date_units[str(args[2]).lower()]
            return delta / pd.to_timedelta(1, unit=unit)
        if name == "is_null" and len(args) == 1:
            return pd.isna(args[0])
        if name == "if_else" and len(args) == 3:
            condition, yes, no = args
            if self._series(condition):
                yes = yes if self._series(yes) else pd.Series(yes, index=condition.index)
                return yes.where(condition.fillna(False), no)
            return yes if not pd.isna(condition) and condition else no
        if name in {"ceil", "floor", "sqrt"} and len(args) == 1:
            import numpy as np

            return {"ceil": np.ceil, "floor": np.floor, "sqrt": np.sqrt}[name](args[0])
        if name == "power" and len(args) == 2:
            return args[0] ** args[1]
        if name in {"min_value", "max_value"} and len(args) >= 1:
            series = (
                [
                    value
                    if self._series(value)
                    else pd.Series(value, index=next(x.index for x in args if self._series(x)))
                    for value in args
                ]
                if any(self._series(x) for x in args)
                else args
            )
            if any(self._series(x) for x in args):
                return (
                    pd.concat(series, axis=1).min(axis=1)
                    if name == "min_value"
                    else pd.concat(series, axis=1).max(axis=1)
                )
            return min(args) if name == "min_value" else max(args)
        raise ValueError(f"不支持的转换函数或参数: {name}")
