import ast

import pandas as pd


class SafeExpressionEvaluator:
    """Small AST interpreter for derived columns; deliberately avoids eval()."""

    def evaluate(self, expression: str, frame: pd.DataFrame):
        return self._node(ast.parse(expression, mode="eval").body, frame)

    def _node(self, node: ast.AST, frame: pd.DataFrame):
        if isinstance(node, ast.Constant): return node.value
        if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
            field = f"{node.value.id}.{node.attr}"
            if field not in frame: raise ValueError(f"衍生列表达式引用了不存在的字段: {field}")
            return frame[field]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = self._node(node.operand, frame)
            return value if isinstance(node.op, ast.UAdd) else -value
        if isinstance(node, ast.BinOp) and isinstance(node.op, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Mod)):
            left, right = self._node(node.left, frame), self._node(node.right, frame)
            operations = {ast.Add: lambda: left + right, ast.Sub: lambda: left - right, ast.Mult: lambda: left * right,
                          ast.Div: lambda: left / right, ast.Mod: lambda: left % right}
            return operations[type(node.op)]()
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            args = [self._node(x, frame) for x in node.args]
            name = node.func.id.lower()
            if name == "coalesce" and len(args) == 2: return args[0].fillna(args[1]) if isinstance(args[0], pd.Series) else args[0] if pd.notna(args[0]) else args[1]
            if name == "abs" and len(args) == 1: return abs(args[0])
            if name == "round" and len(args) in {1, 2}: return args[0].round(int(args[1])) if isinstance(args[0], pd.Series) else round(*args)
            if name == "clip" and len(args) == 3 and isinstance(args[0], pd.Series): return args[0].clip(lower=args[1], upper=args[2])
        raise ValueError("转换表达式仅支持别名.字段、常量、四则运算、%、coalesce、abs、round 和 clip")
