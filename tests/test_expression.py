import unittest
import warnings

import pandas as pd

from excelflow.expression import SafeExpressionEvaluator


class SafeExpressionEvaluatorTest(unittest.TestCase):
    def setUp(self):
        self.evaluator = SafeExpressionEvaluator()
        self.frame = pd.DataFrame({"i.quantity": [2, None], "i.price": [10.5, 8.0]})

    def test_derived_column_with_coalesce(self):
        result = self.evaluator.evaluate("coalesce(i.quantity, 0) * i.price", self.frame)
        self.assertEqual(result.tolist(), [21.0, 0.0])

    def test_rejects_arbitrary_python(self):
        with self.assertRaises(ValueError):
            self.evaluator.evaluate("__import__('os').system('echo unsafe')", self.frame)

    def test_supported_arithmetic_and_functions(self):
        self.assertEqual(self.evaluator.evaluate("-i.price + abs(i.quantity)", self.frame).iloc[0], -8.5)
        self.assertEqual(self.evaluator.evaluate("round(i.price / 3, 2)", self.frame).iloc[0], 3.5)
        self.assertEqual(self.evaluator.evaluate("i.quantity % 2", self.frame).iloc[0], 0)

    def test_rejects_missing_field(self):
        with self.assertRaisesRegex(ValueError, "不存在的字段"):
            self.evaluator.evaluate("i.missing + 1", self.frame)

    def test_signed_window_overrun_days(self):
        frame = pd.DataFrame({
            "v.actual_day": [22, 35, 30],
            "v.plan_day": [28, 28, 28],
            "v.window_days": [3, 3, 3],
        })
        expression = "v.actual_day - clip(v.actual_day, v.plan_day - v.window_days, v.plan_day + v.window_days)"
        result = self.evaluator.evaluate(expression, frame)
        self.assertEqual(result.tolist(), [-3, 4, 0])

    def test_series_division_by_zero_propagates_inf(self):
        # 当前契约: Series / 0 不抛 ZeroDivisionError，而是按 pandas 语义产生 inf，
        # 并向下游聚合传播（sum 仍为 inf）。锁定该行为，使任何"对非有限值抛错"之类
        # 的契约变更都必须显式更新本用例，避免静默把 inf 写进 CSV/JSON。
        frame = pd.DataFrame({"i.price": [100.0, 50.0], "i.qty": [0, 5]})
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", RuntimeWarning)
            result = self.evaluator.evaluate("i.price / i.qty", frame)
        self.assertEqual(result.tolist(), [float("inf"), 10.0])
        self.assertEqual(result.sum(), float("inf"))  # inf 会污染聚合总额


if __name__ == "__main__":
    unittest.main()
