import unittest

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

    def test_rejects_missing_field_and_unsupported_syntax(self):
        with self.assertRaisesRegex(ValueError, "不存在的字段"):
            self.evaluator.evaluate("i.missing + 1", self.frame)
        with self.assertRaises(ValueError):
            self.evaluator.evaluate("i.price > 1", self.frame)


if __name__ == "__main__":
    unittest.main()
