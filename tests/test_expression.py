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


if __name__ == "__main__":
    unittest.main()
