import unittest

import pandas as pd
from pandas.testing import assert_series_equal

from excelflow.expression import SafeExpressionEvaluator


class ExpressionFunctionTest(unittest.TestCase):
    def setUp(self):
        self.evaluator = SafeExpressionEvaluator()
        self.frame = pd.DataFrame({
            "p.first": [" Alice ", "小", None], "p.last": ["Smith", "蓝", "Z"],
            "p.code": ["ab-12", "CN-9", None], "p.number": ["10.5", "-2", "3"],
            "p.date": ["2024-02-29", "2024-03-02", None], "p.amount": [10.2, 20.8, None],
        }, index=[3, 5, 7])

    def evaluate(self, expression): return self.evaluator.evaluate(expression, self.frame)

    def test_string_transformations_and_search(self):
        self.assertEqual(self.evaluate("upper(trim(p.first))").tolist(), ["ALICE", "小", pd.NA])
        self.assertEqual(self.evaluate("lower(p.code)").tolist(), ["ab-12", "cn-9", pd.NA])
        self.assertEqual(self.evaluate('replace(p.code, "-", "")').tolist(), ["ab12", "CN9", pd.NA])
        self.assertEqual(self.evaluate("substring(p.code, 0, 2)").tolist(), ["ab", "CN", pd.NA])
        self.assertEqual(self.evaluate('contains(p.code, "-")').tolist(), [True, True, False])
        self.assertEqual(self.evaluate('startswith(p.code, "CN")').tolist(), [False, True, False])
        self.assertEqual(self.evaluate('endswith(p.code, "12")').tolist(), [True, False, False])
        self.assertEqual(self.evaluate("length(p.last)").tolist(), [5, 1, 1])

    def test_concat_nulls_and_broadcast(self):
        self.assertEqual(self.evaluate('concat(trim(p.first), " ", p.last)').tolist(), ["Alice Smith", "小 蓝", pd.NA])
        self.assertEqual(self.evaluate('concat_ws("-", "ID", p.code)').tolist(), ["ID-ab-12", "ID-CN-9", pd.NA])
        self.assertEqual(self.evaluate('concat_ws("/", "A", "B")'), "A/B")
        self.assertTrue(pd.isna(self.evaluate('concat(None, "x")')))
        self.assertTrue(pd.isna(self.evaluate("to_string(None)")))

    def test_type_and_numeric_functions(self):
        assert_series_equal(self.evaluate("to_number(p.number)"), pd.Series([10.5, -2.0, 3.0], index=[3, 5, 7], name="p.number"))
        self.assertEqual(str(self.evaluate("to_string(p.amount)").dtype), "string")
        self.assertEqual(self.evaluate("ceil(p.amount)").iloc[:2].tolist(), [11.0, 21.0])
        self.assertEqual(self.evaluate("floor(p.amount)").iloc[:2].tolist(), [10.0, 20.0])
        self.assertEqual(self.evaluate("power(p.amount, 2)").iloc[0], 104.03999999999999)
        self.assertEqual(self.evaluate("min_value(p.amount, 15)").iloc[:2].tolist(), [10.2, 15.0])
        self.assertEqual(self.evaluate("max_value(p.amount, 15)").iloc[:2].tolist(), [15.0, 20.8])
        self.assertEqual(self.evaluate("sqrt(9)"), 3.0)
        self.assertEqual(self.evaluate("min_value(3, 1, 2)"), 1)
        self.assertEqual(self.evaluate("max_value(3, 1, 2)"), 3)
        self.assertEqual(self.evaluate('upper("abc")'), "ABC")
        self.assertTrue(self.evaluate('contains("abc", "b")'))
        self.assertEqual(self.evaluate('substring("abcd", 1)'), "bcd")

    def test_date_functions(self):
        formatted = self.evaluate('dateformat(p.date, "%Y/%m/%d")')
        self.assertEqual(formatted.iloc[:2].tolist(), ["2024/02/29", "2024/03/02"])
        self.assertTrue(pd.isna(formatted.iloc[2]))
        self.assertEqual(self.evaluate("year(p.date)").iloc[:2].tolist(), [2024.0, 2024.0])
        self.assertEqual(self.evaluate("month(p.date)").iloc[:2].tolist(), [2.0, 3.0])
        self.assertEqual(self.evaluate("day(p.date)").iloc[:2].tolist(), [29.0, 2.0])
        self.assertEqual(self.evaluate('date_diff(date_add(p.date, 2, "day"), p.date, "day")').iloc[:2].tolist(), [2.0, 2.0])

    def test_comparisons_boolean_and_if_else(self):
        result = self.evaluate('if_else((p.amount >= 20) and not is_null(p.amount), "high", "low")')
        self.assertEqual(result.tolist(), ["low", "high", "low"])
        self.assertEqual(self.evaluate("if_else(True, 1, 0)"), 1)
        self.assertEqual(self.evaluate("if_else(None, 1, 0)"), 0)
        self.assertFalse(self.evaluate("1 > 2 or 3 < 2"))

    def test_bad_calls_syntax_and_conversion(self):
        for expression in ["unknown(p.code)", "upper()", 'date_add(p.date, 1, "month")', "p.code[0]", "p.code.str.upper()", "lambda: 1", "upper(value=p.code)",
                           "[x for x in p.code]", '{"x": p.code}', "(1).__class__", "p.code if True else p.last", "(x := 1)", 'f"{p.code}"']:
            with self.subTest(expression=expression), self.assertRaises(ValueError): self.evaluate(expression)
        with self.assertRaisesRegex(ValueError, "语法错误"): self.evaluate("upper(")
        with self.assertRaises(ValueError): self.evaluate('to_number("bad")')


if __name__ == "__main__": unittest.main()
