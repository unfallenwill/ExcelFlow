import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

import pandas as pd

from excelflow.engine import PandasExtractionEngine
from excelflow.output import OutputWriterFactory
from excelflow.schema import ExtractionSpec, ValidationResult
from excelflow.service import ExtractionService


class EngineConditionTest(unittest.TestCase):
    def setUp(self):
        self.engine = PandasExtractionEngine()
        self.frame = pd.DataFrame({
            "o.number": [1, 2, 3, None],
            "o.text": ["apple", "banana", None, "apricot"],
            "o.date": pd.to_datetime(["2026-01-01", "2026-01-02", "2026-01-03", None]),
        })

    def condition(self, field, operator, value1=None, value2=None):
        return self.engine._condition(self.frame, {"字段": field, "运算符": operator, "值1": value1, "值2": value2})

    def test_filter_operator_families(self):
        self.assertEqual(self.condition("o.number", "IN", "1,3").tolist(), [True, False, True, False])
        self.assertEqual(self.condition("o.number", "NOT IN", "1,3").tolist(), [False, True, False, True])
        self.assertEqual(self.condition("o.number", "BETWEEN", 2, 3).tolist(), [False, True, True, False])
        self.assertEqual(self.condition("o.text", "LIKE", "ap%").tolist(), [True, False, False, True])
        self.assertEqual(self.condition("o.text", "NOT LIKE", "ap%").tolist(), [False, True, True, False])
        self.assertEqual(self.condition("o.number", "IS NULL").tolist(), [False, False, False, True])
        self.assertEqual(self.condition("o.number", "IS NOT NULL").tolist(), [True, True, True, False])

    def test_comparisons_and_type_coercion(self):
        self.assertEqual(self.condition("o.number", ">=", "2").tolist(), [False, True, True, False])
        self.assertEqual(self.condition("o.number", "!=", 2).tolist(), [True, False, True, True])
        self.assertEqual(self.condition("o.date", "<", "2026-01-03").tolist(), [True, True, False, False])

    def test_environment_resolution(self):
        os.environ["EXCELFLOW_TEST_VALUE"] = "2"
        try:
            self.assertEqual(self.condition("o.number", "=", "${EXCELFLOW_TEST_VALUE}").tolist(), [False, True, False, False])
        finally:
            os.environ.pop("EXCELFLOW_TEST_VALUE")
        with self.assertRaisesRegex(ValueError, "未设置"):
            self.condition("o.number", "=", "${EXCELFLOW_MISSING}")


class OutputWriterTest(unittest.TestCase):
    def test_all_writers_and_unknown_format(self):
        frame = pd.DataFrame({"id": [1], "name": ["张三"]})
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for fmt in ("csv", "jsonl", "xlsx"):
                path = root / f"result.{fmt}"
                OutputWriterFactory().create(fmt).write(frame, path)
                self.assertTrue(path.exists())
        with self.assertRaises(ValueError):
            OutputWriterFactory().create("xml")


class ExtractionServiceTest(unittest.TestCase):
    def test_validate_converts_repository_exception(self):
        repository = Mock(); repository.load.side_effect = OSError("broken")
        result = ExtractionService(repository=repository).validate(Path("plan.xlsx"))
        self.assertFalse(result.ok)
        self.assertIn("broken", result.errors[0])

    def test_run_rejects_invalid_and_disabled_tasks(self):
        repository, validator, engine = Mock(), Mock(), Mock()
        repository.load.return_value = ExtractionSpec(plans=[{"任务ID": "task", "启用": "是"}])
        validator.validate.return_value = ValidationResult(errors=["bad plan"])
        service = ExtractionService(repository=repository, validator=validator, engine=engine)
        with self.assertRaisesRegex(ValueError, "bad plan"):
            service.run(Path("plan.xlsx"), "task", Path("source.xlsx"))
        validator.validate.return_value = ValidationResult()
        repository.load.return_value = ExtractionSpec(plans=[{"任务ID": "task", "启用": "否"}])
        with self.assertRaisesRegex(ValueError, "未启用"):
            service.run(Path("plan.xlsx"), "task", Path("source.xlsx"))


if __name__ == "__main__":
    unittest.main()
