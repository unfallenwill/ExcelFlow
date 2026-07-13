import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock

import pandas as pd
from openpyxl import Workbook

from excelflow.engine import PandasExtractionEngine
from excelflow.output import OutputWriterFactory
from excelflow.schema import ExtractionSpec, ValidationResult
from excelflow.service import ExtractionService


class EngineConditionTest(unittest.TestCase):
    def setUp(self):
        self.engine = PandasExtractionEngine()
        self.frame = pd.DataFrame(
            {
                "o.number": [1, 2, 3, None],
                "o.text": ["apple", "banana", None, "apricot"],
                "o.date": pd.to_datetime(["2026-01-01", "2026-01-02", "2026-01-03", None]),
            }
        )

    def condition(self, field, operator, value1=None, value2=None):
        return self.engine._condition(
            self.frame, {"字段": field, "运算符": operator, "值1": value1, "值2": value2}
        )

    def test_filter_operator_families(self):
        self.assertEqual(
            self.condition("o.number", "IN", "1,3").tolist(), [True, False, True, False]
        )
        self.assertEqual(
            self.condition("o.number", "NOT IN", "1,3").tolist(), [False, True, False, True]
        )
        self.assertEqual(
            self.condition("o.number", "BETWEEN", 2, 3).tolist(), [False, True, True, False]
        )
        self.assertEqual(
            self.condition("o.text", "LIKE", "ap%").tolist(), [True, False, False, True]
        )
        self.assertEqual(
            self.condition("o.text", "NOT LIKE", "ap%").tolist(), [False, True, True, False]
        )
        self.assertEqual(
            self.condition("o.number", "IS NULL").tolist(), [False, False, False, True]
        )
        self.assertEqual(
            self.condition("o.number", "IS NOT NULL").tolist(), [True, True, True, False]
        )

    def test_comparisons_and_type_coercion(self):
        self.assertEqual(self.condition("o.number", ">=", "2").tolist(), [False, True, True, False])
        self.assertEqual(self.condition("o.number", "!=", 2).tolist(), [True, False, True, True])
        self.assertEqual(
            self.condition("o.date", "<", "2026-01-03").tolist(), [True, True, False, False]
        )


class EngineFilterSemanticsTest(unittest.TestCase):
    """锁定 _filter 的组内 AND / 组间 OR 语义。

    既有用例只覆盖单个 _condition，从未直接驱动 _filter；若有人交换 `&=` 与 `|=`，
    整套既有测试仍会全绿，但提取出的行集合会整体出错。这里把两种语义分别钉死。
    """

    def setUp(self):
        self.engine = PandasExtractionEngine()
        self.frame = pd.DataFrame({"o.id": [1, 2, 3, 4]})

    def _filter(self, filters):
        return self.engine._filter(self.frame, ExtractionSpec(filters=filters), "t")

    def test_same_group_combines_with_AND(self):
        # 同一条件组内两个条件 -> 组内 AND: (id >= 2) & (id <= 3) -> [2, 3]
        filters = [
            {"任务ID": "t", "条件组": 1, "条件序号": 1, "字段": "o.id", "运算符": ">=", "值1": 2},
            {"任务ID": "t", "条件组": 1, "条件序号": 2, "字段": "o.id", "运算符": "<=", "值1": 3},
        ]
        self.assertEqual(self._filter(filters)["o.id"].tolist(), [2, 3])

    def test_different_groups_combine_with_OR(self):
        # 两个条件组各自一个条件 -> 组间 OR: (id <= 2) | (id >= 3) -> [1, 2, 3, 4]
        # 若误写成组间 AND，结果会塌缩成空集，断言随即失败。
        filters = [
            {"任务ID": "t", "条件组": 1, "条件序号": 1, "字段": "o.id", "运算符": "<=", "值1": 2},
            {"任务ID": "t", "条件组": 2, "条件序号": 1, "字段": "o.id", "运算符": ">=", "值1": 3},
        ]
        self.assertEqual(self._filter(filters)["o.id"].tolist(), [1, 2, 3, 4])


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


class OutputWriterRoundTripTest(unittest.TestCase):
    """把三种 writer 的字节真正读回来，锁定输出契约。

    既有用例只断言 path.exists()；编码(BOM)、JSONL 行向、Excel sheet 名、中文与 NA
    的保真度都没被验证——任何格式回归都会假绿。
    """

    def setUp(self):
        self.frame = pd.DataFrame(
            {
                "id": [1, 2],
                "name": ["张三", None],  # unicode + NA
                "joined": pd.to_datetime(["2026-01-01", "2026-02-01"]),
            }
        )

    def test_csv_round_trip_preserves_unicode_and_columns(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "result.csv"
            OutputWriterFactory().create("csv").write(self.frame, path)
            back = pd.read_csv(path, encoding="utf-8-sig", parse_dates=["joined"])
            self.assertEqual(back.columns.tolist(), ["id", "name", "joined"])
            self.assertEqual(back["id"].tolist(), [1, 2])
            self.assertEqual(back.loc[0, "name"], "张三")  # utf-8-sig BOM 不致乱码
            self.assertTrue(pd.isna(back.loc[1, "name"]))  # NA 保真

    def test_jsonl_round_trip_is_line_oriented_and_unicode(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "result.jsonl"
            OutputWriterFactory().create("jsonl").write(self.frame, path)
            lines = path.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(lines), 2)  # records+lines: 每条一行，而非 JSON 数组
            records = [json.loads(line) for line in lines]
            self.assertEqual(records[0]["id"], 1)
            self.assertEqual(records[0]["name"], "张三")  # force_ascii=False: 中文未转义

    def test_excel_round_trip_uses_data_sheet_name(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "result.xlsx"
            OutputWriterFactory().create("xlsx").write(self.frame, path)
            back = pd.read_excel(path, sheet_name="data")  # 锁定 sheet 名为 "data"
            self.assertEqual(back.columns.tolist(), ["id", "name", "joined"])
            self.assertEqual(back.loc[0, "name"], "张三")


class EngineLoadAndSelectTest(unittest.TestCase):
    def test_load_rejects_missing_sheet_and_duplicate_headers(self):
        engine = PandasExtractionEngine()
        # 源工作簿里不存在引用的 Sheet → 必须给出可读错误，而非 pandas 的原始异常
        missing = ExtractionSpec(
            plans=[{"任务ID": "t", "启用": "是"}],
            objects=[
                {"任务ID": "t", "Sheet名称": "缺失", "对象别名": "o", "表头行": 1, "是否主表": "是"}
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            Workbook().save(source)
            with self.assertRaisesRegex(ValueError, "不存在工作表"):
                engine._load(missing, "t", source)

        # 表头重复 → pandas 会静默重命名为 id.1 导致下游关联错列，必须提前拦截
        dup = ExtractionSpec(
            plans=[{"任务ID": "t", "启用": "是"}],
            objects=[
                {"任务ID": "t", "Sheet名称": "数据", "对象别名": "o", "表头行": 1, "是否主表": "是"}
            ],
        )
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            assert ws is not None
            ws.title = "数据"
            ws.append(["id", "id"])
            ws.append([1, 2])
            wb.save(source)
            with self.assertRaisesRegex(ValueError, "表头不合法"):
                engine._load(dup, "t", source)

    def test_select_passes_through_all_columns_when_no_fields_configured(self):
        # 没有字段映射时，应原样输出关联+过滤后的全部限定名列
        frame = pd.DataFrame({"o.id": [1, 2], "o.name": ["a", "b"]})
        result = PandasExtractionEngine()._select(frame, ExtractionSpec(fields=[]), "t")
        self.assertEqual(result.columns.tolist(), ["o.id", "o.name"])
        self.assertEqual(
            result.to_dict("records"), [{"o.id": 1, "o.name": "a"}, {"o.id": 2, "o.name": "b"}]
        )
        self.assertEqual(result.index.tolist(), [0, 1])


class ExtractionServiceTest(unittest.TestCase):
    def test_validate_converts_repository_exception(self):
        repository = Mock()
        repository.load.side_effect = OSError("broken")
        result = ExtractionService(repository=repository).validate(Path("plan.xlsx"))
        self.assertFalse(result.ok)
        self.assertIn("broken", result.errors[0])

    def test_run_rejects_invalid_and_disabled_tasks(self):
        repository, validator, engine = Mock(), Mock(), Mock()
        repository.load.return_value = ExtractionSpec(plans=[{"任务ID": "task", "启用": "是"}])
        validator.validate.return_value = ValidationResult(errors=["bad plan"])
        service = ExtractionService(repository=repository, validator=validator, engine=engine)
        with self.assertRaisesRegex(ValueError, "bad plan"):
            service.run(Path("plan.xlsx"), "task", Path("source.xlsx"), "csv", Path("out.data"))
        validator.validate.return_value = ValidationResult()
        repository.load.return_value = ExtractionSpec(plans=[{"任务ID": "task", "启用": "否"}])
        with self.assertRaisesRegex(ValueError, "未启用"):
            service.run(Path("plan.xlsx"), "task", Path("source.xlsx"), "csv", Path("out.data"))


if __name__ == "__main__":
    unittest.main()
