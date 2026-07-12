import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from main import create_template, run_task, validate


class ExtractionTest(unittest.TestCase):
    def _plan(self, root: Path, source_type: str, source_path: Path, object_name: str, output_format: str) -> Path:
        plan_path = root / "plan.xlsx"
        create_template(plan_path)
        wb = load_workbook(plan_path)
        plan = wb["抽取计划"]
        plan.delete_rows(2, plan.max_row)
        plan.append(["orders", "是", "source", object_name, "增量", "updated_at", "2026-01-01", "2026-02-01", "amount >= 20", output_format, str(root / f"result.{output_format}"), 2, "test", ""])
        sources = wb["数据源"]
        sources.delete_rows(2, sources.max_row)
        sources.append(["source", source_type, str(source_path), "", "", "", "", "", "{}"])
        fields = wb["字段映射"]
        fields.delete_rows(2, fields.max_row)
        fields.append(["orders", "id", "order_id", "integer", "", 1, ""])
        fields.append(["orders", "amount", "amount", "decimal", "", 2, ""])
        wb.save(plan_path)
        return plan_path

    def test_json_to_csv(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.json"
            source.write_text(json.dumps({"订单": [
                {"id": 1, "amount": 10, "updated_at": "2026-01-02"},
                {"id": 2, "amount": 20, "updated_at": "2026-01-03"},
                {"id": 3, "amount": 30, "updated_at": "2026-03-01"},
            ]}), encoding="utf-8")
            plan = self._plan(root, "json", source, "订单", "csv")
            self.assertTrue(validate(plan).ok)
            count, output = run_task(plan, "orders")
            self.assertEqual(count, 1)
            self.assertIn("2,20", output.read_text(encoding="utf-8-sig"))

    def test_excel_to_jsonl(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "订单"
            ws.append(["id", "amount", "updated_at"])
            ws.append([1, 20, "2026-01-03"])
            wb.save(source)
            plan = self._plan(root, "excel", source, "订单", "jsonl")
            count, output = run_task(plan, "orders")
            self.assertEqual(count, 1)
            self.assertEqual(json.loads(output.read_text()), {"order_id": 1, "amount": 20})


if __name__ == "__main__":
    unittest.main()
