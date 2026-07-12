import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from main import create_template, run_task, validate


class ExtractionTest(unittest.TestCase):
    def _plan(self, root: Path, source_path: Path, object_name: str, output_format: str) -> Path:
        plan_path = root / "plan.xlsx"
        create_template(plan_path)
        wb = load_workbook(plan_path)
        plan = wb["抽取计划"]
        plan.delete_rows(2, plan.max_row)
        plan.append(["orders", "是", object_name, 1, "增量", "updated_at", "2026-01-01", "2026-02-01", output_format, str(root / f"result.{output_format}"), 2, "test", ""])
        fields = wb["字段映射"]
        fields.delete_rows(2, fields.max_row)
        fields.append(["orders", "id", "order_id", "integer", "", 1, ""])
        fields.append(["orders", "amount", "amount", "decimal", "", 2, ""])
        filters = wb["过滤条件"]
        filters.delete_rows(2, filters.max_row)
        filters.append(["orders", 1, 1, "amount", ">=", 20, "", ""])
        wb.save(plan_path)
        return plan_path

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
            plan = self._plan(root, source, "订单", "jsonl")
            count, output = run_task(plan, "orders", source)
            self.assertEqual(count, 1)
            self.assertEqual(json.loads(output.read_text()), {"order_id": 1, "amount": 20})


if __name__ == "__main__":
    unittest.main()
