import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from main import create_template, run_task, validate


class ExtractionTest(unittest.TestCase):
    def _plan(self, root: Path, output_format: str) -> Path:
        plan_path = root / "plan.xlsx"
        create_template(plan_path)
        wb = load_workbook(plan_path)
        plan = wb["抽取计划"]
        plan.delete_rows(2, plan.max_row)
        plan.append(["orders", "是", "增量", "o.updated_at", "2026-01-01", "2026-02-01", output_format, str(root / f"result.{output_format}"), 2, "test", ""])
        objects = wb["数据对象"]
        objects.delete_rows(2, objects.max_row)
        objects.append(["orders", "订单", "o", 1, "是", ""])
        objects.append(["orders", "客户", "c", 1, "否", ""])
        joins = wb["关联关系"]
        joins.delete_rows(2, joins.max_row)
        joins.append(["orders", 1, "LEFT JOIN", "o.customer_id", "c", "c.customer_id", ""])
        fields = wb["字段映射"]
        fields.delete_rows(2, fields.max_row)
        fields.append(["orders", "o.id", "order_id", "integer", "", 1, ""])
        fields.append(["orders", "c.name", "customer_name", "string", "", 2, ""])
        fields.append(["orders", "o.amount", "amount", "decimal", "", 3, ""])
        filters = wb["过滤条件"]
        filters.delete_rows(2, filters.max_row)
        filters.append(["orders", 1, 1, "o.amount", ">=", 20, "", ""])
        wb.save(plan_path)
        return plan_path

    def test_excel_to_jsonl(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "订单"
            ws.append(["id", "customer_id", "amount", "updated_at"])
            ws.append([1, 10, 20, "2026-01-03"])
            customers = wb.create_sheet("客户")
            customers.append(["customer_id", "name"])
            customers.append([10, "张三"])
            wb.save(source)
            plan = self._plan(root, "jsonl")
            count, output = run_task(plan, "orders", source)
            self.assertEqual(count, 1)
            self.assertEqual(json.loads(output.read_text()), {"order_id": 1, "customer_name": "张三", "amount": 20})


if __name__ == "__main__":
    unittest.main()
