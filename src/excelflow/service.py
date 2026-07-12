from pathlib import Path

from .engine import ExtractionEngine, PandasExtractionEngine
from .output import OutputWriterFactory
from .repository import ExcelSpecRepository
from .schema import ValidationResult
from .validator import SpecValidator


class ExtractionService:
    """Application service coordinating repository, validation, engine and writer strategies."""

    def __init__(self, repository=None, validator=None, engine: ExtractionEngine | None = None, writer_factory=None):
        self.repository = repository or ExcelSpecRepository()
        self.validator = validator or SpecValidator()
        self.engine = engine or PandasExtractionEngine()
        self.writer_factory = writer_factory or OutputWriterFactory()

    def validate(self, plan_path: Path) -> ValidationResult:
        try: return self.validator.validate(self.repository.load(plan_path))
        except Exception as exc: return ValidationResult(errors=[str(exc)])

    def preview(self, plan_path: Path, task_id: str) -> str:
        spec, plan = self.repository.load(plan_path), None
        plan = spec.task(task_id)
        objects, joins = spec.for_task(spec.objects, task_id), spec.for_task(spec.joins, task_id)
        lines = [f"Pandas 执行计划: {task_id}", "读取: " + ", ".join(f'{x["Sheet名称"]} -> {x["对象别名"]}' for x in objects)]
        for order in sorted({int(x["关联顺序"]) for x in joins}):
            rows = [x for x in joins if int(x["关联顺序"]) == order]
            lines.append(f'关联{order}: {rows[0]["关联类型"]} {rows[0]["右侧对象"]} ON ' + " AND ".join(f'{x["左侧字段"]}={x["右侧字段"]}' for x in rows))
        lines.extend([f"抽取模式: {plan['抽取模式']}", f"过滤条件: {len(spec.for_task(spec.filters, task_id))} 条", f"输出字段: {len(spec.for_task(spec.fields, task_id))} 个"])
        return "\n".join(lines)

    def run(self, plan_path: Path, task_id: str, source_path: Path) -> tuple[int, Path]:
        spec = self.repository.load(plan_path)
        result = self.validator.validate(spec)
        if not result.ok: raise ValueError("Excel 校验失败:\n" + "\n".join(result.errors))
        plan = spec.task(task_id)
        if str(plan.get("启用")) != "是": raise ValueError(f"任务 {task_id} 未启用")
        frame = self.engine.execute(spec, task_id, source_path)
        output = Path(str(plan["输出路径"])); output.parent.mkdir(parents=True, exist_ok=True)
        self.writer_factory.create(str(plan["输出格式"])).write(frame, output)
        return len(frame), output
