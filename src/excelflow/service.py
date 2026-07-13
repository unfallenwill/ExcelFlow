from pathlib import Path

from .engine import ExtractionEngine, PandasExtractionEngine
from .output import OutputWriterFactory
from .repository import ExcelSpecRepository
from .schema import ExtractionSpec, ValidationResult
from .validator import SpecValidator


class ExtractionService:
    """Application service coordinating repository, validation, engine and writer strategies."""

    def __init__(
        self,
        repository: ExcelSpecRepository | None = None,
        validator: SpecValidator | None = None,
        engine: ExtractionEngine | None = None,
        writer_factory: OutputWriterFactory | None = None,
    ) -> None:
        self.repository = repository or ExcelSpecRepository()
        self.validator = validator or SpecValidator()
        self.engine = engine or PandasExtractionEngine()
        self.writer_factory = writer_factory or OutputWriterFactory()

    def validate(self, plan_path: Path) -> ValidationResult:
        try:
            return self.validator.validate(self.repository.load(plan_path))
        except Exception as exc:
            return ValidationResult(errors=[str(exc)])

    def preview(self, plan_path: Path, task_id: str) -> str:
        spec = self.repository.load(plan_path)
        spec.task(task_id)  # raises ValueError if the task does not exist
        objects, joins = spec.for_task(spec.objects, task_id), spec.for_task(spec.joins, task_id)
        lines = [
            f"Pandas 执行计划: {task_id}",
            "读取: " + ", ".join(f"{x['Sheet名称']} -> {x['对象别名']}" for x in objects),
        ]
        for order in sorted({int(x["关联顺序"]) for x in joins}):
            rows = [x for x in joins if int(x["关联顺序"]) == order]
            lines.append(
                f"关联{order}: {rows[0]['关联类型']} {rows[0]['右侧对象']} ON "
                + " AND ".join(f"{x['左侧字段']}={x['右侧字段']}" for x in rows)
            )
        groups, aggregations = (
            spec.for_task(spec.groups, task_id),
            spec.for_task(spec.aggregations, task_id),
        )
        lines.append(f"过滤条件: {len(spec.for_task(spec.filters, task_id))} 条")
        if aggregations:
            lines.extend([f"分组字段: {len(groups)} 个", f"聚合规则: {len(aggregations)} 条"])
        else:
            lines.append(f"输出字段: {len(spec.for_task(spec.fields, task_id))} 个")
        return "\n".join(lines)

    def _load_validated(self, plan_path: Path) -> ExtractionSpec:
        spec = self.repository.load(plan_path)
        result = self.validator.validate(spec)
        if not result.ok:
            raise ValueError("Excel 校验失败:\n" + "\n".join(result.errors))
        return spec

    def _execute_and_write(
        self,
        spec: ExtractionSpec,
        task_id: str,
        source_path: Path,
        output_format: str,
        output: Path,
    ) -> int:
        frame = self.engine.execute(spec, task_id, source_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        self.writer_factory.create(output_format).write(frame, output)
        return len(frame)

    def run(
        self, plan_path: Path, task_id: str, source_path: Path, output_format: str, output: Path
    ) -> tuple[int, Path]:
        spec = self._load_validated(plan_path)
        plan = spec.task(task_id)
        if str(plan.get("启用")) != "是":
            raise ValueError(f"任务 {task_id} 未启用")
        count = self._execute_and_write(spec, task_id, source_path, output_format, output)
        return count, output

    def run_all(
        self, plan_path: Path, source_path: Path, output_format: str, output_dir: Path
    ) -> list[tuple[str, int, Path]]:
        """Run every enabled task, writing `<task_id>.<format>` into `output_dir`."""
        spec = self._load_validated(plan_path)
        if output_dir.exists() and not output_dir.is_dir():
            raise ValueError(f"多任务输出路径必须是目录: {output_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
        results: list[tuple[str, int, Path]] = []
        for plan in spec.plans:
            if str(plan.get("启用")) != "是":
                continue
            task_id = str(plan.get("任务ID"))
            output = output_dir / f"{task_id}.{output_format}"
            try:
                count = self._execute_and_write(spec, task_id, source_path, output_format, output)
            except Exception as exc:
                # 多任务下给异常加上 task_id 上下文，便于在 N 个任务中定位失败的那个。
                raise RuntimeError(f"任务 {task_id} 失败: {exc}") from exc
            results.append((task_id, count, output))
        return results
