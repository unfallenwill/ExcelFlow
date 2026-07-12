from pathlib import Path

from .service import ExtractionService
from .template import create_template


def validate(path: Path): return ExtractionService().validate(path)
def run_task(path: Path, task_id: str, source_path: Path, output_format: str, output_path: Path):
    return ExtractionService().run(path, task_id, source_path, output_format, output_path)

__all__ = ["ExtractionService", "create_template", "run_task", "validate"]
