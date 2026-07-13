from abc import ABC, abstractmethod
from pathlib import Path

import pandas as pd


class OutputWriter(ABC):
    @abstractmethod
    def write(self, frame: pd.DataFrame, path: Path) -> None: ...


class CsvWriter(OutputWriter):
    def write(self, frame: pd.DataFrame, path: Path) -> None:
        frame.to_csv(path, index=False, encoding="utf-8-sig")


class JsonLinesWriter(OutputWriter):
    def write(self, frame: pd.DataFrame, path: Path) -> None:
        frame.to_json(path, orient="records", lines=True, force_ascii=False, date_format="iso")


class ExcelWriter(OutputWriter):
    def write(self, frame: pd.DataFrame, path: Path) -> None:
        frame.to_excel(path, index=False, sheet_name="data")


class OutputWriterFactory:
    writers = {"csv": CsvWriter, "jsonl": JsonLinesWriter, "xlsx": ExcelWriter}

    def create(self, output_format: str) -> OutputWriter:
        try:
            return self.writers[output_format.lower()]()
        except KeyError as exc:
            raise ValueError(f"不支持的输出格式: {output_format}") from exc
