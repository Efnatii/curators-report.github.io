"""
Combine multiple JSON survey response files into a single Excel workbook.

Usage:
    python combine_json_to_excel.py /path/to/json_dir output.xlsx

Each JSON file should contain a flat mapping of question text to either a
scalar answer (string/number) or a list of strings (representing sub-rows).
When a question contains a list, the corresponding cells in Excel will be
highlighted to distinguish the block of rows taken by that subtable.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
from typing import Dict, Iterable, List, Tuple


def ensure_openpyxl_installed() -> None:
    """Install openpyxl at runtime if it is not already available."""

    if importlib.util.find_spec("openpyxl") is None:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])


ensure_openpyxl_installed()

from openpyxl import Workbook
from openpyxl.styles import PatternFill


ListValue = List[str]
ScalarValue = str | int | float | bool | None
JsonValue = ScalarValue | ListValue
JsonRecord = Dict[str, JsonValue]


HIGHLIGHT_FILL = PatternFill(start_color="FFF8DC", end_color="FFF8DC", fill_type="solid")


def normalize_cell_value(value: JsonValue) -> ScalarValue:
    """Convert any JSON value into something openpyxl can store."""

    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    # Fallback for types like dicts or lists of non-scalar values
    return json.dumps(value, ensure_ascii=False)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Merge JSON files into a single Excel workbook where each question becomes a column.",
    )
    parser.add_argument(
        "input_dir",
        type=Path,
        help="Directory containing JSON files to merge.",
    )
    parser.add_argument(
        "output",
        type=Path,
        nargs="?",
        default=Path("combined.xlsx"),
        help="Path for the generated Excel file (default: combined.xlsx).",
    )
    return parser.parse_args()


def load_json_files(input_dir: Path) -> List[Tuple[Path, JsonRecord]]:
    json_files = sorted(input_dir.glob("*.json"))
    if not json_files:
        raise FileNotFoundError(f"No JSON files found in {input_dir}")

    records: List[Tuple[Path, JsonRecord]] = []
    for path in json_files:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            raise ValueError(f"JSON file {path} does not contain an object at the top level.")
        records.append((path, data))
    return records


def determine_columns(records: Iterable[Tuple[Path, JsonRecord]]) -> List[str]:
    questions: set[str] = set()
    for _, record in records:
        questions.update(record.keys())
    return sorted(questions)


def write_workbook(records: List[Tuple[Path, JsonRecord]], questions: List[str], output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"

    headers = ["Source file", *questions]
    ws.append(headers)

    for file_path, record in records:
        block_height = max(
            (
                len(value) if isinstance(value, list) else 1
                for value in record.values()
            ),
            default=1,
        )
        for row_offset in range(block_height):
            row_values = [file_path.name if row_offset == 0 else ""]
            for question in questions:
                value = record.get(question)
                if isinstance(value, list):
                    if row_offset < len(value):
                        cell_value = normalize_cell_value(value[row_offset])
                    else:
                        cell_value = ""
                else:
                    cell_value = normalize_cell_value(value) if row_offset == 0 else ""
                row_values.append(cell_value)
            ws.append(row_values)

        # Apply highlighting to list ranges per question
        start_row = ws.max_row - block_height + 1
        for col_index, question in enumerate(questions, start=2):
            value = record.get(question)
            if isinstance(value, list) and value:
                for row_index in range(start_row, start_row + len(value)):
                    ws.cell(row=row_index, column=col_index).fill = HIGHLIGHT_FILL

    wb.save(output_path)


def main() -> None:
    args = parse_args()
    records = load_json_files(args.input_dir)
    questions = determine_columns(records)
    write_workbook(records, questions, args.output)
    print(f"Merged {len(records)} JSON files into {args.output}")


if __name__ == "__main__":
    main()
