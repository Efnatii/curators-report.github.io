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
from dataclasses import dataclass
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
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter


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


QUESTION_ORDER: List[Dict[str, object]] = [
    {
        "key": "reporting_period",
        "label": "Отчётный период",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
        },
    },
    {"key": "full_name", "label": "1. Фамилия Имя Отчество"},
    {"key": "job_positions", "label": "2. Должность"},
    {"key": "department", "label": "3. Кафедра"},
    {
        "key": "contact_phone_connected_to_telegram",
        "label": "4. Контактный телефон (подключенный к Telegram)",
    },
    {"key": "telegram_username", "label": "5. Ник в Telegram"},
    {"key": "email", "label": "6. E-mail"},
    {"key": "curated_group_numbers", "label": "7. Номера курируемых групп"},
    {"key": "curator_primary_building", "label": "8. Корпус основного пребывания куратора"},
    {"key": "curator_primary_room", "label": "9. Аудитория основного пребывания куратора"},
    {"key": "institute_or_faculty", "label": "10. Институт/Факультет"},
    {
        "key": "held_minimum_three_curator_sessions_in_reporting_period",
        "label": "11. Проведение не менее трёх кураторских часов за отчётный период",
    },
    {
        "key": "curator_hours_details",
        "label": "12. Даты проведения трёх и более кураторских часов в течение отчётного периода",
        "subfields": {
            "groups": "Группы",
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "topic": "Тема",
            "directions": "Направленность",
            "specialists": "Приглашённые специалисты",
        },
    },
    {"key": "manages_group_chat", "label": "13. Ведение чата с каждой группой или общего чата"},
    {
        "key": "inform_group_about_events",
        "label": "14. Информирование группы о мероприятиях и событиях различного уровня",
    },
    {
        "key": "achievements",
        "label": "15. Призовое место обучающегося во внеучебных мероприятиях",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
            "group": "Группа",
            "student": "ФИО студента",
            "result": "Итог",
        },
    },
    {
        "key": "participated_in_two_events_with_group",
        "label": "16. Совместное участие с группой не менее чем в двух мероприятиях",
    },
    {
        "key": "joint_participation_events",
        "label": "17. Даты проведения двух и более мероприятий в течение отчётного периода",
        "subfields": {
            "groups": "Группы",
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
        },
    },
    {
        "key": "participated_in_two_curator_events",
        "label": "18. Участие не менее чем в двух мероприятиях для кураторов",
    },
    {
        "key": "curator_personal_events",
        "label": "19. Даты участия в мероприятиях для кураторов",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
        },
    },
    {
        "key": "personal_program_participation",
        "label": "20. Личное участие в программах и конкурсах",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
        },
    },
    {
        "key": "mentor_support_events",
        "label": "21. Участие куратора в роли наставника проекта",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
            "group": "Группа",
            "student": "ФИО студента",
            "result": "Итог",
        },
    },
    {
        "key": "scientific_publications",
        "label": "22. Опубликование научной работы",
        "subfields": {
            "description": "Описание",
            "link": "Ссылка",
        },
    },
    {
        "key": "media_materials",
        "label": "23. Интервью и статьи для \"Дзен.Гуап\", соцсетей и сайта ГУАП",
        "subfields": {"link": "Ссылка"},
    },
    {
        "key": "qualification_courses",
        "label": "24. Курсы повышения квалификации",
        "subfields": {
            "date_start": "Дата начала",
            "date_end": "Дата окончания",
            "event": "Мероприятие",
        },
    },
]


@dataclass(frozen=True)
class QuestionColumn:
    key: str
    label: str
    subfields: List[Tuple[str, str]] | None = None


def determine_columns(records: Iterable[Tuple[Path, JsonRecord]]) -> List[QuestionColumn]:
    present_keys = {key for _, record in records for key in record.keys()}
    columns: List[QuestionColumn] = []

    for question in QUESTION_ORDER:
        key = question["key"]  # type: ignore[index]
        if key not in present_keys:
            continue

        label = question["label"]  # type: ignore[index]
        subfields: Dict[str, str] | None = question.get("subfields")  # type: ignore[assignment]

        if subfields:
            columns.append(QuestionColumn(key=key, label=label, subfields=list(subfields.items())))
        else:
            columns.append(QuestionColumn(key=key, label=label))

    return columns


def normalize_reporting_period(value: JsonValue) -> Dict[str, ScalarValue]:
    """Drop legacy keys and split the range into explicit dates."""

    date_start: ScalarValue = None
    date_end: ScalarValue = None

    if isinstance(value, dict):
        date_start = value.get("date_start")  # type: ignore[assignment]
        date_end = value.get("date_end")  # type: ignore[assignment]

        if date_start is None and date_end is None:
            range_value = value.get("range")
            if isinstance(range_value, str):
                start, _, end = range_value.partition(" - ")
                if start and end:
                    date_start = start
                    date_end = end

    return {"date_start": date_start, "date_end": date_end}


def normalize_record_values(record: JsonRecord) -> JsonRecord:
    """Merge scalar lists into comma-separated strings for single-cell output."""

    normalized: JsonRecord = {}
    for key, value in record.items():
        if key == "reporting_period":
            normalized[key] = normalize_reporting_period(value)
            continue

        if isinstance(value, list) and all(not isinstance(item, (dict, list)) for item in value):
            normalized[key] = ", ".join(str(item) for item in value)
        else:
            normalized[key] = value

    return normalized


def write_workbook(
    records: List[Tuple[Path, JsonRecord]],
    questions: List[QuestionColumn],
    output_path: Path,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"

    ws.cell(row=1, column=1, value="Источник")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    flat_columns: List[Tuple[str, str | None]] = []
    column_index = 2
    for question in questions:
        span = len(question.subfields) if question.subfields else 1

        if question.subfields:
            ws.merge_cells(
                start_row=1,
                start_column=column_index,
                end_row=1,
                end_column=column_index + span - 1,
            )
            ws.cell(row=1, column=column_index, value=question.label)
            for offset, (subkey, sublabel) in enumerate(question.subfields):
                ws.cell(row=2, column=column_index + offset, value=sublabel)
                flat_columns.append((question.key, subkey))
        else:
            ws.merge_cells(
                start_row=1,
                start_column=column_index,
                end_row=2,
                end_column=column_index,
            )
            ws.cell(row=1, column=column_index, value=question.label)
            flat_columns.append((question.key, None))

        column_index += span

    alignment_top = Alignment(vertical="top")
    for header_cell in ws[1] + ws[2]:
        if header_cell.value is not None:
            header_cell.alignment = alignment_top

    for file_path, raw_record in records:
        record = normalize_record_values(raw_record)
        block_height = max(
            (
                len(value) if isinstance(value, list) else 1
                for value in record.values()
            ),
            default=1,
        )
        for row_offset in range(block_height):
            row_values = [file_path.name if row_offset == 0 else ""]
            for question_key, subkey in flat_columns:
                value = record.get(question_key)
                if isinstance(value, list):
                    if row_offset < len(value):
                        item = value[row_offset]
                        if subkey is not None and isinstance(item, dict):
                            cell_value = normalize_cell_value(item.get(subkey))
                        elif subkey is None:
                            cell_value = normalize_cell_value(item)
                        else:
                            cell_value = ""
                    else:
                        cell_value = ""
                elif isinstance(value, dict):
                    if subkey is not None:
                        cell_value = normalize_cell_value(value.get(subkey))
                    elif row_offset == 0:
                        cell_value = normalize_cell_value(value)
                    else:
                        cell_value = ""
                else:
                    if subkey is not None:
                        cell_value = ""
                    else:
                        cell_value = normalize_cell_value(value) if row_offset == 0 else ""
                row_values.append(cell_value)
            ws.append(row_values)

        # Apply highlighting to list ranges per question
        start_row = ws.max_row - block_height + 1
        column_start = 2
        for question in questions:
            span = len(question.subfields) if question.subfields else 1
            value = record.get(question.key)
            if isinstance(value, list) and value:
                for offset in range(span):
                    for row_index in range(start_row, start_row + len(value)):
                        ws.cell(row=row_index, column=column_start + offset).fill = HIGHLIGHT_FILL
            column_start += span

    for col_idx, column_cells in enumerate(ws.columns, start=1):
        max_length = 0
        for cell in column_cells:
            if cell.value is None:
                continue
            for line in str(cell.value).split("\n"):
                max_length = max(max_length, len(line))

        if max_length:
            column_letter = get_column_letter(col_idx)
            ws.column_dimensions[column_letter].width = max_length + 2

    wb.save(output_path)


def main() -> None:
    args = parse_args()
    records = load_json_files(args.input_dir)
    questions = determine_columns(records)
    write_workbook(records, questions, args.output)
    print(f"Merged {len(records)} JSON files into {args.output}")


if __name__ == "__main__":
    main()
