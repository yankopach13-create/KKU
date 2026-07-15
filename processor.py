"""Обработка справочника ФИО и файла операций."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any

import pandas as pd

COL_APP = "Приложение"
COL_STATUS = "Зафиксирована/отмена"
COL_PRESENTATION = "Представление данных"

HEADER_ALIASES = {"фио", "ф.и.о.", "ф.и.о", "имя", "сотрудник"}

DOCUMENT_DATE_PATTERN = re.compile(r"\sот\s+\d{1,2}\.\d{1,2}\.\d{2,4}", re.IGNORECASE)


@dataclass
class AuditBlock:
    """Блок аудита: операция + подтверждение + причина."""

    fio: str
    block_no: int
    reason: str
    operation_row_idx: int = -1
    confirmation_row_idx: int | None = None
    source_file: str = ""
    operation_excel_row: int = 0
    confirmation_excel_row: int | None = None
    # Снимок строк для отфильтрованных блоков журнала (не из таблицы операций).
    use_snapshot: bool = False
    snap_app: str = ""
    snap_status: str = ""
    snap_present: str = ""
    snap_conf_present: str = ""


@dataclass
class AuditData:
    """Аудит: принятые и отклонённые блоки."""

    accepted: list[AuditBlock] = field(default_factory=list)
    rejected: list[AuditBlock] = field(default_factory=list)


@dataclass
class PersonResult:
    """Результат по одному сотруднику из справочника."""

    fio: str
    operations: Counter[str] = field(default_factory=Counter)
    fio_rows: int = 0
    header_rows: int = 0

    @property
    def total(self) -> int:
        return sum(self.operations.values())


def _normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().split())


def _cell_display(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _is_filled(value: Any) -> bool:
    return bool(_normalize_text(value))


def _excel_row_number(df_row_index: int) -> int:
    return df_row_index + 2


def _find_column(columns: pd.Index, target: str) -> str | None:
    target_lower = target.lower()
    for col in columns:
        if _normalize_text(col).lower() == target_lower:
            return col
    for col in columns:
        if target_lower in _normalize_text(col).lower():
            return col
    return None


def load_fio_list(df: pd.DataFrame) -> list[str]:
    if df.empty or len(df.columns) == 0:
        return []

    first_col = df.columns[0]
    values: list[str] = []
    seen: set[str] = set()

    for raw in df[first_col]:
        fio = _normalize_text(raw)
        if not fio:
            continue
        if fio.lower() in HEADER_ALIASES:
            continue
        key = fio.lower()
        if key not in seen:
            seen.add(key)
            values.append(fio)

    return values


def _split_fio_parts(text: str) -> list[str]:
    return _normalize_text(text).lower().split()


def _name_parts_match(name_a: str, name_b: str) -> bool:
    if name_a == name_b:
        return True
    if not name_a or not name_b:
        return False
    a = name_a.rstrip(".")
    b = name_b.rstrip(".")
    if a == b:
        return True
    if a[0] != b[0]:
        return False
    if "." in name_a or "." in name_b or len(a) <= 2 or len(b) <= 2:
        return True
    return a.startswith(b) or b.startswith(a)


def _two_word_fio_match(parts_a: list[str], parts_b: list[str]) -> bool:
    a0, a1 = parts_a[0], parts_a[1]
    b0, b1 = parts_b[0], parts_b[1]

    def words_match(left: str, right: str) -> bool:
        return left == right or _name_parts_match(left, right)

    direct = words_match(a0, b0) and words_match(a1, b1)
    swapped = words_match(a0, b1) and words_match(a1, b0)
    return direct or swapped


def _fio_matches(cell_value: Any, fio: str) -> bool:
    cell = _normalize_text(cell_value)
    fio = _normalize_text(fio)
    if not cell or not fio:
        return False

    cell_lower = cell.lower()
    fio_lower = fio.lower()
    if fio_lower in cell_lower or cell_lower in fio_lower:
        return True

    cell_parts = _split_fio_parts(cell)
    fio_parts = _split_fio_parts(fio)
    if len(cell_parts) < 2 or len(fio_parts) < 2:
        return False

    return _two_word_fio_match(cell_parts[:2], fio_parts[:2])


def _is_fixed_status(value: Any) -> bool:
    text = _normalize_text(value).lower()
    if not text:
        return False
    if "отмен" in text:
        return False
    return text == "зафиксирована"


def _is_cancelled_status(value: Any) -> bool:
    text = _normalize_text(value).lower()
    return bool(text) and "отмен" in text


def _is_header_row(
    row: pd.Series,
    col_app: str,
    col_status: str,
    fio: str,
) -> bool:
    return _fio_matches(row[col_app], fio) and _is_fixed_status(row[col_status])


def _is_operation_header_row(
    row: pd.Series,
    col_app: str,
    col_status: str,
    col_present: str,
    fio: str,
) -> bool:
    return (
        _fio_matches(row[col_app], fio)
        and _is_fixed_status(row[col_status])
        and _is_filled(row[col_present])
    )


def _get_confirmation_rejection_reason(
    row: pd.Series,
    col_app: str,
    col_status: str,
    col_present: str,
    fio: str,
    operation_name: str,
) -> str:
    if _is_header_row(row, col_app, col_status, fio):
        return "Следующая строка — новый блок (ФИО + «Зафиксирована»)"

    confirmation = _normalize_text(row[col_present])
    if not confirmation:
        return "Подтверждение пустое"

    if confirmation.lower() == _normalize_text(operation_name).lower():
        return "Текст совпадает с названием операции"

    if not DOCUMENT_DATE_PATTERN.search(confirmation):
        return "Нет даты документа (от ДД.ММ.ГГГГ)"

    return ""


def _append_audit_block(
    target: list[AuditBlock],
    fio: str,
    block_no: int,
    reason: str,
    operation_idx: int,
    confirmation_idx: int | None,
    source_file: str = "",
) -> None:
    target.append(
        AuditBlock(
            fio=fio,
            block_no=block_no,
            reason=reason,
            operation_row_idx=operation_idx,
            confirmation_row_idx=confirmation_idx,
            source_file=source_file,
            operation_excel_row=_excel_row_number(operation_idx),
            confirmation_excel_row=(
                _excel_row_number(confirmation_idx)
                if confirmation_idx is not None
                else None
            ),
        )
    )


def merge_person_results(
    merged: dict[str, PersonResult],
    incoming: list[PersonResult],
) -> None:
    """Суммирует результаты нескольких файлов операций."""
    for person in incoming:
        if person.fio not in merged:
            merged[person.fio] = PersonResult(fio=person.fio)
        target = merged[person.fio]
        target.operations += person.operations
        target.fio_rows += person.fio_rows
        target.header_rows += person.header_rows


def process_operations(
    directory_df: pd.DataFrame,
    operations_df: pd.DataFrame,
    source_file: str = "",
) -> tuple[list[PersonResult], list[str], AuditData]:
    warnings: list[str] = []
    audit = AuditData()
    fio_list = load_fio_list(directory_df)

    if not fio_list:
        warnings.append("Справочник пуст или не содержит ФИО.")
        return [], warnings, audit

    col_app = _find_column(operations_df.columns, COL_APP)
    col_status = _find_column(operations_df.columns, COL_STATUS)
    col_present = _find_column(operations_df.columns, COL_PRESENTATION)

    missing = [
        name
        for name, col in (
            (COL_APP, col_app),
            (COL_STATUS, col_status),
            (COL_PRESENTATION, col_present),
        )
        if col is None
    ]
    if missing:
        warnings.append(
            "В журнале операций не найдены столбцы: " + ", ".join(missing) + "."
        )
        return [], warnings, audit

    persons = {fio: PersonResult(fio=fio) for fio in fio_list}
    block_nos = {fio: 0 for fio in fio_list}
    total_rows = len(operations_df)

    apps = operations_df[col_app].tolist()
    statuses = operations_df[col_status].tolist()
    presents = operations_df[col_present].tolist()

    for row_idx in range(total_rows):
        app_value = apps[row_idx]
        if not _is_filled(app_value):
            continue

        matched_fios = [fio for fio in fio_list if _fio_matches(app_value, fio)]
        if not matched_fios:
            continue

        status_value = statuses[row_idx]
        present_value = presents[row_idx]

        for fio in matched_fios:
            person = persons[fio]
            person.fio_rows += 1

            if _is_cancelled_status(status_value):
                block_nos[fio] += 1
                _append_audit_block(
                    audit.rejected,
                    fio,
                    block_nos[fio],
                    "Не принято: операция отменена",
                    row_idx,
                    None,
                    source_file,
                )
                continue

            if not _is_fixed_status(status_value):
                block_nos[fio] += 1
                _append_audit_block(
                    audit.rejected,
                    fio,
                    block_nos[fio],
                    "Не принято: ФИО без «Зафиксирована»",
                    row_idx,
                    None,
                    source_file,
                )
                continue

            person.header_rows += 1

            if not _is_filled(present_value):
                block_nos[fio] += 1
                _append_audit_block(
                    audit.rejected,
                    fio,
                    block_nos[fio],
                    "Не принято: «Зафиксирована» без названия операции",
                    row_idx,
                    None,
                    source_file,
                )
                continue

            block_nos[fio] += 1
            operation_name = _normalize_text(present_value)

            if row_idx + 1 >= total_rows:
                _append_audit_block(
                    audit.rejected,
                    fio,
                    block_nos[fio],
                    "Не принято: нет строки подтверждения ниже",
                    row_idx,
                    None,
                    source_file,
                )
                continue

            # Для подтверждения достаточно значений следующего ряда.
            next_app = apps[row_idx + 1]
            next_status = statuses[row_idx + 1]
            next_present = presents[row_idx + 1]
            next_row = pd.Series(
                {
                    col_app: next_app,
                    col_status: next_status,
                    col_present: next_present,
                }
            )
            rejection_reason = _get_confirmation_rejection_reason(
                next_row, col_app, col_status, col_present, fio, operation_name
            )

            if rejection_reason:
                _append_audit_block(
                    audit.rejected,
                    fio,
                    block_nos[fio],
                    f"Не принято: {rejection_reason}",
                    row_idx,
                    row_idx + 1,
                    source_file,
                )
                continue

            person.operations[operation_name] += 1
            _append_audit_block(
                audit.accepted,
                fio,
                block_nos[fio],
                "Принято",
                row_idx,
                row_idx + 1,
                source_file,
            )

    results = [persons[fio] for fio in fio_list]
    return results, warnings, audit
