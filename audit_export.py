"""Лёгкая выгрузка отчёта и проверки в Excel (xlsxwriter)."""

from __future__ import annotations

import io
from typing import Any

import pandas as pd

from processor import (
    COL_APP,
    COL_PRESENTATION,
    COL_STATUS,
    AuditBlock,
    PersonResult,
    _cell_display,
    _find_column,
)


def _autofit_xlsxwriter_columns(
    worksheet,
    dataframe: pd.DataFrame,
    max_width: int = 60,
) -> None:
    """Подбирает ширину столбцов по заголовку и содержимому."""
    for col_idx, column_name in enumerate(dataframe.columns):
        max_length = len(str(column_name))
        for value in dataframe.iloc[:, col_idx]:
            if value is None or (isinstance(value, float) and pd.isna(value)):
                continue
            max_length = max(max_length, len(str(value)))
        worksheet.set_column(col_idx, col_idx, min(max_length + 2, max_width))


def build_report_excel(results: list[PersonResult]) -> bytes:
    """Формирует Excel-отчёт: ФИО, операция, количество."""
    rows = [
        {"ФИО": person.fio, "Операция": name, "Количество": count}
        for person in results
        if person.total > 0
        for name, count in person.operations.most_common()
    ]
    dataframe = pd.DataFrame(rows)
    buffer = io.BytesIO()
    sheet_name = "Проведённые операции"
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, index=False, sheet_name=sheet_name)
        _autofit_xlsxwriter_columns(writer.sheets[sheet_name], dataframe)
    return buffer.getvalue()


def _source_values(
    operations_df: pd.DataFrame,
    row_idx: int,
    col_app: str,
    col_status: str,
    col_present: str,
) -> tuple[str, str, str]:
    row = operations_df.iloc[row_idx]
    return (
        _cell_display(row[col_app]),
        _cell_display(row[col_status]),
        _cell_display(row[col_present]),
    )


def _write_triplet(
    worksheet,
    row_idx: int,
    *,
    fio: str,
    block_no: int,
    source_file: str,
    op_row: Any,
    op_app: str,
    op_status: str,
    op_present: str,
    conf_row: Any,
    conf_present: str,
    reason: str,
) -> int:
    """Пишет 3 строки аудита, возвращает следующий индекс строки."""
    worksheet.write_row(
        row_idx,
        0,
        [fio, block_no, source_file, op_row, op_app, op_status, op_present, ""],
    )
    worksheet.write_row(
        row_idx + 1,
        0,
        [fio, block_no, source_file, conf_row, "", "", conf_present, ""],
    )
    worksheet.write_row(
        row_idx + 2,
        0,
        [fio, block_no, source_file, "", "", "", "", reason],
    )
    return row_idx + 3


def _setup_audit_sheet(workbook, sheet_name: str):
    worksheet = workbook.add_worksheet(sheet_name)
    headers = [
        "ФИО",
        "№ блока",
        "Файл",
        "Строка",
        COL_APP,
        COL_STATUS,
        COL_PRESENTATION,
        "Причина",
    ]
    worksheet.write_row(0, 0, headers)
    worksheet.set_column(0, 0, 28)
    worksheet.set_column(1, 1, 10)
    worksheet.set_column(2, 2, 28)
    worksheet.set_column(3, 3, 10)
    worksheet.set_column(4, 6, 36)
    worksheet.set_column(7, 7, 48)
    return worksheet


def build_audit_excel(
    accepted: list[AuditBlock],
    processor_rejected: list[AuditBlock],
    filtered_by_file: dict[str, pd.DataFrame],
    operations_by_file: dict[str, pd.DataFrame],
) -> bytes:
    """Формирует Excel-аудит потоково через xlsxwriter (без merge)."""
    sample_df = next(iter(operations_by_file.values()))
    col_app = _find_column(sample_df.columns, COL_APP)
    col_status = _find_column(sample_df.columns, COL_STATUS)
    col_present = _find_column(sample_df.columns, COL_PRESENTATION)
    if not col_app or not col_status or not col_present:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            pd.DataFrame().to_excel(writer, index=False)
        return buffer.getvalue()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book

        accepted_ws = _setup_audit_sheet(workbook, "Принятые")
        row_idx = 1
        for block_no, block in enumerate(accepted, start=1):
            source_df = operations_by_file[block.source_file]
            op_app, op_status, op_present = _source_values(
                source_df, block.operation_row_idx, col_app, col_status, col_present
            )
            if block.confirmation_row_idx is not None:
                _, _, conf_present = _source_values(
                    source_df,
                    block.confirmation_row_idx,
                    col_app,
                    col_status,
                    col_present,
                )
                conf_row = block.confirmation_excel_row or 0
            else:
                conf_present = ""
                conf_row = ""
            row_idx = _write_triplet(
                accepted_ws,
                row_idx,
                fio=block.fio,
                block_no=block_no,
                source_file=block.source_file,
                op_row=block.operation_excel_row,
                op_app=op_app,
                op_status=op_status,
                op_present=op_present,
                conf_row=conf_row,
                conf_present=conf_present,
                reason=block.reason,
            )

        rejected_ws = _setup_audit_sheet(workbook, "Не принятые")
        row_idx = 1
        block_no = 0

        for file_name, filtered in filtered_by_file.items():
            if filtered.empty:
                continue
            for row in filtered.itertuples(index=False):
                block_no += 1
                presentation = row.snap_conf_present or ""
                row_idx = _write_triplet(
                    rejected_ws,
                    row_idx,
                    fio=row.fio,
                    block_no=block_no,
                    source_file=file_name,
                    op_row=int(row.line_no),
                    op_app=row.snap_app or "",
                    op_status=row.snap_status or "",
                    op_present=row.snap_present or "",
                    conf_row=(int(row.line_no) + 1 if presentation else ""),
                    conf_present=presentation,
                    reason=row.reason,
                )

        for block in processor_rejected:
            block_no += 1
            source_df = operations_by_file[block.source_file]
            op_app, op_status, op_present = _source_values(
                source_df, block.operation_row_idx, col_app, col_status, col_present
            )
            if block.confirmation_row_idx is not None:
                _, _, conf_present = _source_values(
                    source_df,
                    block.confirmation_row_idx,
                    col_app,
                    col_status,
                    col_present,
                )
                conf_row = block.confirmation_excel_row or 0
            else:
                conf_present = ""
                conf_row = ""
            row_idx = _write_triplet(
                rejected_ws,
                row_idx,
                fio=block.fio,
                block_no=block_no,
                source_file=block.source_file,
                op_row=block.operation_excel_row,
                op_app=op_app,
                op_status=op_status,
                op_present=op_present,
                conf_row=conf_row,
                conf_present=conf_present,
                reason=block.reason,
            )

    return buffer.getvalue()
