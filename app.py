"""Streamlit-приложение: справочник ФИО и анализ проведённых операций."""

import io

import pandas as pd
import streamlit as st

from journal_parser import parse_1c_journal_bytes
from processor import (
    COL_APP,
    COL_PRESENTATION,
    COL_STATUS,
    AuditBlock,
    AuditData,
    PersonResult,
    _cell_display,
    _find_column,
    load_fio_list,
    merge_person_results,
    process_operations,
)


def _attach_source_file(blocks: list[AuditBlock], source_file: str) -> list[AuditBlock]:
    """Проставляет имя файла в блоки аудита журнала."""
    for block in blocks:
        block.source_file = source_file
    return blocks

st.set_page_config(
    page_title="Анализ проведённых операций ККУ",
    page_icon="📋",
    layout="wide",
)

st.markdown(
    """
    <style>
    .kku-section-title {
        color: #ffffff;
        font-size: 2.25rem;
        font-weight: 700;
        line-height: 1.2;
        margin: 0.5rem 0 1rem 0;
    }
    .kku-fio-title {
        color: #b388ff;
        font-size: 1.125rem;
        font-weight: 600;
        line-height: 1.3;
        margin: 1rem 0 0.5rem 0;
    }
    div[data-testid="stExpander"] details {
        border: 1px solid #6b6b6b;
        border-radius: 8px;
        background-color: #5a5a5a;
    }
    div[data-testid="stExpander"] details summary {
        background-color: #5a5a5a;
        color: #e0e0e0;
        border-radius: 8px;
        font-weight: 500;
    }
    div[data-testid="stExpander"] details summary:hover {
        background-color: #666666;
        color: #ffffff;
    }
    div[data-testid="stExpander"] details[open] summary {
        border-bottom-left-radius: 0;
        border-bottom-right-radius: 0;
    }
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        background-color: #4f4f4f;
        border-bottom-left-radius: 8px;
        border-bottom-right-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📋 Анализ проведённых операций ККУ")
st.caption(
    "Загрузите справочник с ФИО (xlsx) и один или два файла операций "
    "(журнал 1С, текстовый файл UTF-8)."
)

col_left, col_right = st.columns(2)

with col_left:
    directory_file = st.file_uploader(
        "Справочник",
        type=["xlsx"],
        help="Файл со списком ФИО (первый столбец)",
        key="directory",
    )

with col_right:
    op_col1, op_col2 = st.columns(2)
    with op_col1:
        operations_file_1 = st.file_uploader(
            "Операции 1",
            type=["txt"],
            help="Журнал регистрации 1С: «Текстовый файл UTF-8 (*.txt)»",
            key="operations_1",
        )
    with op_col2:
        operations_file_2 = st.file_uploader(
            "Операции 2",
            type=["txt"],
            help="Второй журнал 1С в UTF-8 (необязательно)",
            key="operations_2",
        )

operations_files = [
    uploaded
    for uploaded in (operations_file_1, operations_file_2)
    if uploaded is not None
]

if directory_file is None or not operations_files:
    st.stop()


def read_directory(uploaded) -> pd.DataFrame:
    """Читает справочник ФИО из xlsx."""
    return pd.read_excel(io.BytesIO(uploaded.getvalue()), sheet_name=0)


def read_operations_journal(uploaded) -> tuple[pd.DataFrame, list[AuditBlock]]:
    """Читает журнал 1С (UTF-8 txt): таблица операций + отфильтрованные блоки."""
    parsed = parse_1c_journal_bytes(uploaded.getvalue())
    filtered = _attach_source_file(parsed.filtered, uploaded.name)
    return parsed.operations, filtered


def process_all_operations(
    directory_df: pd.DataFrame,
    uploaded_files: list,
) -> tuple[list[PersonResult], list[str], AuditData, dict[str, pd.DataFrame]]:
    """Обрабатывает один или два файла операций и объединяет результат."""
    fio_list = load_fio_list(directory_df)
    merged_results = {fio: PersonResult(fio=fio) for fio in fio_list}
    audit = AuditData()
    warnings: list[str] = []
    operations_by_file: dict[str, pd.DataFrame] = {}

    for uploaded in uploaded_files:
        operations_df, filtered_blocks = read_operations_journal(uploaded)
        operations_by_file[uploaded.name] = operations_df
        file_results, file_warnings, file_audit = process_operations(
            directory_df,
            operations_df,
            source_file=uploaded.name,
        )
        merge_person_results(merged_results, file_results)
        audit.accepted.extend(file_audit.accepted)
        # Сначала отфильтрованное парсером, затем отклонения processor.
        audit.rejected.extend(filtered_blocks)
        audit.rejected.extend(file_audit.rejected)
        warnings.extend(file_warnings)

    for index, block in enumerate(audit.rejected, start=1):
        block.block_no = index

    results = [merged_results[fio] for fio in fio_list if fio in merged_results]
    return results, warnings, audit, operations_by_file


def _autofit_worksheet_columns(worksheet, max_width: int = 60) -> None:
    """Подбирает ширину столбцов по содержимому."""
    for column_cells in worksheet.columns:
        column_letter = column_cells[0].column_letter
        max_length = 0
        for cell in column_cells:
            if cell.value is None:
                continue
            max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[column_letter].width = min(max_length + 2, max_width)


def build_report_excel(results: list[PersonResult]) -> bytes:
    """Формирует Excel-отчёт: ФИО, операция, количество."""
    rows = [
        {"ФИО": person.fio, "Операция": name, "Количество": count}
        for person in results
        if person.total > 0
        for name, count in person.operations.most_common()
    ]
    buffer = io.BytesIO()
    sheet_name = "Проведённые операции"
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, sheet_name=sheet_name)
        _autofit_worksheet_columns(writer.sheets[sheet_name])
    return buffer.getvalue()


def _source_row_dict(
    operations_df: pd.DataFrame,
    row_idx: int,
    col_app: str,
    col_status: str,
    col_present: str,
    excel_row: int | None = None,
) -> dict[str, str | int]:
    row = operations_df.iloc[row_idx]
    return {
        "Файл": "",
        "Строка Excel": excel_row if excel_row is not None else row_idx + 2,
        COL_APP: _cell_display(row[col_app]),
        COL_STATUS: _cell_display(row[col_status]),
        COL_PRESENTATION: _cell_display(row[col_present]),
    }


def _snapshot_row_dict(
    block: AuditBlock,
    *,
    is_confirmation: bool,
) -> dict[str, str | int]:
    """Строка аудита из снимка отфильтрованного блока журнала."""
    if is_confirmation:
        return {
            "Файл": block.source_file,
            "Строка Excel": block.confirmation_excel_row or "",
            COL_APP: "",
            COL_STATUS: "",
            COL_PRESENTATION: block.snap_conf_present,
        }
    return {
        "Файл": block.source_file,
        "Строка Excel": block.operation_excel_row,
        COL_APP: block.snap_app,
        COL_STATUS: block.snap_status,
        COL_PRESENTATION: block.snap_present,
    }


def _blocks_to_audit_rows(
    blocks: list[AuditBlock],
    operations_by_file: dict[str, pd.DataFrame],
    col_app: str,
    col_status: str,
    col_present: str,
) -> list[dict[str, str | int]]:
    """Три строки на блок: операция, подтверждение, причина."""
    rows: list[dict[str, str | int]] = []
    for block in blocks:
        if block.use_snapshot:
            operation = _snapshot_row_dict(block, is_confirmation=False)
            has_confirmation = bool(block.snap_conf_present)
            confirmation = (
                _snapshot_row_dict(block, is_confirmation=True)
                if has_confirmation
                else {
                    "Файл": block.source_file,
                    "Строка Excel": "",
                    COL_APP: "",
                    COL_STATUS: "",
                    COL_PRESENTATION: "",
                }
            )
        else:
            source_df = operations_by_file[block.source_file]
            operation = _source_row_dict(
                source_df,
                block.operation_row_idx,
                col_app,
                col_status,
                col_present,
                excel_row=block.operation_excel_row,
            )
            operation["Файл"] = block.source_file
            if block.confirmation_row_idx is not None:
                confirmation = _source_row_dict(
                    source_df,
                    block.confirmation_row_idx,
                    col_app,
                    col_status,
                    col_present,
                    excel_row=block.confirmation_excel_row or 0,
                )
                confirmation["Файл"] = block.source_file
            else:
                confirmation = {
                    "Файл": block.source_file,
                    "Строка Excel": "",
                    COL_APP: "",
                    COL_STATUS: "",
                    COL_PRESENTATION: "",
                }

        rows.append(
            {
                "ФИО": block.fio,
                "№ блока": block.block_no,
                **operation,
                "Причина": "",
            }
        )
        rows.append(
            {
                "ФИО": block.fio,
                "№ блока": block.block_no,
                **confirmation,
                "Причина": "",
            }
        )
        rows.append(
            {
                "ФИО": block.fio,
                "№ блока": block.block_no,
                "Файл": block.source_file,
                "Строка Excel": "",
                COL_APP: "",
                COL_STATUS: "",
                COL_PRESENTATION: "",
                "Причина": block.reason,
            }
        )

    return rows


def _merge_fio_and_block_columns(worksheet, block_count: int) -> None:
    """Объединяет ФИО и № блока для каждой тройки строк."""
    row_start = 2
    for _ in range(block_count):
        row_end = row_start + 2
        worksheet.merge_cells(
            start_row=row_start,
            start_column=1,
            end_row=row_end,
            end_column=1,
        )
        worksheet.merge_cells(
            start_row=row_start,
            start_column=2,
            end_row=row_end,
            end_column=2,
        )
        row_start = row_end + 1


def build_audit_excel(
    audit: AuditData,
    operations_by_file: dict[str, pd.DataFrame],
) -> bytes:
    """Формирует Excel-аудит: 3 строки на блок, листы «Принятые» и «Не принятые»."""
    sample_df = next(iter(operations_by_file.values()))
    col_app = _find_column(sample_df.columns, COL_APP)
    col_status = _find_column(sample_df.columns, COL_STATUS)
    col_present = _find_column(sample_df.columns, COL_PRESENTATION)
    if not col_app or not col_status or not col_present:
        buffer = io.BytesIO()
        pd.DataFrame().to_excel(buffer, index=False)
        return buffer.getvalue()

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, blocks in (
            ("Принятые", audit.accepted),
            ("Не принятые", audit.rejected),
        ):
            sheet_rows = _blocks_to_audit_rows(
                blocks, operations_by_file, col_app, col_status, col_present
            )
            pd.DataFrame(sheet_rows).to_excel(
                writer,
                index=False,
                sheet_name=sheet_name,
            )
            if blocks:
                _merge_fio_and_block_columns(writer.sheets[sheet_name], len(blocks))
            _autofit_worksheet_columns(writer.sheets[sheet_name])

    return buffer.getvalue()


try:
    with st.spinner("Файлы обрабатываются"):
        directory_df = read_directory(directory_file)
        results, warnings, audit, operations_by_file = process_all_operations(
            directory_df,
            operations_files,
        )
except Exception as exc:
    st.error(
        "Не удалось прочитать файл. Справочник — xlsx, операции — журнал 1С в UTF-8 (txt). "
        f"Подробности: {exc}"
    )
    st.stop()

filtered_count = sum(1 for block in audit.rejected if block.use_snapshot)
for uploaded in operations_files:
    rows = len(operations_by_file.get(uploaded.name, []))
    blocks = rows // 2
    st.info(
        f"Журнал 1С «{uploaded.name}»: из UTF-8 txt найдено **{blocks}** "
        f"проведённых операций (события «Данные. Проведение»)."
    )
if filtered_count:
    st.info(
        f"В лист «Не принятые» попало **{filtered_count}** отфильтрованных "
        f"записей журнала (Добавление, отмены, без даты и др.)."
    )

with st.expander("Просмотр загруженных данных", expanded=False):
    tab1, *operation_tabs = st.tabs(
        ["Справочник"] + [f"Операции: {name}" for name in operations_by_file]
    )
    with tab1:
        st.dataframe(directory_df, use_container_width=True, hide_index=True)
    for tab, (name, operations_df) in zip(operation_tabs, operations_by_file.items()):
        with tab:
            st.dataframe(operations_df, use_container_width=True, hide_index=True)

for warning in warnings:
    st.warning(warning)

if warnings and not results:
    st.stop()

st.divider()
header_col, download_col, audit_col = st.columns([2.5, 1, 1])
with header_col:
    st.markdown(
        '<div class="kku-section-title">Проведённые операции</div>',
        unsafe_allow_html=True,
    )

if not results:
    st.info("Нет данных для отображения.")
    st.stop()

has_any_operations = any(person.total > 0 for person in results)

if not has_any_operations:
    st.info("Проведённые операции не найдены ни для одного ФИО из справочника.")
    st.stop()

with download_col:
    st.download_button(
        label="Скачать отчёт Excel",
        data=build_report_excel(results),
        file_name="отчёт_кку.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with audit_col:
    st.download_button(
        label="Скачать проверку",
        data=build_audit_excel(audit, operations_by_file),
        file_name="проверка_кку.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

for person in results:
    if person.total == 0:
        continue

    st.markdown(
        f'<div class="kku-fio-title">{person.fio}</div>',
        unsafe_allow_html=True,
    )
    st.caption(f"Проведённых операций: {person.total}")

    details_df = pd.DataFrame(
        [
            {"Операция": name, "Количество": count}
            for name, count in person.operations.most_common()
        ]
    )
    st.dataframe(details_df, use_container_width=True, hide_index=True)
    st.markdown("")

no_operations = [person.fio for person in results if person.total == 0]
if no_operations:
    with st.expander(f"ФИО без проведённых операций ({len(no_operations)})"):
        for fio in no_operations:
            st.write(fio)
