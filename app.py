"""Streamlit-приложение: справочник ФИО и анализ проведённых операций."""

from __future__ import annotations

import gc
import io
from typing import Any

import pandas as pd
import streamlit as st

from audit_export import build_audit_excel, build_report_excel
from journal_parser import parse_1c_journal_bytes
from processor import (
    AuditBlock,
    PersonResult,
    load_fio_list,
    merge_person_results,
    process_operations,
)

PREVIEW_ROWS = 200

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


@st.cache_data(show_spinner="Файлы обрабатываются…", max_entries=3)
def cached_process_all(
    directory_bytes: bytes,
    operations_payload: tuple[tuple[str, bytes], ...],
) -> tuple[
    list[dict[str, Any]],
    list[str],
    list[AuditBlock],
    list[AuditBlock],
    dict[str, pd.DataFrame],
    dict[str, pd.DataFrame],
]:
    """Кеширует тяжёлую обработку по содержимому файлов."""
    directory_df = pd.read_excel(io.BytesIO(directory_bytes), sheet_name=0)
    fio_list = load_fio_list(directory_df)
    merged_results = {fio: PersonResult(fio=fio) for fio in fio_list}
    warnings: list[str] = []
    operations_by_file: dict[str, pd.DataFrame] = {}
    filtered_by_file: dict[str, pd.DataFrame] = {}
    accepted: list[AuditBlock] = []
    processor_rejected: list[AuditBlock] = []

    for file_name, raw in operations_payload:
        parsed = parse_1c_journal_bytes(raw)
        operations_by_file[file_name] = parsed.operations
        filtered = parsed.filtered.copy()
        if not filtered.empty:
            filtered["source_file"] = file_name
        filtered_by_file[file_name] = filtered

        file_results, file_warnings, file_audit = process_operations(
            directory_df,
            parsed.operations,
            source_file=file_name,
        )
        merge_person_results(merged_results, file_results)
        accepted.extend(file_audit.accepted)
        processor_rejected.extend(file_audit.rejected)
        warnings.extend(file_warnings)
        del parsed

    results = [merged_results[fio] for fio in fio_list if fio in merged_results]
    results_payload = [
        {
            "fio": person.fio,
            "operations": dict(person.operations),
            "total": person.total,
        }
        for person in results
    ]
    gc.collect()
    return (
        results_payload,
        warnings,
        accepted,
        processor_rejected,
        operations_by_file,
        filtered_by_file,
    )


def _results_from_payload(payload: list[dict[str, Any]]) -> list[PersonResult]:
    results: list[PersonResult] = []
    for item in payload:
        person = PersonResult(fio=item["fio"])
        person.operations.update(item["operations"])
        results.append(person)
    return results


directory_bytes = directory_file.getvalue()
operations_payload = tuple(
    (uploaded.name, uploaded.getvalue()) for uploaded in operations_files
)

try:
    (
        results_payload,
        warnings,
        accepted_blocks,
        processor_rejected,
        operations_by_file,
        filtered_by_file,
    ) = cached_process_all(directory_bytes, operations_payload)
    results = _results_from_payload(results_payload)
except Exception as exc:
    st.error(
        "Не удалось прочитать файл. Справочник — xlsx, операции — журнал 1С в UTF-8 (txt). "
        f"Подробности: {exc}"
    )
    st.stop()

accepted_count = len(accepted_blocks)
rejected_count = sum(len(df) for df in filtered_by_file.values()) + len(
    processor_rejected
)
total_count = accepted_count + rejected_count
st.info(
    f"Всего операций: {total_count}  \n"
    f"Отфильтровано в не принятые: {rejected_count}  \n"
    f"Принятые операции: {accepted_count}"
)

with st.expander("Просмотр загруженных данных", expanded=False):
    tab1, *operation_tabs = st.tabs(
        ["Справочник"] + [f"Операции: {name}" for name in operations_by_file]
    )
    with tab1:
        directory_preview = pd.read_excel(io.BytesIO(directory_bytes), sheet_name=0)
        st.dataframe(directory_preview, use_container_width=True, hide_index=True)
    for tab, (name, operations_df) in zip(operation_tabs, operations_by_file.items()):
        with tab:
            total_rows = len(operations_df)
            if total_rows > PREVIEW_ROWS:
                st.caption(
                    f"Показаны первые {PREVIEW_ROWS} из {total_rows} строк "
                    "(полный файл не выводится, чтобы не нагружать память)."
                )
                st.dataframe(
                    operations_df.head(PREVIEW_ROWS),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
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

payload_key = (
    hash(directory_bytes),
    tuple((name, len(raw), hash(raw[:4096])) for name, raw in operations_payload),
)
if st.session_state.get("audit_payload_key") != payload_key:
    st.session_state.pop("audit_excel_bytes", None)
    st.session_state["audit_payload_key"] = payload_key

with download_col:
    st.download_button(
        label="Скачать отчёт Excel",
        data=build_report_excel(results),
        file_name="отчёт_кку.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with audit_col:
    if st.button("Подготовить проверку", use_container_width=True):
        with st.spinner("Формирование файла проверки…"):
            st.session_state["audit_excel_bytes"] = build_audit_excel(
                accepted_blocks,
                processor_rejected,
                filtered_by_file,
                operations_by_file,
            )
            gc.collect()

    audit_bytes = st.session_state.get("audit_excel_bytes")
    if audit_bytes:
        st.download_button(
            label="Скачать проверку",
            data=audit_bytes,
            file_name="проверка_кку.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    else:
        st.caption("Сначала нажмите «Подготовить проверку».")

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
