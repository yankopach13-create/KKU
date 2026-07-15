"""Парсер журнала регистрации 1С (выгрузка «Текстовый файл UTF-8»)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterator

import pandas as pd

from processor import (
    COL_APP,
    COL_PRESENTATION,
    COL_STATUS,
    DOCUMENT_DATE_PATTERN,
)

BLOCK_START = re.compile(r"^\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2}:\d{2}\t")
EVENT_POSTING = "Данные. Проведение"
STATUS_FIXED = "зафиксирована"
TRANSACTION_ID = re.compile(r"\((\d+)\)\s*$")

FILTERED_COLUMNS = [
    "fio",
    "reason",
    "snap_app",
    "snap_status",
    "snap_present",
    "snap_conf_present",
    "line_no",
]


@dataclass
class JournalParseResult:
    """Таблица для подсчёта + отфильтрованные блоки для аудита."""

    operations: pd.DataFrame
    filtered: pd.DataFrame


def _last_filled(parts: list[str]) -> str:
    for part in reversed(parts):
        text = part.strip()
        if text:
            return text
    return ""


def _field(parts: list[str], index: int) -> str:
    if index >= len(parts):
        return ""
    return parts[index].strip()


def _is_posting_event(event: str) -> bool:
    return _normalize(event).lower() == EVENT_POSTING.lower()


def _is_fixed_status(status: str) -> bool:
    text = _normalize(status).lower()
    if not text:
        return False
    if "отмен" in text:
        return False
    return text == STATUS_FIXED


def _is_cancelled_status(status: str) -> bool:
    text = _normalize(status).lower()
    return bool(text) and "отмен" in text


def _normalize(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split())


def _extract_transaction_id(block_lines: list[str]) -> str | None:
    if len(block_lines) < 2:
        return None
    parts = block_lines[1].split("\t")
    for part in parts:
        match = TRANSACTION_ID.search(part.strip())
        if match:
            return match.group(1)
    return None


def _extract_presentation(block_lines: list[str]) -> str:
    """Подтверждение — текст с датой документа «от ДД.ММ.ГГГГ»."""
    for line in block_lines[1:]:
        candidate = _last_filled(line.split("\t"))
        if candidate and DOCUMENT_DATE_PATTERN.search(candidate):
            return candidate
    return ""


def _inspect_block(
    block_lines: list[str],
) -> tuple[str, str, str, str, str | None]:
    """
    Возвращает (user, status, metadata, presentation, reject_reason).

    reject_reason=None — блок можно принять.
    """
    if not block_lines:
        return "", "", "", "", "Пустой блок журнала"

    header_parts = block_lines[0].split("\t")
    user = _field(header_parts, 1)
    event = _field(header_parts, 3)
    status = _field(header_parts, 4)
    metadata = _field(header_parts, 5)
    presentation = _extract_presentation(block_lines)

    if not user and not metadata and not event:
        return user, status, metadata, presentation, "Служебная/пустая запись журнала"

    if not user:
        return user, status, metadata, presentation, "Не принято: нет пользователя (ФИО)"

    if not metadata:
        return (
            user,
            status,
            metadata,
            presentation,
            "Не принято: нет названия операции (метаданные)",
        )

    if not _is_posting_event(event):
        event_label = event or "(пусто)"
        return (
            user,
            status,
            metadata,
            presentation,
            f"Не принято: событие не «Данные. Проведение» (было: {event_label})",
        )

    if _is_cancelled_status(status):
        return (
            user,
            status,
            metadata,
            presentation,
            "Не принято: операция отменена",
        )

    if not _is_fixed_status(status):
        status_label = status or "(пусто)"
        return (
            user,
            status,
            metadata,
            presentation,
            f"Не принято: статус не «Зафиксирована» (было: {status_label})",
        )

    if not presentation:
        return (
            user,
            status,
            metadata,
            presentation,
            "Не принято: нет подтверждения с датой документа (от ДД.ММ.ГГГГ)",
        )

    if not DOCUMENT_DATE_PATTERN.search(presentation):
        return (
            user,
            status,
            metadata,
            presentation,
            "Не принято: нет даты документа (от ДД.ММ.ГГГГ)",
        )

    if _normalize(presentation).lower() == _normalize(metadata).lower():
        return (
            user,
            status,
            metadata,
            presentation,
            "Не принято: текст подтверждения совпадает с названием операции",
        )

    return user, status, metadata, presentation, None


def _iter_blocks(text: str) -> Iterator[tuple[int, list[str]]]:
    """Потоково отдаёт блоки журнала без хранения всего файла как списка блоков."""
    current: list[str] = []
    current_line_no = 0

    for line_no, raw_line in enumerate(text.splitlines(), start=1):
        line = raw_line.rstrip("\r")
        if BLOCK_START.match(line):
            if current:
                yield current_line_no, current
            current = [line]
            current_line_no = line_no
        elif current:
            # Обычно блок 1С = 3 строки; ограничиваем рост при битых данных.
            if len(current) < 8:
                current.append(line)

    if current:
        yield current_line_no, current


def parse_1c_journal_text(text: str) -> JournalParseResult:
    """
    Преобразует журнал 1С (UTF-8 txt) в таблицу для processor.py.

    Принятые блоки «Данные. Проведение» → 2 строки таблицы.
    Отфильтрованные блоки → компактный DataFrame для листа «Не принятые».
    """
    op_users: list[str] = []
    op_statuses: list[str] = []
    op_presents: list[str] = []

    filtered_rows: list[tuple[str, str, str, str, str, str, int]] = []
    seen_transactions: set[str] = set()

    for line_no, block in _iter_blocks(text):
        user, status, metadata, presentation, reject_reason = _inspect_block(block)

        if reject_reason:
            if reject_reason == "Служебная/пустая запись журнала":
                continue
            filtered_rows.append(
                (
                    user or "(без пользователя)",
                    reject_reason,
                    user,
                    status,
                    metadata,
                    presentation,
                    line_no,
                )
            )
            continue

        transaction_id = _extract_transaction_id(block)
        if transaction_id:
            if transaction_id in seen_transactions:
                filtered_rows.append(
                    (
                        user or "(без пользователя)",
                        f"Не принято: дубликат транзакции ({transaction_id})",
                        user,
                        status,
                        metadata,
                        presentation,
                        line_no,
                    )
                )
                continue
            seen_transactions.add(transaction_id)

        op_users.extend((user, ""))
        op_statuses.extend((status, ""))
        op_presents.extend((metadata, presentation))

    operations = pd.DataFrame(
        {
            COL_APP: op_users,
            COL_STATUS: op_statuses,
            COL_PRESENTATION: op_presents,
        }
    )
    filtered = pd.DataFrame(filtered_rows, columns=FILTERED_COLUMNS)
    return JournalParseResult(operations=operations, filtered=filtered)


def parse_1c_journal_bytes(data: bytes) -> JournalParseResult:
    """Читает журнал 1С из байтов (UTF-8 или UTF-8-sig)."""
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return parse_1c_journal_text(data.decode(encoding))
        except UnicodeDecodeError:
            continue
    return parse_1c_journal_text(data.decode("utf-8", errors="replace"))
