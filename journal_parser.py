"""Парсер журнала регистрации 1С (выгрузка «Текстовый файл UTF-8»)."""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from processor import COL_APP, COL_PRESENTATION, COL_STATUS, DOCUMENT_DATE_PATTERN

BLOCK_START = re.compile(r"^\d{2}\.\d{2}\.\d{4}\s+\d{1,2}:\d{2}:\d{2}\t")
EVENT_POSTING = "Данные. Проведение"
STATUS_FIXED = "зафиксирована"
TRANSACTION_ID = re.compile(r"\((\d+)\)\s*$")


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


def _parse_block(block_lines: list[str]) -> tuple[str, str, str, str] | None:
    """Возвращает (фио, статус, операция, подтверждение) или None."""
    if not block_lines:
        return None

    header_parts = block_lines[0].split("\t")
    user = _field(header_parts, 1)
    event = _field(header_parts, 3)
    status = _field(header_parts, 4)
    metadata = _field(header_parts, 5)

    if not user or not metadata:
        return None
    if not _is_posting_event(event):
        return None
    if not _is_fixed_status(status):
        return None

    presentation = ""
    for line in block_lines[1:]:
        candidate = _last_filled(line.split("\t"))
        if candidate and DOCUMENT_DATE_PATTERN.search(candidate):
            presentation = candidate
            break

    if not presentation:
        return None
    if _normalize(presentation).lower() == _normalize(metadata).lower():
        return None

    return user, status, metadata, presentation


def parse_1c_journal_text(text: str) -> pd.DataFrame:
    """
    Преобразует журнал 1С (UTF-8 txt) в таблицу для processor.py.

    Каждый блок «Данные. Проведение» даёт две строки: заголовок операции и подтверждение.
    """
    blocks: list[list[str]] = []
    current: list[str] = []

    for raw_line in text.splitlines():
        line = raw_line.rstrip("\r")
        if BLOCK_START.match(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)

    if current:
        blocks.append(current)

    rows: list[dict[str, str]] = []
    seen_transactions: set[str] = set()

    for block in blocks:
        parsed = _parse_block(block)
        if parsed is None:
            continue

        user, status, metadata, presentation = parsed
        transaction_id = _extract_transaction_id(block)
        if transaction_id:
            if transaction_id in seen_transactions:
                continue
            seen_transactions.add(transaction_id)

        rows.append(
            {
                COL_APP: user,
                COL_STATUS: status,
                COL_PRESENTATION: metadata,
            }
        )
        rows.append(
            {
                COL_APP: "",
                COL_STATUS: "",
                COL_PRESENTATION: presentation,
            }
        )

    return pd.DataFrame(rows, columns=[COL_APP, COL_STATUS, COL_PRESENTATION])


def parse_1c_journal_bytes(data: bytes) -> pd.DataFrame:
    """Читает журнал 1С из байтов (UTF-8 или UTF-8-sig)."""
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return parse_1c_journal_text(data.decode(encoding))
        except UnicodeDecodeError:
            continue
    return parse_1c_journal_text(data.decode("utf-8", errors="replace"))
