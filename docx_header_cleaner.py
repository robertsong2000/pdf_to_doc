#!/usr/bin/env python3
"""
Post-process DOCX files generated from PDFs.

pdf2docx often converts repeated PDF page headers into the first rows of
normal body tables instead of Word header parts. This module removes only
those repeated header rows and keeps the rest of each page table intact.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Sequence

from docx import Document


FIRST_HEADER_ROW_MARKERS = (
    "Geely Automotive Research Institute",
    "Document Type",
    "Document Release Status",
)

SECOND_HEADER_ROW_MARKERS = (
    "Document No",
    "Revision",
    "Volume No",
    "Page No",
)

HEADER_ROWS_TO_REMOVE = 2


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _row_text(row) -> str:
    values = []
    for cell in row.cells:
        text = _normalize_text(cell.text)
        if text:
            values.append(text)
    return " ".join(values)


def _has_markers(text: str, markers: Sequence[str]) -> bool:
    normalized = text.lower()
    return all(marker.lower() in normalized for marker in markers)


def _starts_with_repeated_pdf_header(table) -> bool:
    if len(table.rows) < HEADER_ROWS_TO_REMOVE + 1:
        return False

    first_row = _row_text(table.rows[0])
    second_row = _row_text(table.rows[1])

    return _has_markers(first_row, FIRST_HEADER_ROW_MARKERS) and _has_markers(
        second_row, SECOND_HEADER_ROW_MARKERS
    )


def _remove_first_rows(table, count: int) -> int:
    removed = 0
    for _ in range(min(count, len(table.rows))):
        row_element = table.rows[0]._tr
        parent = row_element.getparent()
        if parent is None:
            break
        parent.remove(row_element)
        removed += 1
    return removed


def remove_repeated_pdf_headers(docx_path: str | Path) -> int:
    """
    Remove repeated PDF page-header rows from a DOCX file.

    Returns the number of rows removed.
    """
    docx_path = Path(docx_path)
    document = Document(str(docx_path))

    removed = 0
    for table in list(document.tables):
        if not _starts_with_repeated_pdf_header(table):
            continue

        removed += _remove_first_rows(table, HEADER_ROWS_TO_REMOVE)

    if removed:
        document.save(str(docx_path))

    return removed
