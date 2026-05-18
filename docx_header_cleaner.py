#!/usr/bin/env python3
"""
Post-process DOCX files generated from PDFs.

pdf2docx often converts repeated PDF page headers into the first rows of
normal body tables instead of Word header parts. This module removes only
those repeated header rows and keeps the rest of each page table intact.
It also normalizes customer-specific references in the final document.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Sequence

from docx import Document
from lxml import etree


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
WORD_TEXT_TAG = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"

GEELY_REFERENCE_REPLACEMENTS = (
    ("Geely Automotive Research Institute (Ningbo)Co.Ltd", "Renault"),
    ("Geely Automotive Research Institute (Ningbo) Co.Ltd", "Renault"),
    ("Geely Automotive Research Institute", "Renault"),
    ("GEEA3.0", "Renault"),
    ("GEEA", "Renault"),
    ("GEELY", "Renault"),
    ("Geely", "Renault"),
    ("geely", "Renault"),
    ("吉利", "Renault"),
)


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _row_text(row) -> str:
    values = [
        element.text or ""
        for element in row._tr.iter()
        if element.tag == WORD_TEXT_TAG and element.text
    ]
    return _normalize_text(" ".join(values))


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


def _replace_in_paragraph(paragraph) -> int:
    replaced = 0
    for run in paragraph.runs:
        text = run.text
        for old, new in GEELY_REFERENCE_REPLACEMENTS:
            count = text.count(old)
            if not count:
                continue
            text = text.replace(old, new)
            replaced += count
        if text != run.text:
            run.text = text
    return replaced


def _iter_document_paragraphs(document):
    for paragraph in document.paragraphs:
        yield paragraph

    seen_cells = set()
    for table in document.tables:
        for row in table.rows:
            try:
                cells = row.cells
            except ValueError:
                continue
            for cell in cells:
                cell_id = id(cell._tc)
                if cell_id in seen_cells:
                    continue
                seen_cells.add(cell_id)
                for paragraph in cell.paragraphs:
                    yield paragraph


def _replace_across_text_nodes(text_nodes) -> int:
    replaced = 0

    for old, new in GEELY_REFERENCE_REPLACEMENTS:
        while True:
            texts = [node.text or "" for node in text_nodes]
            full_text = "".join(texts)
            match_start = full_text.find(old)
            if match_start == -1:
                break

            match_end = match_start + len(old)
            offset = 0
            first_index = None
            last_index = None
            first_inner_start = 0
            last_inner_end = 0

            for index, text in enumerate(texts):
                next_offset = offset + len(text)
                if first_index is None and match_start < next_offset:
                    first_index = index
                    first_inner_start = match_start - offset
                if first_index is not None and match_end <= next_offset:
                    last_index = index
                    last_inner_end = match_end - offset
                    break
                offset = next_offset

            if first_index is None or last_index is None:
                break

            first_text = texts[first_index]
            last_text = texts[last_index]

            if first_index == last_index:
                text_nodes[first_index].text = (
                    first_text[:first_inner_start] + new + first_text[last_inner_end:]
                )
            else:
                text_nodes[first_index].text = first_text[:first_inner_start] + new
                for index in range(first_index + 1, last_index):
                    text_nodes[index].text = ""
                text_nodes[last_index].text = last_text[last_inner_end:]

            replaced += 1

    return replaced


def _replace_references_in_xml_part(xml_content: bytes) -> tuple[bytes, int]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(xml_content, parser)
    text_nodes = [element for element in root.iter() if element.text]
    replaced = _replace_across_text_nodes(text_nodes)
    if not replaced:
        return xml_content, 0
    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=xml_content.lstrip().startswith(b"<?xml"),
        standalone=False,
    ), replaced


def _replace_references_in_all_xml_parts(docx_path: Path) -> int:
    replaced = 0

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for item in source.infolist():
                content = source.read(item.filename)
                if item.filename.endswith(".xml"):
                    try:
                        content, part_replaced = _replace_references_in_xml_part(content)
                    except etree.XMLSyntaxError:
                        part_replaced = 0
                    replaced += part_replaced
                target.writestr(item, content)

        if replaced:
            shutil.move(str(tmp_path), str(docx_path))
        else:
            tmp_path.unlink(missing_ok=True)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise

    return replaced


def replace_geely_references(docx_path: str | Path) -> int:
    """
    Replace Geely/吉利 customer references with Renault across the DOCX.

    Returns the number of text replacements made.
    """
    docx_path = Path(docx_path)
    document = Document(str(docx_path))

    replaced = 0
    for paragraph in _iter_document_paragraphs(document):
        replaced += _replace_in_paragraph(paragraph)

    for section in document.sections:
        for part in (section.header, section.first_page_header, section.even_page_header):
            for paragraph in _iter_document_paragraphs(part):
                replaced += _replace_in_paragraph(paragraph)
        for part in (section.footer, section.first_page_footer, section.even_page_footer):
            for paragraph in _iter_document_paragraphs(part):
                replaced += _replace_in_paragraph(paragraph)

    if replaced:
        document.save(str(docx_path))

    replaced += _replace_references_in_all_xml_parts(docx_path)
    return replaced


def post_process_converted_docx(
    docx_path: str | Path,
    remove_headers: bool = True,
    replace_oem_info: bool = True,
) -> tuple[int, int]:
    """
    Apply all final DOCX cleanup steps.

    Returns (removed_header_rows, replaced_customer_references).
    """
    removed_headers = remove_repeated_pdf_headers(docx_path) if remove_headers else 0
    replaced_references = replace_geely_references(docx_path) if replace_oem_info else 0
    return removed_headers, replaced_references
