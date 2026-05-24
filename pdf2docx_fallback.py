#!/usr/bin/env python3
"""
Service-level fallback detection for pdf2docx conversions.

The fallback is intentionally outside pdf2docx's default behavior: convert once
with the selected mode, inspect the generated DOCX, and only retry with the
fork-only page-frame option when the output looks like a whole-page lattice
table.
"""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from copy import deepcopy
from pathlib import Path
from typing import Any

from lxml import etree
from pdf2docx import Converter

from pdf_conversion_modes import CONVERSION_MODE_DEFAULT


PAGE_FRAME_FALLBACK_OPTION = "ignore_page_frame_tables"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def converter_supports_page_frame_fallback(pdf_path: str | Path) -> bool:
    """Return whether the installed pdf2docx build supports the fork option."""
    converter = None
    try:
        converter = Converter(str(pdf_path))
        return PAGE_FRAME_FALLBACK_OPTION in converter.default_settings
    except Exception:
        return False
    finally:
        if converter:
            converter.close()


def page_frame_fallback_options(conversion_options: dict[str, Any]) -> dict[str, Any]:
    options = dict(conversion_options)
    options[PAGE_FRAME_FALLBACK_OPTION] = True
    return options


def _body_segments(body) -> list[tuple[int, int]]:
    """Split a pdf2docx document body into page-like segments."""
    children = list(body)
    if not children:
        return []

    segments = []
    start = 0
    for index, element in enumerate(children):
        if element.find(".//w:sectPr", NS) is not None or element.tag == f"{{{W_NS}}}sectPr":
            segments.append((start, index + 1))
            start = index + 1
    if start < len(children):
        segments.append((start, len(children)))
    return segments


def _element_text_len(element) -> int:
    return len("".join(text.strip() for text in element.itertext() if text.strip()))


def _has_section_properties(element) -> bool:
    return element.find(".//w:sectPr", NS) is not None or element.tag == f"{{{W_NS}}}sectPr"


def _with_target_segment_boundary(replacement_elements: list[Any], target_elements: list[Any]) -> list[Any]:
    """Use replacement content while preserving the target segment's page boundary."""
    if not replacement_elements or not target_elements:
        return replacement_elements

    if not _has_section_properties(replacement_elements[-1]):
        return replacement_elements
    if not _has_section_properties(target_elements[-1]):
        return replacement_elements

    merged = list(replacement_elements)
    merged[-1] = deepcopy(target_elements[-1])
    return merged


def _segment_page_frame_reason(elements: list[Any]) -> str | None:
    tables = [element for element in elements if element.tag == f"{{{W_NS}}}tbl"]
    if not tables:
        return None

    paragraph_text_len = sum(
        _element_text_len(element)
        for element in elements
        if element.tag == f"{{{W_NS}}}p"
    )
    paragraph_count = sum(
        1
        for element in elements
        if element.tag == f"{{{W_NS}}}p" and _element_text_len(element)
    )

    largest_table = max(
        tables,
        key=lambda table: len(table.findall("./w:tr", NS))
        * max(
            (len(row.findall("./w:tc", NS)) for row in table.findall("./w:tr", NS)),
            default=0,
        ),
    )
    rows = len(largest_table.findall("./w:tr", NS))
    cols = max(
        (len(row.findall("./w:tc", NS)) for row in largest_table.findall("./w:tr", NS)),
        default=0,
    )
    if rows < 6 or cols < 6:
        return None

    table_text_len = _element_text_len(largest_table)
    table_text = " ".join("".join(largest_table.itertext()).split())
    has_page_frame_markers = (
        "Document Name" in table_text
        or "Verification Method" in table_text
        or "Legacy ID" in table_text
    )
    table_dominates_page = (
        table_text_len >= 800
        and has_page_frame_markers
        and paragraph_text_len <= 40
        and (paragraph_count <= 2 or table_text_len >= paragraph_text_len * 3)
    )
    if not table_dominates_page:
        return None

    return (
        f"largest table {rows}x{cols} dominates page "
        f"(table chars={table_text_len}, paragraph chars={paragraph_text_len})"
    )


def detect_page_frame_table_pages(
    docx_path: str | Path,
    conversion_mode: str | None,
) -> tuple[list[int], str]:
    """Return zero-based converted-output page indexes that look like page-frame tables."""
    if conversion_mode and conversion_mode != CONVERSION_MODE_DEFAULT:
        return [], "fallback only applies to layout mode"

    path = Path(docx_path)
    if not path.exists():
        return [], "output DOCX does not exist"

    with zipfile.ZipFile(path, "r") as docx_zip:
        root = etree.fromstring(docx_zip.read("word/document.xml"))

    body = root.find("w:body", NS)
    if body is None:
        return [], "output DOCX has no document body"

    children = list(body)
    if not children:
        return [], "output DOCX body is empty"

    matched = []
    reasons = []
    for page_index, (start, end) in enumerate(_body_segments(body)):
        reason = _segment_page_frame_reason(children[start:end])
        if reason:
            matched.append(page_index)
            reasons.append(f"page {page_index + 1}: {reason}")

    if not matched:
        return [], "no page-frame table pages detected"

    return matched, "; ".join(reasons)


def should_retry_page_frame_fallback(
    docx_path: str | Path,
    conversion_mode: str | None,
) -> tuple[bool, str]:
    """Detect the pdf2docx whole-page table failure shape.

    This is deliberately conservative: only the default layout mode is retried,
    and the generated DOCX must be dominated by one large table. Real documents
    can contain wide tables, so this should remain a retry heuristic rather than
    a change to pdf2docx defaults.
    """
    pages, reason = detect_page_frame_table_pages(docx_path, conversion_mode)
    return bool(pages), reason


def replace_docx_pages(
    target_docx_path: str | Path,
    replacement_docx_by_page: dict[int, str | Path],
) -> None:
    """Replace page-like body segments in target_docx_path with single-page DOCX bodies."""
    target_docx_path = Path(target_docx_path)
    if not replacement_docx_by_page:
        return

    with zipfile.ZipFile(target_docx_path, "r") as target_zip:
        target_xml = target_zip.read("word/document.xml")
        target_root = etree.fromstring(target_xml)

    target_body = target_root.find("w:body", NS)
    if target_body is None:
        raise ValueError("target DOCX has no document body")

    target_children = list(target_body)
    segments = _body_segments(target_body)
    if not segments:
        raise ValueError("target DOCX has no replaceable body segments")

    replacements: dict[int, list[Any]] = {}
    for page_index, replacement_path in replacement_docx_by_page.items():
        if page_index < 0 or page_index >= len(segments):
            raise IndexError(f"page index {page_index} is outside the target DOCX")

        with zipfile.ZipFile(replacement_path, "r") as replacement_zip:
            replacement_root = etree.fromstring(replacement_zip.read("word/document.xml"))
        replacement_body = replacement_root.find("w:body", NS)
        if replacement_body is None:
            raise ValueError(f"replacement DOCX has no body: {replacement_path}")

        replacement_segments = _body_segments(replacement_body)
        if not replacement_segments:
            raise ValueError(f"replacement DOCX has no body segments: {replacement_path}")
        start, end = replacement_segments[0]
        replacement_elements = [
            deepcopy(element)
            for element in list(replacement_body)[start:end]
        ]
        target_start, target_end = segments[page_index]
        replacements[page_index] = _with_target_segment_boundary(
            replacement_elements,
            target_children[target_start:target_end],
        )

    new_children = []
    for page_index, (start, end) in enumerate(segments):
        if page_index in replacements:
            new_children.extend(replacements[page_index])
        else:
            new_children.extend(deepcopy(element) for element in target_children[start:end])

    for element in list(target_body):
        target_body.remove(element)
    for element in new_children:
        target_body.append(element)

    new_document_xml = etree.tostring(
        target_root,
        encoding="UTF-8",
        xml_declaration=target_xml.lstrip().startswith(b"<?xml"),
        standalone=False,
    )

    with tempfile.NamedTemporaryFile(
        suffix=".docx",
        dir=target_docx_path.parent,
        delete=False,
    ) as tmp_file:
        tmp_path = Path(tmp_file.name)

    try:
        with zipfile.ZipFile(target_docx_path, "r") as source, zipfile.ZipFile(
            tmp_path,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as target:
            for item in source.infolist():
                if item.filename == "word/document.xml":
                    target.writestr(item, new_document_xml)
                else:
                    target.writestr(item, source.read(item.filename))
        shutil.move(str(tmp_path), str(target_docx_path))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
