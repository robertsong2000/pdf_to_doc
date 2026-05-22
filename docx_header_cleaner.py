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
from copy import deepcopy
from dataclasses import dataclass
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
LEADING_HEADER_PARAGRAPH_SCAN_LIMIT = 12
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_TEXT_TAG = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
WORD_PARAGRAPH_TAG = f"{{{W_NS}}}p"
WORD_ROW_TAG = f"{{{W_NS}}}tr"
WORD_CELL_TAG = f"{{{W_NS}}}tc"
WORD_RUN_TAG = f"{{{W_NS}}}r"
WORD_RUN_PROPERTIES_TAG = f"{{{W_NS}}}rPr"
WORD_TAB_TAG = f"{{{W_NS}}}tab"
WORD_TABLE_CELL_PROPERTIES_TAG = f"{{{W_NS}}}tcPr"
WORD_GRID_SPAN_TAG = f"{{{W_NS}}}gridSpan"
WORD_VERTICAL_MERGE_TAG = f"{{{W_NS}}}vMerge"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOCUMENT_XML_PART_RE = re.compile(r"^word/document\.xml$")
DOCUMENT_XML = "word/document.xml"
SAFETY_FIELD_LABELS_TEXT = "Legacy ID: FTTI: ASIL (Decomp):"
SAFETY_FIELD_LABEL_TEXTS = ("Legacy ID:", "FTTI:", "ASIL (Decomp):")
SAFETY_FIELD_RE = re.compile(
    r"^(Legacy ID|FTTI|ASIL \(Decomp\)|Comment|Safe state|Verification Method):\s*(.*)$"
)

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


@dataclass(frozen=True)
class SafetyFieldBlock:
    legacy_id: str
    ftti: str
    asil: str


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _xml_text(element) -> str:
    values = [
        text_element.text or ""
        for text_element in element.iter()
        if text_element.tag == WORD_TEXT_TAG and text_element.text
    ]
    return _normalize_text(" ".join(values))


def _set_paragraph_text(paragraph_element, text: str) -> None:
    for child in list(paragraph_element):
        if child.tag != f"{{{W_NS}}}pPr":
            paragraph_element.remove(child)

    run = etree.SubElement(paragraph_element, WORD_RUN_TAG)
    text_element = etree.SubElement(run, WORD_TEXT_TAG)
    text_element.set(XML_SPACE, "preserve")
    text_element.text = text


def _paragraph_with_text(reference_paragraph, text: str):
    paragraph = deepcopy(reference_paragraph)
    _set_paragraph_text(paragraph, text)
    return paragraph


def _text_runs(paragraph_element) -> list:
    return [
        run
        for run in paragraph_element.findall(WORD_RUN_TAG)
        if _xml_text(run)
    ]


def _add_text_run(paragraph_element, text: str, template_run=None) -> None:
    run = etree.SubElement(paragraph_element, WORD_RUN_TAG)
    if template_run is not None:
        run_properties = template_run.find(WORD_RUN_PROPERTIES_TAG)
        if run_properties is not None:
            run.append(deepcopy(run_properties))
    text_element = etree.SubElement(run, WORD_TEXT_TAG)
    text_element.set(XML_SPACE, "preserve")
    text_element.text = text


def _add_tab_run(paragraph_element) -> None:
    run = etree.SubElement(paragraph_element, WORD_RUN_TAG)
    etree.SubElement(run, WORD_TAB_TAG)


def _field_paragraph_from_template(
    template_paragraph,
    label: str,
    value: str,
):
    paragraph = deepcopy(template_paragraph)
    for child in list(paragraph):
        if child.tag != f"{{{W_NS}}}pPr":
            paragraph.remove(child)

    template_runs = _text_runs(template_paragraph)
    label_run = template_runs[0] if template_runs else None
    value_run = template_runs[-1] if len(template_runs) > 1 else label_run

    _add_text_run(paragraph, f"{label}: ", label_run)
    _add_tab_run(paragraph)
    _add_text_run(paragraph, value, value_run)
    return paragraph


def _find_label_paragraphs(container, expected_labels: Sequence[str]) -> list:
    matched = []
    label_index = 0
    for paragraph in container.iter(WORD_PARAGRAPH_TAG):
        text = _normalize_text(_xml_text(paragraph))
        if not text:
            continue

        expected_text = " ".join(expected_labels[label_index:])
        if text == expected_text:
            matched.append(paragraph)
            return matched

        if text == expected_labels[label_index]:
            matched.append(paragraph)
            label_index += 1
            if label_index == len(expected_labels):
                return matched
            continue

        if matched:
            return []

    return []


def _remove_paragraph(paragraph) -> None:
    parent = paragraph.getparent()
    if parent is not None:
        parent.remove(paragraph)


def _remove_first_matching_paragraph(container, text: str) -> bool:
    for paragraph in container.iter(WORD_PARAGRAPH_TAG):
        if _normalize_text(_xml_text(paragraph)) == text:
            _remove_paragraph(paragraph)
            return True
    return False


def _find_following_field_template(table_element):
    if table_element is None:
        return None

    parent = table_element.getparent()
    if parent is None:
        return None

    start_index = parent.index(table_element) + 1
    for sibling in list(parent)[start_index:]:
        if sibling.tag != WORD_PARAGRAPH_TAG:
            continue
        text = _normalize_text(_xml_text(sibling))
        if SAFETY_FIELD_RE.match(text):
            return sibling
    return None


def _find_nearby_field_template(table_element):
    if table_element is None:
        return None

    parent = table_element.getparent()
    if parent is None:
        return None

    table_index = parent.index(table_element)
    siblings = list(parent)
    for sibling in reversed(siblings[:table_index]):
        if sibling.tag != WORD_PARAGRAPH_TAG:
            continue
        text = _normalize_text(_xml_text(sibling))
        if SAFETY_FIELD_RE.match(text):
            return sibling

    return _find_following_field_template(table_element)


def _insert_paragraphs_after_block(block_element, paragraphs: Sequence) -> None:
    parent = block_element.getparent()
    if parent is None:
        return

    insert_index = parent.index(block_element) + 1
    for paragraph in paragraphs:
        parent.insert(insert_index, paragraph)
        insert_index += 1


def _remove_empty_table(table_element) -> None:
    if any(child.tag == WORD_ROW_TAG for child in table_element):
        return

    parent = table_element.getparent()
    if parent is not None:
        parent.remove(table_element)


def _is_document_name_text(text: str) -> bool:
    return text == "Document Name" or text.startswith("Document Name ")


def _is_pdf_header_paragraph_start(text: str) -> bool:
    return "Document Type" in text and "Document Release Status" in text


def _is_pdf_header_table_element(table_element) -> bool:
    text = _xml_text(table_element)
    if not text:
        return False

    return (
        "Document Type" in text
        and "Document Release Status" in text
        and "Document No" in text
        and "Revision" in text
        and "Page No" in text
        and ("NOTE-SWRS" in text or "RELEASED" in text)
    )


def _next_non_empty_block_text(blocks: Sequence, start_index: int) -> str:
    for block in blocks[start_index:]:
        text = _xml_text(block)
        if text:
            return text
    return ""


def _remove_repeated_pdf_header_blocks_in_xml(xml_content: bytes) -> tuple[bytes, int]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(xml_content, parser)
    body = root.find(f"{{{W_NS}}}body")
    if body is None:
        return xml_content, 0

    removed = 0
    index = 0
    while index < len(body):
        block = body[index]
        text = _xml_text(block)

        if block.tag == WORD_PARAGRAPH_TAG and _is_pdf_header_paragraph_start(text):
            end_index = index + 1
            while end_index < len(body):
                candidate_text = _xml_text(body[end_index])
                if _is_document_name_text(candidate_text):
                    break
                end_index += 1

            if end_index < len(body):
                for removable in list(body)[index:end_index]:
                    body.remove(removable)
                    removed += 1
                continue

        if block.tag == f"{{{W_NS}}}tbl" and _is_pdf_header_table_element(block):
            next_text = _next_non_empty_block_text(list(body), index + 1)
            if _is_document_name_text(next_text):
                body.remove(block)
                removed += 1
                continue

        index += 1

    if not removed:
        return xml_content, 0

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=xml_content.lstrip().startswith(b"<?xml"),
        standalone=False,
    ), removed


def _remove_repeated_pdf_header_blocks(docx_path: Path) -> int:
    removed = 0
    with tempfile.NamedTemporaryFile(
        suffix=".docx",
        delete=False,
        dir=docx_path.parent,
    ) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for item in source.infolist():
                content = source.read(item.filename)
                if item.filename == DOCUMENT_XML:
                    content, removed = _remove_repeated_pdf_header_blocks_in_xml(content)
                target.writestr(item, content)

        if removed:
            shutil.move(str(tmp_path), str(docx_path))
        else:
            tmp_path.unlink(missing_ok=True)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise

    return removed


def _group_words_into_lines(words: list[dict], y_tolerance: float = 3.0) -> list[str]:
    lines: list[dict] = []
    for word in sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2, w["x0"])):
        y_center = (word["top"] + word["bottom"]) / 2
        if lines and abs(lines[-1]["y"] - y_center) <= y_tolerance:
            lines[-1]["words"].append(word)
            lines[-1]["y"] = (
                lines[-1]["y"] * (len(lines[-1]["words"]) - 1) + y_center
            ) / len(lines[-1]["words"])
        else:
            lines.append({"y": y_center, "words": [word]})

    return [
        _normalize_text(
            " ".join(
                word["text"] for word in sorted(line["words"], key=lambda w: w["x0"])
            )
        )
        for line in lines
    ]


def _extract_safety_field_blocks(
    pdf_path: str | Path,
    start_page: int | None = None,
    end_page: int | None = None,
) -> list[SafetyFieldBlock]:
    import pdfplumber

    blocks: list[SafetyFieldBlock] = []
    current: dict[str, str] = {}

    def flush_current() -> None:
        nonlocal current
        if current.get("Legacy ID") and current.get("FTTI") and current.get("ASIL (Decomp)"):
            blocks.append(
                SafetyFieldBlock(
                    current["Legacy ID"],
                    current["FTTI"],
                    current["ASIL (Decomp)"],
                )
            )
        current = {}

    with pdfplumber.open(str(pdf_path)) as pdf:
        start_index = max((start_page or 1) - 1, 0)
        end_index = min(end_page, len(pdf.pages)) if end_page else len(pdf.pages)
        for page in pdf.pages[start_index:end_index]:
            words = page.extract_words(
                x_tolerance=1,
                y_tolerance=3,
                keep_blank_chars=False,
            )
            for line in _group_words_into_lines(words):
                match = SAFETY_FIELD_RE.match(line)
                if not match:
                    continue
                label, value = match.groups()
                if label == "Legacy ID":
                    flush_current()
                if current or label == "Legacy ID":
                    current[label] = value.strip()

    flush_current()
    return blocks


def _repair_collapsed_safety_cells_in_xml(
    xml_content: bytes,
    blocks: Sequence[SafetyFieldBlock],
) -> tuple[bytes, int]:
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    root = etree.fromstring(xml_content, parser)
    by_collapsed_values = {
        _normalize_text(f"{block.legacy_id} {block.ftti} {block.asil}"): block
        for block in blocks
    }
    by_collapsed_legacy_ftti = {
        _normalize_text(f"{block.legacy_id} {block.ftti}"): block
        for block in blocks
    }
    repaired = 0

    for row in list(root.iter(WORD_ROW_TAG)):
        cells = [child for child in row if child.tag == WORD_CELL_TAG]
        if len(cells) < 2:
            continue

        for index in range(len(cells) - 1):
            label_cell = cells[index]
            value_cell = cells[index + 1]
            value_text = _xml_text(value_cell)
            block = by_collapsed_values.get(value_text)
            if block is None:
                continue

            label_paragraphs = _find_label_paragraphs(
                label_cell, SAFETY_FIELD_LABEL_TEXTS
            )
            if not label_paragraphs:
                continue

            row_parent = row.getparent()
            table_element = row_parent
            field_template = _find_nearby_field_template(table_element)
            if field_template is None:
                field_template = label_paragraphs[0]

            replacement_paragraphs = []
            label_paragraph_ids = {id(paragraph) for paragraph in label_paragraphs}
            for paragraph in label_cell.iter(WORD_PARAGRAPH_TAG):
                paragraph_text = _normalize_text(_xml_text(paragraph))
                if not paragraph_text or id(paragraph) in label_paragraph_ids:
                    continue
                replacement_paragraphs.append(
                    _paragraph_with_text(paragraph, paragraph_text)
                )

            replacement_paragraphs.extend(
                [
                    _field_paragraph_from_template(
                        field_template, "Legacy ID", block.legacy_id
                    ),
                    _field_paragraph_from_template(field_template, "FTTI", block.ftti),
                    _field_paragraph_from_template(
                        field_template, "ASIL (Decomp)", block.asil
                    ),
                ]
            )

            if row_parent is not None:
                row_parent.remove(row)
                _insert_paragraphs_after_block(table_element, replacement_paragraphs)
                _remove_empty_table(table_element)
            repaired += 1

    for row in list(root.iter(WORD_ROW_TAG)):
        cells = [child for child in row if child.tag == WORD_CELL_TAG]
        if len(cells) < 2:
            continue

        for index in range(len(cells) - 1):
            label_cell = cells[index]
            value_cell = cells[index + 1]
            value_text = _xml_text(value_cell)
            block = by_collapsed_legacy_ftti.get(value_text)
            if block is None:
                continue

            label_paragraphs = _find_label_paragraphs(
                label_cell, ("Legacy ID:", "FTTI:")
            )
            if not label_paragraphs:
                continue

            asil_text = f"ASIL (Decomp): {block.asil}"
            following_cells = []
            seen_current_row = False
            for candidate_row in root.iter(WORD_ROW_TAG):
                if candidate_row is row:
                    seen_current_row = True
                    continue
                if not seen_current_row:
                    continue
                following_cells.extend(
                    child for child in candidate_row if child.tag == WORD_CELL_TAG
                )

            asil_cell = None
            for candidate_cell in following_cells:
                if _remove_first_matching_paragraph(candidate_cell, asil_text):
                    asil_cell = candidate_cell
                    break
            if asil_cell is None:
                continue

            row_parent = row.getparent()
            table_element = row_parent
            field_template = _find_nearby_field_template(table_element)
            if field_template is None:
                field_template = label_paragraphs[0]

            replacement_paragraphs = []
            label_paragraph_ids = {id(paragraph) for paragraph in label_paragraphs}
            for paragraph in label_cell.iter(WORD_PARAGRAPH_TAG):
                paragraph_text = _normalize_text(_xml_text(paragraph))
                if not paragraph_text or id(paragraph) in label_paragraph_ids:
                    continue
                replacement_paragraphs.append(
                    _paragraph_with_text(paragraph, paragraph_text)
                )

            replacement_paragraphs.extend(
                [
                    _field_paragraph_from_template(
                        field_template, "Legacy ID", block.legacy_id
                    ),
                    _field_paragraph_from_template(field_template, "FTTI", block.ftti),
                    _field_paragraph_from_template(
                        field_template, "ASIL (Decomp)", block.asil
                    ),
                ]
            )

            if row_parent is not None:
                row_parent.remove(row)
                _insert_paragraphs_after_block(table_element, replacement_paragraphs)
                _remove_empty_table(table_element)
            repaired += 1
            break

    if not repaired:
        return xml_content, 0

    return etree.tostring(
        root,
        encoding="UTF-8",
        xml_declaration=xml_content.lstrip().startswith(b"<?xml"),
        standalone=False,
    ), repaired


def repair_safety_fields_from_pdf(
    docx_path: str | Path,
    pdf_path: str | Path,
    start_page: int | None = None,
    end_page: int | None = None,
) -> int:
    blocks = _extract_safety_field_blocks(pdf_path, start_page, end_page)
    if not blocks:
        return 0

    docx_path = Path(docx_path)
    repaired = 0
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(docx_path, "r") as source, zipfile.ZipFile(
            tmp_path, "w", zipfile.ZIP_DEFLATED
        ) as target:
            for item in source.infolist():
                content = source.read(item.filename)
                if DOCUMENT_XML_PART_RE.match(item.filename):
                    content, part_repaired = _repair_collapsed_safety_cells_in_xml(
                        content, blocks
                    )
                    repaired += part_repaired
                target.writestr(item, content)

        if repaired:
            shutil.move(str(tmp_path), str(docx_path))
        else:
            tmp_path.unlink(missing_ok=True)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise

    return repaired


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
    if removed:
        _remove_orphan_vertical_merges(table._tbl)
    return removed


def _cell_grid_span(cell_element) -> int:
    properties = cell_element.find(WORD_TABLE_CELL_PROPERTIES_TAG)
    if properties is None:
        return 1

    grid_span = properties.find(WORD_GRID_SPAN_TAG)
    if grid_span is None:
        return 1

    try:
        return max(int(grid_span.get(f"{{{W_NS}}}val", "1")), 1)
    except ValueError:
        return 1


def _remove_orphan_vertical_merges(table_element) -> None:
    """
    Remove vertical-merge continuations whose restart row was deleted.

    Word allows ``w:vMerge`` without a value to mean "continue the merge from
    the row above". If header cleanup deletes that row, the remaining table can
    become invalid and readers may hide or truncate later rows.
    """
    active_merge_offsets: set[int] = set()

    for row in table_element.findall(WORD_ROW_TAG):
        grid_offset = 0
        for cell in row.findall(WORD_CELL_TAG):
            span = _cell_grid_span(cell)
            offsets = set(range(grid_offset, grid_offset + span))
            properties = cell.find(WORD_TABLE_CELL_PROPERTIES_TAG)
            vertical_merge = (
                properties.find(WORD_VERTICAL_MERGE_TAG)
                if properties is not None
                else None
            )

            if vertical_merge is None:
                active_merge_offsets.difference_update(offsets)
            elif vertical_merge.get(f"{{{W_NS}}}val") == "restart":
                active_merge_offsets.update(offsets)
            elif offsets.isdisjoint(active_merge_offsets):
                properties.remove(vertical_merge)
                active_merge_offsets.difference_update(offsets)

            grid_offset += span


def _paragraph_text(paragraph) -> str:
    return _normalize_text(paragraph.text or "")


def _leading_non_empty_paragraphs(document) -> list[tuple[int, object, str]]:
    paragraphs = []
    for index, paragraph in enumerate(document.paragraphs):
        text = _paragraph_text(paragraph)
        if not text:
            continue
        paragraphs.append((index, paragraph, text))
        if len(paragraphs) >= LEADING_HEADER_PARAGRAPH_SCAN_LIMIT:
            break
    return paragraphs


def _leading_pdf_header_paragraph_end(document) -> int | None:
    leading = _leading_non_empty_paragraphs(document)
    if len(leading) < 5:
        return None

    texts = [text for _, _, text in leading]
    scan_text = " ".join(texts[:8])
    has_first_row_markers = (
        "Document Type" in texts[0]
        and "Document Release Status" in texts[0]
    )
    has_institute = any(
        "Geely Automotive Research Institute" in text
        for text in texts[:4]
    )
    has_second_row_markers = all(
        marker in scan_text
        for marker in SECOND_HEADER_ROW_MARKERS
    )
    if not (has_first_row_markers and has_institute and has_second_row_markers):
        return None

    document_name_index = None
    for index, text in enumerate(texts[:10]):
        if text == "Document Name" or text.startswith("Document Name "):
            document_name_index = index
            break

    if document_name_index is None:
        return None

    return document_name_index - 1


def _remove_leading_pdf_header_paragraphs(document) -> int:
    end_index = _leading_pdf_header_paragraph_end(document)
    if end_index is None:
        return 0

    leading = _leading_non_empty_paragraphs(document)
    if end_index >= len(leading):
        return 0

    first_paragraph_index = leading[0][0]
    last_paragraph_index = leading[end_index][0]
    removed = 0
    for paragraph in document.paragraphs[first_paragraph_index:last_paragraph_index + 1]:
        parent = paragraph._p.getparent()
        if parent is not None:
            parent.remove(paragraph._p)
            removed += 1

    return removed


def remove_repeated_pdf_headers(docx_path: str | Path) -> int:
    """
    Remove repeated PDF page-header rows or leading header paragraphs from a DOCX file.

    Returns the number of removed header rows/paragraphs.
    """
    docx_path = Path(docx_path)
    removed = _remove_repeated_pdf_header_blocks(docx_path)
    document = Document(str(docx_path))

    removed += _remove_leading_pdf_header_paragraphs(document)
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
    pdf_path: str | Path | None = None,
    start_page: int | None = None,
    end_page: int | None = None,
    remove_headers: bool = True,
    replace_oem_info: bool = True,
) -> tuple[int, int]:
    """
    Apply all final DOCX cleanup steps.

    Returns (removed_header_rows, repaired_safety_fields, replaced_customer_references).
    """
    removed_headers = remove_repeated_pdf_headers(docx_path) if remove_headers else 0
    repaired_safety_fields = 0
    if pdf_path is not None:
        repaired_safety_fields = repair_safety_fields_from_pdf(
            docx_path, pdf_path, start_page, end_page
        )
    replaced_references = replace_geely_references(docx_path) if replace_oem_info else 0
    return removed_headers, repaired_safety_fields, replaced_references
