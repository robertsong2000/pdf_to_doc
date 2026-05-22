#!/usr/bin/env python3

import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION

from pdf2docx_fallback import (
    detect_page_frame_table_pages,
    repair_page_frame_inner_tables,
    replace_docx_pages,
)


def _add_large_table(document, prefix: str) -> None:
    table = document.add_table(rows=8, cols=8)
    for row_index, row in enumerate(table.rows):
        for col_index, cell in enumerate(row.cells):
            cell.text = f"{prefix} cell {row_index}-{col_index} " + ("content " * 4)
    table.cell(0, 0).text = "Document Name"
    table.cell(1, 0).text = "Verification Method"
    table.cell(2, 0).text = "Legacy ID"


class PageFrameFallbackTests(unittest.TestCase):
    def test_detects_only_page_segment_with_dominating_table(self):
        document = Document()
        document.add_paragraph("Page one normal paragraph")
        document.add_section(WD_SECTION.NEW_PAGE)
        _add_large_table(document, "bad page")
        document.add_section(WD_SECTION.NEW_PAGE)
        document.add_paragraph("Page three normal paragraph")

        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "page-frame.docx"
            document.save(docx_path)

            pages, reason = detect_page_frame_table_pages(docx_path, "layout")

            self.assertEqual(pages, [1])
            self.assertIn("page 2", reason)

    def test_replace_docx_pages_keeps_unmatched_pages(self):
        target = Document()
        target.add_paragraph("page one stays")
        target.add_section(WD_SECTION.NEW_PAGE)
        _add_large_table(target, "bad page")
        target.add_section(WD_SECTION.NEW_PAGE)
        target.add_paragraph("page three stays")

        replacement = Document()
        replacement.add_paragraph("page two replacement")

        with tempfile.TemporaryDirectory() as tmp_dir:
            target_path = Path(tmp_dir) / "target.docx"
            replacement_path = Path(tmp_dir) / "replacement.docx"
            target.save(target_path)
            replacement.save(replacement_path)

            replace_docx_pages(target_path, {1: replacement_path})

            cleaned = Document(str(target_path))
            text = "\n".join(paragraph.text for paragraph in cleaned.paragraphs)
            self.assertIn("page one stays", text)
            self.assertIn("page two replacement", text)
            self.assertIn("page three stays", text)
            self.assertNotIn("bad page cell", text)

    def test_repairs_eeprom_table_from_pdf_page(self):
        test_pdf = Path("test.pdf")
        if not test_pdf.exists():
            self.skipTest("test.pdf fixture is not available")

        document = Document()
        document.add_paragraph("146582 v1 General VMM requirement [CS Released]")
        table = document.add_table(rows=1, cols=7)
        headers = [
            "EEPROM Parameter",
            "Read/Write",
            "Range",
            "Resolution",
            "Unit",
            "Default Value",
            "Description",
        ]
        for index, header in enumerate(headers):
            table.cell(0, index).text = header
        document.add_paragraph("Document Name")

        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "fallback.docx"
            document.save(docx_path)

            repaired = repair_page_frame_inner_tables(docx_path, test_pdf, 0)

            self.assertEqual(repaired, 1)
            repaired_doc = Document(str(docx_path))
            text = "\n".join(
                cell.text
                for table in repaired_doc.tables
                for row in table.rows
                for cell in row.cells
            )
            self.assertIn("EEPROM Parameter", text)
            self.assertIn("WiprFrntSrvPosngRe", text)
            self.assertIn("WiprReSrvPosngReq", text)
            self.assertEqual(len(repaired_doc.tables[0].rows), 3)
            self.assertEqual(len(repaired_doc.tables[0].columns), 7)


if __name__ == "__main__":
    unittest.main()
