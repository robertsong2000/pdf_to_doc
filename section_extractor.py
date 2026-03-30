#!/usr/bin/env python3
"""
PDF章节号提取器

优先读取PDF内嵌书签/大纲（瞬间完成），没有书签时回退到文本扫描。
"""

import re
from typing import Dict, List, Tuple, Optional
import pdfplumber


class PdfSectionExtractor:
    """Extract section numbers from a PDF and map them to page numbers."""

    # Match dotted-decimal section numbers like "5.15.1.1.4.2"
    # At least 2 levels deep (e.g. "5.15") to avoid false positives from plain numbers
    SECTION_PATTERN = re.compile(
        r'(?:^|\s)'
        r'(\d+(?:\.\d+){1,})'
        r'(?:\s|$|[,;。，、])'
    )

    # Match heading lines: section number followed by title (significant whitespace/tab separator)
    HEADING_PATTERN = re.compile(
        r'^(?:\s*)'
        r'(\d+(?:\.\d+)*)'
        r'(?:\s{2,}|\t)'
        r'(.+?)'
        r'(?:\s*$)'
    )

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._sections_cache = None
        self._total_pages = None
        self._source = None  # 'bookmarks' or 'text_scan'

    def extract_sections(self, progress_callback=None) -> Dict[str, dict]:
        """
        Extract sections: try PDF bookmarks first (fast), fall back to text scan.

        Returns {section_number: {page, title}}.
        """
        if self._sections_cache is not None:
            return self._sections_cache

        # Strategy 1: Try PDF bookmarks/outlines (near-instant)
        sections = self._extract_from_bookmarks()
        if sections:
            self._sections_cache = sections
            self._source = 'bookmarks'
            return sections

        # Strategy 2: Fall back to text scanning (slower)
        sections = self._extract_from_text(progress_callback)
        self._sections_cache = sections
        self._source = 'text_scan'
        return sections

    def _extract_from_bookmarks(self) -> Dict[str, dict]:
        """
        Extract sections from PDF bookmarks/outlines.
        This is near-instant as it only reads metadata, no text extraction needed.
        """
        sections = {}

        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                self._total_pages = len(pdf.pages)
                # pdfplumber wraps pdfminer, access outline via the pdfminer object
                outlines = self._get_outlines(pdf)
                if not outlines:
                    return {}

                for outline in outlines:
                    title = outline.get('title', '').strip()
                    page_num = outline.get('page')
                    if not title or page_num is None:
                        continue

                    # Try to extract section number from bookmark title
                    section_num = self._parse_section_from_title(title)
                    if section_num and section_num not in sections:
                        sections[section_num] = {
                            'page': page_num,
                            'title': title
                        }
        except Exception:
            pass

        return sections

    def _get_outlines(self, pdf) -> List[dict]:
        """
        Get PDF outline/bookmarks as a flat list of {title, page} dicts.
        Handles the pdfminer outline format.
        """
        result = []

        def _flatten(outline_list, level=0):
            for item in outline_list:
                if isinstance(item, tuple) and len(item) >= 2:
                    title = item[0] if isinstance(item[0], str) else ''
                    dest = item[1]

                    page_num = self._resolve_page_number(pdf, dest)
                    if page_num is not None and title:
                        result.append({
                            'title': title,
                            'page': page_num,
                            'level': level
                        })

                    # Recurse into children if present
                    if len(item) > 2:
                        children = item[2:]
                        for child in children:
                            if isinstance(child, list):
                                _flatten(child, level + 1)

        try:
            # Access pdfminer's outline via pdfplumber
            pdfminer_pdf = pdf.doc
            outlines = pdfminer_pdf.get_outlines()
            if outlines:
                _flatten(outlines)
        except Exception:
            pass

        return result

    def _resolve_page_number(self, pdf, dest) -> Optional[int]:
        """Resolve a PDF destination to a 1-indexed page number."""
        try:
            if dest is None:
                return None

            # If dest is a page number directly
            if isinstance(dest, int):
                return dest + 1  # pdfminer uses 0-indexed

            # If dest is a string (named destination), resolve it
            if isinstance(dest, str):
                pdftree = pdf.doc
                try:
                    resolved = pdftree.resolve_dest(dest)
                    if resolved and isinstance(resolved, (list, tuple)) and len(resolved) > 0:
                        page_ref = resolved[0]
                        return self._page_ref_to_num(pdf, page_ref)
                except Exception:
                    pass
                return None

            # If dest is a list/tuple (direct destination)
            if isinstance(dest, (list, tuple)) and len(dest) > 0:
                page_ref = dest[0]
                return self._page_ref_to_num(pdf, page_ref)

        except Exception:
            pass

        return None

    def _page_ref_to_num(self, pdf, page_ref) -> Optional[int]:
        """Convert a pdfminer page reference to 1-indexed page number."""
        try:
            # page_ref is often a pdfminer PDFObjRef
            if hasattr(page_ref, 'objid'):
                # Look through pages to find matching object
                pages = pdf.doc.flattened_pages or []
                for i, page in enumerate(pages):
                    if hasattr(page, 'objid') and page.objid == page_ref.objid:
                        return i + 1
            # Direct page object
            if hasattr(page_ref, 'pageid'):
                pages = pdf.doc.flattened_pages or []
                for i, page in enumerate(pages):
                    if hasattr(page, 'pageid') and page.pageid == page_ref.pageid:
                        return i + 1
        except Exception:
            pass
        return None

    @staticmethod
    def _parse_section_from_title(title: str) -> Optional[str]:
        """
        Extract a section number from a bookmark title.
        Handles formats like: "5.15.1.1.4.2 Some Title" or "5.15.1.1.4.2 Some Title"
        """
        # Match section number at the start of the title
        match = re.match(
            r'^\s*(\d+(?:\.\d+){1,})\s*',
            title
        )
        if match:
            return match.group(1)

        # Also try single-level like "5 Title" (only if followed by substantial text)
        match = re.match(
            r'^\s*(\d+)\s{2,}',
            title
        )
        if match:
            return match.group(1)

        return None

    def _extract_from_text(self, progress_callback=None) -> Dict[str, dict]:
        """
        Fall back: scan PDF text on every page to find section numbers.
        Slower but works for PDFs without bookmarks.
        """
        sections = {}

        with pdfplumber.open(self.pdf_path) as pdf:
            self._total_pages = len(pdf.pages)

            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    continue

                page_num = i + 1  # 1-indexed

                for line in text.split('\n'):
                    line = line.strip()
                    if not line:
                        continue

                    # Try heading pattern first (more reliable)
                    heading_match = self.HEADING_PATTERN.match(line)
                    if heading_match:
                        section_num = heading_match.group(1)
                        section_title = heading_match.group(2).strip()
                        if section_num not in sections:
                            sections[section_num] = {
                                'page': page_num,
                                'title': section_title
                            }
                        continue

                    # Try generic section number pattern
                    for match in self.SECTION_PATTERN.finditer(line):
                        section_num = match.group(1)
                        if section_num not in sections:
                            rest = line[match.end():].strip()
                            title = rest[:80] if rest else None
                            sections[section_num] = {
                                'page': page_num,
                                'title': title
                            }

                if progress_callback:
                    progress_callback(page_num, self._total_pages, len(sections))

        return sections

    def resolve_section_range(
        self, start_section: str, end_section: str
    ) -> Tuple[Optional[int], Optional[int]]:
        """
        Given start and end section numbers, return (start_page, end_page) (1-indexed).

        The end_page is determined by finding the next section after end_section
        and using the page before it. If no next section exists, use total page count.
        """
        sections = self.extract_sections()

        if start_section not in sections:
            raise ValueError(f"未找到起始章节: {start_section}")
        if end_section not in sections:
            raise ValueError(f"未找到结束章节: {end_section}")

        start_page = sections[start_section]['page']
        end_page = sections[end_section]['page']

        # Find the next section page after end_page
        all_pages = sorted(set(s['page'] for s in sections.values()))
        next_pages = [p for p in all_pages if p > end_page]

        if next_pages:
            actual_end = next_pages[0] - 1
        else:
            actual_end = self._total_pages or end_page

        return start_page, actual_end

    def search_sections(self, start_section: str, end_section: str) -> dict:
        """
        Search for specific sections on demand and return page range info.

        This is the on-demand entry point: the user supplies section numbers
        directly, and we resolve them to page numbers using bookmarks first
        (fast) and text scan as fallback.
        """
        sections = self.extract_sections()

        if start_section not in sections:
            raise ValueError(f"未找到起始章节: {start_section}")
        if end_section not in sections:
            raise ValueError(f"未找到结束章节: {end_section}")

        start_page, effective_end = self.resolve_section_range(start_section, end_section)

        return {
            'start_page': start_page,
            'end_page': effective_end,
            'start_section_title': sections[start_section].get('title', ''),
            'end_section_title': sections[end_section].get('title', ''),
            'resolved_range': f'第 {start_page} 页 到 第 {effective_end} 页',
            'source': self._source
        }

    def get_sections_tree(self) -> List[dict]:
        """
        Return sections as a sorted list for frontend display.

        Each entry: {"number": str, "page": int, "title": str, "level": int}
        """
        sections = self.extract_sections()

        result = []
        for num, info in sorted(sections.items(), key=lambda x: self._sort_key(x[0])):
            level = len(num.split('.')) - 1
            result.append({
                'number': num,
                'page': info['page'],
                'title': info.get('title', ''),
                'level': level
            })

        return result

    @staticmethod
    def _sort_key(section_num: str) -> Tuple[int, ...]:
        """Sort key for dotted-decimal section numbers."""
        parts = section_num.split('.')
        return tuple(int(p) for p in parts)
