"""
PDF表格提取模块 - 通过PDF→DOCX→CSV两步转换，利用python-docx提取结构化表格

流程: PDF → DOCX (pdf2docx) → CSV (python-docx)
优势: DOCX中表格结构是结构化的，提取准确率远高于pdfplumber直接提取
"""
import csv
import gc
import os
import io
import tempfile
import zipfile
from typing import List, Optional, Tuple
from docx import Document
from pdf2docx import Converter

# Word XML 命名空间
W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


def extract_table_grid(table) -> List[List[str]]:
    """
    从 python-docx 的 Table 对象提取表格数据，正确处理合并单元格。

    核心逻辑：遍历每行的 <w:tc> 元素，通过 gridSpan 和 vMerge
    精确计算每个单元格在网格中的真实列位置和跨度。

    Args:
        table: python-docx 的 Table 对象

    Returns:
        二维列表 [[cell_text, ...], ...]，行列数与视觉表格一致
    """
    # 获取表格网格列数
    tbl_grid = table._tbl.find(f'{W_NS}tblGrid')
    if tbl_grid is not None:
        grid_cols = len(tbl_grid.findall(f'{W_NS}gridCol'))
    else:
        # 回退：取第一行的 cell 数量
        grid_cols = len(table.rows[0].cells) if table.rows else 0

    if grid_cols == 0:
        return []

    # 初始化网格（行数 x 列数）
    num_rows = len(table.rows)
    grid = [[None] * grid_cols for _ in range(num_rows)]

    for row_idx, row in enumerate(table.rows):
        col_pos = 0  # 当前要填充的列位置

        # 找到该行中已被上方垂直合并占据的位置
        while col_pos < grid_cols and grid[row_idx][col_pos] is not None:
            col_pos += 1

        # 遍历该行的每个 <w:tc> 元素（不是 row.cells，避免合并单元格重复）
        tr_element = row._tr
        tc_elements = tr_element.findall(f'{W_NS}tc')

        for tc_element in tc_elements:
            # 跳过已被垂直合并占据的列
            while col_pos < grid_cols and grid[row_idx][col_pos] is not None:
                col_pos += 1

            if col_pos >= grid_cols:
                break

            # 读取 gridSpan（水平跨度）
            tc_pr = tc_element.find(f'{W_NS}tcPr')
            grid_span = 1
            v_merge_val = None  # None=不是合并, 'restart'=合并起点, 'continue'=合并续接

            if tc_pr is not None:
                gs_elem = tc_pr.find(f'{W_NS}gridSpan')
                if gs_elem is not None:
                    grid_span = int(gs_elem.get(f'{W_NS}val', 1))

                # 读取 vMerge（垂直合并）
                vm_elem = tc_pr.find(f'{W_NS}vMerge')
                if vm_elem is not None:
                    v_merge_val = vm_elem.get(f'{W_NS}val')
                    # vMerge 存在但无 val 属性 = "continue"
                    if v_merge_val is None:
                        v_merge_val = 'continue'

            # 提取单元格文本
            text = _extract_tc_text(tc_element)

            if v_merge_val == 'continue':
                # 垂直合并续接：从上方行复制文本
                for r in range(row_idx - 1, -1, -1):
                    if grid[r][col_pos] is not None:
                        text = grid[r][col_pos]
                        break

            # 填充网格：水平跨度内的所有位置都填入文本
            for c in range(col_pos, min(col_pos + grid_span, grid_cols)):
                grid[row_idx][c] = text

            col_pos += grid_span

    # 清理 None（不应存在，但以防万一）
    for row in grid:
        for i in range(len(row)):
            if row[i] is None:
                row[i] = ''

    return grid


def _extract_table_grid_from_element(tbl_element) -> List[List[str]]:
    """
    从 XML <w:tbl> 元素直接提取表格数据，无需 python-docx Table 对象。
    支持 extract_table_grid 相同的合并单元格处理逻辑。
    """
    tbl_grid = tbl_element.find(f'{W_NS}tblGrid')
    if tbl_grid is not None:
        grid_cols = len(tbl_grid.findall(f'{W_NS}gridCol'))
    else:
        # 回退：取第一行的 cell 数量
        first_row = tbl_element.find(f'{W_NS}tr')
        if first_row is not None:
            grid_cols = len(first_row.findall(f'{W_NS}tc'))
        else:
            grid_cols = 0

    if grid_cols == 0:
        return []

    rows = tbl_element.findall(f'{W_NS}tr')
    num_rows = len(rows)
    grid = [[None] * grid_cols for _ in range(num_rows)]

    for row_idx, row in enumerate(rows):
        col_pos = 0

        while col_pos < grid_cols and grid[row_idx][col_pos] is not None:
            col_pos += 1

        tc_elements = row.findall(f'{W_NS}tc')

        for tc_element in tc_elements:
            while col_pos < grid_cols and grid[row_idx][col_pos] is not None:
                col_pos += 1

            if col_pos >= grid_cols:
                break

            tc_pr = tc_element.find(f'{W_NS}tcPr')
            grid_span = 1
            v_merge_val = None

            if tc_pr is not None:
                gs_elem = tc_pr.find(f'{W_NS}gridSpan')
                if gs_elem is not None:
                    grid_span = int(gs_elem.get(f'{W_NS}val', 1))

                vm_elem = tc_pr.find(f'{W_NS}vMerge')
                if vm_elem is not None:
                    v_merge_val = vm_elem.get(f'{W_NS}val')
                    if v_merge_val is None:
                        v_merge_val = 'continue'

            # 提取单元格文本（仅当前单元格，不递归进入嵌套表格）
            text = _extract_tc_text_no_nested(tc_element)

            if v_merge_val == 'continue':
                for r in range(row_idx - 1, -1, -1):
                    if grid[r][col_pos] is not None:
                        text = grid[r][col_pos]
                        break

            for c in range(col_pos, min(col_pos + grid_span, grid_cols)):
                grid[row_idx][c] = text

            col_pos += grid_span

    for row in grid:
        for i in range(len(row)):
            if row[i] is None:
                row[i] = ''

    return grid


def _extract_tc_text_no_nested(tc_element) -> str:
    """从 <w:tc> 元素提取文本，排除嵌套表格中的文本"""
    texts = []
    for p in tc_element.findall(f'{W_NS}p'):
        para_texts = []
        for r in p.findall(f'{W_NS}r'):
            for t in r.findall(f'{W_NS}t'):
                if t.text:
                    para_texts.append(t.text)
        texts.append(''.join(para_texts))
    text = '\n'.join(texts).strip()
    text = text.replace('\n', ' ').replace('\r', '')
    return text


def _extract_tc_text(tc_element) -> str:
    """从 <w:tc> XML 元素中提取纯文本"""
    texts = []
    for p in tc_element.findall(f'.//{W_NS}p'):
        para_texts = []
        for r in p.findall(f'.//{W_NS}r'):
            for t in r.findall(f'{W_NS}t'):
                if t.text:
                    para_texts.append(t.text)
        texts.append(''.join(para_texts))
    text = '\n'.join(texts).strip()
    text = text.replace('\n', ' ').replace('\r', '')
    return text


class DocxTableExtractor:
    """DOCX表格提取器 - 直接从DOCX文件提取表格为CSV"""

    def __init__(self, docx_path: str):
        self.docx_path = docx_path

    def extract_tables(self, include_nested: bool = True) -> List[Tuple[int, int, List[List[str]]]]:
        """
        从DOCX文件中提取所有表格，包括嵌套表格。

        Args:
            include_nested: 是否提取嵌套在单元格内的子表格

        Returns:
            列表，每个元素为元组: (1, 表格索引, 表格数据)
        """
        doc = Document(self.docx_path)
        tables = []
        table_idx = 0

        body = doc.element.body
        # 使用XML遍历找到所有表格（包括嵌套的）
        all_tbl_elements = body.findall(f'.//{W_NS}tbl')

        for tbl_element in all_tbl_elements:
            # 如果不包含嵌套表格，跳过非顶层表格
            if not include_nested:
                parent = tbl_element.getparent()
                is_top_level = True
                while parent is not None:
                    if parent.tag == f'{W_NS}tbl':
                        is_top_level = False
                        break
                    parent = parent.getparent()
                if not is_top_level:
                    continue

            grid = _extract_table_grid_from_element(tbl_element)
            if grid and any(any(cell for cell in row) for row in grid):
                # 跳过页眉表格（识别特征：包含Document Name/Document Type/Document No等元数据）
                if self._is_header_table(grid):
                    continue
                tables.append((1, table_idx, grid))
                table_idx += 1

        return tables

    def _is_header_table(self, grid: List[List[str]]) -> bool:
        """
        检测是否为页眉表格。

        页眉表格的特征：包含典型的文档元数据关键词，如
        Document Name, Document Type, Document No, Revision, Page No 等。
        """
        if not grid:
            return False

        header_keywords = [
            'Document Type', 'Document No', 'Document Name',
            'Document Release Status', 'Page No', 'Volume No'
        ]

        # 检查表格中是否包含多个页眉关键词
        all_text = ' '.join(cell for row in grid for cell in row)
        match_count = sum(1 for kw in header_keywords if kw in all_text)

        # 如果匹配3个或以上关键词，认为是页眉表格
        return match_count >= 3

    def save_as_single_csv(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str
    ):
        """将所有表格保存为一个CSV文件"""
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            for page_num, table_idx, table_data in tables:
                writer.writerow([f'--- 表格{table_idx + 1} ---'])
                for row in table_data:
                    writer.writerow(row)
                writer.writerow([])

    def save_as_zip(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str,
        base_name: str = 'table'
    ):
        """将每个表格保存为单独的CSV文件，打包成ZIP"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for page_num, table_idx, table_data in tables:
                csv_filename = f'{base_name}_表格{table_idx + 1}.csv'
                csv_buffer = io.StringIO()
                writer = csv.writer(csv_buffer)
                for row in table_data:
                    writer.writerow(row)
                zf.writestr(csv_filename, csv_buffer.getvalue())


class PdfTableExtractor:
    """PDF表格提取器 - 基于pdf2docx + python-docx两步转换"""

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path

    def _convert_pdf_to_docx(
        self,
        output_docx_path: str,
        start_page: Optional[int] = None,
        end_page: Optional[int] = None,
        progress_callback=None
    ):
        """将PDF转换为DOCX（使用pdf2docx）"""
        cv = Converter(self.pdf_path)
        convert_start = (start_page - 1) if start_page else 0
        convert_end = end_page if end_page else None

        if progress_callback:
            def _page_progress(page, total):
                progress_callback(page, total, 0)

            if hasattr(cv, 'set_progress_callback'):
                cv.set_progress_callback(_page_progress)

        cv.convert(output_docx_path, start=convert_start, end=convert_end)
        cv.close()
        del cv
        gc.collect()

    def extract_tables(
        self,
        start_page: Optional[int] = None,
        end_page: Optional[int] = None,
        progress_callback=None
    ) -> List[Tuple[int, int, List[List[str]]]]:
        """
        提取PDF中的所有表格。

        流程: PDF → DOCX (pdf2docx) → 表格提取 (python-docx)

        Args:
            start_page: 起始页码（1-indexed，包含）
            end_page: 结束页码（1-indexed，包含）
            progress_callback: 回调函数(current_page, total_pages, tables_found)

        Returns:
            列表，每个元素为元组: (页码, 表格索引, 表格数据)
        """
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
            tmp_docx_path = tmp.name

        try:
            # Step 1: PDF → DOCX
            if progress_callback:
                progress_callback(0, 1, 0)

            self._convert_pdf_to_docx(
                tmp_docx_path,
                start_page=start_page,
                end_page=end_page,
                progress_callback=progress_callback
            )

            # Step 2: DOCX → 表格数据（复用 DocxTableExtractor）
            extractor = DocxTableExtractor(tmp_docx_path)
            tables = extractor.extract_tables()
            del extractor
            gc.collect()

            if progress_callback:
                progress_callback(1, 1, len(tables))

            return tables

        finally:
            if os.path.exists(tmp_docx_path):
                os.unlink(tmp_docx_path)

    def save_as_single_csv(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str
    ):
        """将所有表格保存为一个CSV文件"""
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            for page_num, table_idx, table_data in tables:
                writer.writerow([f'--- 表格{table_idx + 1} ---'])
                for row in table_data:
                    writer.writerow(row)
                writer.writerow([])

    def save_as_zip(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str,
        base_name: str = 'table'
    ):
        """将每个表格保存为单独的CSV文件，打包成ZIP"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for page_num, table_idx, table_data in tables:
                csv_filename = f'{base_name}_表格{table_idx + 1}.csv'
                csv_buffer = io.StringIO()
                writer = csv.writer(csv_buffer)
                for row in table_data:
                    writer.writerow(row)
                zf.writestr(csv_filename, csv_buffer.getvalue())
