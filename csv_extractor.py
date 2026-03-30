"""
PDF表格提取模块 - 使用pdfplumber提取PDF中的表格并保存为CSV
"""
import csv
import os
import io
import zipfile
from typing import List, Optional, Tuple


class PdfTableExtractor:
    """PDF表格提取器 - 基于pdfplumber"""

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path

    def extract_tables(
        self,
        start_page: Optional[int] = None,
        end_page: Optional[int] = None,
        progress_callback=None
    ) -> List[Tuple[int, int, List[List[str]]]]:
        """
        提取PDF中的所有表格。

        Args:
            start_page: 起始页码（1-indexed，包含）
            end_page: 结束页码（1-indexed，包含）
            progress_callback: 回调函数(current_page, total_pages, tables_found)

        Returns:
            列表，每个元素为元组: (页码, 表格索引, 表格数据)
            表格数据为二维列表，每个元素是单元格字符串
        """
        import pdfplumber

        tables = []
        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)

            # 页码范围（1-indexed 转 0-indexed）
            page_start = (start_page - 1) if start_page else 0
            page_end = end_page if end_page else total_pages

            pages_to_process = pdf.pages[page_start:page_end]

            for i, page in enumerate(pages_to_process):
                actual_page_num = page_start + i + 1  # 1-indexed

                page_tables = page.extract_tables()
                for j, table in enumerate(page_tables):
                    if table:
                        cleaned_table = [
                            [cell if cell is not None else '' for cell in row]
                            for row in table
                        ]
                        tables.append((actual_page_num, j, cleaned_table))

                if progress_callback:
                    progress_callback(actual_page_num, total_pages, len(tables))

        return tables

    def save_as_single_csv(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str
    ):
        """
        将所有表格保存为一个CSV文件，表格之间用分隔行标注页码和表格编号。
        """
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            for page_num, table_idx, table_data in tables:
                writer.writerow([f'--- 第{page_num}页 表格{table_idx + 1} ---'])
                for row in table_data:
                    writer.writerow(row)
                writer.writerow([])

    def save_as_zip(
        self,
        tables: List[Tuple[int, int, List[List[str]]]],
        output_path: str,
        base_name: str = 'table'
    ):
        """
        将每个表格保存为单独的CSV文件，打包成ZIP。
        """
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for page_num, table_idx, table_data in tables:
                csv_filename = f'{base_name}_第{page_num}页_表格{table_idx + 1}.csv'
                csv_buffer = io.StringIO()
                writer = csv.writer(csv_buffer)
                for row in table_data:
                    writer.writerow(row)
                zf.writestr(csv_filename, csv_buffer.getvalue())
