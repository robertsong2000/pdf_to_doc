#!/usr/bin/env python3
"""
PDF转换工作进程脚本
这个脚本在一个独立的进程中运行PDF转换，以便主进程可以终止它
"""

import sys
import os
import json
import time
import tempfile
from pdf2docx import Converter
from pdf2docx.converter import Converter as CVConverter
from docx_header_cleaner import post_process_converted_docx
from pdf_conversion_modes import CONVERSION_MODE_DEFAULT, get_conversion_options
from pdf2docx_fallback import (
    PAGE_FRAME_FALLBACK_OPTION,
    converter_supports_page_frame_fallback,
    detect_page_frame_table_pages,
    page_frame_fallback_options,
    repair_page_frame_inner_tables,
    replace_docx_pages,
)

def fix_pdf2docx_compatibility():
    """修复pdf2docx在Docker中的兼容性问题"""
    os.environ['DISPLAY'] = ':99'
    os.environ['PDF2DOCV_SKIP_CHECK'] = '1'

def convert_pdf_to_docx(
    task_id,
    pdf_path,
    output_path,
    status_file,
    start_page=None,
    end_page=None,
    remove_headers=True,
    replace_oem_info=True,
    conversion_mode=CONVERSION_MODE_DEFAULT,
):
    """执行PDF到DOCX的转换"""
    start_time = time.time()
    
    try:
        # 初始化状态
        update_status(status_file, {
            'status': 'converting',
            'progress': 0,
            'message': 'Starting conversion...',
            'step': 'initialization',
            'error': None
        })

        # Step 1: Upload complete
        update_status(status_file, {
            'progress': 5,
            'message': 'File uploaded successfully',
            'step': 'upload_complete'
        })

        # Step 2: Initialize converter
        update_status(status_file, {
            'progress': 10,
            'message': 'Initializing PDF converter...',
            'step': 'initializing'
        })

        # Step 3: Opening document
        update_status(status_file, {
            'progress': 20,
            'message': 'Opening PDF document...',
            'step': 'opening_document'
        })

        # 修复pdf2docx兼容性问题
        fix_pdf2docx_compatibility()

        # Step 4: Analyzing document structure
        update_status(status_file, {
            'progress': 30,
            'message': 'Analyzing document structure...',
            'step': 'analyzing_structure'
        })

        # Step 5: Preparing conversion
        update_status(status_file, {
            'progress': 40,
            'message': 'Preparing conversion parameters...',
            'step': 'preparing_conversion'
        })

        print(f"Start to convert {pdf_path} to {output_path}")

        # Calculate page range for pdf2docx (0-indexed)
        convert_start = (start_page - 1) if start_page else 0
        convert_end = end_page if end_page else None
        conversion_options = get_conversion_options(conversion_mode)

        if start_page or end_page:
            page_info = f" (pages {start_page or 'start'} to {end_page or 'end'})"
            update_status(status_file, {
                'progress': 42,
                'message': f'指定转换页面范围: {page_info}',
                'step': 'preparing_conversion'
            })
            print(f"Converting page range: start={convert_start}, end={convert_end}")

        # Step 6: Converting elements
        update_status(status_file, {
            'progress': 55,
            'message': 'Converting PDF elements to Word format...',
            'step': 'converting_elements'
        })

        # Step 7: Preserving formatting
        update_status(status_file, {
            'progress': 60,
            'message': 'Preserving original formatting...',
            'step': 'preserving_formatting'
        })

        # Step 8: Processing content
        update_status(status_file, {
            'progress': 65,
            'message': 'Processing document content...',
            'step': 'processing_content'
        })

        def run_pdf2docx_conversion(
            target_path,
            options,
            progress_mode='primary',
            page_start=convert_start,
            page_end=convert_end,
        ):
            # 初始化转换器
            cv = Converter(pdf_path)

            # 定义进度回调
            def progress_callback(page, total):
                # 检查是否被取消
                status = read_status(status_file)
                if status and status.get('status') == 'cancelled':
                    print(f"[DEBUG] Task {task_id} was cancelled in progress callback at page {page}")
                    raise Exception("Task was cancelled by user")
                
                # 更新进度
                if total > 0:
                    if progress_mode == 'page_frame_fallback':
                        page_progress = 82 + int((page / total) * 6)
                        message = f'重新处理页面 {page}/{total} (整页表格 fallback)...'
                        step = 'page_frame_table_fallback'
                    else:
                        page_progress = 50 + int((page / total) * 30)  # Map page progress to 50-80% range
                        message = f'处理页面 {page}/{total}...'
                        step = 'processing_pages'
                    update_status(status_file, {
                        'progress': page_progress,
                        'message': message,
                        'step': step,
                        'current_page': page,
                        'total_pages': total,
                        'eta': calculate_eta(start_time, page, total)
                    })
            
            # 设置进度回调
            if hasattr(cv, 'set_progress_callback'):
                cv.set_progress_callback(progress_callback)
            
            try:
                # 执行转换
                cv.convert(
                    target_path,
                    start=page_start,
                    end=page_end,
                    **options,
                )
            finally:
                cv.close()

        # 执行转换
        try:
            run_pdf2docx_conversion(output_path, conversion_options)

        except AttributeError as ae:
            print(f"pdf2docx attribute error: {str(ae)}")
            if 'get_area' in str(ae):
                print("Detected known pdf2docx compatibility issue. Trying alternative approach...")
                try:
                    if 'cv' in locals():
                        cv.close()
                    
                    # 重新初始化转换器
                    cv = CVConverter(pdf_path, strict=False)
                    
                    # 定义备用进度回调
                    def fallback_progress_callback(page, total):
                        # 检查是否被取消
                        status = read_status(status_file)
                        if status and status.get('status') == 'cancelled':
                            print(f"[DEBUG] Task {task_id} was cancelled in fallback progress callback at page {page}")
                            raise Exception("Task was cancelled by user")
                        
                        # 更新进度
                        if total > 0:
                            page_progress = 55 + int((page / total) * 25)  # Map page progress to 55-80% range
                            update_status(status_file, {
                                'progress': page_progress,
                                'message': f'处理页面 {page}/{total} (备用模式)...',
                                'step': 'processing_pages_fallback',
                                'current_page': page,
                                'total_pages': total,
                                'eta': calculate_eta(start_time, page, total)
                            })
                    
                    # 设置进度回调
                    if hasattr(cv, 'set_progress_callback'):
                        cv.set_progress_callback(fallback_progress_callback)
                    
                    # 执行转换
                    cv.convert(output_path, start=convert_start, end=convert_end,
                              multi_processing=False,
                              debug=False,
                              keep_layout=True,
                              **conversion_options)
                    cv.close()
                    
                except Exception as fallback_error:
                    print(f"Fallback conversion also failed: {str(fallback_error)}")
                    raise Exception(f"PDF conversion failed due to compatibility issue: {str(ae)}")
            else:
                raise Exception(f"PDF conversion failed: {str(ae)}")
        
        except Exception as e:
            print(f"pdf2docx conversion failed: {str(e)}")
            raise Exception(f"PDF conversion failed: {str(e)}")

        page_frame_fallback_applied = False
        page_frame_fallback_reason = None
        page_frame_fallback_pages, retry_reason = detect_page_frame_table_pages(
            output_path,
            conversion_mode,
        )
        if page_frame_fallback_pages:
            page_frame_fallback_reason = retry_reason
            print(f"Detected likely whole-page table output: {retry_reason}")
            if converter_supports_page_frame_fallback(pdf_path):
                update_status(status_file, {
                    'progress': 82,
                    'message': '检测到整页大表格，正在按页使用规格书 fallback 重新转换...',
                    'step': 'page_frame_table_fallback',
                    'page_frame_table_fallback': True,
                    'page_frame_table_pages': [page + 1 for page in page_frame_fallback_pages],
                })
                output_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
                fallback_output_paths = {}
                try:
                    fallback_options = page_frame_fallback_options(conversion_options)
                    repaired_inner_tables = 0
                    for page_index in page_frame_fallback_pages:
                        pdf_page_start = convert_start + page_index
                        pdf_page_end = pdf_page_start + 1
                        fd, fallback_output_path = tempfile.mkstemp(
                            suffix='.docx',
                            prefix=f'{task_id}_page_frame_fallback_p{page_index + 1}_',
                            dir=output_dir,
                        )
                        os.close(fd)
                        fallback_output_paths[page_index] = fallback_output_path
                        run_pdf2docx_conversion(
                            fallback_output_path,
                            fallback_options,
                            progress_mode='page_frame_fallback',
                            page_start=pdf_page_start,
                            page_end=pdf_page_end,
                        )
                        repaired_inner_tables += repair_page_frame_inner_tables(
                            fallback_output_path,
                            pdf_path,
                            pdf_page_start,
                        )

                    replace_docx_pages(output_path, fallback_output_paths)
                    page_frame_fallback_applied = True
                    print(
                        "Page-frame table fallback conversion applied to pages: "
                        f"{[page + 1 for page in page_frame_fallback_pages]} "
                        f"(repaired inner tables: {repaired_inner_tables})"
                    )
                except Exception as fallback_error:
                    print(f"Page-frame table fallback failed; keeping first output: {fallback_error}")
                finally:
                    try:
                        for fallback_output_path in fallback_output_paths.values():
                            if os.path.exists(fallback_output_path):
                                os.remove(fallback_output_path)
                    except Exception:
                        pass
            else:
                print(
                    "Installed pdf2docx does not support "
                    f"{PAGE_FRAME_FALLBACK_OPTION}; "
                    "keeping first output."
                )

        # Step 9: Finalizing
        update_status(status_file, {
            'progress': 90,
            'message': 'Finalizing output document...',
            'step': 'finalizing'
        })

        post_process_warning = None
        if os.path.exists(output_path):
            try:
                (
                    removed_headers,
                    repaired_safety_fields,
                    replaced_references,
                ) = post_process_converted_docx(
                    output_path,
                    pdf_path=pdf_path,
                    start_page=start_page,
                    end_page=end_page,
                    remove_headers=remove_headers,
                    replace_oem_info=replace_oem_info,
                )
                if removed_headers or repaired_safety_fields or replaced_references:
                    update_status(status_file, {
                        'progress': 95,
                        'message': (
                            f'已删除 {removed_headers} 行重复页头，'
                            f'修复 {repaired_safety_fields} 处安全字段，'
                            f'替换 {replaced_references} 处 OEM 信息...'
                        ),
                        'step': 'cleaning_headers',
                        'removed_headers': removed_headers,
                        'repaired_safety_fields': repaired_safety_fields,
                        'replaced_oem_references': replaced_references
                    })
            except Exception as post_process_error:
                post_process_warning = str(post_process_error)
                print(f"Post-processing failed but conversion output will be kept: {post_process_warning}")
                update_status(status_file, {
                    'progress': 95,
                    'message': f'DOCX已生成，后处理失败但文件可下载: {post_process_warning}',
                    'step': 'post_process_warning',
                    'warning': post_process_warning
                })

        # 检查输出文件是否创建成功
        if os.path.exists(output_path):
            completed_status = {
                'status': 'completed',
                'progress': 100,
                'message': 'Conversion completed successfully!',
                'step': 'completed',
                'output_file': os.path.basename(output_path)
            }
            if post_process_warning:
                completed_status['warning'] = post_process_warning
                completed_status['message'] = 'Conversion completed with post-processing warning.'
            if page_frame_fallback_reason:
                completed_status['page_frame_table_detected'] = True
                completed_status['page_frame_table_reason'] = page_frame_fallback_reason
                completed_status['page_frame_table_fallback_applied'] = page_frame_fallback_applied
            update_status(status_file, completed_status)
        else:
            raise Exception("Output file was not created")

    except Exception as e:
        update_status(status_file, {
            'status': 'error',
            'message': f'Conversion failed: {str(e)}',
            'step': 'error',
            'error': str(e)
        })
        
        # Keep the uploaded PDF for troubleshooting/retry. Only remove partial output.
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except:
            pass

def update_status(status_file, data):
    """更新状态文件"""
    try:
        # 读取现有状态
        status = read_status(status_file) or {}
        
        # 更新状态
        status.update(data)
        
        # 写入状态文件
        with open(status_file, 'w') as f:
            json.dump(status, f)
    except Exception as e:
        print(f"Error updating status: {e}")

def read_status(status_file):
    """读取状态文件"""
    try:
        with open(status_file, 'r') as f:
            return json.load(f)
    except:
        return None

def calculate_eta(start_time, current_page, total_pages):
    """计算预计剩余时间"""
    if current_page <= 0:
        return ""
    
    elapsed_time = time.time() - start_time
    avg_time_per_page = elapsed_time / current_page
    remaining_pages = total_pages - current_page
    eta_seconds = avg_time_per_page * remaining_pages
    return f"预计剩余时间: {int(eta_seconds)}秒"

if __name__ == "__main__":
    if len(sys.argv) < 5:
        print("Usage: python conversion_worker.py <task_id> <pdf_path> <output_path> <status_file> [start_page] [end_page] [remove_headers] [replace_oem_info] [conversion_mode]")
        sys.exit(1)

    task_id = sys.argv[1]
    pdf_path = sys.argv[2]
    output_path = sys.argv[3]
    status_file = sys.argv[4]

    start_page = int(sys.argv[5]) if len(sys.argv) > 5 and sys.argv[5] else None
    end_page = int(sys.argv[6]) if len(sys.argv) > 6 and sys.argv[6] else None
    remove_headers = len(sys.argv) <= 7 or sys.argv[7].lower() != 'false'
    replace_oem_info = len(sys.argv) <= 8 or sys.argv[8].lower() != 'false'
    conversion_mode = sys.argv[9] if len(sys.argv) > 9 and sys.argv[9] else CONVERSION_MODE_DEFAULT

    convert_pdf_to_docx(
        task_id,
        pdf_path,
        output_path,
        status_file,
        start_page,
        end_page,
        remove_headers,
        replace_oem_info,
        conversion_mode,
    )
