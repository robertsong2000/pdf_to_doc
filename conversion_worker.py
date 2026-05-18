#!/usr/bin/env python3
"""
PDF转换工作进程脚本
这个脚本在一个独立的进程中运行PDF转换，以便主进程可以终止它
"""

import sys
import os
import json
import time
from pdf2docx import Converter
from pdf2docx.converter import Converter as CVConverter
from docx_header_cleaner import post_process_converted_docx

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

        # 执行转换
        try:
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
                    page_progress = 50 + int((page / total) * 30)  # Map page progress to 50-80% range
                    update_status(status_file, {
                        'progress': page_progress,
                        'message': f'处理页面 {page}/{total}...',
                        'step': 'processing_pages',
                        'current_page': page,
                        'total_pages': total,
                        'eta': calculate_eta(start_time, page, total)
                    })
            
            # 设置进度回调
            if hasattr(cv, 'set_progress_callback'):
                cv.set_progress_callback(progress_callback)
            
            # 执行转换
            cv.convert(output_path, start=convert_start, end=convert_end)
            cv.close()

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
                              keep_layout=True)
                    cv.close()
                    
                except Exception as fallback_error:
                    print(f"Fallback conversion also failed: {str(fallback_error)}")
                    raise Exception(f"PDF conversion failed due to compatibility issue: {str(ae)}")
            else:
                raise Exception(f"PDF conversion failed: {str(ae)}")
        
        except Exception as e:
            print(f"pdf2docx conversion failed: {str(e)}")
            raise Exception(f"PDF conversion failed: {str(e)}")

        # Step 9: Finalizing
        update_status(status_file, {
            'progress': 90,
            'message': 'Finalizing output document...',
            'step': 'finalizing'
        })

        post_process_warning = None
        if os.path.exists(output_path) and (remove_headers or replace_oem_info):
            try:
                removed_headers, replaced_references = post_process_converted_docx(
                    output_path,
                    remove_headers=remove_headers,
                    replace_oem_info=replace_oem_info,
                )
                if removed_headers or replaced_references:
                    update_status(status_file, {
                        'progress': 95,
                        'message': f'已删除 {removed_headers} 行重复页头，替换 {replaced_references} 处 OEM 信息...',
                        'step': 'cleaning_headers',
                        'removed_headers': removed_headers,
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
        print("Usage: python conversion_worker.py <task_id> <pdf_path> <output_path> <status_file> [start_page] [end_page] [remove_headers] [replace_oem_info]")
        sys.exit(1)

    task_id = sys.argv[1]
    pdf_path = sys.argv[2]
    output_path = sys.argv[3]
    status_file = sys.argv[4]

    start_page = int(sys.argv[5]) if len(sys.argv) > 5 and sys.argv[5] else None
    end_page = int(sys.argv[6]) if len(sys.argv) > 6 and sys.argv[6] else None
    remove_headers = len(sys.argv) <= 7 or sys.argv[7].lower() != 'false'
    replace_oem_info = len(sys.argv) <= 8 or sys.argv[8].lower() != 'false'

    convert_pdf_to_docx(
        task_id,
        pdf_path,
        output_path,
        status_file,
        start_page,
        end_page,
        remove_headers,
        replace_oem_info,
    )
