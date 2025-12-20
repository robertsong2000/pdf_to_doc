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

def fix_pdf2docx_compatibility():
    """修复pdf2docx在Docker中的兼容性问题"""
    os.environ['DISPLAY'] = ':99'
    os.environ['PDF2DOCV_SKIP_CHECK'] = '1'

def convert_pdf_to_docx(task_id, pdf_path, output_path, status_file):
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
            cv.convert(output_path, start=0, end=None)
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
                    cv.convert(output_path, start=0, end=None,
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

        # 检查输出文件是否创建成功
        if os.path.exists(output_path):
            update_status(status_file, {
                'status': 'completed',
                'progress': 100,
                'message': 'Conversion completed successfully!',
                'step': 'completed',
                'output_file': os.path.basename(output_path)
            })
        else:
            raise Exception("Output file was not created")

    except Exception as e:
        update_status(status_file, {
            'status': 'error',
            'message': f'Conversion failed: {str(e)}',
            'step': 'error',
            'error': str(e)
        })
        
        # 清理文件
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
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
    if len(sys.argv) != 5:
        print("Usage: python conversion_worker.py <task_id> <pdf_path> <output_path> <status_file>")
        sys.exit(1)
    
    task_id = sys.argv[1]
    pdf_path = sys.argv[2]
    output_path = sys.argv[3]
    status_file = sys.argv[4]
    
    convert_pdf_to_docx(task_id, pdf_path, output_path, status_file)