#!/usr/bin/env python3
"""
Flask Web Application for PDF to DOCX Conversion

This web application provides a simple interface for converting PDF files to DOCX format
using the pdf2docx library.

Requirements:
    pip install flask pdf2docx
"""

import os
import uuid
import argparse
from flask import Flask, request, jsonify, send_file, render_template, send_from_directory
from werkzeug.utils import secure_filename
import threading
import multiprocessing
import subprocess
import time
import json
from pathlib import Path
from pdf2docx import Converter

app = Flask(__name__)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 80 * 1024 * 1024  # 80MB max file size
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Change this in production

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Allowed extensions
ALLOWED_EXTENSIONS = {'pdf'}

# Global dictionary to track conversion status
conversion_status = {}

# Global dictionary to track frontend heartbeat status
frontend_heartbeat = {}

# Global dictionary to track conversion processes
conversion_processes = {}

def monitor_conversion_process(task_id):
    """监控转换进程并更新状态"""
    try:
        process_info = conversion_processes.get(task_id)
        if not process_info:
            return
        
        process = process_info['process']
        status_file = process_info['status_file']
        
        # 监控进程直到完成
        while process.poll() is None:
            # 读取状态文件并更新全局状态
            try:
                with open(status_file, 'r') as f:
                    status = json.load(f)
                    conversion_status[task_id] = status
            except:
                pass
            
            # 等待一段时间再检查
            time.sleep(1)
        
        # 进程已完成，读取最终状态
        try:
            with open(status_file, 'r') as f:
                final_status = json.load(f)
                conversion_status[task_id] = final_status
        except:
            # 如果无法读取状态文件，检查进程退出码
            if process.returncode != 0:
                conversion_status[task_id] = {
                    'status': 'error',
                    'message': f'Process exited with code {process.returncode}',
                    'step': 'error',
                    'error': f'Process exited with code {process.returncode}'
                }
        
        # 清理资源
        if task_id in conversion_processes:
            del conversion_processes[task_id]
        
        # 删除状态文件
        try:
            os.remove(status_file)
        except:
            pass
            
    except Exception as e:
        print(f"[ERROR] Error monitoring process for task {task_id}: {e}")
        conversion_status[task_id] = {
            'status': 'error',
            'message': f'Monitoring error: {str(e)}',
            'step': 'error',
            'error': str(e)
        }

def allowed_file(filename):
    """Check if the file has an allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def is_frontend_active(task_id, timeout_seconds=10):
    """Check if frontend is still active based on heartbeat"""
    if task_id not in frontend_heartbeat:
        return False
    
    last_heartbeat = frontend_heartbeat[task_id]
    current_time = time.time()
    return (current_time - last_heartbeat) <= timeout_seconds

def fix_pdf2docx_compatibility():
    """Fix pdf2docx compatibility issues in Docker"""
    import os
    os.environ['OPENCV_IO_ENABLE_OPENEXR'] = '1'
    # Set up virtual display for headless operations
    os.environ['DISPLAY'] = ':99'
    # Disable problematic pdf2docx features
    os.environ['PDF2DOCV_SKIP_CHECK'] = '1'

def convert_pdf_to_docx_task(task_id, pdf_path, output_path):
    """Background task for PDF conversion"""
    start_time = time.time()
    
    # Initialize heartbeat for this task
    frontend_heartbeat[task_id] = time.time()
    
    try:
        conversion_status[task_id] = {
            'status': 'converting',
            'progress': 0,
            'message': 'Starting conversion...',
            'step': 'initialization',
            'error': None
        }

        # Check if frontend is still active
        if not is_frontend_active(task_id):
            raise Exception("Frontend disconnected, cancelling conversion")
            
        # Check if task was cancelled
        if conversion_status[task_id].get('status') == 'cancelled':
            raise Exception("Task was cancelled by user")

        # Step 1: Upload complete
        conversion_status[task_id]['progress'] = 5
        conversion_status[task_id]['message'] = 'File uploaded successfully'
        conversion_status[task_id]['step'] = 'upload_complete'

        # Check if frontend is still active
        if not is_frontend_active(task_id):
            raise Exception("Frontend disconnected, cancelling conversion")

        # Step 2: Initialize converter
        conversion_status[task_id]['progress'] = 10
        conversion_status[task_id]['message'] = 'Initializing PDF converter...'
        conversion_status[task_id]['step'] = 'initializing'

        # Step 3: Opening document
        conversion_status[task_id]['progress'] = 20
        conversion_status[task_id]['message'] = 'Opening PDF document...'
        conversion_status[task_id]['step'] = 'opening_document'

        # Create converter
        cv = Converter(pdf_path)

        # Step 4: Analyzing document structure
        conversion_status[task_id]['progress'] = 30
        conversion_status[task_id]['message'] = 'Analyzing document structure and layout...'
        conversion_status[task_id]['step'] = 'analyzing_document'

        # Fix compatibility before conversion
        fix_pdf2docx_compatibility()

        # Step 5: Extracting content
        conversion_status[task_id]['progress'] = 40
        conversion_status[task_id]['message'] = 'Extracting text, tables, and images...'
        conversion_status[task_id]['step'] = 'extracting_content'

        # Try pdf2docx conversion with enhanced error handling
        print(f"Start to convert {pdf_path} to {output_path}")

        # Check if frontend is still active before starting conversion
        if not is_frontend_active(task_id):
            raise Exception("Frontend disconnected, cancelling conversion")

        try:
            # Step 6: Converting elements
            conversion_status[task_id]['progress'] = 55
            conversion_status[task_id]['message'] = 'Converting PDF elements to Word format...'
            conversion_status[task_id]['step'] = 'converting_elements'

            # Step 7: Preserving formatting
            conversion_status[task_id]['progress'] = 70
            conversion_status[task_id]['message'] = 'Preserving formatting and layout...'
            conversion_status[task_id]['step'] = 'preserving_formatting'

            # Step 8: Generating Word document
            conversion_status[task_id]['progress'] = 80
            conversion_status[task_id]['message'] = 'Generating Word document structure...'
            conversion_status[task_id]['step'] = 'generating_docx'

            # Perform conversion with progress tracking
            def progress_callback(page, total):
                # Check if task was cancelled
                if conversion_status[task_id].get('status') == 'cancelled':
                    print(f"[DEBUG] Task {task_id} was cancelled in progress callback at page {page}")
                    raise Exception("Task was cancelled by user")
                
                # Check if frontend is still active
                if not is_frontend_active(task_id):
                    raise Exception("Frontend disconnected, cancelling conversion")
                    
                # Update progress during conversion (50% to 80% range)
                if total > 0:
                    page_progress = 50 + int((page / total) * 30)  # Map page progress to 50-80% range
                    conversion_status[task_id]['progress'] = page_progress
                    conversion_status[task_id]['message'] = f'处理页面 {page}/{total}...'
                    conversion_status[task_id]['step'] = 'processing_pages'
                    conversion_status[task_id]['current_page'] = page
                    conversion_status[task_id]['total_pages'] = total
                    
                    # Calculate estimated time remaining
                    elapsed_time = time.time() - start_time
                    if page > 0:
                        avg_time_per_page = elapsed_time / page
                        remaining_pages = total - page
                        eta_seconds = avg_time_per_page * remaining_pages
                        conversion_status[task_id]['eta'] = f"预计剩余时间: {int(eta_seconds)}秒"
            
            # Set progress callback if supported
            if hasattr(cv, 'set_progress_callback'):
                cv.set_progress_callback(progress_callback)
            
            cv.convert(output_path, start=0, end=None)
            cv.close()

            print(f"pdf2docx conversion completed for {pdf_path}")

        except AttributeError as ae:
            print(f"pdf2docx attribute error: {str(ae)}")
            if 'get_area' in str(ae):
                print("Detected known pdf2docx compatibility issue. Trying alternative approach...")
                # Try conversion with different parameters to work around the issue
                try:
                    if 'cv' in locals():
                        cv.close()

                    # 设置环境变量以减少兼容性问题
                    os.environ['PDF2DOCV_SKIP_CHECK'] = '1'

                    # Re-initialize with different parameters
                    cv = Converter(pdf_path, strict=False)

                    # Try conversion with minimal layout analysis and progress tracking
                    def fallback_progress_callback(page, total):
                        # Check if task was cancelled
                        if conversion_status[task_id].get('status') == 'cancelled':
                            print(f"[DEBUG] Task {task_id} was cancelled in fallback progress callback at page {page}")
                            raise Exception("Task was cancelled by user")
                        
                        # Update progress during conversion (55% to 80% range)
                        if total > 0:
                            page_progress = 55 + int((page / total) * 25)  # Map page progress to 55-80% range
                            conversion_status[task_id]['progress'] = page_progress
                            conversion_status[task_id]['message'] = f'处理页面 {page}/{total} (备用模式)...'
                            conversion_status[task_id]['step'] = 'processing_pages_fallback'
                            conversion_status[task_id]['current_page'] = page
                            conversion_status[task_id]['total_pages'] = total
                            
                            # Calculate estimated time remaining
                            elapsed_time = time.time() - start_time
                            if page > 0:
                                avg_time_per_page = elapsed_time / page
                                remaining_pages = total - page
                                eta_seconds = avg_time_per_page * remaining_pages
                                conversion_status[task_id]['eta'] = f"预计剩余时间: {int(eta_seconds)}秒"
                    
                    # Set progress callback if supported
                    if hasattr(cv, 'set_progress_callback'):
                        cv.set_progress_callback(fallback_progress_callback)
                    
                    cv.convert(output_path, start=0, end=None,
                              multi_processing=False,
                              debug=False,
                              keep_layout=True)
                    cv.close()
                    print(f"Alternative pdf2docx conversion completed for {pdf_path}")

                except Exception as fallback_error:
                    print(f"Fallback conversion also failed: {str(fallback_error)}")
                    raise Exception(f"PDF conversion failed due to compatibility issue: {str(ae)}")
            else:
                raise Exception(f"PDF conversion failed: {str(ae)}")

        except Exception as e:
            print(f"pdf2docx conversion failed: {str(e)}")
            print(f"Error details: {type(e).__name__}")

            # Try to close converter if it was opened
            try:
                if 'cv' in locals():
                    cv.close()
            except:
                pass

            # Re-raise the exception to trigger error handling
            raise Exception(f"PDF conversion failed: {str(e)}")

        # Step 9: Finalizing
        conversion_status[task_id]['progress'] = 90
        conversion_status[task_id]['message'] = 'Finalizing output document...'
        conversion_status[task_id]['step'] = 'finalizing'

        # Check if output file was created
        if os.path.exists(output_path):
            conversion_status[task_id]['status'] = 'completed'
            conversion_status[task_id]['progress'] = 100
            conversion_status[task_id]['message'] = 'Conversion completed successfully!'
            conversion_status[task_id]['step'] = 'completed'
            conversion_status[task_id]['output_file'] = os.path.basename(output_path)
        else:
            raise Exception("Output file was not created")
            
        # Clean up heartbeat record
        if task_id in frontend_heartbeat:
            del frontend_heartbeat[task_id]

    except Exception as e:
        conversion_status[task_id]['status'] = 'error'
        conversion_status[task_id]['message'] = f'Conversion failed: {str(e)}'
        conversion_status[task_id]['step'] = 'error'
        conversion_status[task_id]['error'] = str(e)
        
        # Clean up heartbeat record
        if task_id in frontend_heartbeat:
            del frontend_heartbeat[task_id]

        # Clean up files on error
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(output_path):
                os.remove(output_path)
        except:
            pass

@app.route('/')
def index():
    """Serve the main page"""
    return render_template('index.html')

@app.route('/api/health')
def health_check():
    """Health check endpoint"""
    try:
        import sys
        import os
        import pdf2docx

        return jsonify({
            'status': 'healthy',
            'python_version': sys.version,
            'platform': os.uname().sysname if hasattr(os, 'uname') else 'unknown',
            'pdf2docx_available': True,
            'working_directory': os.getcwd(),
            'uploads_dir': app.config['UPLOAD_FOLDER'],
            'outputs_dir': app.config['OUTPUT_FOLDER'],
            'directories_exist': {
                'uploads': os.path.exists(app.config['UPLOAD_FOLDER']),
                'outputs': os.path.exists(app.config['OUTPUT_FOLDER']),
                'templates': os.path.exists('templates')
            }
        })
    except Exception as e:
        return jsonify({
            'status': 'unhealthy',
            'error': str(e)
        }), 500

@app.route('/api/convert', methods=['POST'])
def convert_pdf():
    """Convert PDF to DOCX"""
    try:
        print(f"[DEBUG] Convert request received from {request.remote_addr}")

        # Safely check request files without causing errors
        try:
            files_keys = list(request.files.keys()) if request.files else []
            form_data = dict(request.form) if request.form else {}
            print(f"[DEBUG] Request files: {files_keys}")
            print(f"[DEBUG] Request form: {form_data}")
        except Exception as debug_error:
            print(f"[DEBUG] Error accessing request data: {debug_error}")
            return jsonify({'error': 'Invalid request format'}), 400

        # Check if file is in request
        if not request.files or 'file' not in request.files:
            print("[DEBUG] No file in request")
            return jsonify({'error': 'No file provided'}), 400

        file = request.files['file']

        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file(file.filename):
            return jsonify({'error': 'Only PDF files are allowed'}), 400

        # Generate unique task ID
        task_id = str(uuid.uuid4())

        # Secure filename and create paths
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_{filename}")
        output_filename = filename.rsplit('.', 1)[0] + '.docx'
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], f"{task_id}_{output_filename}")

        # Save uploaded file
        file.save(pdf_path)

        # Start conversion in background process
        print(f"[DEBUG] Starting conversion process for task {task_id}")
        try:
            # Create status file for process communication
            status_file = os.path.join(app.config['UPLOAD_FOLDER'], f"{task_id}_status.json")
            
            # Initialize status file
            with open(status_file, 'w') as f:
                json.dump({
                    'status': 'converting',
                    'progress': 0,
                    'message': 'Starting conversion...',
                    'step': 'initialization',
                    'error': None
                }, f)
            
            # Start conversion worker process
            process = subprocess.Popen([
                'python', 'conversion_worker.py',
                task_id,
                pdf_path,
                output_path,
                status_file
            ])
            
            # Store process reference
            conversion_processes[task_id] = {
                'process': process,
                'status_file': status_file,
                'pdf_path': pdf_path,
                'output_path': output_path
            }
            
            # Start a thread to monitor the process and update status
            monitor_thread = threading.Thread(
                target=monitor_conversion_process,
                args=(task_id,)
            )
            monitor_thread.daemon = True
            monitor_thread.start()
            
            print(f"[DEBUG] Conversion process started successfully with PID {process.pid}")
        except Exception as process_error:
            print(f"[ERROR] Failed to start conversion process: {process_error}")
            import traceback
            print(f"[ERROR] Process error traceback: {traceback.format_exc()}")
            return jsonify({'error': f'Failed to start conversion: {str(process_error)}'}), 500

        return jsonify({
            'task_id': task_id,
            'message': 'File uploaded successfully. Conversion started.',
            'filename': filename
        })

    except Exception as e:
        import traceback
        error_details = f"Upload failed: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(f"[ERROR] {error_details}")
        return jsonify({'error': f'Upload failed: {str(e)}', 'details': traceback.format_exc()}), 500

@app.route('/api/status/<task_id>')
def get_status(task_id):
    """Get conversion status"""
    if task_id not in conversion_status:
        return jsonify({'error': 'Task not found'}), 404

    # If task is running in a process, try to get the latest status from the status file
    if task_id in conversion_processes:
        process_info = conversion_processes[task_id]
        status_file = process_info['status_file']
        
        try:
            with open(status_file, 'r') as f:
                file_status = json.load(f)
                # Update global status with latest from file
                conversion_status[task_id].update(file_status)
        except:
            pass  # If we can't read the file, use the cached status

    return jsonify(conversion_status[task_id])

@app.route('/api/download/<task_id>')
def download_file(task_id):
    """Download converted file"""
    if task_id not in conversion_status:
        return jsonify({'error': 'Task not found'}), 404

    status = conversion_status[task_id]

    if status['status'] != 'completed':
        return jsonify({'error': 'File not ready for download'}), 400

    output_filename = status.get('output_file')
    if not output_filename:
        return jsonify({'error': 'Output file not found'}), 404

    # Find the file in output folder
    for filename in os.listdir(app.config['OUTPUT_FOLDER']):
        if filename.startswith(task_id):
            return send_file(
                os.path.join(app.config['OUTPUT_FOLDER'], filename),
                as_attachment=True,
                download_name=output_filename
            )

    return jsonify({'error': 'File not found on server'}), 404


@app.route('/api/cleanup/<task_id>', methods=['DELETE'])
def cleanup_files(task_id):
    """Clean up uploaded and output files"""
    try:
        # Remove uploaded file
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            if filename.startswith(task_id):
                os.remove(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        # Remove output file
        for filename in os.listdir(app.config['OUTPUT_FOLDER']):
            if filename.startswith(task_id):
                os.remove(os.path.join(app.config['OUTPUT_FOLDER'], filename))

        # Remove from status
        if task_id in conversion_status:
            del conversion_status[task_id]

        return jsonify({'message': 'Files cleaned up successfully'})

    except Exception as e:
        return jsonify({'error': f'Cleanup failed: {str(e)}'}), 500

@app.route('/api/heartbeat/<task_id>', methods=['POST'])
def heartbeat(task_id):
    """Receive heartbeat from frontend to indicate it's still active"""
    frontend_heartbeat[task_id] = time.time()
    return jsonify({'status': 'received'})

@app.route('/api/cancel/<task_id>', methods=['POST'])
def cancel_task(task_id):
    """Cancel a running conversion task"""
    if task_id not in conversion_status:
        return jsonify({'error': 'Task not found'}), 404
    
    # Mark task as cancelled
    conversion_status[task_id]['status'] = 'cancelled'
    conversion_status[task_id]['message'] = 'Task cancelled by user'
    conversion_status[task_id]['step'] = 'cancelled'
    
    # Remove heartbeat record
    if task_id in frontend_heartbeat:
        del frontend_heartbeat[task_id]
    
    # If task is running in a process, terminate it
    if task_id in conversion_processes:
        process_info = conversion_processes[task_id]
        process = process_info['process']
        status_file = process_info['status_file']
        
        # Update status file to notify the process
        try:
            with open(status_file, 'w') as f:
                json.dump({
                    'status': 'cancelled',
                    'message': 'Task cancelled by user',
                    'step': 'cancelled'
                }, f)
        except:
            pass
        
        # Give the process a moment to exit gracefully
        time.sleep(1)
        
        # If process is still running, terminate it
        if process.poll() is None:
            print(f"[DEBUG] Terminating process {process.pid} for task {task_id}")
            try:
                process.terminate()
                
                # If terminate doesn't work, force kill
                time.sleep(2)
                if process.poll() is None:
                    print(f"[DEBUG] Force killing process {process.pid} for task {task_id}")
                    process.kill()
            except Exception as e:
                print(f"[ERROR] Failed to terminate process for task {task_id}: {e}")
        
        # Clean up process info
        if task_id in conversion_processes:
            del conversion_processes[task_id]
    
    return jsonify({'status': 'cancelled'})

# ============================================================================
# DOCX Merge API Routes
# ============================================================================

from merger import DocxMerger

# Global dictionary to track merge status
merge_status = {}

def allowed_docx_file(filename):
    """Check if the file has a .docx extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'docx'

@app.route('/api/merge/health', methods=['GET'])
def merge_health_check():
    """Health check endpoint for merge service"""
    try:
        return jsonify({
            'status': 'ok',
            'message': 'DOCX合并服务运行中'
        })
    except Exception as e:
        return jsonify({
            'status': 'error',
            'error': str(e)
        }), 500

@app.route('/api/merge', methods=['POST'])
def merge_docx():
    """Merge multiple DOCX files into one"""
    try:
        # Check if files are in request
        if not request.files or 'files' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        files = request.files.getlist('files')

        if not files or len(files) == 0:
            return jsonify({'error': '没有选择文件'}), 400

        # Generate unique session ID
        session_id = str(uuid.uuid4())

        # Create session directory
        session_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'merge_{session_id}')
        os.makedirs(session_dir, exist_ok=True)

        # Validate and save files
        valid_files = []
        for file in files:
            if file and file.filename and allowed_docx_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(session_dir, filename)
                file.save(filepath)
                valid_files.append(filepath)

        if not valid_files:
            return jsonify({'error': '没有有效的docx文件'}), 400

        if len(valid_files) < 2:
            return jsonify({'error': '请至少上传2个docx文件进行合并'}), 400

        # Get parameters
        page_break = request.form.get('page_break', 'true').lower() == 'true'
        output_name = request.form.get('output_name', 'merged.docx')

        if not output_name.endswith('.docx'):
            output_name += '.docx'

        # Perform merge
        output_file = os.path.join(app.config['OUTPUT_FOLDER'], f'merge_{session_id}_{output_name}')

        try:
            # Use first file as base document
            first_file = valid_files[0]
            merger = DocxMerger(output_file, first_file=first_file)
            merger.merge_documents(valid_files, add_page_break=page_break)
            merger.save()

            return jsonify({
                'success': True,
                'message': f'成功合并 {len(valid_files)} 个文件',
                'download_url': f'/api/merge/download/merge_{session_id}_{output_name}',
                'filename': output_name
            })

        except Exception as e:
            return jsonify({'error': f'合并失败: {str(e)}'}), 500

    except Exception as e:
        import traceback
        error_details = f"Merge failed: {str(e)}\n\nTraceback:\n{traceback.format_exc()}"
        print(f"[ERROR] {error_details}")
        return jsonify({'error': f'合并失败: {str(e)}'}), 500

@app.route('/api/merge/download/<filename>', methods=['GET'])
def download_merged(filename):
    """Download merged DOCX file"""
    # Security check: ensure filename starts with 'merge_'
    if not filename.startswith('merge_'):
        return jsonify({'error': '无效的文件名'}), 400

    filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)

    if not os.path.exists(filepath):
        return jsonify({'error': '文件不存在或已过期'}), 404

    # Remove session_id prefix from download name
    download_name = filename
    # Find the last underscore before the actual filename
    parts = filename.split('_', 2)
    if len(parts) >= 3:
        download_name = parts[2]

    return send_file(
        filepath,
        as_attachment=True,
        download_name=download_name
    )

@app.route('/api/merge/cleanup/<session_id>', methods=['DELETE'])
def cleanup_merge_files(session_id):
    """Clean up merged files"""
    try:
        # Remove session directory
        session_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'merge_{session_id}')
        if os.path.exists(session_dir):
            import shutil
            shutil.rmtree(session_dir)

        # Remove output file
        for filename in os.listdir(app.config['OUTPUT_FOLDER']):
            if filename.startswith(f'merge_{session_id}'):
                os.remove(os.path.join(app.config['OUTPUT_FOLDER'], filename))

        return jsonify({'message': '文件清理成功'})

    except Exception as e:
        return jsonify({'error': f'清理失败: {str(e)}'}), 500

# ============================================================================
# Error Handlers
# ============================================================================

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 80MB'}), 413


def parse_args():
    """解析命令行参数"""
    parser = argparse.ArgumentParser(
        description='文档工具集 - PDF转DOCX & DOCX合并服务',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 使用默认端口5000启动
  python app.py

  # 使用端口8080启动
  python app.py -p 8080

  # 使用端口3000并关闭调试模式
  python app.py -p 3000 --no-debug

  # 指定host和port
  python app.py -p 5000 --host 127.0.0.1
        """
    )

    parser.add_argument(
        '-p', '--port',
        type=int,
        default=5000,
        help='服务端口 (默认: 5000)'
    )

    parser.add_argument(
        '--host',
        type=str,
        default='0.0.0.0',
        help='服务主机 (默认: 0.0.0.0)'
    )

    parser.add_argument(
        '--no-debug',
        action='store_true',
        help='关闭调试模式'
    )

    return parser.parse_args()


if __name__ == '__main__':
    # 解析命令行参数
    args = parse_args()

    # Clean up old files on startup
    try:
        for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                # Remove files older than 1 hour
                if os.path.getctime(file_path) < time.time() - 3600:
                    os.remove(file_path)
    except:
        pass

    print("=" * 60)
    print("文档工具集启动中...")
    print(f"  - PDF转DOCX转换")
    print(f"  - DOCX文件合并")
    print(f"访问地址: http://{args.host}:{args.port}")
    print("=" * 60)

    app.run(debug=not args.no_debug, host=args.host, port=args.port)