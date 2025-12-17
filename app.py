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
from flask import Flask, request, jsonify, send_file, render_template, send_from_directory
from werkzeug.utils import secure_filename
import threading
import time
from pathlib import Path
from pdf2docx import Converter

app = Flask(__name__)

# Configuration
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
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

def allowed_file(filename):
    """Check if the file has an allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

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
    try:
        conversion_status[task_id] = {
            'status': 'converting',
            'progress': 0,
            'message': 'Starting conversion...',
            'step': 'initialization',
            'error': None
        }

        # Step 1: Upload complete
        conversion_status[task_id]['progress'] = 5
        conversion_status[task_id]['message'] = 'File uploaded successfully'
        conversion_status[task_id]['step'] = 'upload_complete'

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

            # Perform conversion
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

                    # Try conversion with minimal layout analysis
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

    except Exception as e:
        conversion_status[task_id]['status'] = 'error'
        conversion_status[task_id]['message'] = f'Conversion failed: {str(e)}'
        conversion_status[task_id]['step'] = 'error'
        conversion_status[task_id]['error'] = str(e)

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
    return render_template('index_local.html')

@app.route('/api/convert', methods=['POST'])
def convert_pdf():
    """Convert PDF to DOCX"""
    try:
        # Check if file is in request
        if 'file' not in request.files:
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

        # Start conversion in background
        thread = threading.Thread(
            target=convert_pdf_to_docx_task,
            args=(task_id, pdf_path, output_path)
        )
        thread.daemon = True
        thread.start()

        return jsonify({
            'task_id': task_id,
            'message': 'File uploaded successfully. Conversion started.',
            'filename': filename
        })

    except Exception as e:
        return jsonify({'error': f'Upload failed: {str(e)}'}), 500

@app.route('/api/status/<task_id>')
def get_status(task_id):
    """Get conversion status"""
    if task_id not in conversion_status:
        return jsonify({'error': 'Task not found'}), 404

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

@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File too large. Maximum size is 16MB'}), 413

if __name__ == '__main__':
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

    app.run(debug=True, host='0.0.0.0', port=5000)