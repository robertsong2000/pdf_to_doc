#!/usr/bin/env python3
"""
Test script to check PDF conversion functionality
"""

import os
import sys
import tempfile
from pathlib import Path

# Add app directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_pdf2docx_import():
    """Test if pdf2docx can be imported"""
    try:
        from pdf2docx import Converter
        print("✓ pdf2docx imported successfully")
        return True
    except ImportError as e:
        print(f"✗ Failed to import pdf2docx: {e}")
        return False

def test_opencv_import():
    """Test if OpenCV can be imported"""
    try:
        import cv2
        print(f"✓ OpenCV imported successfully (version: {cv2.__version__})")
        return True
    except ImportError as e:
        print(f"✗ Failed to import OpenCV: {e}")
        return False

def test_conversion():
    """Test PDF conversion with a simple file"""
    try:
        from pdf2docx import Converter

        # Create a simple test PDF path (you need to provide one)
        test_pdf = "/app/uploads/test.pdf"  # You'll need to upload a file first

        if not os.path.exists(test_pdf):
            print(f"⚠ Test PDF file not found at {test_pdf}")
            print("Please upload a PDF file first and try again")
            return False

        # Test conversion
        output_docx = "/app/outputs/test.docx"
        print(f"Converting {test_pdf} to {output_docx}")

        cv = Converter(test_pdf)
        cv.convert(output_docx)
        cv.close()

        if os.path.exists(output_docx):
            print(f"✓ Conversion successful! Output file: {output_docx}")
            print(f"File size: {os.path.getsize(output_docx)} bytes")
            return True
        else:
            print("✗ Conversion failed - no output file created")
            return False

    except Exception as e:
        print(f"✗ Conversion failed with error: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("🔍 Testing PDF to DOCX conversion...")
    print("=" * 50)

    # Test imports
    opencv_ok = test_opencv_import()
    pdf2docx_ok = test_pdf2docx_import()

    if not opencv_ok or not pdf2docx_ok:
        print("\n❌ Basic imports failed - cannot proceed with conversion test")
        return False

    print("\n✅ All imports successful!")

    # Test conversion
    print("\n🔄 Testing conversion...")
    conversion_ok = test_conversion()

    if conversion_ok:
        print("\n🎉 All tests passed!")
        return True
    else:
        print("\n💥 Conversion test failed!")
        return False

if __name__ == "__main__":
    main()