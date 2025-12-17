#!/usr/bin/env python3
"""
PDF to DOC Converter Script

This script converts PDF files to DOCX format using the pdf2docx library.
It supports converting single files or batch processing multiple files.

Requirements:
    pip install pdf2docx

Usage:
    python pdf_to_doc_converter.py input.pdf output.docx
    python pdf_to_doc_converter.py input.pdf                    # Auto-generates output name
    python pdf_to_doc_converter.py --batch *.pdf                # Batch convert multiple PDFs
"""

import sys
import os
import argparse
import glob
from pathlib import Path
from pdf2docx import Converter


def convert_pdf_to_docx(pdf_path, docx_path=None):
    """
    Convert a single PDF file to DOCX format.

    Args:
        pdf_path (str): Path to the input PDF file
        docx_path (str, optional): Path for the output DOCX file

    Returns:
        str: Path to the created DOCX file
    """
    try:
        pdf_path = Path(pdf_path)

        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        if not pdf_path.suffix.lower() == '.pdf':
            raise ValueError(f"File must have .pdf extension: {pdf_path}")

        # Generate output filename if not provided
        if docx_path is None:
            docx_path = pdf_path.with_suffix('.docx')
        else:
            docx_path = Path(docx_path)
            if not docx_path.suffix.lower() == '.docx':
                docx_path = docx_path.with_suffix('.docx')

        print(f"Converting: {pdf_path.name} -> {docx_path.name}")

        # Create converter and perform conversion
        cv = Converter(str(pdf_path))
        cv.convert(str(docx_path), start=0, end=None)
        cv.close()

        print(f"✓ Successfully converted: {docx_path}")
        return str(docx_path)

    except Exception as e:
        print(f"✗ Error converting {pdf_path}: {str(e)}")
        return None


def batch_convert_pdf_to_docx(pdf_paths, output_dir=None):
    """
    Convert multiple PDF files to DOCX format.

    Args:
        pdf_paths (list): List of PDF file paths
        output_dir (str, optional): Directory for output files

    Returns:
        tuple: (successful_conversions, failed_conversions)
    """
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)

    successful = []
    failed = []

    for pdf_path in pdf_paths:
        try:
            pdf_path = Path(pdf_path)

            if output_dir:
                docx_path = output_dir / pdf_path.with_suffix('.docx').name
            else:
                docx_path = pdf_path.with_suffix('.docx')

            result = convert_pdf_to_docx(pdf_path, docx_path)

            if result:
                successful.append(result)
            else:
                failed.append(str(pdf_path))

        except Exception as e:
            print(f"✗ Error processing {pdf_path}: {str(e)}")
            failed.append(str(pdf_path))

    return successful, failed


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF files to DOCX format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s file.pdf                          # Convert single file
  %(prog)s file.pdf output.docx              # Convert with custom output name
  %(prog)s --batch *.pdf                     # Batch convert multiple PDFs
  %(prog)s --batch *.pdf --output-dir docs/  # Batch convert to specific directory
        """
    )

    parser.add_argument('input', nargs='?', help='Input PDF file or pattern (when using --batch)')
    parser.add_argument('output', nargs='?', help='Output DOCX file (optional)')
    parser.add_argument('--batch', action='store_true', help='Batch convert multiple PDF files')
    parser.add_argument('--output-dir', '-o', help='Output directory for converted files (batch mode)')

    args = parser.parse_args()

    if not args.input:
        parser.error("Input file or pattern is required")

    try:
        if args.batch:
            # Handle batch conversion
            if '*' in args.input or '?' in args.input:
                # Handle glob patterns
                pdf_files = glob.glob(args.input)
            elif os.path.isdir(args.input):
                # Handle directory
                pdf_files = glob.glob(os.path.join(args.input, '*.pdf'))
            else:
                # Handle space-separated list
                pdf_files = args.input.split()

            if not pdf_files:
                print("No PDF files found matching the pattern.")
                return 1

            print(f"Found {len(pdf_files)} PDF files to convert...")
            successful, failed = batch_convert_pdf_to_docx(pdf_files, args.output_dir)

            print(f"\nConversion Summary:")
            print(f"✓ Successful: {len(successful)}")
            print(f"✗ Failed: {len(failed)}")

            if failed:
                print(f"Failed files:")
                for file in failed:
                    print(f"  - {file}")

        else:
            # Handle single file conversion
            result = convert_pdf_to_docx(args.input, args.output)
            if not result:
                return 1

    except KeyboardInterrupt:
        print("\nConversion interrupted by user.")
        return 1
    except Exception as e:
        print(f"Error: {str(e)}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())