#!/usr/bin/env python3
"""
PDF to DOCX Web Application Startup Script

This script provides a convenient way to start the web application with proper
environment setup and dependency checking.
"""

import sys
import os
import subprocess
import importlib

def check_dependencies():
    """Check if all required dependencies are installed"""
    print("🔍 Checking dependencies...")

    required_packages = [
        'flask',
        'werkzeug',
        'pdf2docx',
        'pathlib2',
        'pdfplumber'
    ]

    missing_packages = []

    for package in required_packages:
        try:
            # Convert package name to import name if needed
            import_name = {
                'werkzeug': 'werkzeug',
                'pdf2docx': 'pdf2docx',
                'pathlib2': 'pathlib2'
            }.get(package.lower().replace('-', '_'), package.lower())

            importlib.import_module(import_name)
            print(f"  ✓ {package}")
        except ImportError:
            print(f"  ✗ {package} (missing)")
            missing_packages.append(package)

    if missing_packages:
        print(f"\n❌ Missing dependencies: {', '.join(missing_packages)}")
        print("Please run: pip install -r requirements.txt")
        return False

    print("✅ All dependencies are installed!")
    return True

def create_directories():
    """Create necessary directories"""
    print("📁 Creating directories...")

    directories = [
        'uploads',
        'outputs'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"  ✓ Created {directory}/")
        else:
            print(f"  ✓ {directory}/ exists")

    print("✅ Directories ready!")

def main():
    """Main startup function"""
    print("🚀 Starting PDF to DOCX Web Application")
    print("=" * 50)

    # Parse command line arguments
    import argparse
    parser = argparse.ArgumentParser(description='PDF to DOCX Web Application')
    parser.add_argument('--port', '-p', type=int, default=5500,
                       help='Port to run the server on (default: 5500)')
    parser.add_argument('--host', default='0.0.0.0',
                       help='Host to bind the server to (default: 0.0.0.0)')
    parser.add_argument('--debug', action='store_true', default=True,
                       help='Enable debug mode (default: True)')
    args = parser.parse_args()

    # Check dependencies
    if not check_dependencies():
        sys.exit(1)

    # Create directories
    create_directories()

    # Print startup info
    print(f"\n🌐 Server Information:")
    print(f"  • Local URL: http://localhost:{args.port}")
    print(f"  • Network URL: http://{args.host}:{args.port}")
    print(f"  • Debug Mode: {'ON' if args.debug else 'OFF'}")
    if args.debug:
        print("  • Auto-reload: ON")
    print(f"\n⚠️  Press Ctrl+C to stop the server")
    print("=" * 50)

    try:
        # Start the Flask app
        from app import app
        app.run(debug=args.debug, host=args.host, port=args.port)
    except KeyboardInterrupt:
        print("\n\n👋 Server stopped by user")
    except Exception as e:
        print(f"\n❌ Failed to start server: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()