#!/usr/bin/env python3
"""
Setup validation script for pbi-to-exec-deck
Checks that Python and all required dependencies are installed.
"""

import sys
import subprocess

def check_python_version():
    """Check if Python version is 3.8+"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"[X] Python 3.8+ required (you have {version.major}.{version.minor}.{version.micro})")
        print("    Download from: https://www.python.org/downloads/")
        return False
    print(f"[OK] Python {version.major}.{version.minor}.{version.micro} detected")
    return True

def check_package(package_name, import_name=None):
    """Check if a Python package is installed"""
    if import_name is None:
        import_name = package_name

    try:
        __import__(import_name)
        print(f"[OK] {package_name} installed")
        return True
    except ImportError:
        print(f"[X] {package_name} not installed")
        return False

def main():
    print("=" * 70)
    print("Power BI to Executive Deck - Setup Validation")
    print("=" * 70)
    print()

    all_good = True

    # Check Python version
    print("Checking Python version...")
    if not check_python_version():
        all_good = False
    print()

    # Check required packages
    print("Checking required packages...")
    packages = [
        ("python-pptx", "pptx"),
        ("Pillow", "PIL"),
        ("PyMuPDF", "fitz"),
        ("markitdown", "markitdown"),
    ]

    missing_packages = []
    for package_name, import_name in packages:
        if not check_package(package_name, import_name):
            missing_packages.append(package_name)
            all_good = False

    print()
    print("=" * 70)

    if all_good:
        print("[OK] All dependencies installed! You're ready to go.")
        print()
        print("Next step: Run the converter")
        print("  python convert_dashboard_claude.py --source your_dashboard.pdf")
    else:
        print("[X] Setup incomplete. Please install missing dependencies:")
        print()
        if missing_packages:
            print("Run this command to install all dependencies:")
            print()
            print("  pip install -r requirements.txt")
            print()
            print("Or install individually:")
            for package in missing_packages:
                print(f"  pip install {package}")

    print("=" * 70)

    return 0 if all_good else 1

if __name__ == "__main__":
    sys.exit(main())
