#!/usr/bin/env python3
"""
Setup validation script for pbi-to-exec-deck
Checks that Python and all required dependencies are installed.
"""

import sys
import subprocess
import argparse

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



def _build_package_list(profile: str):
    packages = [
        ("python-pptx", "pptx"),
        ("Pillow", "PIL"),
        ("markitdown", "markitdown"),
    ]

    if profile == "claude":
        packages.append(("PyMuPDF", "fitz"))
    elif profile == "copilot":
        packages.append(("pypdfium2", "pypdfium2"))
        packages.append(("easyocr", "easyocr"))

    return packages


def _check_packages(packages):
    missing = []
    for package_name, import_name in packages:
        if not check_package(package_name, import_name):
            missing.append(package_name)
    return missing


def main():
    ap = argparse.ArgumentParser(
        description="Power BI to Executive Deck - Setup Validation"
    )
    ap.add_argument(
        "--auto-install",
        action="store_true",
        help="Auto-install missing Python dependencies for the selected profile",
    )
    ap.add_argument(
        "--profile",
        choices=["claude", "copilot"],
        default="claude",
        help="Dependency profile to validate (claude or copilot)",
    )
    args = ap.parse_args()

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
    packages = _build_package_list(args.profile)
    missing_packages = _check_packages(packages)
    if missing_packages:
        all_good = False

    print()
    print("=" * 70)

    if all_good:
        print("[OK] All dependencies installed! You're ready to go.")
        print()
        print("Next step: Run the converter")
        print("  python convert_dashboard.py --source your_dashboard.pdf")
    else:
        print("[X] Setup incomplete. Please install missing dependencies:")
        if args.profile == "copilot":
            print("    (pypdfium2 is required for Copilot PDF text extraction)")
        print()
        if missing_packages:
            print("Run this command to install all dependencies:")
            print()
            req_file = "requirements-copilot.txt" if args.profile == "copilot" else "requirements.txt"
            print(f"  pip install -r {req_file}")
            print()
            print("Or install individually:")
            for package in missing_packages:
                print(f"  pip install {package}")

        if args.auto_install and missing_packages:
            print()
            print("Auto-installing missing Python dependencies...")
            try:
                req_file = "requirements-copilot.txt" if args.profile == "copilot" else "requirements.txt"
                subprocess.check_call([
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "-r",
                    req_file,
                ])
                print("[OK] Dependencies installed. Re-checking...")
                print()
                print("Checking required packages...")
                missing_packages = _check_packages(packages)
                if missing_packages:
                    print("[X] Some packages are still missing:")
                    for package in missing_packages:
                        print(f"  {package}")
                    return 1
                print("[OK] All dependencies installed! You're ready to go.")
                return 0
            except subprocess.CalledProcessError as exc:
                print(f"[X] Auto-install failed with exit code {exc.returncode}")

    print("=" * 70)

    return 0 if all_good else 1

if __name__ == "__main__":
    sys.exit(main())
