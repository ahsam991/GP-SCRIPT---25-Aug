#!/usr/bin/env python
"""POS Daily Report - Dependency Check & Launcher.

Run this file (python run.py) and it will:
  1. Check whether every required Python library is installed on this PC.
  2. Auto-install the latest version of any library that is missing.
  3. Start the GUI (pos_daily_report_automation_gui.py).

Extra: pass --check-only to only check/install without launching the GUI.
"""

import importlib
import os
import subprocess
import sys

# (pip package name, import module name)
REQUIRED_PACKAGES = [
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("SQLAlchemy", "sqlalchemy"),
    ("psycopg2-binary", "psycopg2"),
]

# Needed only for sending emails (Google Drive upload) and .xls recipient files.
OPTIONAL_PACKAGES = [
    ("pydrive2", "pydrive2"),
    ("xlrd", "xlrd"),
]


def is_installed(import_name):
    """Return True if the module can be imported."""
    try:
        importlib.import_module(import_name)
        return True
    except Exception:
        return False


def pip_install(package_name):
    """Install the latest version of a package using the current interpreter."""
    print(f"  -> Installing {package_name} (latest)...")
    attempts = [
        [sys.executable, "-m", "pip", "install", package_name],
        [sys.executable, "-m", "pip", "install", "--user", package_name],
    ]
    for cmd in attempts:
        try:
            subprocess.check_call(cmd)
            return True
        except subprocess.CalledProcessError:
            continue
        except Exception as e:
            print(f"     (error: {e})")
            continue
    return False


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    gui_script = os.path.join(script_dir, "pos_daily_report_automation_gui.py")

    print("=" * 62)
    print("POS Daily Report - Dependency Check & Launcher")
    print("=" * 62)
    print(f"Python interpreter : {sys.executable}")
    print()

    # ---- Phase 1: check ----
    missing_required = []
    for pip_name, import_name in REQUIRED_PACKAGES:
        if is_installed(import_name):
            print(f"[OK]      {pip_name}")
        else:
            print(f"[MISSING] {pip_name}")
            missing_required.append(pip_name)

    missing_optional = []
    for pip_name, import_name in OPTIONAL_PACKAGES:
        if is_installed(import_name):
            print(f"[OK]      {pip_name}")
        else:
            print(f"[MISSING] {pip_name} (optional - needed for sending emails / .xls files)")
            missing_optional.append(pip_name)

    # ---- Phase 2: auto-install ----
    if missing_required:
        print("\nAuto-installing missing required libraries...")
        for pkg in missing_required:
            if not pip_install(pkg):
                print(f"[FAILED]  Could not install {pkg}. Please check your internet "
                      f"connection and try again.")
                input("\nPress Enter to exit...")
                sys.exit(1)

    if missing_optional:
        print("\nAuto-installing missing optional libraries...")
        for pkg in missing_optional:
            if not pip_install(pkg):
                print(f"[WARN]    Could not install optional library {pkg}. "
                      f"Sending emails will not work until it is installed.")

    # ---- Phase 3: verify ----
    still_missing = [p for _, p in REQUIRED_PACKAGES if not is_installed(p)]
    if still_missing:
        print("\nSome required libraries are still missing:")
        for m in still_missing:
            print(f"  - {m}")
        print("Please install them manually: python -m pip install " + " ".join(still_missing))
        input("\nPress Enter to exit...")
        sys.exit(1)

    print("\nAll required libraries are ready.")

    if "--check-only" in sys.argv:
        print("Dependency check completed. (GUI not launched because --check-only was used.)")
        sys.exit(0)

    print("Starting the GUI...")
    print()
    rc = subprocess.call([sys.executable, gui_script])
    sys.exit(rc)


if __name__ == "__main__":
    main()
