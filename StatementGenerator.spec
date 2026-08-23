# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller specification for Statement Generator.

Build on Windows:
    pyinstaller StatementGenerator.spec

The assets folder is intentionally kept OUTSIDE the executable.
The final distribution should be:

    StatementGenerator/
        StatementGenerator.exe
        assets/
            GHR.png
            CED.png
            EBIZ.png (or CPlus.png)
            calibri.ttf
            calibri-bold.ttf
            calibri-italic.ttf
            calibri-bold-italic.ttf
"""

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules


PROJECT_DIR = Path(SPECPATH).resolve()
APP_NAME = "StatementGenerator"


# ---------------------------------------------------------------------------
# Hidden imports
# ---------------------------------------------------------------------------

hiddenimports = [
    "openpyxl",
]

# Pandas/reportlab dependencies are normally detected automatically.
# Collecting reportlab submodules makes the build more robust.
hiddenimports += collect_submodules("reportlab")


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

a = Analysis(
    [str(PROJECT_DIR / "app_tkinter.py")],
    pathex=[str(PROJECT_DIR)],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)


# ---------------------------------------------------------------------------
# Python bytecode archive
# ---------------------------------------------------------------------------

pyz = PYZ(
    a.pure,
    a.zipped_data,
)


# ---------------------------------------------------------------------------
# Windows executable
# ---------------------------------------------------------------------------

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)
