# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for building the Follow-Up app.

Usage:
  pyinstaller generator.spec

Notes:
- This spec is written to work on both Linux and Windows (when run on each platform).
- PyInstaller does not cross-compile; to build a Windows exe you must run this on Windows.
- It bundles all Python dependencies and includes the project data files.
"""

from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Root of the project (where this spec lives)
project_root = Path(__file__).resolve().parent

# Make sure PyInstaller can find your source packages
pathex = [
    str(project_root),
    str(project_root / "src"),
]

# Include all submodules under the src package (ensures dynamic imports are collected)
hiddenimports = collect_submodules("src")

# Include user-facing config/data files that are loaded at runtime
datas = [
    (str(project_root / "configs.json"), "."),
    (str(project_root / "emails_cc.json"), "."),
    # Include the UI theme file used by the frontend
    (str(project_root / "src" / "frontend" / "theme.qss"), "src/frontend"),
]

# Collect any package data under src (if any exist)
datas += collect_data_files("src")

a = Analysis(
    [str(project_root / "main.py")],
    pathex=pathex,
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="follow-up",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="follow-up",
)
