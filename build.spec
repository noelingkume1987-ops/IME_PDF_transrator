# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Excel to PDF Processor Pro.
Usage:  pyinstaller build.spec
"""

a = Analysis(
    ['excel_pdf_processor.py'],
    pathex=[],
    binaries=[],
    datas=[('README.md', '.')],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ExcelPdfProcessorPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # windowed mode (no console)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
