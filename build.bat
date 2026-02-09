@echo off
REM ============================================================
REM  Build script for PLC Excel PDF Converter (16-Sheet)
REM  Creates a standalone .exe using PyInstaller
REM ============================================================

echo ============================================
echo  Building PLC Excel PDF Converter (EXE)
echo ============================================

REM Install dependencies
pip install -r requirements.txt

REM Build single-file executable
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "PLCExcelPdfConverter" ^
    --add-data "README.md;." ^
    --hidden-import pymcprotocol ^
    --hidden-import config ^
    --hidden-import plc_comm ^
    --hidden-import core ^
    --noconfirm ^
    excel_pdf_processor.py

echo.
echo Build complete!
echo Output: dist\PLCExcelPdfConverter.exe
echo.
pause
