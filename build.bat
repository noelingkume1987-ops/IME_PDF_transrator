@echo off
REM ============================================================
REM  Build script for Excel to PDF Processor Pro
REM  Creates a standalone .exe using PyInstaller
REM ============================================================

echo ============================================
echo  Building Excel to PDF Processor Pro (EXE)
echo ============================================

REM Install dependencies
pip install -r requirements.txt

REM Build single-file executable
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "ExcelPdfProcessorPro" ^
    --add-data "README.md;." ^
    --noconfirm ^
    excel_pdf_processor.py

echo.
echo Build complete!
echo Output: dist\ExcelPdfProcessorPro.exe
echo.
pause
