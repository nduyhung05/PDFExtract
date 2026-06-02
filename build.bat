@echo off
REM Build Invoice Extractor v3
echo === Cai dat thu vien ===
pip install pdfplumber openpyxl pyinstaller

echo.
echo === Dong goi thanh .exe ===
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "InvoiceExtractor" ^
    --add-data "config.py;." ^
    main.py

echo.
echo === HOAN THANH: dist\InvoiceExtractor.exe ===
pause