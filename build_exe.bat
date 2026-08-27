@echo off
title Build SmartAccounting EXE
cd /d "%~dp0"

echo =======================================
echo Building SmartAccounting EXE...
echo =======================================

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller --noconfirm --clean --onefile --windowed --name SmartAccounting ^
  --add-data "accounting.db;." ^
  --add-data "sample_import.xlsx;." ^
  main.py

echo.
echo =======================================
echo Done.
echo EXE path: dist\SmartAccounting.exe
echo =======================================
pause
