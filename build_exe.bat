@echo off
title Build SmartAccounting EXE
cd /d "%~dp0"

echo =======================================
echo Building SmartAccounting EXE (onefile)...
echo =======================================

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller --noconfirm --clean SmartAccounting.spec

echo.
echo =======================================
echo Done.
echo EXE path: dist\SmartAccounting.exe
echo =======================================
pause
