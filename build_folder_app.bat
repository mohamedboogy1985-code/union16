@echo off
title Build SmartAccounting Folder App
cd /d "%~dp0"

echo =======================================
echo Building SmartAccounting folder app (onedir)...
echo =======================================

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller --noconfirm --clean SmartAccounting_folder.spec

echo.
echo =======================================
echo Done.
echo App folder path: dist\SmartAccounting\
echo =======================================
pause
