# -*- mode: python ; coding: utf-8 -*-
# spec للبناء كمجلد تطبيق (onedir) - الأكثر استقراراً على Windows
import os

datas = []
for f in ["accounting.db", "sample_import.xlsx", "sample_payroll.xlsx"]:
    if os.path.exists(f):
        datas.append((f, "."))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'openpyxl',
        'reportlab',
        'arabic_reshaper',
        'bidi',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SmartAccounting',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SmartAccounting',
)
