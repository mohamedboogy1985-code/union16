# -*- mode: python ; coding: utf-8 -*-
# spec للبناء كملف EXE واحد (onefile)
# قاعدة البيانات (accounting.db) تُنشأ تلقائياً بجوار ملف التنفيذ أول تشغيل، لذا لا تُضمَّن هنا.

datas = [("sample_import.xlsx", "."), ("sample_payroll.xlsx", ".")]

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
    a.binaries,
    a.datas,
    [],
    name='SmartAccounting',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
