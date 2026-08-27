# بناء ملف EXE

## الأسهل
- شغّل الملف: `build_exe.bat`
- بعد انتهاء البناء ستجد الملف هنا:
  `dist\SmartAccounting.exe`

## النسخة الأكثر استقراراً
- شغّل الملف: `build_folder_app.bat`
- بعد انتهاء البناء ستجد البرنامج هنا:
  `dist\SmartAccounting\`

## ملاحظات
- استخدم `python -m PyInstaller` وليس `pyinstaller` مباشرة، لأن جهازك سبق وظهر فيه أن المسار ليس مضافاً إلى PATH.
- إذا كان Windows Defender يمنع التشغيل، شغّل البرنامج من مجلد `dist`.
- إذا ظهر خطأ مع Python 3.14، فالأفضل البناء باستخدام Python 3.12.
