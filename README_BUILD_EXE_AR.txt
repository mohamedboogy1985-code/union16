# بناء ملف EXE

## الأسهل
- شغّل الملف: `build_exe.bat`
- بعد انتهاء البناء ستجد الملف هنا:
  `dist\SmartAccounting.exe`

## النسخة الأكثر استقراراً
- شغّل الملف: `build_folder_app.bat`
- بعد انتهاء البناء ستجد البرنامج هنا:
  `dist\SmartAccounting\`

## البناء من GitHub (بدون جهازك)
- افتح المستودع على GitHub → تبويب Actions → «Build Windows EXE» → Run workflow.
- أو ارفع tag باسم `v1.0` مثلاً، وسيُبنى الـ EXE ويُرفق تلقائياً في صفحة Releases.
- نزّل الملف من صفحة Release أو من Artifacts الخاصة بالعملية.

## ملاحظات
- استخدم `python -m PyInstaller` وليس `pyinstaller` مباشرة، لأن جهازك سبق وظهر فيه أن المسار ليس مضافاً إلى PATH.
- سكربتا البناء يستخدمان الآن ملفات spec جاهزة: `SmartAccounting.spec` (ملف واحد) و `SmartAccounting_folder.spec` (مجلد تطبيق).
- إذا كان Windows Defender يمنع التشغيل، شغّل البرنامج من مجلد `dist`.
- إذا ظهر خطأ مع Python 3.14، فالأفضل البناء باستخدام Python 3.12.
