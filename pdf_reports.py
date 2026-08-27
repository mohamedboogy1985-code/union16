"""
وحدة طباعة التقارير والمستندات PDF (عربي RTL).

تستعمل هذه الوحدة في كل طباعات النظام:
- تقارير جدولية عامة (build_table_pdf)
- سندات: سند قيد / إذن صرف / فاتورة (build_voucher_pdf)
- مذكرة التسوية البنكية (build_settlement_pdf)

تعتمد على مكتبة reportlab مع إعادة تشكيل النص العربي
(arabic_reshaper + python-bidi) وخط عربي من خطوط النظام.
"""

import os
from datetime import datetime

PDF_AVAILABLE = False
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.units import cm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    import arabic_reshaper
    from bidi.algorithm import get_display
    PDF_AVAILABLE = True
except Exception:
    pass

_FONT_REGISTERED = None
_FONT_READY = False


def pdf_available():
    return PDF_AVAILABLE


def rtl_text(text):
    """يعيد تشكيل النص العربي ليعرض بشكل صحيح داخل PDF."""
    text = str(text or "")
    if PDF_AVAILABLE:
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


def pdf_font_name():
    """يسجل خطاً يدعم العربية من خطوط النظام ويعيد اسمه (أو Helvetica كبديل)."""
    global _FONT_REGISTERED, _FONT_READY
    if _FONT_READY:
        return _FONT_REGISTERED
    if not PDF_AVAILABLE:
        _FONT_READY = True
        _FONT_REGISTERED = "Helvetica"
        return _FONT_REGISTERED

    windir = os.environ.get("WINDIR", "C:/Windows")
    local_fonts = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
    candidates = [
        os.path.join(local_fonts, "Amiri-Regular.ttf"),
        os.path.join(local_fonts, "Cairo-Regular.ttf"),
        os.path.join(windir, "Fonts", "arial.ttf"),
        os.path.join(windir, "Fonts", "arialuni.ttf"),
        os.path.join(windir, "Fonts", "Tahoma.ttf"),
        os.path.join(windir, "Fonts", "times.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("ArabicUI", path))
                _FONT_READY = True
                _FONT_REGISTERED = "ArabicUI"
                return _FONT_REGISTERED
            except Exception:
                continue
    _FONT_READY = True
    _FONT_REGISTERED = "Helvetica"
    return _FONT_REGISTERED


def _base_style(font_name, size, leading, align):
    return ParagraphStyle(
        "style_ar",
        fontName=font_name,
        fontSize=size,
        leading=leading,
        alignment=align,
    )


def build_table_pdf(path, title, subtitle="", headers=None, rows=None,
                    total_row=None, col_widths=None, landscape_page=False):
    """
    ينشئ تقريراً جدولياً PDF بعنوان عربي RTL.
    headers: عناوين الأعمدة | rows: صفوف البيانات | total_row: صف الإجماليات (اختياري)
    """
    if not PDF_AVAILABLE:
        raise RuntimeError("مكتبة reportlab غير متوفرة")
    font_name = pdf_font_name()
    page = landscape(A4) if landscape_page else A4
    doc = SimpleDocTemplate(
        path, pagesize=page,
        rightMargin=1.2 * cm, leftMargin=1.2 * cm,
        topMargin=1.4 * cm, bottomMargin=1.4 * cm,
        title=title,
    )
    title_style = _base_style(font_name, 17, 23, TA_CENTER)
    sub_style = _base_style(font_name, 10, 14, TA_CENTER)
    head_style = _base_style(font_name, 9, 12, TA_CENTER)
    cell_style = _base_style(font_name, 8.5, 12, TA_RIGHT)

    story = [Paragraph(rtl_text(title), title_style)]
    if subtitle:
        story.append(Spacer(1, 4))
        story.append(Paragraph(rtl_text(subtitle), sub_style))
    story.append(Spacer(1, 10))

    rows = rows or []
    headers = headers or []
    data = [[Paragraph(rtl_text(h), head_style) for h in headers]]
    for r in rows:
        data.append([Paragraph(rtl_text(c), cell_style) for c in r])
    if total_row:
        data.append([Paragraph(rtl_text(c), head_style) for c in total_row])

    usable = page[0] - 2.4 * cm
    if not col_widths:
        n = max(len(headers), 1)
        col_widths = [usable / n] * n
    col_widths = [min(w, usable) for w in col_widths]

    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f4c81")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#8EAADB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    if total_row:
        last = len(data) - 1
        commands += [
            ("BACKGROUND", (0, last), (-1, last), colors.HexColor("#DCE6F1")),
            ("FONTNAME", (0, last), (-1, last), font_name),
            ("FONTSIZE", (0, last), (-1, last), 9),
        ]
    tbl.setStyle(TableStyle(commands))
    story.append(tbl)
    story.append(Spacer(1, 14))
    story.append(Paragraph(rtl_text(f"تاريخ الطباعة: {datetime.now().strftime('%Y-%m-%d %H:%M')}"), sub_style))
    doc.build(story)


def build_voucher_pdf(path, doc_title, head_fields, lines, total_debit, total_credit, note=""):
    """
    ينشئ مستنداً رسمياً (سند قيد / إذن صرف / فاتورة).
    head_fields: قائمة أزواج (التسمية، القيمة) تظهر أعلى المستند
    lines: قائمة صفوف (الحساب، مدين، دائن، البيان)
    """
    if not PDF_AVAILABLE:
        raise RuntimeError("مكتبة reportlab غير متوفرة")
    font_name = pdf_font_name()
    doc = SimpleDocTemplate(
        path, pagesize=A4,
        rightMargin=1.6 * cm, leftMargin=1.6 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
        title=doc_title,
    )
    title_style = _base_style(font_name, 20, 26, TA_CENTER)
    meta_style = _base_style(font_name, 10, 14, TA_RIGHT)
    head_style = _base_style(font_name, 9.5, 13, TA_CENTER)
    cell_style = _base_style(font_name, 9.5, 13, TA_RIGHT)

    story = []
    story.append(Paragraph(rtl_text(doc_title), title_style))
    story.append(Spacer(1, 0.3 * cm))

    head_rows = []
    half = (len(head_fields) + 1) // 2
    for i in range(half):
        left = head_fields[i] if i < len(head_fields) else ("", "")
        right = head_fields[i + half] if i + half < len(head_fields) else ("", "")
        head_rows.append([
            Paragraph(rtl_text(str(right[1])), cell_style),
            Paragraph(rtl_text(right[0]), head_style),
            Paragraph(rtl_text(str(left[1])), cell_style),
            Paragraph(rtl_text(left[0]), head_style),
        ])
    head_tbl = Table(head_rows, colWidths=[5.2 * cm, 3.3 * cm, 5.2 * cm, 3.3 * cm])
    head_tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#8EAADB")),
        ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#DCE6F1")),
        ("BACKGROUND", (3, 0), (3, -1), colors.HexColor("#DCE6F1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(head_tbl)
    story.append(Spacer(1, 0.4 * cm))

    data = [[Paragraph(rtl_text(h), head_style) for h in ("الحساب", "مدين", "دائن", "البيان")]]
    for account, debit, credit, desc in lines:
        data.append([
            Paragraph(rtl_text(account), cell_style),
            Paragraph(rtl_text(debit), cell_style),
            Paragraph(rtl_text(credit), cell_style),
            Paragraph(rtl_text(desc), cell_style),
        ])
    data.append([
        Paragraph(rtl_text("الإجمالي"), head_style),
        Paragraph(rtl_text(total_debit), head_style),
        Paragraph(rtl_text(total_credit), head_style),
        Paragraph(rtl_text(""), head_style),
    ])
    tbl = Table(data, colWidths=[6.2 * cm, 2.7 * cm, 2.7 * cm, 5.4 * cm], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0f4c81")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#8EAADB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    last = len(data) - 1
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, last), (-1, last), colors.HexColor("#DCE6F1")),
        ("FONTNAME", (0, last), (-1, last), font_name),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.4 * cm))

    if note:
        story.append(Paragraph(rtl_text(f"ملاحظات: {note}"), meta_style))
        story.append(Spacer(1, 0.3 * cm))

    sign_tbl = Table(
        [[Paragraph(rtl_text("المُعِد"), cell_style), Paragraph(rtl_text("المراجعة"), cell_style), Paragraph(rtl_text("الاعتماد"), cell_style)]],
        colWidths=[5.2 * cm, 5.2 * cm, 5.2 * cm], hAlign="CENTER",
    )
    sign_tbl.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.7, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.7, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 26),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 26),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(sign_tbl)
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph(rtl_text(f"تاريخ الطباعة: {datetime.now().strftime('%Y-%m-%d %H:%M')}"), meta_style))
    doc.build(story)


def build_settlement_pdf(path, payload):
    """مذكرة تسوية بنكية (منقولة من main.py لتعمل كوحدة مستقلة)."""
    if not PDF_AVAILABLE:
        raise RuntimeError("مكتبة reportlab غير متوفرة")
    font_name = pdf_font_name()
    doc = SimpleDocTemplate(
        path, pagesize=A4,
        rightMargin=1.6 * cm, leftMargin=1.6 * cm,
        topMargin=1.5 * cm, bottomMargin=1.5 * cm,
    )
    title_style = _base_style(font_name, 18, 24, TA_CENTER)
    meta_style = _base_style(font_name, 10, 14, TA_RIGHT)
    cell_style = _base_style(font_name, 11, 15, TA_RIGHT)

    data_pairs = [
        ("اسم البنك", payload.get("bank_name")),
        ("السنة", payload.get("year")),
        ("الرصيد الافتتاحي", payload.get("opening_balance")),
        ("إجمالي المقبوضات", payload.get("receipts")),
        ("إجمالي المدفوعات", payload.get("payments")),
        ("الرصيد الدفتري آخر الفترة", payload.get("book_balance_end")),
        ("شيكات تحت التحصيل", payload.get("checks_under_collection")),
        ("شيك مسحوب ولم يصرف 1", payload.get("uncashed_check1")),
        ("شيك مسحوب ولم يصرف 2", payload.get("uncashed_check2")),
        ("الرصيد بعد التسوية", payload.get("final_bank_balance")),
        ("رصيد كشف الحساب", payload.get("bank_statement_balance")),
        ("فرق التسوية", payload.get("discrepancy")),
        ("ملاحظات", payload.get("notes") or "-"),
    ]

    story = []
    story.append(Paragraph(rtl_text(payload.get("title", "مذكرة تسوية بنكية")), title_style))
    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph(rtl_text(f"تاريخ الطباعة: {payload.get('generated_at', '')}"), meta_style))
    story.append(Spacer(1, 0.35 * cm))

    table_rows = []
    for label, value in data_pairs:
        table_rows.append([
            Paragraph(rtl_text(str(value)), cell_style),
            Paragraph(rtl_text(label), cell_style),
        ])

    tbl = Table(table_rows, colWidths=[10.5 * cm, 6.0 * cm], hAlign="RIGHT")
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
        ("BACKGROUND", (1, 0), (1, -1), colors.HexColor("#DCE6F1")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("LEADING", (0, 0), (-1, -1), 15),
        ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#8EAADB")),
        ("BOX", (0, 0), (-1, -1), 1.0, colors.HexColor("#5B9BD5")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.45 * cm))
    sign_tbl = Table(
        [[Paragraph(rtl_text("إعداد / مراجعة"), cell_style), Paragraph(rtl_text("اعتماد"), cell_style)]],
        colWidths=[8.0 * cm, 8.0 * cm], hAlign="CENTER",
    )
    sign_tbl.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.7, colors.grey),
        ("INNERGRID", (0, 0), (-1, -1), 0.7, colors.grey),
        ("TOPPADDING", (0, 0), (-1, -1), 24),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 24),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(sign_tbl)
    doc.build(story)
