
import os
import shutil
import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from openpyxl import Workbook

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
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
    PDF_AVAILABLE = False

from db import init_db, get_connection, DB_PATH
from importer import import_excel_file
from payroll_importer import import_payroll_excel
from pdf_reports import (
    pdf_available as pdf_lib_available,
    build_table_pdf,
    build_voucher_pdf,
    build_settlement_pdf,
)
from services import (
    dashboard_summary,
    list_entries,
    get_entry_lines,
    debtors_people,
    debtors_person_report,
    revenue_expense_accounts,
    revenue_expense_report,
    revenue_expense_final_summary,
    latest_payroll_import,
    payroll_rows,
    workers_advances_report,
    bank_loans_report,
    smart_vouchers,
    ledger_accounts,
    ledger_for_account,
    trial_balance,
    account_statement_report,
    bank_settlement_report,
    bank_settlement_history,
    save_bank_settlement,
    post_bank_settlement_adjustment,
    cost_centers_list,
    add_cost_center,
    rename_cost_center,
    delete_cost_center,
    cost_center_report,
)


APP_TITLE = "النظام المحاسبي المكتبي المتكامل"

def fmt(v):
    try:
        return f"{float(v or 0):,.2f}"
    except Exception:
        return "0.00"


class EntryEditor(tk.Toplevel):
    def __init__(self, master, entry_id=None):
        super().__init__(master)
        self.master = master
        self.entry_id = entry_id
        self.title("إضافة / تعديل قيد")
        self.geometry("1080x650")
        self.transient(master)
        self.grab_set()
        self.lines = []

        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        self.date_var = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        self.ref_var = tk.StringVar()
        self.desc_var = tk.StringVar()

        ttk.Label(top, text="التاريخ").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.date_var, width=14).pack(side="right", padx=4)
        ttk.Label(top, text="المرجع").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.ref_var, width=18).pack(side="right", padx=4)
        ttk.Label(top, text="البيان").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.desc_var, width=50).pack(side="right", padx=4)

        tools = ttk.Frame(self)
        tools.pack(fill="x", padx=10)
        ttk.Button(tools, text="إضافة سطر", command=self.add_line).pack(side="left", padx=4)
        ttk.Button(tools, text="حذف السطر المحدد", command=self.remove_selected).pack(side="left", padx=4)
        ttk.Button(tools, text="حفظ", command=self.save).pack(side="left", padx=4)

        self.tree = ttk.Treeview(self, columns=("account", "debit", "credit", "desc", "cost"), show="headings", height=14)
        for c, h, w in [
            ("account", "الحساب", 380), ("debit", "مدين", 110), ("credit", "دائن", 110), ("desc", "بيان السطر", 220), ("cost", "مركز التكلفة", 150)
        ]:
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=8)

        editor = ttk.LabelFrame(self, text="تحرير السطر")
        editor.pack(fill="x", padx=10, pady=(0, 10))
        self.line_account = tk.StringVar()
        self.line_debit = tk.StringVar()
        self.line_credit = tk.StringVar()
        self.line_desc = tk.StringVar()
        self.line_cost = tk.StringVar()

        ttk.Label(editor, text="الحساب").pack(side="right", padx=4)
        self.account_combo = ttk.Combobox(editor, textvariable=self.line_account, width=36, values=ledger_accounts())
        self.account_combo.pack(side="right", padx=4)
        ttk.Label(editor, text="مدين").pack(side="right", padx=4)
        ttk.Entry(editor, textvariable=self.line_debit, width=11).pack(side="right", padx=4)
        ttk.Label(editor, text="دائن").pack(side="right", padx=4)
        ttk.Entry(editor, textvariable=self.line_credit, width=11).pack(side="right", padx=4)
        ttk.Label(editor, text="بيان السطر").pack(side="right", padx=4)
        ttk.Entry(editor, textvariable=self.line_desc, width=22).pack(side="right", padx=4)
        ttk.Label(editor, text="مركز التكلفة").pack(side="right", padx=4)
        ttk.Combobox(editor, textvariable=self.line_cost, width=16, values=[""] + [c["name"] for c in cost_centers_list()]).pack(side="right", padx=4)
        ttk.Button(editor, text="تطبيق على السطر المحدد", command=self.apply_line).pack(side="left", padx=4)

        self.status = ttk.Label(self, text="")
        self.status.pack(anchor="e", padx=10, pady=(0, 8))

        self.tree.bind("<<TreeviewSelect>>", self.load_selected)
        if entry_id:
            self.load_entry()
        else:
            self.add_line()
            self.add_line()

    def add_line(self):
        iid = self.tree.insert("", "end", values=("", "0.00", "0.00", "", self.line_cost.get()))
        self.tree.selection_set(iid)
        self.load_selected()

    def remove_selected(self):
        for iid in self.tree.selection():
            self.tree.delete(iid)
        self.update_status()

    def load_selected(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        self.line_account.set(vals[0] if len(vals) > 0 else "")
        self.line_debit.set(vals[1] if len(vals) > 1 else "0.00")
        self.line_credit.set(vals[2] if len(vals) > 2 else "0.00")
        self.line_desc.set(vals[3] if len(vals) > 3 else "")
        self.line_cost.set(vals[4] if len(vals) > 4 else "")

    def apply_line(self):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        self.tree.item(iid, values=(
            self.line_account.get().strip(),
            self.line_debit.get().strip() or "0.00",
            self.line_credit.get().strip() or "0.00",
            self.line_desc.get().strip(),
            self.line_cost.get().strip()
        ))
        self.update_status()

    def update_status(self):
        total_debit = 0.0
        total_credit = 0.0
        for iid in self.tree.get_children():
            vals = self.tree.item(iid, "values")
            total_debit += float(vals[1] or 0)
            total_credit += float(vals[2] or 0)
        self.status.config(text=f"إجمالي المدين: {fmt(total_debit)} | إجمالي الدائن: {fmt(total_credit)}")

    def load_entry(self):
        conn = get_connection()
        cur = conn.cursor()
        head = cur.execute("""
            SELECT id, entry_date, reference, description
            FROM journal_entries WHERE id=?
        """, (self.entry_id,)).fetchone()
        lines = cur.execute("""
            SELECT a.name, jl.debit, jl.credit, jl.line_description, cc.name
            FROM journal_lines jl
            JOIN accounts a ON a.id = jl.account_id
            LEFT JOIN cost_centers cc ON cc.id = jl.cost_center_id
            WHERE jl.entry_id=?
            ORDER BY jl.id
        """, (self.entry_id,)).fetchall()
        conn.close()

        if head:
            self.date_var.set(head["entry_date"] or "")
            self.ref_var.set(head["reference"] or "")
            self.desc_var.set(head["description"] or "")
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for row in lines:
            self.tree.insert("", "end", values=(row[0], f"{float(row[1] or 0):.2f}", f"{float(row[2] or 0):.2f}", row[3] or "", row[4] or ""))
        self.update_status()

    def save(self):
        rows = []
        total_debit = total_credit = 0.0
        for iid in self.tree.get_children():
            account, debit, credit, desc, cost = self.tree.item(iid, "values")
            if not account.strip():
                continue
            d = float(debit or 0)
            c = float(credit or 0)
            rows.append((account.strip(), d, c, desc.strip(), cost.strip()))
            total_debit += d
            total_credit += c

        if not rows:
            messagebox.showerror("خطأ", "أدخل سطور القيد أولاً")
            return
        if round(total_debit, 2) != round(total_credit, 2):
            messagebox.showerror("خطأ", "القيد غير متوازن")
            return

        conn = get_connection()
        cur = conn.cursor()

        def ensure_account(name):
            row = cur.execute("SELECT id FROM accounts WHERE name=?", (name,)).fetchone()
            if row:
                return row[0]
            code = f"U{int(datetime.now().timestamp())%100000}"
            category = "أخرى"
            normal_side = "debit"
            cur.execute("INSERT INTO accounts(code,name,category,normal_side) VALUES (?,?,?,?)",
                        (code, name, category, normal_side))
            return cur.lastrowid

        def resolve_cost_center(name):
            if not name:
                return None
            row = cur.execute("SELECT id FROM cost_centers WHERE name=?", (name,)).fetchone()
            return row[0] if row else None

        if self.entry_id:
            cur.execute("DELETE FROM journal_lines WHERE entry_id=?", (self.entry_id,))
            cur.execute("""
                UPDATE journal_entries
                SET entry_date=?, reference=?, description=?, total_debit=?, total_credit=?
                WHERE id=?
            """, (self.date_var.get(), self.ref_var.get(), self.desc_var.get(), total_debit, total_credit, self.entry_id))
            entry_id = self.entry_id
        else:
            cur.execute("""
                INSERT INTO journal_entries(entry_date, reference, description, total_debit, total_credit, source_row, import_id)
                VALUES (?,?,?,?,?,?,?)
            """, (self.date_var.get(), self.ref_var.get(), self.desc_var.get(), total_debit, total_credit, None, None))
            entry_id = cur.lastrowid

        for account_name, d, c, desc, cost_name in rows:
            account_id = ensure_account(account_name)
            cost_id = resolve_cost_center(cost_name)
            cur.execute("""
                INSERT INTO journal_lines(entry_id, account_id, debit, credit, line_description, cost_center_id)
                VALUES (?,?,?,?,?,?)
            """, (entry_id, account_id, d, c, desc, cost_id))

        conn.commit()
        conn.close()
        self.master.refresh_all()
        messagebox.showinfo("تم", "تم حفظ القيد بنجاح")
        self.destroy()


class AccountingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1450x900")
        self.minsize(1200, 760)
        init_db()
        self._style()
        self._header()
        self._build_ui()
        self.refresh_all()

    def _style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TNotebook.Tab", font=("Tahoma", 10, "bold"), padding=(10, 7))
        style.configure("Treeview.Heading", font=("Tahoma", 10, "bold"))
        style.configure("Treeview", rowheight=25, font=("Tahoma", 10))
        style.configure("Section.TLabel", font=("Tahoma", 12, "bold"))
        style.configure("Header.TLabel", font=("Tahoma", 18, "bold"), background="#0f4c81", foreground="white")
        style.configure("Card.TFrame", background="white")
        style.configure("KPI.TLabel", font=("Tahoma", 13, "bold"), background="white")

    def _header(self):
        top = tk.Frame(self, bg="#0f4c81", height=60)
        top.pack(fill="x")
        top.pack_propagate(False)
        ttk.Label(top, text="نظام محاسبي مكتبي متكامل", style="Header.TLabel").pack(side="right", padx=16, pady=12)
        btns = tk.Frame(top, bg="#0f4c81")
        btns.pack(side="left", padx=16)
        for txt, cmd in [
            ("تحديث", self.refresh_all),
            ("استيراد اليومية", self.import_journal),
            ("استيراد المرتبات", self.import_payroll),
            ("نسخة احتياطية", self.backup_db),
        ]:
            ttk.Button(btns, text=txt, command=cmd).pack(side="left", padx=4)

    def _build_ui(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)

        self.tabs = {}
        for key, title in [
            ("home", "الرئيسية"),
            ("entries", "القيود"),
            ("accounts", "الحسابات"),
            ("debtors", "المدينون والسلف"),
            ("revexp", "الإيرادات والمصروفات"),
            ("cost", "مراكز التكلفة"),
            ("settlements", "التسويات البنكية"),
            ("payroll", "المرتبات"),
            ("reports", "التقارير"),
            ("settings", "الإعدادات"),
        ]:
            frame = ttk.Frame(self.nb)
            self.nb.add(frame, text=title)
            self.tabs[key] = frame

        self._build_home()
        self._build_entries()
        self._build_accounts()
        self._build_debtors()
        self._build_revexp()
        self._build_cost()
        self._build_settlements()
        self._build_payroll()
        self._build_reports()
        self._build_settings()

    def _subnb(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        return nb

    def _screen(self, nb, title):
        f = ttk.Frame(nb)
        nb.add(f, text=title)
        return f

    def _tree(self, parent, columns, headings, widths):
        frame = ttk.Frame(parent)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        ys = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        xs = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        for c, h, w in zip(columns, headings, widths):
            tree.heading(c, text=h)
            tree.column(c, width=w, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return frame, tree

    def _kpi_row(self, parent):
        box = ttk.Frame(parent)
        box.pack(fill="x", pady=(4, 8))
        labels = {}
        for key, title in [("debit", "إجمالي المدين"), ("credit", "إجمالي الدائن"), ("balance", "الرصيد النهائي")]:
            card = ttk.Frame(box, style="Card.TFrame")
            card.pack(side="right", fill="x", expand=True, padx=4)
            ttk.Label(card, text=title).pack(anchor="e", padx=10, pady=(8, 2))
            lbl = ttk.Label(card, text="0.00", style="KPI.TLabel")
            lbl.pack(anchor="e", padx=10, pady=(0, 8))
            labels[key] = lbl
        return labels

    def _fill_tree(self, tree, rows):
        for iid in tree.get_children():
            tree.delete(iid)
        for row in rows:
            tree.insert("", "end", values=row)

    def export_simple_excel(self, rows, headers, title):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=title)
        if not path:
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "البيانات"
        ws.append(headers)
        for r in rows:
            ws.append(list(r))
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(14, min(40, max(len(str(c.value or "")) for c in col) + 2))
        wb.save(path)
        messagebox.showinfo("تم", "تم التصدير بنجاح")


    def _rtl_text(self, text):
        from pdf_reports import rtl_text
        return rtl_text(text)

    def _pdf_font_name(self):
        from pdf_reports import pdf_font_name
        return pdf_font_name()

    def _create_settlement_pdf(self, path, payload):
        build_settlement_pdf(path, payload)

    def import_journal(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            msg = import_excel_file(path)
            self.refresh_all()
            messagebox.showinfo("نجاح", msg)
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def import_payroll(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            msg = import_payroll_excel(path)
            self.refresh_all()
            messagebox.showinfo("نجاح", msg)
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def backup_db(self):
        path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite", "*.db")], initialfile=f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
        if not path:
            return
        shutil.copyfile(DB_PATH, path)
        messagebox.showinfo("تم", "تم حفظ النسخة الاحتياطية")

    def restore_db(self):
        path = filedialog.askopenfilename(filetypes=[("SQLite", "*.db")])
        if not path:
            return
        if not messagebox.askyesno("تأكيد", "سيتم استبدال قاعدة البيانات الحالية. هل تريد المتابعة؟"):
            return
        shutil.copyfile(path, DB_PATH)
        self.refresh_all()
        messagebox.showinfo("تم", "تم استرجاع النسخة الاحتياطية")

    def _build_home(self):
        f = self.tabs["home"]
        self.home_kpis = self._kpi_row(f)
        self.home_top = ttk.Label(f, text="ملخص عام", style="Section.TLabel")
        self.home_top.pack(anchor="e", padx=10, pady=4)

        wrap = ttk.Panedwindow(f, orient="horizontal")
        wrap.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.LabelFrame(wrap, text="أحدث القيود")
        right = ttk.LabelFrame(wrap, text="أكبر الحسابات")
        wrap.add(left, weight=1)
        wrap.add(right, weight=1)

        lf, self.home_recent_tree = self._tree(left, ("date", "ref", "desc", "debit", "credit"),
                                               ("التاريخ", "المرجع", "البيان", "مدين", "دائن"),
                                               (100, 120, 320, 110, 110))
        lf.pack(fill="both", expand=True, padx=6, pady=6)
        rf, self.home_accounts_tree = self._tree(right, ("account", "debit", "credit", "balance"),
                                                 ("الحساب", "مدين", "دائن", "الرصيد"),
                                                 (340, 110, 110, 120))
        rf.pack(fill="both", expand=True, padx=6, pady=6)

    def _build_entries(self):
        nb = self._subnb(self.tabs["entries"])
        self.entries_nb = nb

        # import screen
        imp = self._screen(nb, "استيراد اليومية الأمريكية")
        ttk.Label(imp, text="استيراد ملف اليومية الأمريكية", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(imp, text="اختيار ملف واستيراد", command=self.import_journal).pack(anchor="e", padx=10)

        # journal
        journal = self._screen(nb, "دفتر اليومية")
        top = ttk.Frame(journal); top.pack(fill="x", padx=10, pady=8)
        self.journal_search = tk.StringVar()
        ttk.Button(top, text="تحديث", command=self.refresh_journal).pack(side="left", padx=4)
        ttk.Button(top, text="إضافة قيد", command=lambda: EntryEditor(self)).pack(side="left", padx=4)
        ttk.Button(top, text="تعديل القيد المحدد", command=self.open_selected_entry).pack(side="left", padx=4)
        ttk.Button(top, text="حذف القيد المحدد", command=self.delete_selected_entry).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير", command=self.export_journal).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة سند القيد", command=self.print_entry_voucher).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_journal_pdf).pack(side="left", padx=4)
        ttk.Label(top, text="بحث").pack(side="right", padx=4)
        ent = ttk.Entry(top, textvariable=self.journal_search, width=40); ent.pack(side="right", padx=4)
        ent.bind("<KeyRelease>", lambda e: self.refresh_journal())
        tf, self.journal_tree = self._tree(journal, ("id", "date", "ref", "desc", "debit", "credit"),
                                           ("رقم", "التاريخ", "المرجع", "البيان", "مدين", "دائن"),
                                           (70, 100, 120, 500, 120, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.journal_tree.bind("<<TreeviewSelect>>", lambda e: self.refresh_entry_lines())
        lines_box = ttk.LabelFrame(journal, text="سطور القيد")
        lines_box.pack(fill="both", expand=False, padx=10, pady=(0, 8))
        lf, self.entry_lines_tree = self._tree(lines_box, ("account", "debit", "credit", "desc", "cost"),
                                               ("الحساب", "مدين", "دائن", "بيان السطر", "مركز التكلفة"),
                                               (360, 110, 110, 300, 140))
        lf.pack(fill="both", expand=True, padx=6, pady=6)

        add_screen = self._screen(nb, "إضافة قيد جديد")
        ttk.Label(add_screen, text="إضافة قيد جديد", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(add_screen, text="فتح شاشة الإضافة", command=lambda: EntryEditor(self)).pack(anchor="e", padx=10)

        edit_screen = self._screen(nb, "تعديل قيد")
        ttk.Label(edit_screen, text="اختر قيداً من شاشة دفتر اليومية ثم اضغط تعديل", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(edit_screen, text="تعديل القيد المحدد من دفتر اليومية", command=self.open_selected_entry).pack(anchor="e", padx=10)

        delete_screen = self._screen(nb, "حذف قيد")
        ttk.Label(delete_screen, text="اختر قيداً من شاشة دفتر اليومية ثم اضغط حذف", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(delete_screen, text="حذف القيد المحدد", command=self.delete_selected_entry).pack(anchor="e", padx=10)

        post = self._screen(nb, "ترحيل القيود")
        ttk.Label(post, text="هذه النسخة ترحّل القيود تلقائياً داخل قاعدة البيانات عند الحفظ أو الاستيراد.", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)

    def _build_accounts(self):
        nb = self._subnb(self.tabs["accounts"])

        ledger = self._screen(nb, "الأستاذ العام")
        top = ttk.Frame(ledger); top.pack(fill="x", padx=10, pady=8)
        self.ledger_account_var = tk.StringVar()
        self.ledger_combo = ttk.Combobox(top, textvariable=self.ledger_account_var, width=50)
        self.ledger_combo.pack(side="right", padx=4)
        self.ledger_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_ledger())
        ttk.Label(top, text="الحساب").pack(side="right", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_ledger).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_ledger_pdf).pack(side="left", padx=4)
        tf, self.ledger_tree = self._tree(ledger, ("date", "ref", "desc", "debit", "credit", "balance"),
                                          ("التاريخ", "المرجع", "البيان", "مدين", "دائن", "الرصيد"),
                                          (100, 120, 420, 110, 110, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        trial = self._screen(nb, "ميزان المراجعة")
        ttk.Button(trial, text="تصدير", command=self.export_trial).pack(anchor="w", padx=10, pady=8)
        ttk.Button(trial, text="طباعة PDF", command=self.print_trial_pdf).pack(anchor="w", padx=10)
        tf, self.trial_tree = self._tree(trial, ("account", "debit", "credit", "net_debit", "net_credit"),
                                         ("الحساب", "إجمالي مدين", "إجمالي دائن", "رصيد مدين", "رصيد دائن"),
                                         (440, 120, 120, 120, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        chart = self._screen(nb, "دليل الحسابات")
        ttk.Button(chart, text="تحديث", command=self.refresh_chart).pack(anchor="w", padx=10, pady=8)
        tf, self.chart_tree = self._tree(chart, ("code", "name", "category", "normal"),
                                         ("الكود", "اسم الحساب", "الفئة", "الطبيعة"),
                                         (120, 440, 200, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

    def _build_debtors(self):
        nb = self._subnb(self.tabs["debtors"])

        debt = self._screen(nb, "المدينون المتنوعون")
        top = ttk.Frame(debt); top.pack(fill="x", padx=10, pady=8)
        self.debtor_search = tk.StringVar()
        ttk.Label(top, text="اسم الشخص").pack(side="right", padx=4)
        ent = ttk.Entry(top, textvariable=self.debtor_search, width=35); ent.pack(side="right", padx=4)
        ent.bind("<KeyRelease>", lambda e: self.refresh_debtors_people())
        ttk.Button(top, text="تحديث", command=self.refresh_debtors_people).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير", command=self.export_debtors).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_debtor_pdf).pack(side="left", padx=4)
        body = ttk.Panedwindow(debt, orient="horizontal"); body.pack(fill="both", expand=True, padx=10, pady=8)
        left = ttk.LabelFrame(body, text="الأشخاص"); right = ttk.Frame(body)
        body.add(left, weight=1); body.add(right, weight=4)
        self.debtors_list = tk.Listbox(left, exportselection=False, font=("Tahoma", 10))
        self.debtors_list.pack(fill="both", expand=True, padx=6, pady=6)
        self.debtors_list.bind("<<ListboxSelect>>", lambda e: self.refresh_debtor_statement())
        self.debtor_kpis = self._kpi_row(right)
        tf, self.debtor_tree = self._tree(right, ("date", "ref", "account", "desc", "debit", "credit", "balance"),
                                          ("التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"),
                                          (100, 120, 250, 380, 100, 100, 120))
        tf.pack(fill="both", expand=True, padx=6, pady=6)

        adv = self._screen(nb, "سلف العاملين")
        top = ttk.Frame(adv); top.pack(fill="x", padx=10, pady=8)
        self.adv_search = tk.StringVar()
        ttk.Label(top, text="بحث").pack(side="right", padx=4)
        e = ttk.Entry(top, textvariable=self.adv_search, width=35); e.pack(side="right", padx=4)
        e.bind("<KeyRelease>", lambda ev: self.refresh_advances())
        ttk.Button(top, text="تصدير", command=self.export_advances).pack(side="left", padx=4)
        self.adv_total = ttk.Label(adv, text="إجمالي السلف: 0.00", style="Section.TLabel"); self.adv_total.pack(anchor="e", padx=10)
        tf, self.adv_tree = self._tree(adv, ("no", "name", "month", "amount"),
                                       ("الرقم", "الاسم", "الفترة", "قسط السلف"),
                                       (100, 360, 200, 150))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        bank = self._screen(nb, "قرض البنك")
        top = ttk.Frame(bank); top.pack(fill="x", padx=10, pady=8)
        self.bank_search = tk.StringVar()
        ttk.Label(top, text="بحث").pack(side="right", padx=4)
        e = ttk.Entry(top, textvariable=self.bank_search, width=35); e.pack(side="right", padx=4)
        e.bind("<KeyRelease>", lambda ev: self.refresh_bank())
        ttk.Button(top, text="تصدير", command=self.export_bank).pack(side="left", padx=4)
        self.bank_total = ttk.Label(bank, text="إجمالي قرض البنك: 0.00", style="Section.TLabel"); self.bank_total.pack(anchor="e", padx=10)
        tf, self.bank_tree = self._tree(bank, ("no", "name", "month", "amount"),
                                        ("الرقم", "الاسم", "الفترة", "قسط سلفة البنك"),
                                        (100, 360, 200, 150))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        vouchers = self._screen(nb, "أذون الصرف الذكية")
        ttk.Button(vouchers, text="تصدير", command=self.export_vouchers).pack(anchor="w", padx=10, pady=8)
        ttk.Button(vouchers, text="طباعة أذن الصرف المحدد", command=self.print_selected_voucher).pack(anchor="w", padx=10, pady=8)
        tf, self.vouchers_tree = self._tree(vouchers, ("id", "date", "ref", "desc", "source", "target", "amount"),
                                            ("رقم", "التاريخ", "المرجع", "البيان", "من", "إلى", "المبلغ"),
                                            (70, 100, 120, 320, 240, 240, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

    def _build_revexp(self):
        nb = self._subnb(self.tabs["revexp"])

        det = self._screen(nb, "حساب الإيرادات والمصروفات")
        top = ttk.Frame(det); top.pack(fill="x", padx=10, pady=8)
        self.revexp_account_var = tk.StringVar()
        self.revexp_filter_var = tk.StringVar()
        ttk.Label(top, text="الحساب").pack(side="right", padx=4)
        self.revexp_combo = ttk.Combobox(top, textvariable=self.revexp_account_var, width=50)
        self.revexp_combo.pack(side="right", padx=4)
        self.revexp_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_revexp())
        ttk.Label(top, text="فلتر البيان").pack(side="right", padx=4)
        e = ttk.Entry(top, textvariable=self.revexp_filter_var, width=30); e.pack(side="right", padx=4)
        e.bind("<KeyRelease>", lambda ev: self.refresh_revexp())
        ttk.Button(top, text="تحديث", command=self.refresh_revexp).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير", command=self.export_revexp).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_revexp_pdf).pack(side="left", padx=4)
        self.revexp_kpis = self._kpi_row(det)
        tf, self.revexp_tree = self._tree(det, ("date", "ref", "account", "desc", "debit", "credit", "balance"),
                                          ("التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"),
                                          (100, 120, 260, 360, 100, 100, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        sumf = self._screen(nb, "ملخص الفائض / العجز")
        self.rev_sum_labels = self._kpi_row(sumf)
        box = ttk.Frame(sumf)
        box.pack(fill="x", padx=10, pady=4)
        self.rev_status = ttk.Label(box, text="الحالة: -", style="Section.TLabel")
        self.rev_status.pack(anchor="e")
        tf, self.revsum_tree = self._tree(sumf, ("account", "type", "debit", "credit", "balance"),
                                          ("الحساب", "النوع", "مدين", "دائن", "الرصيد الطبيعي"),
                                          (420, 120, 120, 120, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)



    def _build_cost(self):
        nb = self._subnb(self.tabs["cost"])

        mgmt = self._screen(nb, "إدارة مراكز التكلفة")
        top = ttk.Frame(mgmt); top.pack(fill="x", padx=10, pady=8)
        self.cost_name_var = tk.StringVar()
        self.cost_desc_var = tk.StringVar()
        ttk.Label(top, text="اسم المركز").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.cost_name_var, width=26).pack(side="right", padx=4)
        ttk.Label(top, text="الوصف").pack(side="right", padx=4)
        ttk.Entry(top, textvariable=self.cost_desc_var, width=34).pack(side="right", padx=4)
        ttk.Button(top, text="إضافة", command=self.add_cost_center).pack(side="left", padx=4)
        ttk.Button(top, text="إعادة تسمية المحدد", command=self.rename_cost_center).pack(side="left", padx=4)
        ttk.Button(top, text="حذف المحدد", command=self.delete_cost_center).pack(side="left", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_cost_centers).pack(side="left", padx=4)
        ttk.Label(mgmt, text="لتوزيع الحركات على المراكز: افتح القيد من دفتر اليومية ثم اختر المركز لكل سطر قبل الحفظ.",
                  style="Section.TLabel").pack(anchor="e", padx=10, pady=4)
        tf, self.cost_centers_tree = self._tree(mgmt, ("id", "name", "desc", "created", "movements"),
                                                ("رقم", "اسم المركز", "الوصف", "تاريخ الإنشاء", "عدد الحركات"),
                                                (70, 260, 320, 170, 110))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.cost_centers_tree.bind("<<TreeviewSelect>>", lambda e: self.load_selected_cost_center())

        rep = self._screen(nb, "تقرير حركة المراكز")
        top = ttk.Frame(rep); top.pack(fill="x", padx=10, pady=8)
        self.cost_report_center_var = tk.StringVar(value="كل المراكز")
        self.cost_report_combo = ttk.Combobox(top, textvariable=self.cost_report_center_var, width=34, state="readonly")
        self.cost_report_combo.pack(side="right", padx=4)
        self.cost_report_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_cost_report())
        ttk.Label(top, text="المركز").pack(side="right", padx=4)
        ttk.Label(top, text="من تاريخ").pack(side="right", padx=4)
        self.cost_from_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.cost_from_var, width=12).pack(side="right", padx=4)
        ttk.Label(top, text="إلى تاريخ").pack(side="right", padx=4)
        self.cost_to_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.cost_to_var, width=12).pack(side="right", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_cost_report).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير Excel", command=self.export_cost_report).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_cost_report).pack(side="left", padx=4)
        self.cost_kpis = self._kpi_row(rep)
        tf, self.cost_report_tree = self._tree(rep, ("date", "ref", "desc", "account", "debit", "credit", "balance"),
                                               ("التاريخ", "المرجع", "البيان", "الحساب", "مدين", "دائن", "الرصيد"),
                                               (100, 120, 320, 260, 100, 100, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

    def _build_settlements(self):
        nb = self._subnb(self.tabs["settlements"])
        self.settlement_widgets = {}
        self._create_settlement_screen(nb, "bank_misr", "مذكرة تسوية بنك مصر")
        self._create_settlement_screen(nb, "idb", "مذكرة تسوية بنك العمال")

    def _create_settlement_screen(self, nb, bank_key, title):
        f = self._screen(nb, title)
        top = ttk.Frame(f)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="تحديث", command=lambda k=bank_key: self.refresh_settlement(k)).pack(side="left", padx=4)
        ttk.Button(top, text="حفظ التسوية", command=lambda k=bank_key: self.save_settlement(k)).pack(side="left", padx=4)
        ttk.Button(top, text="ترحيل الفروقات", command=lambda k=bank_key: self.post_settlement(k)).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير Excel", command=lambda k=bank_key: self.export_settlement(k)).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=lambda k=bank_key: self.export_settlement_pdf(k)).pack(side="left", padx=4)

        year_var = tk.StringVar(value=str(datetime.today().year))
        ttk.Label(top, text="السنة").pack(side="right", padx=4)
        year_combo = ttk.Combobox(
            top,
            textvariable=year_var,
            values=[str(y) for y in range(datetime.today().year - 5, datetime.today().year + 2)],
            width=8,
            state="readonly",
        )
        year_combo.pack(side="right", padx=4)
        year_combo.bind("<<ComboboxSelected>>", lambda e, k=bank_key: self.refresh_settlement(k))

        info = ttk.Label(f, text="", style="Section.TLabel")
        info.pack(anchor="e", padx=10)

        body = ttk.Frame(f)
        body.pack(fill="both", expand=True, padx=10, pady=6)

        left = ttk.LabelFrame(body, text="بيانات التسوية")
        right = ttk.LabelFrame(body, text="ملخص التسوية")
        left.pack(side="right", fill="both", expand=True, padx=5)
        right.pack(side="right", fill="both", expand=True, padx=5)

        inputs = {}
        for key, label in [
            ("checks_under_collection", "شيكات تحت التحصيل"),
            ("uncashed_check1", "شيك مسحوب ولم يصرف (1)"),
            ("uncashed_check2", "شيك مسحوب ولم يصرف (2)"),
            ("bank_statement_balance", "رصيد كشف الحساب الفعلي"),
        ]:
            row = ttk.Frame(left)
            row.pack(fill="x", padx=10, pady=8)
            ttk.Label(row, text=label).pack(side="right", padx=4)
            var = tk.StringVar(value="0")
            ent = ttk.Entry(row, textvariable=var, width=18)
            ent.pack(side="left", padx=4)
            ent.bind("<KeyRelease>", lambda e, k=bank_key: self.refresh_settlement_preview(k))
            inputs[key] = var

        notes_var = tk.StringVar()
        nrow = ttk.Frame(left)
        nrow.pack(fill="x", padx=10, pady=8)
        ttk.Label(nrow, text="ملاحظات").pack(side="right", padx=4)
        notes_ent = ttk.Entry(nrow, textvariable=notes_var, width=60)
        notes_ent.pack(side="left", fill="x", expand=True, padx=4)

        stats = {}
        for key, label in [
            ("opening_balance", "الرصيد الافتتاحي"),
            ("receipts", "إجمالي المقبوضات"),
            ("payments", "إجمالي المدفوعات"),
            ("book_balance_end", "الرصيد الدفتري آخر الفترة"),
            ("final_bank_balance", "الرصيد بعد التسوية"),
            ("discrepancy", "فرق التسوية"),
        ]:
            row = ttk.Frame(right)
            row.pack(fill="x", padx=10, pady=8)
            ttk.Label(row, text=label).pack(side="right", padx=4)
            lbl = ttk.Label(row, text="0.00", style="Section.TLabel")
            lbl.pack(side="left", padx=4)
            stats[key] = lbl

        live_box = ttk.LabelFrame(f, text="حركة الحساب")
        live_box.pack(fill="both", expand=True, padx=10, pady=6)
        tf, tree = self._tree(live_box, ("date", "ref", "desc", "debit", "credit"),
                              ("التاريخ", "المرجع", "البيان", "مدين", "دائن"),
                              (100, 180, 420, 120, 120))
        tf.pack(fill="both", expand=True, padx=6, pady=6)

        hist_box = ttk.LabelFrame(f, text="سجل التسويات السابقة")
        hist_box.pack(fill="both", expand=False, padx=10, pady=(0, 8))
        hf, history_tree = self._tree(hist_box,
                                      ("year", "stmt", "checks", "u1", "u2", "notes", "updated"),
                                      ("السنة", "رصيد الكشف", "تحت التحصيل", "شيك 1", "شيك 2", "ملاحظات", "آخر تحديث"),
                                      (90, 120, 120, 110, 110, 360, 160))
        hf.pack(fill="both", expand=True, padx=6, pady=6)

        self.settlement_widgets[bank_key] = {
            "year": year_var,
            "info": info,
            "inputs": inputs,
            "notes": notes_var,
            "stats": stats,
            "tree": tree,
            "history_tree": history_tree,
            "title": title,
        }

    def _safe_float(self, value):
        try:
            return float(str(value).replace(",", "").strip() or 0)
        except Exception:
            return 0.0

    def refresh_settlement_preview(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if not widgets:
            return
        year = int(widgets["year"].get() or datetime.today().year)
        report = bank_settlement_report(bank_key, year)
        checks = self._safe_float(widgets["inputs"]["checks_under_collection"].get())
        u1 = self._safe_float(widgets["inputs"]["uncashed_check1"].get())
        u2 = self._safe_float(widgets["inputs"]["uncashed_check2"].get())
        stmt = self._safe_float(widgets["inputs"]["bank_statement_balance"].get())
        book_end = float(report.get("opening_balance", 0)) + float(report.get("receipts", 0)) - float(report.get("payments", 0))
        final_bal = book_end + checks + u1 + u2
        discrepancy = stmt - final_bal
        widgets["stats"]["opening_balance"].config(text=fmt(report.get("opening_balance", 0)))
        widgets["stats"]["receipts"].config(text=fmt(report.get("receipts", 0)))
        widgets["stats"]["payments"].config(text=fmt(report.get("payments", 0)))
        widgets["stats"]["book_balance_end"].config(text=fmt(book_end))
        widgets["stats"]["final_bank_balance"].config(text=fmt(final_bal))
        widgets["stats"]["discrepancy"].config(text=fmt(discrepancy))

    def refresh_settlement(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if not widgets:
            return
        year = int(widgets["year"].get() or datetime.today().year)
        report = bank_settlement_report(bank_key, year)
        widgets["info"].config(
            text=f"الحساب: {report.get('bank_name', '-')}" + ("" if report.get("account_found") else " - لم يتم العثور على الحساب")
        )
        widgets["inputs"]["checks_under_collection"].set(str(report.get("checks_under_collection", 0)))
        widgets["inputs"]["uncashed_check1"].set(str(report.get("uncashed_check1", 0)))
        widgets["inputs"]["uncashed_check2"].set(str(report.get("uncashed_check2", 0)))
        widgets["inputs"]["bank_statement_balance"].set(str(report.get("bank_statement_balance", 0)))
        widgets["notes"].set(report.get("notes", ""))
        self.refresh_settlement_preview(bank_key)

        if report.get("account_found"):
            account_report = account_statement_report([report.get("bank_name")])
            rows = [
                (r["entry_date"], r["reference"], r["description"], fmt(r["debit"]), fmt(r["credit"]))
                for r in account_report.get("rows", [])[-200:]
            ]
            self._fill_tree(widgets["tree"], rows)
        else:
            self._fill_tree(widgets["tree"], [])

        history_rows = []
        for r in bank_settlement_history(bank_key):
            history_rows.append((
                r.get("settlement_year", ""),
                fmt(r.get("bank_statement_balance", 0)),
                fmt(r.get("checks_under_collection", 0)),
                fmt(r.get("uncashed_check1", 0)),
                fmt(r.get("uncashed_check2", 0)),
                r.get("notes", ""),
                r.get("updated_at", ""),
            ))
        self._fill_tree(widgets["history_tree"], history_rows)

    def save_settlement(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if not widgets:
            return
        msg = save_bank_settlement(
            bank_key=bank_key,
            settlement_year=int(widgets["year"].get() or datetime.today().year),
            checks_under_collection=self._safe_float(widgets["inputs"]["checks_under_collection"].get()),
            uncashed_check1=self._safe_float(widgets["inputs"]["uncashed_check1"].get()),
            uncashed_check2=self._safe_float(widgets["inputs"]["uncashed_check2"].get()),
            bank_statement_balance=self._safe_float(widgets["inputs"]["bank_statement_balance"].get()),
            notes=widgets["notes"].get().strip(),
        )
        self.refresh_settlement(bank_key)
        messagebox.showinfo("تم", msg)

    def post_settlement(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if widgets:
            save_bank_settlement(
                bank_key=bank_key,
                settlement_year=int(widgets["year"].get() or datetime.today().year),
                checks_under_collection=self._safe_float(widgets["inputs"]["checks_under_collection"].get()),
                uncashed_check1=self._safe_float(widgets["inputs"]["uncashed_check1"].get()),
                uncashed_check2=self._safe_float(widgets["inputs"]["uncashed_check2"].get()),
                bank_statement_balance=self._safe_float(widgets["inputs"]["bank_statement_balance"].get()),
                notes=widgets["notes"].get().strip(),
            )
        try:
            msg = post_bank_settlement_adjustment(bank_key, int(widgets["year"].get() or datetime.today().year))
            self.refresh_all()
            messagebox.showinfo("تم", msg)
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def export_settlement(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if not widgets:
            return
        report = bank_settlement_report(bank_key, int(widgets["year"].get() or datetime.today().year))
        rows = [
            ("الحساب", report.get("bank_name", "")),
            ("السنة", report.get("settlement_year", "")),
            ("الرصيد الافتتاحي", fmt(report.get("opening_balance", 0))),
            ("إجمالي المقبوضات", fmt(report.get("receipts", 0))),
            ("إجمالي المدفوعات", fmt(report.get("payments", 0))),
            ("الرصيد الدفتري آخر الفترة", fmt(report.get("book_balance_end", 0))),
            ("شيكات تحت التحصيل", widgets["inputs"]["checks_under_collection"].get()),
            ("شيك مسحوب ولم يصرف 1", widgets["inputs"]["uncashed_check1"].get()),
            ("شيك مسحوب ولم يصرف 2", widgets["inputs"]["uncashed_check2"].get()),
            ("الرصيد بعد التسوية", widgets["stats"]["final_bank_balance"].cget("text")),
            ("رصيد كشف الحساب", widgets["inputs"]["bank_statement_balance"].get()),
            ("فرق التسوية", widgets["stats"]["discrepancy"].cget("text")),
            ("ملاحظات", widgets["notes"].get()),
        ]
        initial = f"{bank_key}_settlement.xlsx"
        self.export_simple_excel(rows, ["البيان", "القيمة"], initial)

    def export_settlement_pdf(self, bank_key):
        widgets = self.settlement_widgets.get(bank_key)
        if not widgets:
            return
        if not PDF_AVAILABLE:
            messagebox.showerror("خطأ", "مكتبات PDF غير مثبتة. نفذ pip install -r requirements.txt")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"{bank_key}_settlement.pdf")
        if not path:
            return
        report = bank_settlement_report(bank_key, int(widgets["year"].get() or datetime.today().year))
        payload = {
            "title": widgets.get("title", "مذكرة تسوية بنكية"),
            "bank_name": report.get("bank_name", ""),
            "year": str(report.get("settlement_year", "")),
            "opening_balance": fmt(report.get("opening_balance", 0)),
            "receipts": fmt(report.get("receipts", 0)),
            "payments": fmt(report.get("payments", 0)),
            "book_balance_end": fmt(report.get("book_balance_end", 0)),
            "checks_under_collection": widgets["inputs"]["checks_under_collection"].get(),
            "uncashed_check1": widgets["inputs"]["uncashed_check1"].get(),
            "uncashed_check2": widgets["inputs"]["uncashed_check2"].get(),
            "final_bank_balance": widgets["stats"]["final_bank_balance"].cget("text"),
            "bank_statement_balance": widgets["inputs"]["bank_statement_balance"].get(),
            "discrepancy": widgets["stats"]["discrepancy"].cget("text"),
            "notes": widgets["notes"].get().strip(),
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        }
        self._create_settlement_pdf(path, payload)
        messagebox.showinfo("تم", "تم إنشاء ملف PDF بنجاح")

    def _build_payroll(self):
        nb = self._subnb(self.tabs["payroll"])

        payroll = self._screen(nb, "شاشة المرتبات")
        top = ttk.Frame(payroll); top.pack(fill="x", padx=10, pady=8)
        self.payroll_search_var = tk.StringVar()
        ttk.Button(top, text="استيراد ملف المرتبات", command=self.import_payroll).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير بيانات المرتبات", command=self.export_payroll).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_payroll_pdf).pack(side="left", padx=4)
        ttk.Label(top, text="بحث").pack(side="right", padx=4)
        ent = ttk.Entry(top, textvariable=self.payroll_search_var, width=35); ent.pack(side="right", padx=4)
        ent.bind("<KeyRelease>", lambda e: self.refresh_payroll())
        self.payroll_info = ttk.Label(payroll, text="", style="Section.TLabel")
        self.payroll_info.pack(anchor="e", padx=10)
        tf, self.payroll_tree = self._tree(payroll, ("no", "name", "month", "gross", "ded", "net", "advance", "bank", "ins", "tax"),
                                           ("الرقم", "الاسم", "الفترة", "الجملة", "المستقطعات", "الصافي", "قسط السلف", "قسط البنك", "تأمين", "ضريبة"),
                                           (80, 260, 160, 110, 110, 110, 110, 110, 100, 100))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        imp = self._screen(nb, "استيراد ملف المرتبات")
        ttk.Label(imp, text="استيراد ملف المرتبات وربطه بالنظام", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(imp, text="اختيار ملف واستيراد", command=self.import_payroll).pack(anchor="e", padx=10)

        ex = self._screen(nb, "تصدير بيانات المرتبات")
        ttk.Label(ex, text="تصدير بيانات المرتبات الحالية إلى Excel", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(ex, text="تصدير الآن", command=self.export_payroll).pack(anchor="e", padx=10)

    def _build_reports(self):
        nb = self._subnb(self.tabs["reports"])

        # generic report creators
        person = self._screen(nb, "تقرير كشف حساب شخص")
        top = ttk.Frame(person); top.pack(fill="x", padx=10, pady=8)
        self.rep_person_var = tk.StringVar()
        self.rep_person_combo = ttk.Combobox(top, textvariable=self.rep_person_var, width=50, state="readonly")
        self.rep_person_combo.pack(side="right", padx=4)
        self.rep_person_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_person_report())
        ttk.Label(top, text="الشخص").pack(side="right", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_person_report).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير", command=self.export_person_report).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_person_pdf).pack(side="left", padx=4)
        tf, self.rep_person_tree = self._tree(person, ("date", "ref", "account", "desc", "debit", "credit", "balance"),
                                              ("التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"),
                                              (100, 120, 240, 360, 100, 100, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        account = self._screen(nb, "تقرير كشف حساب حساب")
        top = ttk.Frame(account); top.pack(fill="x", padx=10, pady=8)
        self.rep_account_var = tk.StringVar()
        self.rep_account_combo = ttk.Combobox(top, textvariable=self.rep_account_var, width=50, state="readonly")
        self.rep_account_combo.pack(side="right", padx=4)
        self.rep_account_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_account_report())
        ttk.Label(top, text="الحساب").pack(side="right", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_account_report).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير", command=self.export_account_report).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_account_pdf).pack(side="left", padx=4)
        tf, self.rep_account_tree = self._tree(account, ("date", "ref", "desc", "debit", "credit", "balance"),
                                               ("التاريخ", "المرجع", "البيان", "مدين", "دائن", "الرصيد"),
                                               (100, 120, 420, 100, 100, 120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        movement = self._screen(nb, "تقرير حركة حساب")
        ttk.Label(movement, text="يعرض نفس كشف الحساب المختار في الشاشة السابقة بصورة حركة حساب.", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(movement, text="نسخ من كشف الحساب", command=self.refresh_account_report).pack(anchor="e", padx=10)

        adv = self._screen(nb, "تقرير السلف")
        ttk.Button(adv, text="تحديث", command=self.refresh_advances).pack(anchor="w", padx=10, pady=8)
        tf, self.rep_adv_tree = self._tree(adv, ("no", "name", "month", "amount"),
                                           ("الرقم", "الاسم", "الفترة", "المبلغ"),
                                           (100, 340, 180, 140))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        bank = self._screen(nb, "تقرير قرض البنك")
        ttk.Button(bank, text="تحديث", command=self.refresh_bank).pack(anchor="w", padx=10, pady=8)
        tf, self.rep_bank_tree = self._tree(bank, ("no", "name", "month", "amount"),
                                            ("الرقم", "الاسم", "الفترة", "المبلغ"),
                                            (100, 340, 180, 140))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        rev = self._screen(nb, "تقرير الإيرادات")
        ttk.Button(rev, text="تحديث", command=lambda: self.refresh_pattern_report("revenues")).pack(anchor="w", padx=10, pady=8)
        ttk.Button(rev, text="طباعة PDF", command=lambda: self.print_pattern_pdf("revenues")).pack(anchor="w", padx=10)
        tf, self.rep_revenues_tree = self._tree(rev, ("date", "ref", "account", "desc", "debit", "credit", "balance"),
                                                ("التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"),
                                                (100,120,260,360,100,100,120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        exp = self._screen(nb, "تقرير المصروفات")
        ttk.Button(exp, text="تحديث", command=lambda: self.refresh_pattern_report("expenses")).pack(anchor="w", padx=10, pady=8)
        ttk.Button(exp, text="طباعة PDF", command=lambda: self.print_pattern_pdf("expenses")).pack(anchor="w", padx=10)
        tf, self.rep_expenses_tree = self._tree(exp, ("date", "ref", "account", "desc", "debit", "credit", "balance"),
                                                ("التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"),
                                                (100,120,260,360,100,100,120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        jr = self._screen(nb, "تقرير يومية مفصل")
        ttk.Button(jr, text="تحديث", command=self.refresh_journal).pack(anchor="w", padx=10, pady=8)
        ttk.Button(jr, text="طباعة PDF", command=self.print_rep_journal_pdf).pack(anchor="w", padx=10)
        tf, self.rep_journal_tree = self._tree(jr, ("id", "date", "ref", "desc", "debit", "credit"),
                                               ("رقم", "التاريخ", "المرجع", "البيان", "مدين", "دائن"),
                                               (70,100,120,460,110,110))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        td = self._screen(nb, "تقرير ميزان مراجعة تفصيلي")
        ttk.Button(td, text="تحديث", command=self.refresh_trial).pack(anchor="w", padx=10, pady=8)
        ttk.Button(td, text="طباعة PDF", command=self.print_trial_pdf).pack(anchor="w", padx=10)
        tf, self.rep_trial_tree = self._tree(td, ("account", "debit", "credit", "net_debit", "net_credit"),
                                             ("الحساب", "إجمالي مدين", "إجمالي دائن", "رصيد مدين", "رصيد دائن"),
                                             (440,120,120,120,120))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        ts = self._screen(nb, "تقرير ميزان مراجعة إجمالي")
        self.trial_summary_label = ttk.Label(ts, text="الإجماليات: -", style="Section.TLabel")
        self.trial_summary_label.pack(anchor="e", padx=10, pady=10)

    def _build_settings(self):
        nb = self._subnb(self.tabs["settings"])

        people = self._screen(nb, "إدارة أسماء الأشخاص")
        ttk.Label(people, text="الأسماء المستخرجة من البيانات الحالية", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        tf, self.people_tree = self._tree(people, ("name",), ("الاسم",), (500,))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        imp = self._screen(nb, "استيراد البيانات من Excel")
        ttk.Label(imp, text="استيراد اليومية الأمريكية أو المرتبات", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(imp, text="استيراد اليومية", command=self.import_journal).pack(anchor="e", padx=10, pady=4)
        ttk.Button(imp, text="استيراد المرتبات", command=self.import_payroll).pack(anchor="e", padx=10, pady=4)

        exp = self._screen(nb, "تصدير البيانات إلى Excel")
        ttk.Label(exp, text="تصدير التقارير من الشاشات التشغيلية المختلفة", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(exp, text="تصدير دفتر اليومية", command=self.export_journal).pack(anchor="e", padx=10, pady=4)
        ttk.Button(exp, text="تصدير الميزان", command=self.export_trial).pack(anchor="e", padx=10, pady=4)
        ttk.Button(exp, text="تصدير المرتبات", command=self.export_payroll).pack(anchor="e", padx=10, pady=4)

        backup = self._screen(nb, "النسخ الاحتياطي")
        ttk.Label(backup, text="حفظ نسخة احتياطية من قاعدة البيانات", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(backup, text="إنشاء نسخة احتياطية", command=self.backup_db).pack(anchor="e", padx=10)

        restore = self._screen(nb, "استرجاع النسخة الاحتياطية")
        ttk.Label(restore, text="استرجاع نسخة قاعدة البيانات", style="Section.TLabel").pack(anchor="e", padx=10, pady=10)
        ttk.Button(restore, text="استرجاع قاعدة بيانات", command=self.restore_db).pack(anchor="e", padx=10)

        sysf = self._screen(nb, "إعدادات النظام")
        txt = (
            "إعدادات النظام الحالية:\n"
            "- قاعدة بيانات SQLite محلية\n"
            "- استيراد يومية أمريكية من Excel\n"
            "- استيراد مرتبات من Excel\n"
            "- تصدير الشاشات الرئيسية إلى Excel\n"
            "- طباعة PDF لكل الشاشات: اليومية، الميزان، الأستاذ العام، كشوف المدينين، الإيرادات والمصروفات، المرتبات\n"
            "- طباعة سندات رسمية: سند قيد وأذن صرف من الشاشة مباشرة\n"
            "- مراكز تكلفة: إدارة المراكز وتوزيع سطور القيود عليها مع تقرير حركة\n"
            "- يمكن تطوير الصلاحيات ومتعدد المستخدمين لاحقاً"
        )
        ttk.Label(sysf, text=txt, justify="right").pack(anchor="e", padx=10, pady=10)

    # ---------- cost centers actions ----------
    def load_selected_cost_center(self):
        sel = self.cost_centers_tree.selection()
        if not sel:
            return
        vals = self.cost_centers_tree.item(sel[0], "values")
        self.cost_name_var.set(vals[1] if len(vals) > 1 else "")
        self.cost_desc_var.set(vals[2] if len(vals) > 2 else "")

    def add_cost_center(self):
        result = add_cost_center(self.cost_name_var.get(), self.cost_desc_var.get())
        self.refresh_cost_centers()
        messagebox.showinfo("مراكز التكلفة", result)

    def rename_cost_center(self):
        sel = self.cost_centers_tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر مركز تكلفة أولاً")
            return
        center_id = int(self.cost_centers_tree.item(sel[0], "values")[0])
        result = rename_cost_center(center_id, self.cost_name_var.get())
        self.refresh_cost_centers()
        messagebox.showinfo("مراكز التكلفة", result)

    def delete_cost_center(self):
        sel = self.cost_centers_tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر مركز تكلفة أولاً")
            return
        center_id = int(self.cost_centers_tree.item(sel[0], "values")[0])
        if not messagebox.askyesno("تأكيد", "حذف المركز؟ سيتم فك ارتباطه من كل القيود دون حذف القيود."):
            return
        result = delete_cost_center(center_id)
        self.refresh_cost_centers()
        self.refresh_cost_report()
        messagebox.showinfo("مراكز التكلفة", result)

    def refresh_cost_centers(self):
        centers = cost_centers_list()
        self._fill_tree(self.cost_centers_tree, [
            (c["id"], c["name"], c["description"] or "", c["created_at"] or "", c["movements"])
            for c in centers
        ])
        names = ["كل المراكز"] + [c["name"] for c in centers]
        self.cost_report_combo["values"] = names
        if not self.cost_report_center_var.get() or self.cost_report_center_var.get() not in names:
            self.cost_report_center_var.set("كل المراكز")

    def _selected_cost_center_id(self):
        name = self.cost_report_center_var.get().strip()
        if not name or name == "كل المراكز":
            return None
        for c in cost_centers_list():
            if c["name"] == name:
                return c["id"]
        return None

    def refresh_cost_report(self):
        data = cost_center_report(
            center_id=self._selected_cost_center_id(),
            date_from=self.cost_from_var.get().strip() or None,
            date_to=self.cost_to_var.get().strip() or None,
        )
        self.current_cost_rows = data["rows"]
        self.cost_kpis["debit"].config(text=fmt(data["total_debit"]))
        self.cost_kpis["credit"].config(text=fmt(data["total_credit"]))
        self.cost_kpis["balance"].config(text=fmt(data["final_balance"]))
        self._fill_tree(self.cost_report_tree, [
            (r["entry_date"], r["reference"], r["description"], r["account_name"],
             fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in data["rows"]
        ])

    def export_cost_report(self):
        rows = [self.cost_report_tree.item(i, "values") for i in self.cost_report_tree.get_children()]
        self.export_simple_excel(rows, ["التاريخ", "المرجع", "البيان", "الحساب", "مدين", "دائن", "الرصيد"], "cost_centers_report.xlsx")

    def print_cost_report(self):
        rows = [self.cost_report_tree.item(i, "values") for i in self.cost_report_tree.get_children()]
        self.print_rows_pdf(
            "تقرير حركة مراكز التكلفة",
            f"المركز: {self.cost_report_center_var.get()}",
            ["التاريخ", "المرجع", "البيان", "الحساب", "مدين", "دائن", "الرصيد"],
            rows, "cost_centers_report.pdf", landscape_page=True,
        )

    # ---------- PDF printing helpers ----------
    def print_rows_pdf(self, title, subtitle, headers, rows, default_name, landscape_page=False, col_widths=None):
        if not pdf_lib_available():
            messagebox.showerror("خطأ", "مكتبات PDF غير مثبتة. نفذ pip install -r requirements.txt")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=default_name)
        if not path:
            return
        try:
            build_table_pdf(path, title, subtitle, headers, rows, col_widths=col_widths, landscape_page=landscape_page)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر إنشاء ملف PDF:\n{e}")
            return
        messagebox.showinfo("تم", "تم إنشاء ملف PDF بنجاح")

    def print_entry_voucher(self):
        """طباعة القيد المحدد كمستند سند قيد رسمي."""
        entry_id = self.get_selected_entry_id()
        if not entry_id:
            messagebox.showwarning("تنبيه", "اختر قيداً أولاً من دفتر اليومية")
            return
        if not pdf_lib_available():
            messagebox.showerror("خطأ", "مكتبات PDF غير مثبتة. نفذ pip install -r requirements.txt")
            return
        conn = get_connection(); cur = conn.cursor()
        head = cur.execute(
            "SELECT id, entry_date, reference, description, total_debit, total_credit FROM journal_entries WHERE id=?",
            (entry_id,)).fetchone()
        conn.close()
        if not head:
            return
        lines = get_entry_lines(entry_id)
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"voucher_{entry_id}.pdf")
        if not path:
            return
        try:
            build_voucher_pdf(
                path,
                "سند قيد",
                [
                    ("رقم السند", f"قيد رقم {entry_id}"),
                    ("التاريخ", head["entry_date"]),
                    ("المرجع", head["reference"] or "-"),
                    ("البيان", head["description"] or "-"),
                ],
                [(r["account_name"], fmt(r["debit"]), fmt(r["credit"]), r["line_description"] or "") for r in lines],
                fmt(head["total_debit"]),
                fmt(head["total_credit"]),
            )
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر إنشاء ملف PDF:\n{e}")
            return
        messagebox.showinfo("تم", "تم إنشاء سند القيد PDF بنجاح")

    def print_selected_voucher(self):
        """طباعة أذن الصرف المحدد كمستند رسمي."""
        sel = self.vouchers_tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر أذن صرف أولاً من الجدول")
            return
        if not pdf_lib_available():
            messagebox.showerror("خطأ", "مكتبات PDF غير مثبتة. نفذ pip install -r requirements.txt")
            return
        vals = self.vouchers_tree.item(sel[0], "values")
        entry_id = int(vals[0])
        lines = get_entry_lines(entry_id)
        total_debit = sum(float(r["debit"] or 0) for r in lines)
        total_credit = sum(float(r["credit"] or 0) for r in lines)
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"payment_voucher_{entry_id}.pdf")
        if not path:
            return
        try:
            build_voucher_pdf(
                path,
                "إذن صرف",
                [
                    ("رقم الأذن", f"{entry_id}"),
                    ("التاريخ", vals[1]),
                    ("المرجع", vals[2]),
                    ("البيان", vals[3]),
                    ("المستفيد", vals[3]),
                    ("من حساب", vals[4]),
                    ("إلى حساب", vals[5]),
                    ("المبلغ", f"{vals[6]}"),
                ],
                [(r["account_name"], fmt(r["debit"]), fmt(r["credit"]), r["line_description"] or "") for r in lines],
                fmt(total_debit),
                fmt(total_credit),
            )
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر إنشاء ملف PDF:\n{e}")
            return
        messagebox.showinfo("تم", "تم إنشاء إذن الصرف PDF بنجاح")

    def print_journal_pdf(self):
        rows = [self.journal_tree.item(i, "values") for i in self.journal_tree.get_children()]
        self.print_rows_pdf(
            "دفتر اليومية",
            f"عدد القيود المعروضة: {len(rows)}",
            ["رقم", "التاريخ", "المرجع", "البيان", "مدين", "دائن"],
            rows, "journal.pdf", landscape_page=True,
        )

    def print_rep_journal_pdf(self):
        rows = [self.rep_journal_tree.item(i, "values") for i in self.rep_journal_tree.get_children()]
        self.print_rows_pdf(
            "تقرير يومية مفصل",
            f"عدد القيود: {len(rows)}",
            ["رقم", "التاريخ", "المرجع", "البيان", "مدين", "دائن"],
            rows, "journal_detail.pdf", landscape_page=True,
        )

    def print_trial_pdf(self):
        rows = [self.trial_tree.item(i, "values") for i in self.trial_tree.get_children()]
        self.print_rows_pdf(
            "ميزان المراجعة",
            f"عدد الحسابات: {len(rows)}",
            ["الحساب", "إجمالي مدين", "إجمالي دائن", "رصيد مدين", "رصيد دائن"],
            rows, "trial_balance.pdf",
        )

    def print_ledger_pdf(self):
        account = self.ledger_account_var.get().strip()
        rows = [self.ledger_tree.item(i, "values") for i in self.ledger_tree.get_children()]
        self.print_rows_pdf(
            "كشف حساب (الأستاذ العام)",
            f"الحساب: {account or '-'} | عدد الحركات: {len(rows)}",
            ["التاريخ", "المرجع", "البيان", "مدين", "دائن", "الرصيد"],
            rows, "ledger.pdf",
        )

    def print_debtor_pdf(self):
        sel = self.debtors_list.curselection()
        person = self.debtors_list.get(sel[0]) if sel else "-"
        rows = [self.debtor_tree.item(i, "values") for i in self.debtor_tree.get_children()]
        self.print_rows_pdf(
            "كشف حساب مدين",
            f"الشخص: {person}",
            ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"],
            rows, "debtor_statement.pdf", landscape_page=True,
        )

    def print_revexp_pdf(self):
        account = self.revexp_account_var.get().strip() or "كل الحسابات"
        rows = [self.revexp_tree.item(i, "values") for i in self.revexp_tree.get_children()]
        self.print_rows_pdf(
            "حساب الإيرادات والمصروفات",
            f"الحساب: {account}",
            ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"],
            rows, "revexp_statement.pdf", landscape_page=True,
        )

    def print_payroll_pdf(self):
        rows = [self.payroll_tree.item(i, "values") for i in self.payroll_tree.get_children()]
        self.print_rows_pdf(
            "بيان المرتبات",
            f"عدد العاملين: {len(rows)}",
            ["الرقم", "الاسم", "الفترة", "الجملة", "المستقطعات", "الصافي", "قسط السلف", "قسط البنك", "تأمين", "ضريبة"],
            rows, "payroll.pdf", landscape_page=True,
        )

    def print_person_pdf(self):
        person = self.rep_person_var.get().strip()
        rows = [self.rep_person_tree.item(i, "values") for i in self.rep_person_tree.get_children()]
        self.print_rows_pdf(
            "تقرير كشف حساب شخص",
            f"الشخص: {person or '-'}",
            ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"],
            rows, "person_report.pdf", landscape_page=True,
        )

    def print_account_pdf(self):
        account = self.rep_account_var.get().strip()
        rows = [self.rep_account_tree.item(i, "values") for i in self.rep_account_tree.get_children()]
        self.print_rows_pdf(
            "تقرير كشف حساب حساب",
            f"الحساب: {account or '-'}",
            ["التاريخ", "المرجع", "البيان", "مدين", "دائن", "الرصيد"],
            rows, "account_report.pdf",
        )

    def print_pattern_pdf(self, which):
        if which == "revenues":
            tree = self.rep_revenues_tree
            title = "تقرير الإيرادات"
        else:
            tree = self.rep_expenses_tree
            title = "تقرير المصروفات"
        rows = [tree.item(i, "values") for i in tree.get_children()]
        self.print_rows_pdf(
            title,
            f"عدد الحركات: {len(rows)}",
            ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"],
            rows, f"{which}_report.pdf", landscape_page=True,
        )

    # ---------- refresh methods ----------


    def refresh_all(self):
        self.refresh_dashboard()
        self.refresh_journal()
        self.refresh_chart()
        self.refresh_ledger()
        self.refresh_trial()
        self.refresh_debtors_people()
        self.refresh_debtor_statement()
        self.refresh_advances()
        self.refresh_bank()
        self.refresh_vouchers()
        self.refresh_revexp()
        self.refresh_rev_summary()
        self.refresh_cost_centers()
        self.refresh_cost_report()
        self.refresh_settlement("bank_misr")
        self.refresh_settlement("idb")
        self.refresh_payroll()
        self.refresh_person_report()
        self.refresh_account_report()
        self.refresh_pattern_report("revenues")
        self.refresh_pattern_report("expenses")
        self.refresh_people_names()

    def refresh_dashboard(self):
        data = dashboard_summary()
        self.home_kpis["debit"].config(text=fmt(data["debits"]))
        self.home_kpis["credit"].config(text=fmt(data["credits"]))
        self.home_kpis["balance"].config(text=fmt(data["debtors_balance"]))
        self._fill_tree(self.home_recent_tree, [
            (r["entry_date"], r["reference"], r["description"], fmt(r["total_debit"]), fmt(r["total_credit"]))
            for r in data["recent_entries"]
        ])
        self._fill_tree(self.home_accounts_tree, [
            (r["account_name"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["balance"]))
            for r in data["top_accounts"]
        ])

    def refresh_journal(self):
        rows = list_entries(self.journal_search.get().strip())
        self.current_journal_rows = rows
        self._fill_tree(self.journal_tree, [
            (r["id"], r["entry_date"], r["reference"], r["description"], fmt(r["total_debit"]), fmt(r["total_credit"]))
            for r in rows
        ])
        self._fill_tree(self.rep_journal_tree, [
            (r["id"], r["entry_date"], r["reference"], r["description"], fmt(r["total_debit"]), fmt(r["total_credit"]))
            for r in rows
        ])
        self.refresh_entry_lines()

    def get_selected_entry_id(self):
        sel = self.journal_tree.selection()
        if not sel:
            return None
        return int(self.journal_tree.item(sel[0], "values")[0])

    def refresh_entry_lines(self):
        entry_id = self.get_selected_entry_id()
        if not entry_id:
            self._fill_tree(self.entry_lines_tree, [])
            return
        rows = get_entry_lines(entry_id)
        self._fill_tree(self.entry_lines_tree, [
            (r["account_name"], fmt(r["debit"]), fmt(r["credit"]), r["line_description"] or "", r["cost_center"] or "")
            for r in rows
        ])

    def open_selected_entry(self):
        entry_id = self.get_selected_entry_id()
        if not entry_id:
            messagebox.showwarning("تنبيه", "اختر قيداً أولاً من دفتر اليومية")
            return
        EntryEditor(self, entry_id)

    def delete_selected_entry(self):
        entry_id = self.get_selected_entry_id()
        if not entry_id:
            messagebox.showwarning("تنبيه", "اختر قيداً أولاً")
            return
        if not messagebox.askyesno("تأكيد", "هل تريد حذف القيد المحدد؟"):
            return
        conn = get_connection()
        cur = conn.cursor()
        cur.execute("DELETE FROM journal_lines WHERE entry_id=?", (entry_id,))
        cur.execute("DELETE FROM journal_entries WHERE id=?", (entry_id,))
        conn.commit(); conn.close()
        self.refresh_all()

    def refresh_chart(self):
        conn = get_connection(); cur = conn.cursor()
        rows = cur.execute("SELECT code, name, category, normal_side FROM accounts ORDER BY name").fetchall()
        conn.close()
        self._fill_tree(self.chart_tree, rows)
        accs = [r[1] for r in rows]
        self.ledger_combo["values"] = accs
        self.rep_account_combo["values"] = accs
        self.revexp_combo["values"] = revenue_expense_accounts()
        people = debtors_people()
        self.rep_person_combo["values"] = people

        if accs and not self.rep_account_var.get().strip():
            self.rep_account_var.set(accs[0])
        if people and not self.rep_person_var.get().strip():
            self.rep_person_var.set(people[0])
        if hasattr(self, "ledger_var") and accs and not self.ledger_var.get().strip():
            self.ledger_var.set(accs[0])
        revs = revenue_expense_accounts()
        if revs and not self.revexp_account_var.get().strip():
            self.revexp_account_var.set(revs[0])

    def refresh_ledger(self):
        account = self.ledger_account_var.get().strip()
        rows = ledger_for_account(account) if account else []
        self._fill_tree(self.ledger_tree, [
            (r["entry_date"], r["reference"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in rows
        ])

    def refresh_trial(self):
        rows = trial_balance()
        view = [(r["account_name"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["net_debit"]), fmt(r["net_credit"])) for r in rows]
        self._fill_tree(self.trial_tree, view)
        self._fill_tree(self.rep_trial_tree, view)
        total_d = sum(float(r["debit"] or 0) for r in rows)
        total_c = sum(float(r["credit"] or 0) for r in rows)
        total_nd = sum(float(r["net_debit"] or 0) for r in rows)
        total_nc = sum(float(r["net_credit"] or 0) for r in rows)
        self.trial_summary_label.config(text=f"إجمالي مدين: {fmt(total_d)} | إجمالي دائن: {fmt(total_c)} | رصيد مدين: {fmt(total_nd)} | رصيد دائن: {fmt(total_nc)}")

    def refresh_debtors_people(self):
        people = debtors_people()
        q = self.debtor_search.get().strip()
        if q:
            people = [p for p in people if q in p]
        self.debtors_list.delete(0, "end")
        for p in people:
            self.debtors_list.insert("end", p)
        if people and not self.debtors_list.curselection():
            self.debtors_list.selection_set(0)
        self.refresh_debtor_statement()

    def refresh_debtor_statement(self):
        sel = self.debtors_list.curselection()
        person = self.debtors_list.get(sel[0]) if sel else ""
        report = debtors_person_report(person_name=person)
        self.current_debtor_rows = report["rows"]
        self.debtor_kpis["debit"].config(text=fmt(report["total_debit"]))
        self.debtor_kpis["credit"].config(text=fmt(report["total_credit"]))
        self.debtor_kpis["balance"].config(text=fmt(report["final_balance"]))
        self._fill_tree(self.debtor_tree, [
            (r["entry_date"], r["reference"], r["account_name"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in report["rows"]
        ])

    def refresh_advances(self):
        data = workers_advances_report(self.adv_search.get().strip())
        self.current_adv_rows = data["rows"]
        self.adv_total.config(text=f"إجمالي السلف: {fmt(data['total'])}")
        rows = [(r["employee_no"], r["employee_name"], r["payroll_month"], fmt(r["amount"])) for r in data["rows"]]
        self._fill_tree(self.adv_tree, rows)
        self._fill_tree(self.rep_adv_tree, rows)

    def refresh_bank(self):
        data = bank_loans_report(self.bank_search.get().strip())
        self.current_bank_rows = data["rows"]
        self.bank_total.config(text=f"إجمالي قرض البنك: {fmt(data['total'])}")
        rows = [(r["employee_no"], r["employee_name"], r["payroll_month"], fmt(r["amount"])) for r in data["rows"]]
        self._fill_tree(self.bank_tree, rows)
        self._fill_tree(self.rep_bank_tree, rows)

    def refresh_vouchers(self):
        rows = smart_vouchers()
        self.current_voucher_rows = rows
        self._fill_tree(self.vouchers_tree, [
            (r["entry_id"], r["entry_date"], r["reference"], r["description"], r["source_account"], r["target_account"], fmt(r["amount"]))
            for r in rows
        ])

    def refresh_revexp(self):
        account = self.revexp_account_var.get().strip() or None
        report = revenue_expense_report(account_name=account, description_filter=self.revexp_filter_var.get().strip())
        self.current_rev_rows = report["rows"]
        self.revexp_kpis["debit"].config(text=fmt(report["total_debit"]))
        self.revexp_kpis["credit"].config(text=fmt(report["total_credit"]))
        self.revexp_kpis["balance"].config(text=fmt(report["final_balance"]))
        self._fill_tree(self.revexp_tree, [
            (r["entry_date"], r["reference"], r["account_name"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in report["rows"]
        ])

    def refresh_rev_summary(self):
        data = revenue_expense_final_summary()
        self.rev_sum_labels["debit"].config(text=fmt(data["total_revenues"]))
        self.rev_sum_labels["credit"].config(text=fmt(data["total_expenses"]))
        self.rev_sum_labels["balance"].config(text=fmt(data["net_result"]))
        state = "فائض" if data["net_result"] >= 0 else "عجز"
        self.rev_status.config(text=f"الحالة: {state}")
        self._fill_tree(self.revsum_tree, [
            (r["account_name"], r["account_type"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["natural_balance"]))
            for r in data["rows"]
        ])

    def refresh_payroll(self):
        info = latest_payroll_import()
        self.payroll_info.config(
            text=f"الفترة: {info.get('payroll_month','-')} | العاملون: {info.get('employees_count',0)} | إجمالي السلف: {fmt(info.get('salary_advance_total',0))} | إجمالي البنك: {fmt(info.get('bank_loan_total',0))}"
        )
        rows = payroll_rows(self.payroll_search_var.get().strip())
        self.current_payroll_rows = rows
        self._fill_tree(self.payroll_tree, [
            (r["employee_no"], r["employee_name"], r["payroll_month"], fmt(r["gross_total"]), fmt(r["total_deductions"]),
             fmt(r["net_pay"]), fmt(r["salary_advance_installment"]), fmt(r["bank_loan_installment"]), fmt(r["insurance_employee"]), fmt(r["tax_amount"]))
            for r in rows
        ])

    def refresh_person_report(self):
        people = list(self.rep_person_combo.cget("values")) if hasattr(self, "rep_person_combo") else []
        person = self.rep_person_var.get().strip()
        if not person and people:
            person = people[0]
            self.rep_person_var.set(person)
        elif person and people and person not in people:
            matches = [p for p in people if person in p]
            if matches:
                person = matches[0]
                self.rep_person_var.set(person)
        rows = debtors_person_report(person_name=person)["rows"] if person else []
        self.current_person_report = rows
        self._fill_tree(self.rep_person_tree, [
            (r["entry_date"], r["reference"], r["account_name"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in rows
        ])

    def refresh_account_report(self):
        accounts = list(self.rep_account_combo.cget("values")) if hasattr(self, "rep_account_combo") else []
        account = self.rep_account_var.get().strip()
        if not account and accounts:
            account = accounts[0]
            self.rep_account_var.set(account)
        elif account and accounts and account not in accounts:
            matches = [a for a in accounts if account in a]
            if matches:
                account = matches[0]
                self.rep_account_var.set(account)
        rows = ledger_for_account(account) if account else []
        self.current_account_report = rows
        self._fill_tree(self.rep_account_tree, [
            (r["entry_date"], r["reference"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
            for r in rows
        ])

    def refresh_pattern_report(self, which):
        if which == "revenues":
            report = account_statement_report(["%إيراد%", "%ايراد%", "%الإيرادات%"])
            self.current_rev_report = report["rows"]
            self._fill_tree(self.rep_revenues_tree, [
                (r["entry_date"], r["reference"], r["account_name"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
                for r in report["rows"]
            ])
        else:
            report = account_statement_report(["%مصروف%", "%نفقة%"])
            self.current_exp_report = report["rows"]
            self._fill_tree(self.rep_expenses_tree, [
                (r["entry_date"], r["reference"], r["account_name"], r["description"], fmt(r["debit"]), fmt(r["credit"]), fmt(r["running_balance"]))
                for r in report["rows"]
            ])

    def refresh_people_names(self):
        people = debtors_people()
        self._fill_tree(self.people_tree, [(p,) for p in people])

    # ---------- exports ----------
    def export_journal(self):
        rows = [(r["id"], r["entry_date"], r["reference"], r["description"], r["total_debit"], r["total_credit"]) for r in getattr(self, "current_journal_rows", [])]
        self.export_simple_excel(rows, ["رقم", "التاريخ", "المرجع", "البيان", "مدين", "دائن"], "journal.xlsx")

    def export_trial(self):
        rows = [self.trial_tree.item(i, "values") for i in self.trial_tree.get_children()]
        self.export_simple_excel(rows, ["الحساب", "إجمالي مدين", "إجمالي دائن", "رصيد مدين", "رصيد دائن"], "trial_balance.xlsx")

    def export_debtors(self):
        rows = [self.debtor_tree.item(i, "values") for i in self.debtor_tree.get_children()]
        self.export_simple_excel(rows, ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"], "debtor_statement.xlsx")

    def export_advances(self):
        rows = [self.adv_tree.item(i, "values") for i in self.adv_tree.get_children()]
        self.export_simple_excel(rows, ["الرقم", "الاسم", "الفترة", "قسط السلف"], "workers_advances.xlsx")

    def export_bank(self):
        rows = [self.bank_tree.item(i, "values") for i in self.bank_tree.get_children()]
        self.export_simple_excel(rows, ["الرقم", "الاسم", "الفترة", "قسط البنك"], "bank_loans.xlsx")

    def export_vouchers(self):
        rows = [self.vouchers_tree.item(i, "values") for i in self.vouchers_tree.get_children()]
        self.export_simple_excel(rows, ["رقم", "التاريخ", "المرجع", "البيان", "من", "إلى", "المبلغ"], "smart_vouchers.xlsx")

    def export_revexp(self):
        rows = [self.revexp_tree.item(i, "values") for i in self.revexp_tree.get_children()]
        self.export_simple_excel(rows, ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"], "revexp_statement.xlsx")

    def export_payroll(self):
        rows = [self.payroll_tree.item(i, "values") for i in self.payroll_tree.get_children()]
        self.export_simple_excel(rows, ["الرقم", "الاسم", "الفترة", "الجملة", "المستقطعات", "الصافي", "قسط السلف", "قسط البنك", "تأمين", "ضريبة"], "payroll.xlsx")

    def export_person_report(self):
        rows = [self.rep_person_tree.item(i, "values") for i in self.rep_person_tree.get_children()]
        self.export_simple_excel(rows, ["التاريخ", "المرجع", "الحساب", "البيان", "مدين", "دائن", "الرصيد"], "person_report.xlsx")

    def export_account_report(self):
        rows = [self.rep_account_tree.item(i, "values") for i in self.rep_account_tree.get_children()]
        self.export_simple_excel(rows, ["التاريخ", "المرجع", "البيان", "مدين", "دائن", "الرصيد"], "account_report.xlsx")


if __name__ == "__main__":
    app = AccountingApp()
    app.mainloop()
