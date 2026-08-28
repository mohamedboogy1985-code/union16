
import os
import shutil
import sqlite3
import sys
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
import app_settings as APPS
from app_settings import (
    ROLES,
    PERMISSION_LABELS,
    init_settings_tables,
    get_all_settings,
    set_settings,
    set_setting,
    get_json_setting,
    verify_user,
    list_users,
    add_user,
    update_user,
    can as user_can,
    log_action,
    list_audit,
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
    monthly_summary,
    available_years,
)

import compliance as CMP
from compliance import (
    check_entry as compliance_check_entry,
    save_violations_for_entry,
    list_violations,
    violations_summary,
    acknowledge_violation,
    scan_all_entries,
    check_cash_on_hand,
    get_settings as compliance_get_settings,
    save_settings as compliance_save_settings,
    reset_settings as compliance_reset_settings,
    SETTING_LABELS,
    init_compliance_tables,
)


APP_TITLE = "النظام المحاسبي المكتبي المتكامل"

def fmt(v):
    try:
        return f"{float(v or 0):,.2f}"
    except Exception:
        return "0.00"


class ComplianceWarning(tk.Toplevel):
    """نافذة إنذار بعدم مطابقة القيد للائحة المالية."""

    def __init__(self, master, message, has_violations=True):
        super().__init__(master)
        self.title("إنذار: مخالفة اللائحة المالية")
        self.geometry("720x520")
        self.transient(master)
        self.grab_set()
        self.confirmed = False

        header = tk.Label(self, text="⛔ تنبيه الرقابة المالية",
                          font=("Tahoma", 15, "bold"),
                          fg="white", bg="#b02a2a" if has_violations else "#b8860b")
        header.pack(fill="x")
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=12, pady=10)

        txt = tk.Text(body, wrap="word", font=("Tahoma", 11), spacing1=2)
        ys = ttk.Scrollbar(body, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=ys.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        txt.insert("1.0", message)
        txt.configure(state="disabled")

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=10)
        ttk.Button(btns, text="الرجوع للتعديل", command=self._cancel).pack(side="right", padx=6)
        ttk.Button(btns, text="حفظ رغم التنبيه (تسجيل للمراجعة)",
                   command=self._confirm).pack(side="left", padx=6)

    def _confirm(self):
        self.confirmed = True
        self.destroy()

    def _cancel(self):
        self.confirmed = False
        self.destroy()


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

        # ===== رقابة اللائحة المالية: فحص القيد قبل الحفظ =====
        draft_lines = [
            {"account": acc, "debit": d, "credit": c, "desc": desc}
            for acc, d, c, desc, cost in rows
        ]
        cmp_results = compliance_check_entry(
            date=self.date_var.get(), reference=self.ref_var.get(),
            description=self.desc_var.get(), draft_lines=draft_lines)
        blocking = [r for r in cmp_results if r["severity"] == "violation"]
        approvals = [r for r in cmp_results if r["severity"] == "approval"]
        if cmp_results:
            msg_lines = ["⚠️  نتيجة فحص اللائحة المالية لهذا القيد:", ""]
            for r in cmp_results:
                tag = {"violation": "❌ مخالفة", "approval": "🟡 يلزم اعتماد",
                       "info": "ℹ️ تنبيه"}.get(r["severity"], "")
                msg_lines.append(f"{tag} — {r['article']}")
                msg_lines.append(r["message"])
                msg_lines.append("")
            if blocking or approvals:
                msg_lines.append("هل تريد حفظ القيد رغم ذلك مع تسجيل الملاحظات للمراجعة؟")
                win = ComplianceWarning(self, "\n".join(msg_lines), has_violations=bool(blocking))
                self.wait_window(win)
                if not win.confirmed:
                    return
            else:
                messagebox.showinfo("فحص اللائحة المالية", "\n".join(msg_lines))

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

        # تسجيل ملاحظات اللائحة في سجل الرقابة
        if cmp_results:
            try:
                save_violations_for_entry(entry_id, cmp_results)
            except Exception:
                pass
        app = getattr(self.master, "app_ref", None) or self.master
        if hasattr(app, "log_audit"):
            app.log_audit(
                "تعديل قيد" if self.entry_id else "إضافة قيد",
                f"قيد رقم {entry_id} بتاريخ {self.date_var.get()} بمبلغ {fmt(total_debit)}")
        self.master.refresh_all()
        messagebox.showinfo("تم", "تم حفظ القيد بنجاح")
        self.destroy()


class AccountingApp(tk.Tk):
    def __init__(self, current_user=None):
        super().__init__()
        self.current_user = current_user
        self.title(APP_TITLE)
        self.geometry("1450x900")
        self.minsize(1200, 760)
        init_db()
        init_compliance_tables()
        init_settings_tables()
        self._style()
        self._header()
        self._build_ui()
        self._apply_permissions()
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
        user = self.current_user or {}
        role_label = ROLES.get(user.get("role"), {}).get("label", user.get("role") or "-")
        user_info = f"{user.get('full_name') or user.get('username') or 'مستخدم'} ({role_label})"
        ttk.Label(top, text=user_info, style="Header.TLabel", font=("Tahoma", 10)).pack(side="left", padx=16)
        ttk.Button(top, text="تسجيل الخروج", command=self.logout).pack(side="left", padx=4)
        btns = tk.Frame(top, bg="#0f4c81")
        btns.pack(side="left", padx=16)
        for txt, cmd in [
            ("تحديث", self.refresh_all),
            ("استيراد اليومية", self.import_journal),
            ("استيراد المرتبات", self.import_payroll),
            ("نسخة احتياطية", self.backup_db),
        ]:
            ttk.Button(btns, text=txt, command=cmd).pack(side="left", padx=4)
        # حفظ مرجع أزرار الهيدر لتقييدها حسب الصلاحية
        self.header_buttons = {
            "import_journal": self._find_header_button(btns, "استيراد اليومية"),
            "import_payroll": self._find_header_button(btns, "استيراد المرتبات"),
            "backup": self._find_header_button(btns, "نسخة احتياطية"),
        }

    @staticmethod
    def _find_header_button(parent, text):
        for child in parent.winfo_children():
            try:
                if str(child.cget("text")) == text:
                    return child
            except Exception:
                continue
        return None

    def logout(self):
        if messagebox.askyesno("تسجيل الخروج", "هل تريد تسجيل الخروج؟"):
            self.destroy()

    def log_audit(self, action, details=""):
        uname = self.current_user["username"] if self.current_user else "غير معروف"
        log_action(uname, action, details)

    def can(self, permission):
        if not self.current_user:
            return False
        return user_can(self.current_user["role"], permission)

    def deny(self, message="لا تملك صلاحية لهذا الإجراء."):
        messagebox.showerror("صلاحية مرفوضة", message)

    def _guard_button(self, btn, permission):
        # اعتراض الضغطة: منع تنفيذ أمر الزر إذا لم تتوفر الصلاحية
        def on_click(event):
            if not self.can(permission):
                self.deny(f"دورك لا يسمح بـ ({PERMISSION_LABELS.get(permission, permission)})")
                return "break"
        btn.bind("<Button-1>", on_click)

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
            ("compliance", "الرقابة المالية"),
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
        self._build_compliance()
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
        if not self.can("import"):
            self.deny("دورك لا يسمح باستيراد ملفات اليومية.")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            msg = import_excel_file(path)
            self.log_audit("استيراد يومية أمريكية", os.path.basename(path))
            self.refresh_all()
            messagebox.showinfo("نجاح", msg)
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def import_payroll(self):
        if not self.can("import"):
            self.deny("دورك لا يسمح باستيراد ملفات المرتبات.")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        try:
            msg = import_payroll_excel(path)
            self.log_audit("استيراد مرتبات", os.path.basename(path))
            self.refresh_all()
            messagebox.showinfo("نجاح", msg)
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def backup_db(self):
        if not self.can("backup"):
            self.deny("دورك لا يسمح بإنشاء نسخة احتياطية.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite", "*.db")], initialfile=f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
        if not path:
            return
        shutil.copyfile(DB_PATH, path)
        self.log_audit("نسخة احتياطية", os.path.basename(path))
        messagebox.showinfo("تم", "تم حفظ النسخة الاحتياطية")

    def restore_db(self):
        if not self.can("restore"):
            self.deny("دورك لا يسمح باسترجاع نسخة احتياطية.")
            return
        path = filedialog.askopenfilename(filetypes=[("SQLite", "*.db")])
        if not path:
            return
        if not messagebox.askyesno("تأكيد", "سيتم استبدال قاعدة البيانات الحالية. هل تريد المتابعة؟"):
            return
        shutil.copyfile(path, DB_PATH)
        self.log_audit("استرجاع نسخة احتياطية", os.path.basename(path))
        self.refresh_all()
        messagebox.showinfo("تم", "تم استرجاع النسخة الاحتياطية")

    def _build_home(self):
        f = self.tabs["home"]
        self.compliance_badge = ttk.Label(
            f, text="✅ رقابة اللائحة: جاري التحميل...",
            style="Section.TLabel", foreground="#1b7a32")
        self.compliance_badge.pack(anchor="w", padx=14, pady=(6, 0))
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
        ttk.Button(box, text="طباعة PDF", command=self.print_rev_summary_pdf).pack(anchor="w", padx=10)
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
        self.cost_admin_buttons = []
        for txt, cmd in [
            ("إضافة", self.add_cost_center),
            ("إعادة تسمية المحدد", self.rename_cost_center),
            ("حذف المحدد", self.delete_cost_center),
        ]:
            btn = ttk.Button(top, text=txt, command=cmd)
            btn.pack(side="left", padx=4)
            self.cost_admin_buttons.append(btn)
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
        self.settlement_admin_buttons = getattr(self, "settlement_admin_buttons", {})
        btns = {}
        for txt, cmd in [
            ("حفظ التسوية", lambda k=bank_key: self.save_settlement(k)),
            ("ترحيل الفروقات", lambda k=bank_key: self.post_settlement(k)),
        ]:
            btn = ttk.Button(top, text=txt, command=cmd)
            btn.pack(side="left", padx=4)
            btns[txt] = btn
        self.settlement_admin_buttons[bank_key] = btns
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

        msum = self._screen(nb, "الملخص الشهري (فائض / عجز)")
        top = ttk.Frame(msum); top.pack(fill="x", padx=10, pady=8)
        self.monthly_year_var = tk.StringVar()
        self.monthly_year_combo = ttk.Combobox(top, textvariable=self.monthly_year_var, width=10, state="readonly")
        self.monthly_year_combo.pack(side="right", padx=4)
        self.monthly_year_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_monthly())
        ttk.Label(top, text="السنة").pack(side="right", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_monthly).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير Excel", command=self.export_monthly).pack(side="left", padx=4)
        ttk.Button(top, text="طباعة PDF", command=self.print_monthly).pack(side="left", padx=4)
        self.monthly_kpis = self._kpi_row(msum)
        self.monthly_status = ttk.Label(msum, text="", style="Section.TLabel")
        self.monthly_status.pack(anchor="e", padx=10)
        tf, self.monthly_tree = self._tree(msum, ("month", "revenues", "expenses", "net"),
                                           ("الشهر", "الإيرادات", "المصروفات", "صافي الفائض / العجز"),
                                           (260, 160, 160, 200))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

    def _build_compliance(self):
        nb = self._subnb(self.tabs["compliance"])

        # ---------- شاشة: فحص القيود / سجل المخالفات ----------
        log = self._screen(nb, "إنذارات اللائحة المالية")
        top = ttk.Frame(log); top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="إعادة فحص كل القيود", command=self.rescan_all_compliance).pack(side="left", padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh_compliance_log).pack(side="left", padx=4)
        ttk.Button(top, text="اعتماد المحدد كاستثناء", command=self.ack_selected_violation).pack(side="left", padx=4)
        ttk.Button(top, text="فتح القيد", command=self.open_violation_entry).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير Excel", command=self.export_compliance).pack(side="left", padx=4)
        self.cmp_filter_var = tk.StringVar(value="الكل")
        ttk.Label(top, text="النوع").pack(side="right", padx=4)
        filt = ttk.Combobox(top, textvariable=self.cmp_filter_var, width=22, state="readonly",
                            values=["الكل", "مخالفات فقط", "يلزم اعتماد", "تنبيهات", "المعتمدة كاستثناء"])
        filt.pack(side="right", padx=4)
        filt.bind("<<ComboboxSelected>>", lambda e: self.refresh_compliance_log())

        self.cmp_summary_lbl = ttk.Label(log, text="", style="Section.TLabel")
        self.cmp_summary_lbl.pack(anchor="e", padx=10)

        tf, self.cmp_tree = self._tree(
            log,
            ("id", "date", "ref", "article", "severity", "message", "amount", "checked", "ack"),
            ("#", "تاريخ القيد", "المرجع", "المادة", "النوع", "نص الإنذار / الملاحظة", "المبلغ", "وقت الفحص", "الحالة"),
            (50, 95, 100, 90, 110, 520, 100, 120, 110))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.cmp_tree.tag_configure("violation", background="#f8d7da")
        self.cmp_tree.tag_configure("approval", background="#fff3cd")
        self.cmp_tree.tag_configure("info", background="#d1ecf1")
        self.cmp_tree.tag_configure("ack", background="#e2e2e2", foreground="#666")

        # ---------- شاشة: فحص فوري لقيد ----------
        live = self._screen(nb, "فحص قيد فوري")
        ltop = ttk.Frame(live); ltop.pack(fill="x", padx=10, pady=8)
        ttk.Label(ltop, text="التاريخ").pack(side="right", padx=4)
        self.cmp_live_date = tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        ttk.Entry(ltop, textvariable=self.cmp_live_date, width=12).pack(side="right", padx=4)
        ttk.Label(ltop, text="المرجع").pack(side="right", padx=4)
        self.cmp_live_ref = tk.StringVar()
        ttk.Entry(ltop, textvariable=self.cmp_live_ref, width=16).pack(side="right", padx=4)
        ttk.Label(ltop, text="البيان / الغرض").pack(side="right", padx=4)
        self.cmp_live_desc = tk.StringVar()
        ttk.Entry(ltop, textvariable=self.cmp_live_desc, width=44).pack(side="right", padx=4)
        ttk.Button(ltop, text="فحص", command=self.refresh_compliance_live).pack(side="left", padx=4)

        mid = ttk.LabelFrame(live, text="سطور القيد (الحساب - مدين - دائن - بيان)")
        mid.pack(fill="both", expand=False, padx=10, pady=6)
        tools = ttk.Frame(mid); tools.pack(fill="x", padx=6, pady=4)
        ttk.Button(tools, text="إضافة سطر", command=self.cmp_live_add_row).pack(side="left", padx=4)
        ttk.Button(tools, text="حذف سطر", command=self.cmp_live_del_row).pack(side="left", padx=4)
        self.cmp_live_tree = ttk.Treeview(mid, columns=("account", "debit", "credit", "desc"),
                                          show="headings", height=6)
        for c, h, w in [("account", "الحساب", 360), ("debit", "مدين", 110),
                        ("credit", "دائن", 110), ("desc", "بيان السطر", 320)]:
            self.cmp_live_tree.heading(c, text=h)
            self.cmp_live_tree.column(c, width=w, anchor="center")
        self.cmp_live_tree.pack(fill="x", padx=6, pady=4)
        edit = ttk.Frame(mid); edit.pack(fill="x", padx=6, pady=4)
        self.cmp_row_account = tk.StringVar()
        self.cmp_row_debit = tk.StringVar(value="0")
        self.cmp_row_credit = tk.StringVar(value="0")
        self.cmp_row_desc = tk.StringVar()
        ttk.Label(edit, text="الحساب").pack(side="right", padx=3)
        ttk.Combobox(edit, textvariable=self.cmp_row_account, width=34, values=ledger_accounts()).pack(side="right", padx=3)
        ttk.Label(edit, text="مدين").pack(side="right", padx=3)
        ttk.Entry(edit, textvariable=self.cmp_row_debit, width=11).pack(side="right", padx=3)
        ttk.Label(edit, text="دائن").pack(side="right", padx=3)
        ttk.Entry(edit, textvariable=self.cmp_row_credit, width=11).pack(side="right", padx=3)
        ttk.Label(edit, text="بيان").pack(side="right", padx=3)
        ttk.Entry(edit, textvariable=self.cmp_row_desc, width=28).pack(side="right", padx=3)
        ttk.Button(edit, text="تطبيق على السطر", command=self.cmp_live_apply_row).pack(side="left", padx=3)
        self.cmp_live_tree.bind("<<TreeviewSelect>>", self.cmp_live_load_row)

        res_box = ttk.LabelFrame(live, text="نتيجة الفحص")
        res_box.pack(fill="both", expand=True, padx=10, pady=6)
        self.cmp_live_result = tk.Text(res_box, wrap="word", font=("Tahoma", 11), height=10)
        rys = ttk.Scrollbar(res_box, orient="vertical", command=self.cmp_live_result.yview)
        self.cmp_live_result.configure(yscrollcommand=rys.set)
        self.cmp_live_result.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        rys.pack(side="right", fill="y", pady=6)
        self.cmp_live_result.tag_configure("violation", foreground="#b00020", font=("Tahoma", 11, "bold"))
        self.cmp_live_result.tag_configure("approval", foreground="#9c6500")
        self.cmp_live_result.tag_configure("info", foreground="#0b5b78")
        self.cmp_live_result.tag_configure("ok", foreground="#1b7a32", font=("Tahoma", 12, "bold"))

        # ---------- شاشة: حدود اللائحة (الإعدادات) ----------
        rules = self._screen(nb, "حدود اللائحة المالية")
        rtop = ttk.Frame(rules); rtop.pack(fill="x", padx=10, pady=8)
        ttk.Button(rtop, text="حفظ الحدود", command=self.save_compliance_limits).pack(side="left", padx=4)
        ttk.Button(rtop, text="استرجاع حدود اللائحة الافتراضية", command=self.reset_compliance_limits).pack(side="left", padx=4)
        ttk.Label(rtop, text="نطاق الجهة").pack(side="right", padx=6)
        self.cmp_scope_var = tk.StringVar()
        scope_combo = ttk.Combobox(rtop, textvariable=self.cmp_scope_var, width=18, state="readonly",
                                   values=["union: نقابة عامة", "committee: لجنة نقابية"])
        scope_combo.pack(side="right", padx=6)
        tf, self.cmp_rules_tree = self._tree(
            rules, ("key", "group", "label", "value", "ref"),
            ("م", "المجموعة", "البند", "الحد الحالي", "المرجع / النص"),
            (40, 120, 260, 120, 600))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.cmp_rules_tree.bind("<Double-1>", self.edit_compliance_limit)

        # ---------- شاشة: رصيد الخزينة ----------
        cash = self._screen(nb, "رقابة رصيد الخزينة (مادة 6)")
        ctop = ttk.Frame(cash); ctop.pack(fill="x", padx=10, pady=8)
        ttk.Button(ctop, text="فحص رصيد الخزينة الآن", command=self.refresh_compliance_cash).pack(side="left", padx=4)
        tf, self.cmp_cash_tree = self._tree(
            cash, ("account", "balance", "cap", "status"),
            ("الحساب", "الرصيد الحالي", "سقف اللائحة", "الحالة"),
            (300, 160, 160, 500))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.cmp_cash_tree.tag_configure("violation", background="#f8d7da")
        self.cmp_cash_tree.tag_configure("ok", background="#d4edda")

    # ----- لقطة الفحص الفوري -----
    def cmp_live_add_row(self):
        iid = self.cmp_live_tree.insert("", "end", values=("", "0", "0", ""))
        self.cmp_live_tree.selection_set(iid)

    def cmp_live_del_row(self):
        for iid in self.cmp_live_tree.selection():
            self.cmp_live_tree.delete(iid)

    def cmp_live_load_row(self, event=None):
        sel = self.cmp_live_tree.selection()
        if not sel:
            return
        v = self.cmp_live_tree.item(sel[0], "values")
        self.cmp_row_account.set(v[0])
        self.cmp_row_debit.set(v[1])
        self.cmp_row_credit.set(v[2])
        self.cmp_row_desc.set(v[3])

    def cmp_live_apply_row(self):
        sel = self.cmp_live_tree.selection()
        if not sel:
            return
        self.cmp_live_tree.item(sel[0], values=(
            self.cmp_row_account.get().strip(),
            self.cmp_row_debit.get().strip() or "0",
            self.cmp_row_credit.get().strip() or "0",
            self.cmp_row_desc.get().strip()))

    def refresh_compliance_live(self):
        draft = []
        for iid in self.cmp_live_tree.get_children():
            acc, d, c, desc = self.cmp_live_tree.item(iid, "values")
            if not str(acc).strip():
                continue
            draft.append({"account": str(acc).strip(),
                          "debit": float(d or 0), "credit": float(c or 0),
                          "desc": str(desc or "")})
        res = compliance_check_entry(date=self.cmp_live_date.get(),
                                     reference=self.cmp_live_ref.get(),
                                     description=self.cmp_live_desc.get(),
                                     draft_lines=draft)
        self.cmp_live_result.configure(state="normal")
        self.cmp_live_result.delete("1.0", "end")
        if not draft:
            self.cmp_live_result.insert("end", "أدخل سطور القيد أولاً ثم اضغط (فحص).", "info")
        elif not res:
            self.cmp_live_result.insert("end", "✅ القيد مطابق للائحة المالية — لا توجد ملاحظات.", "ok")
        else:
            for r in res:
                tag = r["severity"]
                tag_txt = {"violation": "❌ مخالفة للائحة", "approval": "🟡 يلزم موافقة / اعتماد",
                           "info": "ℹ️ تنبيه إرشادي"}.get(tag, tag)
                self.cmp_live_result.insert("end", f"{tag_txt} — {r['article']}\n", tag)
                self.cmp_live_result.insert("end", r["message"] + "\n\n")
        self.cmp_live_result.configure(state="disabled")

    # ----- سجل المخالفات -----
    def rescan_all_compliance(self):
        if not self.can("compliance_rescan"):
            self.deny("دورك لا يسمح بإعادة الفحص الشامل للائحة (متاح للمراجع ومدير النظام).")
            return
        if not messagebox.askyesno("تأكيد", "سيتم إعادة فحص جميع قيود اليومية مقابل اللائحة المالية. متابعة؟"):
            return
        entries = self._all_entry_ids()
        count = 0
        for eid in entries:
            res = compliance_check_entry(entry_id=eid)
            if res:
                save_violations_for_entry(eid, res)
                count += 1
        if hasattr(self, "log_audit"):
            self.log_audit("فحص شامل للائحة", f"فحص {len(entries)} قيد، قيود عليها ملاحظات: {count}")
        self.refresh_compliance_log()
        self.refresh_compliance_cash()
        self.refresh_compliance_badge()
        messagebox.showinfo("تم", f"اكتمل الفحص الشامل.\nقيود عليها ملاحظات: {count} من {len(entries)} قيد.")

    def _all_entry_ids(self):
        conn = get_connection()
        ids = [r[0] for r in conn.execute("SELECT id FROM journal_entries ORDER BY id").fetchall()]
        conn.close()
        return ids

    def refresh_compliance_log(self):
        f = self.cmp_filter_var.get()
        include_ack = (f == "المعتمدة كاستثناء")
        rows = list_violations(include_ack=include_ack)
        if f == "مخالفات فقط":
            rows = [r for r in rows if r["severity"] == "violation"]
        elif f == "يلزم اعتماد":
            rows = [r for r in rows if r["severity"] == "approval"]
        elif f == "تنبيهات":
            rows = [r for r in rows if r["severity"] == "info"]
        sev_txt = {"violation": "❌ مخالفة", "approval": "🟡 يلزم اعتماد", "info": "ℹ️ تنبيه"}
        for iid in self.cmp_tree.get_children():
            self.cmp_tree.delete(iid)
        for r in rows:
            tag = "ack" if r["acknowledged"] else r["severity"]
            self.cmp_tree.insert(
                "", "end", iid=f"v{r['id']}",
                values=(r["id"], r["entry_date"] or "", r["entry_ref"] or "",
                        r["article"] or "",
                        "معتمد كاستثناء" if r["acknowledged"] else sev_txt.get(r["severity"], r["severity"]),
                        r["message"] or "", fmt(r["amount"]),
                        r["checked_at"] or "",
                        (r["ack_note"] or "استثناء معتمد") if r["acknowledged"] else "قائم"),
                tags=(tag,))
        summ = violations_summary()
        self.cmp_summary_lbl.config(
            text=(f"المخالفات: {summ['violation']}  |  يلزم اعتماد: {summ['approval']}  |  "
                  f"تنبيهات: {summ['info']}"))
        self.cmp_rows = rows

    def ack_selected_violation(self):
        if not self.can("compliance_ack"):
            self.deny("دورك لا يسمح باعتماد مخالفات اللائحة كاستثناء (متاح للمراجع ومدير النظام).")
            return
        sel = self.cmp_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر ملاحظة من القائمة أولاً")
            return
        note = "معتمدة من المسؤول كاستثناء"
        ids = []
        for iid in sel:
            vid = int(iid[1:])
            acknowledge_violation(vid, note)
            ids.append(str(vid))
        if hasattr(self, "log_audit"):
            self.log_audit("اعتماد مخالفة لائحة", f"ملاحظات أرقام: {', '.join(ids)}")
        self.refresh_compliance_log()
        self.refresh_compliance_badge()

    def open_violation_entry(self):
        sel = self.cmp_tree.selection()
        if not sel:
            return
        vid = int(sel[0][1:])
        row = next((r for r in getattr(self, "cmp_rows", []) if r["id"] == vid), None)
        if row and row["entry_id"]:
            EntryEditor(self, entry_id=row["entry_id"])

    def export_compliance(self):
        rows = getattr(self, "cmp_rows", [])
        data = [(r["id"], r["entry_date"], r["entry_ref"], r["article"], r["severity"],
                 r["message"], r["amount"], r["checked_at"],
                 "معتمد" if r["acknowledged"] else "قائم") for r in rows]
        self.export_simple_excel(
            data,
            ["#", "التاريخ", "المرجع", "المادة", "النوع", "الرسالة", "المبلغ", "وقت الفحص", "الحالة"],
            "تقرير_الرقابة_المالية")

    # ----- حدود اللائحة -----
    def _populate_rules_tree(self, tree):
        settings = compliance_get_settings()
        for iid in tree.get_children():
            tree.delete(iid)
        for i, (key, (label, ref, group)) in enumerate(SETTING_LABELS.items(), start=1):
            if key == "entity_scope":
                continue
            val = settings.get(key, "")
            tree.insert("", "end", iid=f"r_{key}", values=(i, group, label, val, ref))

    def refresh_compliance_rules(self):
        settings = compliance_get_settings()
        scope = settings.get("entity_scope", "union")
        self.cmp_scope_var.set(f"{scope}: {'نقابة عامة' if scope == 'union' else 'لجنة نقابية'}")
        self._populate_rules_tree(self.cmp_rules_tree)
        if hasattr(self, "st_rules_tree"):
            self._populate_rules_tree(self.st_rules_tree)

    def _edit_limit_on_tree(self, tree, event=None):
        if not self.can("settings"):
            self.deny("دورك لا يسمح بتعديل حدود اللائحة.")
            return
        sel = tree.selection()
        if not sel:
            return
        iid = sel[0]
        vals = tree.item(iid, "values")
        key = iid[2:]
        from tkinter import simpledialog
        new_val = simpledialog.askstring("تعديل الحد",
                                         f"{vals[2]}\nالمرجع: {vals[4]}\nأدخل القيمة الجديدة:",
                                         initialvalue=str(vals[3]), parent=self)
        if new_val is None:
            return
        compliance_save_settings({key: new_val.strip()})
        if hasattr(self, "log_audit"):
            self.log_audit("تعديل حد لائحة", f"{vals[2]}: {vals[3]} -> {new_val}")
        self.refresh_compliance_rules()
        messagebox.showinfo("تم", "تم حفظ الحد الجديد")

    def edit_compliance_limit(self, event=None):
        self._edit_limit_on_tree(self.cmp_rules_tree, event)

    def save_compliance_limits(self):
        if not self.can("settings"):
            self.deny("دورك لا يسمح بتعديل إعدادات اللائحة.")
            return
        scope_raw = self.cmp_scope_var.get()
        scope = "committee" if scope_raw.startswith("committee") else "union"
        compliance_save_settings({"entity_scope": scope})
        if hasattr(self, "log_audit"):
            self.log_audit("تعديل نطاق الجهة", scope)
        self.refresh_compliance_rules()
        self.refresh_compliance_log()
        if hasattr(self, "st_scope_var"):
            self.st_scope_var.set(self.cmp_scope_var.get())
        messagebox.showinfo("تم", "تم حفظ حدود اللائحة ونطاق الجهة")

    def reset_compliance_limits(self):
        if not self.can("settings"):
            self.deny("دورك لا يسمح بتعديل حدود اللائحة.")
            return
        if not messagebox.askyesno("تأكيد", "سيتم استرجاع جميع حدود اللائحة الافتراضية. متابعة؟"):
            return
        compliance_reset_settings()
        self.refresh_compliance_rules()
        messagebox.showinfo("تم", "تم استرجاع الحدود الافتراضية")

    # ----- رصيد الخزينة -----
    def refresh_compliance_cash(self):
        for iid in self.cmp_cash_tree.get_children():
            self.cmp_cash_tree.delete(iid)
        settings = compliance_get_settings()
        is_union = settings.get("entity_scope", "union") != "committee"
        cap = settings["cash_on_hand_union"] if is_union else settings["cash_on_hand_committee"]
        results = check_cash_on_hand()
        if not results:
            conn = get_connection()
            rows = conn.execute("""
                SELECT a.name, COALESCE(SUM(jl.debit),0)-COALESCE(SUM(jl.credit),0) bal
                FROM accounts a JOIN journal_lines jl ON jl.account_id=a.id
                WHERE a.name LIKE '%خزينة%' OR a.name LIKE '%صندوق%' OR a.name LIKE '%نقد%'
                GROUP BY a.id HAVING bal > 0
            """).fetchall()
            conn.close()
            if not rows:
                self.cmp_cash_tree.insert("", "end", values=("لا توجد أرصدة خزينة ظاهرة", "-", fmt(cap), "-"))
            for r in rows:
                self.cmp_cash_tree.insert("", "end",
                                          values=(r[0], fmt(r[1]), fmt(cap), "✅ ضمن الحد"),
                                          tags=("ok",))
        for r in results:
            self.cmp_cash_tree.insert("", "end",
                                      values=(r["account"], fmt(r["amount"]), fmt(cap), r["message"]),
                                      tags=("violation",))

    # ----- شارة المخالفات في الرئيسية -----
    def refresh_compliance_badge(self):
        try:
            summ = violations_summary()
            total = summ["violation"] + summ["approval"]
            if total > 0:
                self.compliance_badge.config(
                    text=f"⚠️ رقابة اللائحة: {summ['violation']} مخالفة و {summ['approval']} بانتظار اعتماد",
                    foreground="#b00020")
            else:
                self.compliance_badge.config(
                    text="✅ رقابة اللائحة: لا توجد مخالفات قائمة", foreground="#1b7a32")
        except Exception:
            pass

    def _build_settings(self):
        nb = self._subnb(self.tabs["settings"])

        # ================================================================
        # 1) إعدادات النظام العامة
        # ================================================================
        gen = self._screen(nb, "النظام العامة")
        gf = ttk.LabelFrame(gen, text="بيانات المنشأة")
        gf.pack(fill="x", padx=12, pady=8)
        self.st_general_vars = {}
        for key, label in [
            ("org_name", "اسم المنشأة"),
            ("tax_registration", "رقم التسجيل الضريبي"),
            ("fiscal_year", "السنة المالية"),
            ("currency", "العملة"),
        ]:
            row = ttk.Frame(gf); row.pack(fill="x", padx=10, pady=4)
            ttk.Label(row, text=label, width=22).pack(side="right", padx=4)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var, width=60).pack(side="right", fill="x", expand=True, padx=4)
            self.st_general_vars[key] = var

        ttk.Button(gf, text="اختيار الشعار", command=self.pick_logo).pack(side="right", padx=10, pady=6)
        self.st_logo_lbl = ttk.Label(gf, text="")
        self.st_logo_lbl.pack(side="right", padx=4)

        nf = ttk.LabelFrame(gen, text="التنبيهات")
        nf.pack(fill="x", padx=12, pady=8)
        self.st_notify_vars = {}
        for key, label in [
            ("notify_cash_cap", "تنبيه تجاوز سقف الخزينة (مادة 6)"),
            ("notify_violations", "تنبيه مخالفات اللائحة المالية"),
            ("notify_monthly_close", "تنبيه الجرد الشهري وإقفال الشهر (مادة 16)"),
        ]:
            var = tk.IntVar()
            ttk.Checkbutton(nf, text=label, variable=var).pack(anchor="e", padx=12, pady=3)
            self.st_notify_vars[key] = var
        ttk.Button(gen, text="حفظ إعدادات النظام العامة", command=self.save_general_settings).pack(anchor="e", padx=16, pady=8)

        # ================================================================
        # 2) إعدادات اللائحة المالية
        # ================================================================
        reg = self._screen(nb, "اللائحة المالية")
        top = ttk.Frame(reg); top.pack(fill="x", padx=10, pady=6)
        ttk.Button(top, text="حفظ حدود اللائحة", command=self.save_compliance_limits).pack(side="left", padx=4)
        ttk.Button(top, text="استرجاع الحدود الافتراضية", command=self.reset_compliance_limits).pack(side="left", padx=4)
        ttk.Label(top, text="نطاق الجهة").pack(side="right", padx=6)
        self.st_scope_var = tk.StringVar()
        ttk.Combobox(top, textvariable=self.st_scope_var, width=20, state="readonly",
                     values=["union: نقابة عامة", "committee: لجنة نقابية"]).pack(side="right", padx=6)
        tf, self.st_rules_tree = self._tree(
            reg, ("key", "group", "label", "value", "ref"),
            ("م", "المجموعة", "البند", "الحد الحالي", "المرجع / نص المادة"),
            (40, 120, 260, 120, 620))
        tf.pack(fill="both", expand=True, padx=10, pady=8)
        self.st_rules_tree.bind("<Double-1>", lambda e: self._edit_limit_on_tree(self.st_rules_tree, e))
        ttk.Label(reg, text="لتعديل أي حد: نقرة مزدوجة على الصف.", foreground="#777").pack(anchor="e", padx=14)

        # ================================================================
        # 3) إعدادات الحسابات والضرائب
        # ================================================================
        tax = self._screen(nb, "الحسابات والضرائب")
        tlf = ttk.LabelFrame(tax, text="الفاتورة الإلكترونية")
        tlf.pack(fill="x", padx=12, pady=8)
        self.st_tax_vars = {}
        for key, label in [
            ("einvoice_client_id", "Client ID"),
            ("einvoice_client_secret", "Client Secret"),
            ("einvoice_activity_code", "كود النشاط"),
        ]:
            row = ttk.Frame(tlf); row.pack(fill="x", padx=10, pady=3)
            ttk.Label(row, text=label, width=20).pack(side="right", padx=4)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var, width=46).pack(side="right", padx=4)
            self.st_tax_vars[key] = var
        row = ttk.Frame(tlf); row.pack(fill="x", padx=10, pady=3)
        ttk.Label(row, text="بيئة الفاتورة الإلكترونية", width=24).pack(side="right", padx=4)
        self.st_env_var = tk.StringVar()
        ttk.Combobox(row, textvariable=self.st_env_var, width=20, state="readonly",
                     values=["test", "production"]).pack(side="right", padx=4)
        self.st_einvoice_var = tk.IntVar()
        ttk.Checkbutton(tlf, text="تفعيل الربط مع الفاتورة الإلكترونية",
                        variable=self.st_einvoice_var).pack(anchor="e", padx=12, pady=4)

        rates = ttk.LabelFrame(tax, text="النسب الضريبية")
        rates.pack(fill="x", padx=12, pady=8)
        for key, label in [
            ("vat_rate", "ضريبة القيمة المضافة %"),
            ("withholding_fees_pct", "خصم تحت حساب الضريبة - أتعاب مهنية %"),
            ("withholding_commissions_pct", "خصم تحت حساب الضريبة - عمولات وسمسرة %"),
            ("withholding_supply_pct", "خصم تحت حساب الضريبة - مقاولات وتوريدات %"),
            ("withholding_min_amount", "حد تطبيق الخصم (جنيه)"),
        ]:
            row = ttk.Frame(rates); row.pack(fill="x", padx=10, pady=3)
            ttk.Label(row, text=label, width=44).pack(side="right", padx=4)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var, width=14).pack(side="right", padx=4)
            self.st_tax_vars[key] = var
        ttk.Button(tax, text="حفظ إعدادات الحسابات والضرائب",
                   command=self.save_tax_settings).pack(anchor="e", padx=16, pady=8)

        # ================================================================
        # 4) الموارد البشرية والمرتبات
        # ================================================================
        hr = self._screen(nb, "الموارد البشرية والمرتبات")
        ins = ttk.LabelFrame(hr, text="التأمينات الاجتماعية")
        ins.pack(fill="x", padx=12, pady=8)
        self.st_hr_vars = {}
        for key, label in [
            ("ins_employee_pct", "حصة العامل %"),
            ("ins_employer_pct", "حصة المنشأة %"),
            ("ins_min_wage", "الحد الأدنى للأجر التأميني"),
            ("ins_max_wage", "الحد الأقصى للأجر التأميني"),
        ]:
            row = ttk.Frame(ins); row.pack(fill="x", padx=10, pady=3)
            ttk.Label(row, text=label, width=26).pack(side="right", padx=4)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var, width=14).pack(side="right", padx=4)
            self.st_hr_vars[key] = var

        brk = ttk.LabelFrame(hr, text="شرائح ضريبة كسب العمل (استمارة 2)")
        brk.pack(fill="both", expand=False, padx=12, pady=8)
        ttk.Label(brk, text="من (جنيه/سنة)   |   إلى (0 = بدون حد)   |   النسبة %   |   البيان").pack(anchor="e", padx=10)
        self.st_bracket_vars = []
        self.brackets_box = ttk.Frame(brk); self.brackets_box.pack(fill="x", padx=10, pady=4)
        ttk.Button(brk, text="+ إضافة شريحة", command=self.add_tax_bracket_row).pack(anchor="e", padx=10, pady=4)
        row = ttk.Frame(hr); row.pack(fill="x", padx=12)
        ttk.Label(row, text="التخفيض الضريبي السنوي").pack(side="right", padx=4)
        self.st_hr_vars["tax_annual_relief"] = tk.StringVar()
        ttk.Entry(row, textvariable=self.st_hr_vars["tax_annual_relief"], width=12).pack(side="right", padx=4)
        ttk.Button(hr, text="حفظ إعدادات المرتبات", command=self.save_hr_settings).pack(anchor="e", padx=16, pady=10)

        # ================================================================
        # 5) المستخدمون والصلاحيات
        # ================================================================
        usr = self._screen(nb, "المستخدمون والصلاحيات")
        top = ttk.Frame(usr); top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="تحديث", command=self.refresh_settings_users).pack(side="left", padx=4)
        ttk.Button(top, text="إضافة مستخدم", command=self.add_settings_user).pack(side="left", padx=4)
        ttk.Button(top, text="إعادة تعيين كلمة المرور", command=self.reset_user_password).pack(side="left", padx=4)
        ttk.Button(top, text="تفعيل/إيقاف", command=self.toggle_user_active).pack(side="left", padx=4)
        tf, self.st_users_tree = self._tree(
            usr, ("id", "username", "full_name", "role", "active", "created"),
            ("#", "اسم المستخدم", "الاسم بالكامل", "الدور", "الحالة", "تاريخ الإنشاء"),
            (50, 180, 260, 180, 100, 160))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        roles_box = ttk.LabelFrame(usr, text="الأدوار والصلاحيات")
        roles_box.pack(fill="x", padx=10, pady=8)
        for rkey, rdata in ROLES.items():
            rframe = ttk.Frame(roles_box); rframe.pack(fill="x", padx=8, pady=3)
            ttk.Label(rframe, text=rdata["label"] + ":", width=18,
                      font=("Tahoma", 10, "bold")).pack(side="right", padx=4)
            perms = " | ".join(PERMISSION_LABELS.get(p, p) for p in sorted(rdata["permissions"]))
            ttk.Label(rframe, text=perms, foreground="#444").pack(side="right", padx=4)

        # ================================================================
        # 6) سجل التدقيق والمراجعة
        # ================================================================
        aud = self._screen(nb, "سجل التدقيق")
        top = ttk.Frame(aud); top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="تحديث", command=self.refresh_audit_log).pack(side="left", padx=4)
        ttk.Button(top, text="تصدير Excel", command=self.export_audit_log).pack(side="left", padx=4)
        tf, self.audit_tree = self._tree(
            aud, ("ts", "user", "action", "details"),
            ("التاريخ/الوقت", "المستخدم", "العملية", "التفاصيل"),
            (160, 160, 260, 700))
        tf.pack(fill="both", expand=True, padx=10, pady=8)

        # ================================================================
        # 7) البيانات (الأشخاص / استيراد / تصدير / نسخ احتياطي)
        # ================================================================
        data = self._screen(nb, "البيانات والنسخ الاحتياطي")
        people = ttk.LabelFrame(data, text="إدارة أسماء الأشخاص")
        people.pack(fill="x", padx=10, pady=6)
        ttk.Label(people, text="الأسماء المستخرجة من البيانات الحالية").pack(anchor="e", padx=10, pady=4)
        tf, self.people_tree = self._tree(people, ("name",), ("الاسم",), (500,))
        tf.pack(fill="x", padx=10, pady=6)

        ops = ttk.LabelFrame(data, text="الاستيراد والتصدير والنسخ الاحتياطي")
        ops.pack(fill="x", padx=10, pady=6)
        ttk.Button(ops, text="استيراد اليومية الأمريكية", command=self.import_journal).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="استيراد المرتبات", command=self.import_payroll).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="تصدير دفتر اليومية", command=self.export_journal).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="تصدير الميزان", command=self.export_trial).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="تصدير المرتبات", command=self.export_payroll).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="إنشاء نسخة احتياطية", command=self.backup_db).pack(anchor="e", padx=10, pady=3)
        ttk.Button(ops, text="استرجاع قاعدة بيانات", command=self.restore_db).pack(anchor="e", padx=10, pady=3)

    # ---- إعدادات عامة ----
    def pick_logo(self):
        path = filedialog.askopenfilename(filetypes=[("صور", "*.png *.jpg *.jpeg *.bmp")])
        if not path:
            return
        set_setting("org_logo", path)
        self.st_logo_lbl.config(text=os.path.basename(path))
        self.log_audit("اختيار شعار", os.path.basename(path))

    def load_general_settings(self):
        if not hasattr(self, "st_general_vars"):
            return  # شاشة الإعدادات لم تُبنَ بعد
        s = get_all_settings()
        for key, var in self.st_general_vars.items():
            var.set(str(s.get(key, "")))
        for key, var in getattr(self, "st_tax_vars", {}).items():
            var.set(str(s.get(key, "")))
        if hasattr(self, "st_logo_lbl"):
            self.st_logo_lbl.config(text=os.path.basename(s.get("org_logo", "")) if s.get("org_logo") else "لا يوجد شعار")
        for key, var in getattr(self, "st_notify_vars", {}).items():
            var.set(1 if str(s.get(key, "1")) in ("1", "True", "نعم") else 0)
        if hasattr(self, "st_einvoice_var"):
            self.st_einvoice_var.set(1 if str(s.get("einvoice_enabled", "0")) == "1" else 0)
            self.st_env_var.set(str(s.get("einvoice_env", "test")))
        for key, var in getattr(self, "st_hr_vars", {}).items():
            var.set(str(s.get(key, "")))
        if hasattr(self, "st_scope_var"):
            cs = compliance_get_settings()
            sc = cs.get("entity_scope", "union")
            self.st_scope_var.set(f"{sc}: {'نقابة عامة' if sc == 'union' else 'لجنة نقابية'}")
        self._load_tax_brackets()

    def _load_tax_brackets(self):
        if not hasattr(self, "brackets_box") or not hasattr(self, "st_bracket_vars"):
            return
        for child in self.brackets_box.winfo_children():
            child.destroy()
        self.st_bracket_vars = []
        brackets = get_json_setting("tax_brackets", [])
        for b in brackets:
            self.add_tax_bracket_row(b.get("from", 0), b.get("to", 0), b.get("rate", 0), b.get("label", ""))

    def add_tax_bracket_row(self, frm="", to="", rate="", label=""):
        row = ttk.Frame(self.brackets_box); row.pack(fill="x", pady=2)
        vf = tk.StringVar(value=str(frm)); vt = tk.StringVar(value=str(to))
        vr = tk.StringVar(value=str(rate)); vl = tk.StringVar(value=str(label))
        ttk.Entry(row, textvariable=vl, width=26).pack(side="right", padx=3)
        ttk.Entry(row, textvariable=vr, width=8).pack(side="right", padx=3)
        ttk.Entry(row, textvariable=vt, width=12).pack(side="right", padx=3)
        ttk.Entry(row, textvariable=vf, width=12).pack(side="right", padx=3)
        ttk.Button(row, text="حذف", width=6, command=row.destroy).pack(side="right", padx=3)
        self.st_bracket_vars.append((vf, vt, vr, vl))

    def save_general_settings(self):
        if not self.can("settings"):
            self.deny(); return
        mapping = {k: v.get() for k, v in self.st_general_vars.items()}
        for k, v in self.st_notify_vars.items():
            mapping[k] = str(v.get())
        set_settings(mapping)
        self.log_audit("تعديل إعدادات عامة", str(mapping))
        messagebox.showinfo("تم", "تم حفظ إعدادات النظام العامة")

    def save_tax_settings(self):
        if not self.can("settings"):
            self.deny(); return
        mapping = {k: v.get() for k, v in self.st_tax_vars.items()}
        mapping["einvoice_enabled"] = str(self.st_einvoice_var.get())
        mapping["einvoice_env"] = self.st_env_var.get()
        set_settings(mapping)
        self.log_audit("تعديل إعدادات الضرائب", str(mapping))
        messagebox.showinfo("تم", "تم حفظ إعدادات الحسابات والضرائب")

    def save_hr_settings(self):
        if not self.can("settings"):
            self.deny(); return
        import json as _json
        mapping = {k: v.get() for k, v in self.st_hr_vars.items()}
        brackets = []
        for vf, vt, vr, vl in self.st_bracket_vars:
            try:
                brackets.append({"from": float(vf.get() or 0), "to": float(vt.get() or 0),
                                 "rate": float(vr.get() or 0), "label": vl.get()})
            except ValueError:
                continue
        mapping["tax_brackets"] = _json.dumps(brackets, ensure_ascii=False)
        set_settings(mapping)
        self.log_audit("تعديل إعدادات المرتبات", f"{len(brackets)} شريحة ضريبية")
        messagebox.showinfo("تم", "تم حفظ إعدادات الموارد البشرية")

    # ---- المستخدمون ----
    def refresh_settings_users(self):
        if not hasattr(self, "st_users_tree"):
            return
        for iid in self.st_users_tree.get_children():
            self.st_users_tree.delete(iid)
        for u in list_users():
            self.st_users_tree.insert(
                "", "end", iid=f"u{u['id']}",
                values=(u["id"], u["username"], u["full_name"] or "",
                        ROLES.get(u["role"], {}).get("label", u["role"]),
                        "نشط" if u["active"] else "موقوف",
                        u["created_at"] or ""))

    def add_settings_user(self):
        if not self.can("users_manage"):
            self.deny("إدارة المستخدمين متاحة لمدير النظام فقط.")
            return
        from tkinter import simpledialog
        uname = simpledialog.askstring("مستخدم جديد", "اسم المستخدم:", parent=self)
        if not uname:
            return
        full = simpledialog.askstring("مستخدم جديد", "الاسم بالكامل:", parent=self) or ""
        pwd = simpledialog.askstring("مستخدم جديد", "كلمة المرور:", parent=self, show="*")
        if not pwd:
            return
        role = simpledialog.askstring("مستخدم جديد",
                                      "الدور (admin / accountant / reviewer / viewer):",
                                      initialvalue="viewer", parent=self) or "viewer"
        try:
            add_user(uname.strip(), full.strip(), pwd, role.strip())
            self.log_audit("إضافة مستخدم", uname)
            self.refresh_settings_users()
            messagebox.showinfo("تم", "تم إضافة المستخدم")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def _selected_user_id(self):
        sel = self.st_users_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مستخدماً أولاً")
            return None
        return int(sel[0][1:])

    def reset_user_password(self):
        if not self.can("users_manage"):
            self.deny("إدارة المستخدمين متاحة لمدير النظام فقط.")
            return
        uid = self._selected_user_id()
        if not uid:
            return
        from tkinter import simpledialog
        pwd = simpledialog.askstring("كلمة مرور جديدة", "كلمة المرور الجديدة:",
                                     parent=self, show="*")
        if not pwd:
            return
        update_user(uid, password=pwd)
        self.log_audit("إعادة تعيين كلمة مرور", f"مستخدم رقم {uid}")
        messagebox.showinfo("تم", "تم تحديث كلمة المرور")

    def toggle_user_active(self):
        if not self.can("users_manage"):
            self.deny("إدارة المستخدمين متاحة لمدير النظام فقط.")
            return
        uid = self._selected_user_id()
        if not uid:
            return
        if self.current_user and self.current_user.get("id") == uid:
            messagebox.showwarning("تنبيه", "لا يمكنك إيقاف حسابك الحالي.")
            return
        row = next((u for u in list_users() if u["id"] == uid), None)
        new_state = not bool(row["active"]) if row else True
        update_user(uid, active=new_state)
        self.log_audit("تفعيل/إيقاف مستخدم", f"{row['username']} -> {'نشط' if new_state else 'موقوف'}")
        self.refresh_settings_users()

    # ---- سجل التدقيق ----
    def refresh_audit_log(self):
        if not hasattr(self, "audit_tree"):
            return
        for iid in self.audit_tree.get_children():
            self.audit_tree.delete(iid)
        for r in list_audit():
            self.audit_tree.insert("", "end", values=(r["ts"], r["username"], r["action"], r["details"] or ""))

    def export_audit_log(self):
        rows = [(r["ts"], r["username"], r["action"], r["details"] or "") for r in list_audit(5000)]
        self.export_simple_excel(rows, ["التاريخ/الوقت", "المستخدم", "العملية", "التفاصيل"],
                                 "سجل_التدقيق")

    # ---------- permissions ----------
    def _apply_permissions(self):
        user = self.current_user
        if not user or user.get("role") == "admin":
            return

        disabled_texts = []
        if not user_can(user.get("role"), "import"):
            disabled_texts += ["استيراد اليومية", "استيراد المرتبات", "استيراد ملف المرتبات",
                               "اختيار ملف واستيراد", "استيراد اليومية الأمريكية"]
        if not user_can(user.get("role"), "entries_edit"):
            disabled_texts += ["إضافة قيد", "تعديل القيد المحدد",
                               "فتح شاشة الإضافة", "إضافة قيد جديد",
                               "حفظ التسوية", "ترحيل الفروقات"]
        if not user_can(user.get("role"), "entries_delete"):
            disabled_texts += ["حذف القيد المحدد"]
        if not user_can(user.get("role"), "export"):
            disabled_texts += ["تصدير", "طباعة"]
        if not user_can(user.get("role"), "backup"):
            disabled_texts += ["نسخة احتياطية"]
        if not user_can(user.get("role"), "restore"):
            disabled_texts += ["استرجاع"]
        if not user_can(user.get("role"), "compliance_rescan"):
            disabled_texts += ["إعادة فحص كل القيود"]
        if not user_can(user.get("role"), "compliance_ack"):
            disabled_texts += ["اعتماد المحدد كاستثناء"]
        if not user_can(user.get("role"), "settings"):
            disabled_texts += ["حفظ حدود اللائحة", "حفظ الحدود", "حفظ إعدادات النظام العامة",
                               "حفظ إعدادات الحسابات والضرائب", "حفظ إعدادات المرتبات",
                               "استرجاع الحدود الافتراضية", "اختيار الشعار"]
        if not user_can(user.get("role"), "users_manage"):
            disabled_texts += ["إضافة مستخدم", "إعادة تعيين كلمة المرور", "تفعيل/إيقاف"]

        # تعطيل أزرار مراكز التكلفة الإدارية
        if not user_can(user.get("role"), "entries_edit"):
            for btn in getattr(self, "cost_admin_buttons", []):
                try:
                    btn.state(["disabled"])
                except Exception:
                    pass

        # تعطيل أزرار التسوية الإدارية
        for bank_key, btns in getattr(self, "settlement_admin_buttons", {}).items():
            for btn in btns.values():
                try:
                    btn.state(["disabled"])
                except Exception:
                    pass

        # تعطيل عام: كل الأزرار التي يطابق نصها قائمة الممنوعات
        def walk(widget):
            for child in widget.winfo_children():
                try:
                    text = str(child.cget("text"))
                    if any(block in text for block in disabled_texts):
                        try:
                            child.state(["disabled"])
                        except Exception:
                            pass
                except Exception:
                    pass
                walk(child)

        walk(self)

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

    # ---------- monthly summary ----------
    def refresh_monthly(self):
        years = available_years()
        self.monthly_year_combo["values"] = years
        current = self.monthly_year_var.get().strip()
        if current not in years:
            current = years[0] if years else str(datetime.today().year)
            self.monthly_year_var.set(current)
        data = monthly_summary(current or None)
        self.current_monthly = data
        self.monthly_kpis["debit"].config(text=fmt(data["total_revenues"]))
        self.monthly_kpis["credit"].config(text=fmt(data["total_expenses"]))
        self.monthly_kpis["balance"].config(text=fmt(data["net_result"]))
        state = "فائض" if data["net_result"] >= 0 else "عجز"
        self.monthly_status.config(
            text=f"سنة {data['year']} | الإيرادات: {fmt(data['total_revenues'])} | المصروفات: {fmt(data['total_expenses'])} | النتيجة: {fmt(data['net_result'])} ({state})"
        )
        self._fill_tree(self.monthly_tree, [
            (m["month_name"], fmt(m["revenues"]), fmt(m["expenses"]), fmt(m["net"]))
            for m in data["rows"]
        ])

    def export_monthly(self):
        rows = [self.monthly_tree.item(i, "values") for i in self.monthly_tree.get_children()]
        self.export_simple_excel(rows, ["الشهر", "الإيرادات", "المصروفات", "صافي الفائض / العجز"], "monthly_summary.xlsx")

    def print_monthly(self):
        rows = [self.monthly_tree.item(i, "values") for i in self.monthly_tree.get_children()]
        year = self.monthly_year_var.get().strip()
        self.print_rows_pdf(
            "الملخص الشهري للفائض / العجز",
            f"سنة {year or '-'}",
            ["الشهر", "الإيرادات", "المصروفات", "صافي الفائض / العجز"],
            rows, "monthly_summary.pdf",
        )

    def print_rev_summary_pdf(self):
        """طباعة ملخص الفائض / العجز النهائي من شاشة الإيرادات والمصروفات."""
        rows = [self.revsum_tree.item(i, "values") for i in self.revsum_tree.get_children()]
        self.print_rows_pdf(
            "ملخص الفائض / العجز النهائي",
            "إيرادات ومصروفات حسب الحساب",
            ["الحساب", "النوع", "مدين", "دائن", "الرصيد الطبيعي"],
            rows, "surplus_deficit_summary.pdf",
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
        self.refresh_monthly()
        self.refresh_compliance_log()
        self.refresh_compliance_rules()
        self.refresh_compliance_cash()
        self.refresh_compliance_live()
        self.refresh_compliance_badge()
        self.refresh_settings_users()

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
        if not self.can("entries_delete"):
            self.deny("دورك لا يسمح بحذف القيود.")
            return
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
        self.log_audit("حذف قيد", f"قيد رقم {entry_id}")
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


class LoginDialog(tk.Toplevel):
    """نافذة تسجيل الدخول التي تظهر قبل فتح النظام."""
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.result = None
        self.title("تسجيل الدخول - النظام المحاسبي")
        self.geometry("420x260")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        box = ttk.Frame(self)
        box.pack(fill="both", expand=True, padx=24, pady=20)

        ttk.Label(box, text="النظام المحاسبي المكتبي المتكامل", style="Section.TLabel").pack(pady=(0, 4))
        ttk.Label(box, text="يرجى تسجيل الدخول للمتابعة").pack(pady=(0, 14))

        ttk.Label(box, text="اسم المستخدم").pack(anchor="e")
        self.username_var = tk.StringVar()
        ent = ttk.Entry(box, textvariable=self.username_var, width=32)
        ent.pack(fill="x", pady=(2, 8))
        ent.focus_set()

        ttk.Label(box, text="كلمة المرور").pack(anchor="e")
        self.password_var = tk.StringVar()
        pent = ttk.Entry(box, textvariable=self.password_var, width=32, show="•")
        pent.pack(fill="x", pady=(2, 12))

        ttk.Label(box, text="المستخدم الافتراضي عند أول تشغيل: admin / admin", foreground="#777").pack(pady=(0, 8))

        btns = ttk.Frame(box)
        btns.pack(fill="x")
        ttk.Button(btns, text="دخول", command=self.submit).pack(side="left", padx=4)
        ttk.Button(btns, text="خروج", command=self.cancel).pack(side="left", padx=4)

        self.error = ttk.Label(box, text="", foreground="red")
        self.error.pack(pady=(6, 0))

        self.bind("<Return>", lambda e: self.submit())
        self.bind("<Escape>", lambda e: self.cancel())
        self.protocol("WM_DELETE_WINDOW", self.cancel)

    def submit(self):
        user = verify_user(self.username_var.get(), self.password_var.get())
        if user:
            self.result = user
            self.destroy()
        else:
            self.error.config(text="اسم المستخدم أو كلمة المرور غير صحيحة")

    def cancel(self):
        self.result = None
        self.destroy()


def show_login():
    """يعرض نافذة الدخول ويعيد بيانات المستخدم أو None عند الخروج."""
    root = tk.Tk()
    root.withdraw()
    dialog = LoginDialog(root)
    root.wait_window(dialog)
    user = dialog.result
    root.destroy()
    return user


if __name__ == "__main__":
    current_user = show_login()
    if current_user is None:
        sys.exit(0)
    app = AccountingApp(current_user)
    app.mainloop()
