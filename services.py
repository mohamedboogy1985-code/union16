
from db import get_connection
import re



def _normalize_arabic_text(text):
    text = str(text or "").strip()
    text = text.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    text = text.replace("ى", "ي").replace("ة", "ه")
    text = re.sub(r"[()\[\]{}\-_/+]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def extract_person_name(description):
    text = _normalize_arabic_text(description)
    if not text or "القيد الافتتاح" in text:
        return None

    tokens = text.split()
    if len(tokens) < 2:
        return None

    stop_words = {
        "عهده", "عهده", "تسويه", "تسويه", "تسويه", "تسويه", "تسوية",
        "مصروفات", "مصروف", "بدل", "وجبات", "شراء", "لشراء", "للشراء",
        "للصرف", "صرف", "سداد", "لحين", "مرور", "مكتب", "اللجنه", "اللجنة",
        "اعانات", "اعانه", "مساهمه", "مساهمة", "مرتبات", "مكافاه", "مكافأة",
        "رحله", "رحلة", "فواتير", "فاتوره", "فاتورة", "شحن", "ترخيص", "تراخيص",
        "اصلاح", "صيانه", "صيانه", "صيانة"
    }

    name_tokens = []
    for tok in tokens:
        if tok in stop_words:
            break
        if re.search(r"\d", tok):
            break
        name_tokens.append(tok)
        if len(name_tokens) >= 5:
            break

    if len(name_tokens) < 2:
        name_tokens = tokens[:min(4, len(tokens))]

    # remove trailing connectors that are not part of names
    while name_tokens and name_tokens[-1] in {"تسويه", "تسوية", "مصروفات", "بدل", "وجبات"}:
        name_tokens.pop()

    person = " ".join(name_tokens).strip()
    if len(person) < 3:
        return None
    return person

def debtors_people():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT DISTINCT je.description
        FROM journal_entries je
        JOIN journal_lines jl ON jl.entry_id = je.id
        JOIN accounts a ON a.id = jl.account_id
        WHERE a.name LIKE '%مدينون متنوعون%' OR a.name LIKE '%عهد%'
        ORDER BY je.description
    """).fetchall()
    conn.close()

    people = {}
    for r in rows:
        name = extract_person_name(r[0])
        if name:
            key = _normalize_arabic_text(name)
            people[key] = name

    return sorted(people.values())

def debtors_person_report(person_name=None, description_filter=""):
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT a.name AS account_name,
               je.entry_date,
               je.reference,
               je.description,
               jl.debit,
               jl.credit
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        JOIN journal_entries je ON je.id = jl.entry_id
        WHERE (a.name LIKE '%مدينون متنوعون%' OR a.name LIKE '%عهد%')
        ORDER BY je.entry_date, je.id, a.name
    """).fetchall()
    conn.close()

    normalized_person = _normalize_arabic_text(person_name) if person_name else ""
    normalized_filter = _normalize_arabic_text(description_filter) if description_filter else ""

    output = []
    running = 0.0
    total_debit = 0.0
    total_credit = 0.0

    for r in rows:
        item = dict(r)
        desc_norm = _normalize_arabic_text(item.get("description", ""))
        person = extract_person_name(item.get("description", ""))
        person_norm = _normalize_arabic_text(person) if person else ""
        if normalized_person and person_norm != normalized_person:
            continue
        if normalized_filter and normalized_filter not in desc_norm:
            continue

        debit = float(item.get("debit") or 0)
        credit = float(item.get("credit") or 0)
        running += debit - credit
        total_debit += debit
        total_credit += credit
        item["person_name"] = person or "-"
        item["running_balance"] = running
        output.append(item)

    return {
        "rows": output,
        "total_debit": total_debit,
        "total_credit": total_credit,
        "final_balance": running,
    }


def dashboard_summary():
    conn = get_connection()
    cur = conn.cursor()

    entries = cur.execute("SELECT COUNT(*) FROM journal_entries").fetchone()[0]
    lines = cur.execute("SELECT COUNT(*) FROM journal_lines").fetchone()[0]
    accounts = cur.execute("SELECT COUNT(*) FROM accounts").fetchone()[0]
    totals = cur.execute("SELECT COALESCE(SUM(total_debit),0), COALESCE(SUM(total_credit),0) FROM journal_entries").fetchone()
    debtors_balance = cur.execute("""
        SELECT COALESCE(SUM(jl.debit - jl.credit),0)
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        WHERE a.name LIKE '%مدينون متنوعون%'
    """).fetchone()[0]

    recent_entries = cur.execute("""
        SELECT id, entry_date, reference, description, total_debit, total_credit
        FROM journal_entries
        ORDER BY entry_date DESC, id DESC
        LIMIT 12
    """).fetchall()

    top_accounts = cur.execute("""
        SELECT a.name AS account_name,
               ROUND(SUM(jl.debit),2) AS debit,
               ROUND(SUM(jl.credit),2) AS credit,
               ROUND(SUM(jl.debit - jl.credit),2) AS balance
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        GROUP BY a.name
        ORDER BY ABS(SUM(jl.debit - jl.credit)) DESC
        LIMIT 12
    """).fetchall()

    conn.close()
    return {
        "entries": entries,
        "lines": lines,
        "accounts": accounts,
        "debits": totals[0],
        "credits": totals[1],
        "debtors_balance": debtors_balance,
        "recent_entries": [dict(x) for x in recent_entries],
        "top_accounts": [dict(x) for x in top_accounts],
    }

def list_entries(search=""):
    conn = get_connection()
    cur = conn.cursor()
    if search:
        rows = cur.execute("""
            SELECT id, entry_date, reference, description, total_debit, total_credit
            FROM journal_entries
            WHERE description LIKE ? OR reference LIKE ?
            ORDER BY entry_date, id
            LIMIT 500
        """, (f"%{search}%", f"%{search}%")).fetchall()
    else:
        rows = cur.execute("""
            SELECT id, entry_date, reference, description, total_debit, total_credit
            FROM journal_entries
            ORDER BY entry_date, id
            LIMIT 500
        """).fetchall()
    conn.close()
    return [dict(x) for x in rows]

def get_entry_lines(entry_id):
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT a.name AS account_name, jl.debit, jl.credit
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        WHERE jl.entry_id = ?
        ORDER BY a.name
    """, (entry_id,)).fetchall()
    conn.close()
    return [dict(x) for x in rows]

def debtors_accounts():
    return debtors_people()

def debtors_report(account_name=None, description_filter=""):
    return debtors_person_report(account_name, description_filter)

def debtors_account_groups():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT name
        FROM accounts
        WHERE name LIKE '%مدينون متنوعون%' OR name LIKE '%عهد%'
        ORDER BY name
    """).fetchall()
    conn.close()
    return [r[0] for r in rows]

def debtors_report_by_account(account_name=None, description_filter=""):
    return account_statement_report(["%مدينون متنوعون%", "%عهد%"], account_name, description_filter)


def revenue_expense_accounts():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT name
        FROM accounts
        WHERE name LIKE '%إيراد%' OR name LIKE '%ايراد%' OR name LIKE '%الإيرادات%'
           OR name LIKE '%مصروف%' OR name LIKE '%نفقة%'
        ORDER BY name
    """).fetchall()
    conn.close()
    return [r[0] for r in rows]


def revenue_expense_report(account_name=None, description_filter=""):
    return account_statement_report([
        "%إيراد%", "%ايراد%", "%الإيرادات%", "%مصروف%", "%نفقة%"
    ], account_name, description_filter)


def revenue_expense_final_summary():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT a.name AS account_name,
               ROUND(SUM(jl.debit),2) AS debit,
               ROUND(SUM(jl.credit),2) AS credit
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        WHERE a.name LIKE '%إيراد%' OR a.name LIKE '%ايراد%' OR a.name LIKE '%الإيرادات%'
           OR a.name LIKE '%مصروف%' OR a.name LIKE '%نفقة%'
        GROUP BY a.name
        ORDER BY a.name
    """).fetchall()
    conn.close()

    result_rows = []
    total_revenues = 0.0
    total_expenses = 0.0
    for r in rows:
        item = dict(r)
        name = item["account_name"]
        debit = float(item["debit"] or 0)
        credit = float(item["credit"] or 0)
        is_expense = ("مصروف" in name) or ("نفقة" in name)
        natural_balance = (debit - credit) if is_expense else (credit - debit)
        account_type = "مصروفات" if is_expense else "إيرادات"
        if is_expense:
            total_expenses += natural_balance
        else:
            total_revenues += natural_balance
        result_rows.append({
            "account_name": name,
            "account_type": account_type,
            "debit": debit,
            "credit": credit,
            "natural_balance": natural_balance,
        })

    return {
        "rows": result_rows,
        "total_revenues": total_revenues,
        "total_expenses": total_expenses,
        "net_result": total_revenues - total_expenses,
    }

def account_statement_report(name_patterns, account_name=None, description_filter=""):
    conn = get_connection()
    cur = conn.cursor()

    where_patterns = " OR ".join(["a.name LIKE ?" for _ in name_patterns])
    params = list(name_patterns)
    sql = f"""
        SELECT a.name AS account_name,
               je.entry_date,
               je.reference,
               je.description,
               jl.debit,
               jl.credit
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        JOIN journal_entries je ON je.id = jl.entry_id
        WHERE ({where_patterns})
    """

    if account_name:
        sql += " AND a.name = ?"
        params.append(account_name)

    if description_filter:
        sql += " AND COALESCE(je.description, '') LIKE ?"
        params.append(f"%{description_filter}%")

    sql += " ORDER BY a.name, je.entry_date, je.id LIMIT 5000"
    rows = cur.execute(sql, params).fetchall()
    conn.close()

    running = 0
    total_debit = 0
    total_credit = 0
    output = []
    for r in rows:
        item = dict(r)
        debit = item["debit"] or 0
        credit = item["credit"] or 0
        running += debit - credit
        total_debit += debit
        total_credit += credit
        item["running_balance"] = running
        output.append(item)

    return {
        "rows": output,
        "total_debit": total_debit,
        "total_credit": total_credit,
        "final_balance": running,
    }

def smart_vouchers():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT je.id AS entry_id,
               je.entry_date,
               je.reference,
               je.description,
               MAX(CASE WHEN src.debit = 0 AND src.credit > 0 THEN a1.name END) AS source_account,
               MAX(CASE WHEN dst.debit > 0 THEN a2.name END) AS target_account,
               MAX(CASE WHEN dst.debit > 0 THEN dst.debit END) AS amount
        FROM journal_entries je
        JOIN journal_lines src ON src.entry_id = je.id
        JOIN accounts a1 ON a1.id = src.account_id
        JOIN journal_lines dst ON dst.entry_id = je.id
        JOIN accounts a2 ON a2.id = dst.account_id
        WHERE (a1.name LIKE '%بنك%' OR a1.name LIKE '%خزينة%')
          AND src.credit > 0
          AND dst.debit > 0
        GROUP BY je.id, je.entry_date, je.reference, je.description
        ORDER BY je.entry_date DESC, je.id DESC
        LIMIT 500
    """).fetchall()
    conn.close()

    result = []
    for r in rows:
        item = dict(r)
        item["beneficiary"] = item["description"]
        result.append(item)
    return result

def ledger_accounts():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("SELECT name FROM accounts ORDER BY name").fetchall()
    conn.close()
    return [r[0] for r in rows]

def ledger_for_account(account_name):
    if not account_name:
        return []
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT je.entry_date, je.reference, je.description, jl.debit, jl.credit
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.entry_id
        JOIN accounts a ON a.id = jl.account_id
        WHERE a.name = ?
        ORDER BY je.entry_date, je.id
    """, (account_name,)).fetchall()
    conn.close()

    running = 0
    output = []
    for r in rows:
        running += (r["debit"] or 0) - (r["credit"] or 0)
        item = dict(r)
        item["running_balance"] = running
        output.append(item)
    return output

def trial_balance():
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT a.name AS account_name,
               ROUND(SUM(jl.debit),2) AS debit,
               ROUND(SUM(jl.credit),2) AS credit,
               CASE WHEN SUM(jl.debit - jl.credit) > 0 THEN ROUND(SUM(jl.debit - jl.credit),2) ELSE 0 END AS net_debit,
               CASE WHEN SUM(jl.debit - jl.credit) < 0 THEN ROUND(ABS(SUM(jl.debit - jl.credit)),2) ELSE 0 END AS net_credit
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        GROUP BY a.name
        HAVING ROUND(SUM(jl.debit),2) <> 0 OR ROUND(SUM(jl.credit),2) <> 0
        ORDER BY a.name
    """).fetchall()
    conn.close()
    return [dict(x) for x in rows]


def latest_payroll_import():
    conn = get_connection()
    cur = conn.cursor()
    row = cur.execute("""
        SELECT id, source_file, imported_at, payroll_month, employees_count,
               gross_total, deductions_total, net_total, salary_advance_total, bank_loan_total
        FROM payroll_imports
        ORDER BY id DESC
        LIMIT 1
    """).fetchone()
    conn.close()
    return dict(row) if row else {
        "payroll_month": "-",
        "employees_count": 0,
        "gross_total": 0,
        "deductions_total": 0,
        "net_total": 0,
        "salary_advance_total": 0,
        "bank_loan_total": 0,
    }

def payroll_rows(search=""):
    conn = get_connection()
    cur = conn.cursor()
    if search:
        rows = cur.execute("""
            SELECT employee_no, employee_name, payroll_month, gross_total, total_deductions,
                   net_pay, salary_advance_installment, bank_loan_installment,
                   insurance_employee, tax_amount
            FROM payroll_rows
            WHERE employee_name LIKE ?
            ORDER BY employee_no, employee_name
        """, (f"%{search}%",)).fetchall()
    else:
        rows = cur.execute("""
            SELECT employee_no, employee_name, payroll_month, gross_total, total_deductions,
                   net_pay, salary_advance_installment, bank_loan_installment,
                   insurance_employee, tax_amount
            FROM payroll_rows
            ORDER BY employee_no, employee_name
        """).fetchall()
    conn.close()
    return [dict(x) for x in rows]

def workers_advances_report(search=""):
    conn = get_connection()
    cur = conn.cursor()
    sql = """
        SELECT employee_no, employee_name, payroll_month,
               salary_advance_installment AS amount
        FROM payroll_rows
        WHERE salary_advance_installment > 0
    """
    params = []
    if search:
        sql += " AND employee_name LIKE ?"
        params.append(f"%{search}%")
    sql += " ORDER BY employee_no, employee_name"
    rows = cur.execute(sql, params).fetchall()
    conn.close()
    total = sum(float(r["amount"] or 0) for r in rows)
    return {"rows": [dict(x) for x in rows], "total": total}

def bank_loans_report(search=""):
    conn = get_connection()
    cur = conn.cursor()
    sql = """
        SELECT employee_no, employee_name, payroll_month,
               bank_loan_installment AS amount
        FROM payroll_rows
        WHERE bank_loan_installment > 0
    """
    params = []
    if search:
        sql += " AND employee_name LIKE ?"
        params.append(f"%{search}%")
    sql += " ORDER BY employee_no, employee_name"
    rows = cur.execute(sql, params).fetchall()
    conn.close()
    total = sum(float(r["amount"] or 0) for r in rows)
    return {"rows": [dict(x) for x in rows], "total": total}


def _get_or_create_account(cur, name, category="أخرى", normal_side="debit"):
    row = cur.execute("SELECT id, name, code, category, normal_side FROM accounts WHERE name = ?", (name,)).fetchone()
    if row:
        return dict(row)
    code = f"A{cur.execute('SELECT COALESCE(MAX(id),0)+1 FROM accounts').fetchone()[0]:04d}"
    cur.execute(
        "INSERT INTO accounts(code, name, category, normal_side) VALUES (?, ?, ?, ?)",
        (code, name, category, normal_side)
    )
    row = cur.execute("SELECT id, name, code, category, normal_side FROM accounts WHERE id = ?", (cur.lastrowid,)).fetchone()
    return dict(row)

def settlement_bank_mapping():
    return {
        "bank_misr": {"title": "بنك مصر", "match": ["بنك مصر"], "fallback_codes": ["1002", "A0001"]},
        "idb": {"title": "بنك العمال", "match": ["بنك العمال", "التنمية الصناعية", "بنك التنمية", "التنمية"], "fallback_codes": ["1003"]},
    }

def bank_account_info(bank_key):
    mapping = settlement_bank_mapping().get(bank_key)
    if not mapping:
        return None
    conn = get_connection()
    cur = conn.cursor()
    row = None
    for phrase in mapping["match"]:
        row = cur.execute(
            "SELECT id, code, name, category, normal_side FROM accounts WHERE name LIKE ? ORDER BY id LIMIT 1",
            (f"%{phrase}%",)
        ).fetchone()
        if row:
            break
    if not row:
        for code in mapping.get("fallback_codes", []):
            row = cur.execute(
                "SELECT id, code, name, category, normal_side FROM accounts WHERE code = ? ORDER BY id LIMIT 1",
                (code,)
            ).fetchone()
            if row:
                break
    conn.close()
    return dict(row) if row else None

def bank_settlement_report(bank_key, settlement_year=None):
    from datetime import datetime
    year = int(settlement_year or datetime.now().year)
    start = f"{year}-01-01"
    end = f"{year}-12-31"
    bank = bank_account_info(bank_key)
    if not bank:
        return {
            "bank_key": bank_key, "bank_name": settlement_bank_mapping().get(bank_key, {}).get("title", bank_key),
            "settlement_year": year, "account_found": False,
            "opening_balance": 0, "receipts": 0, "payments": 0, "book_balance_end": 0,
            "checks_under_collection": 0, "uncashed_check1": 0, "uncashed_check2": 0,
            "final_bank_balance": 0, "bank_statement_balance": 0, "discrepancy": 0,
        }
    conn = get_connection()
    cur = conn.cursor()
    opening = cur.execute("""
        SELECT COALESCE(SUM(jl.debit - jl.credit), 0)
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.entry_id
        WHERE jl.account_id = ? AND je.entry_date < ?
    """, (bank["id"], start)).fetchone()[0] or 0
    receipts = cur.execute("""
        SELECT COALESCE(SUM(jl.debit), 0)
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.entry_id
        WHERE jl.account_id = ? AND je.entry_date BETWEEN ? AND ?
    """, (bank["id"], start, end)).fetchone()[0] or 0
    payments = cur.execute("""
        SELECT COALESCE(SUM(jl.credit), 0)
        FROM journal_lines jl
        JOIN journal_entries je ON je.id = jl.entry_id
        WHERE jl.account_id = ? AND je.entry_date BETWEEN ? AND ?
    """, (bank["id"], start, end)).fetchone()[0] or 0
    saved = cur.execute("""
        SELECT checks_under_collection, uncashed_check1, uncashed_check2, bank_statement_balance, notes, updated_at
        FROM bank_settlements
        WHERE bank_key = ? AND settlement_year = ?
        ORDER BY id DESC LIMIT 1
    """, (bank_key, year)).fetchone()
    checks = float(saved["checks_under_collection"] or 0) if saved else 0.0
    u1 = float(saved["uncashed_check1"] or 0) if saved else 0.0
    u2 = float(saved["uncashed_check2"] or 0) if saved else 0.0
    stmt = float(saved["bank_statement_balance"] or 0) if saved else 0.0
    notes = saved["notes"] if saved else ""
    updated_at = saved["updated_at"] if saved else ""
    conn.close()
    book_end = float(opening) + float(receipts) - float(payments)
    final_bal = book_end + checks + u1 + u2
    discrepancy = stmt - final_bal
    return {
        "bank_key": bank_key,
        "bank_name": bank["name"],
        "account_found": True,
        "account_id": bank["id"],
        "account_code": bank["code"],
        "settlement_year": year,
        "opening_balance": float(opening),
        "receipts": float(receipts),
        "payments": float(payments),
        "book_balance_end": float(book_end),
        "checks_under_collection": checks,
        "uncashed_check1": u1,
        "uncashed_check2": u2,
        "final_bank_balance": float(final_bal),
        "bank_statement_balance": stmt,
        "discrepancy": float(discrepancy),
        "notes": notes,
        "updated_at": updated_at,
    }

def save_bank_settlement(bank_key, settlement_year, checks_under_collection=0, uncashed_check1=0, uncashed_check2=0, bank_statement_balance=0, notes=""):
    info = settlement_bank_mapping().get(bank_key, {"title": bank_key})
    conn = get_connection()
    cur = conn.cursor()
    updated_at = __import__("datetime").datetime.now().isoformat(timespec="seconds")
    cur.execute("""
        INSERT INTO bank_settlements(bank_key, bank_name, settlement_year, checks_under_collection, uncashed_check1, uncashed_check2, bank_statement_balance, notes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(bank_key, settlement_year) DO UPDATE SET
            bank_name=excluded.bank_name,
            checks_under_collection=excluded.checks_under_collection,
            uncashed_check1=excluded.uncashed_check1,
            uncashed_check2=excluded.uncashed_check2,
            bank_statement_balance=excluded.bank_statement_balance,
            notes=excluded.notes,
            updated_at=excluded.updated_at
    """, (bank_key, info["title"], int(settlement_year), float(checks_under_collection or 0), float(uncashed_check1 or 0), float(uncashed_check2 or 0), float(bank_statement_balance or 0), notes or "", updated_at))
    conn.commit()
    conn.close()
    return "تم حفظ التسوية"


def bank_settlement_history(bank_key):
    conn = get_connection()
    cur = conn.cursor()
    rows = cur.execute("""
        SELECT bank_key, bank_name, settlement_year, checks_under_collection, uncashed_check1,
               uncashed_check2, bank_statement_balance, notes, updated_at
        FROM bank_settlements
        WHERE bank_key = ?
        ORDER BY settlement_year DESC, updated_at DESC
    """, (bank_key,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def post_bank_settlement_adjustment(bank_key, settlement_year=None):
    report = bank_settlement_report(bank_key, settlement_year)
    if not report.get("account_found"):
        raise ValueError("لم يتم العثور على حساب البنك داخل دليل الحسابات")
    discrepancy = float(report.get("discrepancy") or 0)
    if abs(discrepancy) < 0.01:
        raise ValueError("لا يوجد فرق يحتاج إلى ترحيل")
    conn = get_connection()
    cur = conn.cursor()
    bank_row = cur.execute("SELECT id, name, code, category, normal_side FROM accounts WHERE id = ?", (report["account_id"],)).fetchone()
    bank = dict(bank_row)
    revenue_acc = _get_or_create_account(cur, "فروق تسوية بنكية دائنة", "إيرادات", "credit")
    expense_acc = _get_or_create_account(cur, "فروق تسوية بنكية مدينة", "مصروفات", "debit")
    today = __import__("datetime").datetime.now().strftime("%Y-%m-%d")
    ref = f"BANK-SETTLE-{bank_key.upper()}-{report['settlement_year']}"
    desc = f"قيد تسوية بنكية - {report['bank_name']} - سنة {report['settlement_year']}"
    amount = abs(discrepancy)
    cur.execute("""
        INSERT INTO journal_entries(entry_date, reference, description, total_debit, total_credit, source_row, import_id)
        VALUES (?, ?, ?, ?, ?, NULL, NULL)
    """, (today, ref, desc, amount, amount))
    entry_id = cur.lastrowid
    if discrepancy > 0:
        lines = [
            (entry_id, bank["id"], amount, 0.0, "إيداعات بنكية غير مسجلة"),
            (entry_id, revenue_acc["id"], 0.0, amount, "تسوية رصيد البنك"),
        ]
    else:
        lines = [
            (entry_id, expense_acc["id"], amount, 0.0, "تسوية رصيد البنك"),
            (entry_id, bank["id"], 0.0, amount, "مصاريف بنكية غير مسجلة"),
        ]
    cur.executemany("""
        INSERT INTO journal_lines(entry_id, account_id, debit, credit, line_description)
        VALUES (?, ?, ?, ?, ?)
    """, lines)
    conn.commit()
    conn.close()
    return f"تم ترحيل قيد التسوية بمبلغ {amount:,.2f}"

