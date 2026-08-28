# -*- coding: utf-8 -*-
"""
محرك الرقابة على اللائحة المالية
================================
يقوم هذا الموديول بفحص عمليات الصرف والإيراد وفقاً لأحكام اللائحة المالية
للنقابة العامة للعاملين بصناعات البناء والأخشاب وصنع مواد البناء ولجانها النقابية.

كل قاعدة في اللائحة لها كود ووصف ومادة مرجعية، وحدودها قابلة للتعديل من
شاشة «الرقابة المالية» دون الحاجة لتعديل الكود.

نتيجة الفحص:
    []                      -> القيد سليم ومطابق للائحة
    [{"level": "warning", ...}]  -> مخالفات تستوجب إنذار على الشاشة
    [{"level": "info", ...}]     -> تنبيهات إرشادية (موافقات مطلوبة)
"""

import os
import sqlite3
from datetime import datetime, timedelta

from db import get_connection

# ---------------------------------------------------------------------------
# 1) التعريفات الافتراضية للحدود المنصوص عليها في اللائحة المالية
#    القيم بالجنيه المصري ما لم يُذكر خلاف ذلك.
#    entity = "union"  -> النقابة العامة
#    entity = "committee" -> اللجنة النقابية
# ---------------------------------------------------------------------------

DEFAULT_SETTINGS = {
    # نطاق الجهة: نقابة عامة أم لجنة نقابية (يؤثر على جميع الحدود)
    "entity_scope": "union",          # union | committee

    # مادة (6): سقف الرصيد النقدي بالخزينة كسلفة مستديمة
    "cash_on_hand_union": 50000.0,
    "cash_on_hand_committee": 20000.0,

    # مادة (9): سقف الصرف النقدي في غرض واحد
    "cash_payment_per_purpose_union": 20000.0,
    "cash_payment_per_purpose_committee": 10000.0,

    # مادة (37): بدل السفر الداخلي عن الليلة (حد أدنى) + أقصى زيادة 100%
    "travel_allowance_daily_union": 200.0,
    "travel_allowance_daily_committee": 100.0,
    "travel_allowance_max_increase_pct": 100.0,   # المادة 37: لا تزيد الزيادة عن 100%
    # مادة (37/ب): تخفيض 25% إذا تحملت المنظمة المبيت
    "travel_housing_discount_pct": 25.0,
    # مادة (36): السفر بالطائرة السياحية يتطلب موافقة رئيس مجلس الإدارة

    # مادة (39): بدل الانتقال الشهري الثابت
    "monthly_transport_allowance_cap": 300.0,
    # مادة (40): بدل الأعباء الشهري
    "monthly_burden_allowance_cap": 500.0,
    # مادة (44): مكافأة التفرغ بحد أقصى 30% من الأجر الأساسي
    "secondment_bonus_pct_cap": 30.0,

    # مادة (50): هدايا الوفود المسافرة للخارج
    "delegation_gifts_cap": 5000.0,
    "delegation_gifts_exception_cap": 10000.0,   # بقرار رئيس المنظمة
    # مادة (51): مصروف جيب الوفد المستضاف (بالدولار عن الليلة)
    "hosted_pocket_money_usd_cap": 100.0,

    # مادة (61): حدود الشراء وتنفيذ الأعمال — النقابة العامة
    "proc_direct_order_union": 50000.0,          # أمر مباشر حتى 50 ألف
    "proc_practice_union": 200000.0,             # ممارسة حتى 200 ألف
    "proc_limited_tender_union": 500000.0,       # مناقصة محدودة حتى 500 ألف
    # أكثر من ذلك: مناقصة عامة
    # مادة (61): حدود الشراء — اللجنة النقابية
    "proc_direct_order_committee": 20000.0,      # أمر مباشر حتى 20 ألف
    "proc_practice_committee": 100000.0,         # ممارسة حتى 100 ألف
    "proc_limited_tender_committee": 250000.0,   # مناقصة محدودة حتى 250 ألف
    # أكثر من ذلك: مناقصة عامة

    # مادة (72): الدفعة المقدمة للمقاول بحد أقصى 25% مقابل خطاب ضمان
    "contract_advance_pct_cap": 25.0,
    # مادة (72): الدفعات تحت الحساب بحد أقصى 95% من الأعمال المنفذة
    "contract_progress_payment_pct_cap": 95.0,
    # مادة (73): غرامة التأخير لا تتجاوز 15% للمقاولات / 4% للتوريد (مرجعية)
    "delay_penalty_contracts_pct_cap": 15.0,
    "delay_penalty_supply_pct_cap": 4.0,

    # مادة (77): سداد 30% من ثمن المنقولات المباعة فور رسو المزاد
    "auction_moveable_down_payment_pct": 30.0,
    # مادة (78): سداد 10% من ثمن العقارات فور رسو المزاد والباقي خلال 3 أشهر
    "auction_realestate_down_payment_pct": 10.0,

    # مادة (2): توزيع حصيلة الاشتراكات (للرقابة على الإيرادات)
    "subscription_union_share_pct": 10.0,        # توريد للاتحاد النقابي إن وُجد
    "subscription_committee_share_pct": 60.0,
    "subscription_general_share_pct": 20.0,

    # مادة (13): تحويل الشيكات والحوالات للبنك في اليوم التالي على الأكثر
    "check_deposit_deadline_days": 1,
    # مادة (18): لا تُبقى إيصالات معلقة بالخزينة أكثر من شهر
    "pending_receipt_max_days": 30,
}

# ترتيب الحدود الخاص بالنقابة العامة / اللجنة النقابية لعرضها في الشاشة
SETTING_LABELS = {
    "entity_scope": ("نطاق الجهة", "نوع الجهة التي تعمل عليها (نقابة عامة / لجنة نقابية)", "عام"),
    "cash_on_hand_union": ("سقف الخزينة - نقابة عامة", "مادة (6): لا يزيد الرصيد النقدي بالخزينة كسلفة مستديمة عن 50 ألف جنيه", "خزينة"),
    "cash_on_hand_committee": ("سقف الخزينة - لجنة نقابية", "مادة (6): 20 ألف جنيه بخزينة اللجنة النقابية", "خزينة"),
    "cash_payment_per_purpose_union": ("سقف الصرف النقدي لغرض واحد - نقابة", "مادة (9): لا يزيد مجموع المنصرف نقداً في غرض واحد على 20 ألف جنيه", "صرف نقدي"),
    "cash_payment_per_purpose_committee": ("سقف الصرف النقدي لغرض واحد - لجنة", "مادة (9): 10 آلاف جنيه للجنة النقابية", "صرف نقدي"),
    "travel_allowance_daily_union": ("بدل السفر اليومي - نقابة", "مادة (37): حد أدنى 200 جنيه عن الليلة", "بدلات"),
    "travel_allowance_daily_committee": ("بدل السفر اليومي - لجنة", "مادة (37): 100 جنيه للجان النقابية عن الليلة", "بدلات"),
    "travel_allowance_max_increase_pct": ("أقصى زيادة في بدل السفر %", "مادة (37): لا تزيد الزيادة على 100% من الحد الأدنى بقرار مجلس الإدارة", "بدلات"),
    "travel_housing_discount_pct": ("تخفيض البدل عند تحمل المبيت %", "مادة (37/ب): تخفيض 25% عند تحمل المنظمة للمبيت", "بدلات"),
    "monthly_transport_allowance_cap": ("بدل الانتقال الشهري", "مادة (39): لا يجاوز 300 جنيه شهرياً", "بدلات"),
    "monthly_burden_allowance_cap": ("بدل الأعباء الشهري", "مادة (40): لا يجاوز 500 جنيه شهرياً", "بدلات"),
    "secondment_bonus_pct_cap": ("مكافأة التفرغ % من الأجر الأساسي", "مادة (44): بما لا يجاوز 30% من الأجر الأساسي", "بدلات"),
    "delegation_gifts_cap": ("هدايا الوفود للخارج", "مادة (50): لا تتجاوز 5000 جنيه للوفد", "علاقات دولية"),
    "delegation_gifts_exception_cap": ("هدايا الوفود - حد استثنائي", "مادة (50): بقرار رئيس المنظمة وبما لا يجاوز 10000 جنيه", "علاقات دولية"),
    "hosted_pocket_money_usd_cap": ("مصروف جيب الوفد المستضاف (دولار/ليلة)", "مادة (51): بحد أقصى 100 دولار عن الليلة", "علاقات دولية"),
    "proc_direct_order_union": ("حد الأمر المباشر - نقابة", "مادة (61): حتى 50 ألف جنيه", "مشتريات"),
    "proc_practice_union": ("حد الممارسة - نقابة", "مادة (61): أكثر من 50 ألف حتى 200 ألف جنيه", "مشتريات"),
    "proc_limited_tender_union": ("حد المناقصة المحدودة - نقابة", "مادة (61): أكثر من 200 ألف حتى 500 ألف جنيه، وبعدها مناقصة عامة", "مشتريات"),
    "proc_direct_order_committee": ("حد الأمر المباشر - لجنة", "مادة (61): حتى 20 ألف جنيه", "مشتريات"),
    "proc_practice_committee": ("حد الممارسة - لجنة", "مادة (61): أكثر من 20 ألف حتى 100 ألف جنيه", "مشتريات"),
    "proc_limited_tender_committee": ("حد المناقصة المحدودة - لجنة", "مادة (61): أكثر من 100 ألف حتى 250 ألف جنيه، وبعدها مناقصة عامة", "مشتريات"),
    "contract_advance_pct_cap": ("الدفعة المقدمة للمقاول %", "مادة (72): لا تزيد على 25% من قيمة التعاقد مقابل خطاب ضمان", "تعاقدات"),
    "contract_progress_payment_pct_cap": ("الدفعة تحت الحساب %", "مادة (72): بحد أقصى 95% من قيمة الأعمال المنفذة فعلاً", "تعاقدات"),
    "delay_penalty_contracts_pct_cap": ("غرامة تأخير المقاولات %", "مادة (73): لا يجاوز مجموع الغرامة 15%", "تعاقدات"),
    "delay_penalty_supply_pct_cap": ("غرامة تأخير التوريد %", "مادة (73): 4% لعقود التوريد", "تعاقدات"),
    "auction_moveable_down_payment_pct": ("دفعة المزاد للمنقولات %", "مادة (77): سداد 30% فور رسو المزاد", "مبيعات"),
    "auction_realestate_down_payment_pct": ("دفعة المزاد للعقارات %", "مادة (78): سداد 10% فور الرسو والباقي خلال 3 أشهر", "مبيعات"),
    "subscription_union_share_pct": ("نصيب الاتحاد من الاشتراكات %", "مادة (2): توريد 10% للاتحاد النقابي إن وجد", "إيرادات"),
    "subscription_committee_share_pct": ("نصيب اللجنة من الاشتراكات %", "مادة (2): 60% للجنة النقابية", "إيرادات"),
    "subscription_general_share_pct": ("نصيب النقابة العامة من الاشتراكات %", "مادة (2): 20% للنقابة العامة", "إيرادات"),
    "check_deposit_deadline_days": ("مهلة إيداع الشيكات بالبنك (أيام)", "مادة (13): اليوم التالي على الأكثر", "إيرادات"),
    "pending_receipt_max_days": ("مدة تعليق الإيصالات بالخزينة (أيام)", "مادة (18): لا تزيد على شهر", "إيرادات"),
}

# ---------------------------------------------------------------------------
# 2) إعدادات الحدود (جدول compliance_settings)
# ---------------------------------------------------------------------------

def init_compliance_tables():
    """إنشاء جداول الرقابة المالية وتعبئة الحدود الافتراضية لأول مرة."""
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS compliance_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS compliance_violations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER,
            entry_date TEXT,
            entry_ref TEXT,
            rule_code TEXT,
            article TEXT,
            severity TEXT,            -- violation | approval | info
            message TEXT,
            amount REAL,
            checked_at TEXT,
            acknowledged INTEGER DEFAULT 0,
            ack_note TEXT
        )""")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compl_entry ON compliance_violations(entry_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_compl_date ON compliance_violations(entry_date)")

    existing = {r[0] for r in cur.execute("SELECT key FROM compliance_settings").fetchall()}
    for k, v in DEFAULT_SETTINGS.items():
        if k not in existing:
            cur.execute("INSERT INTO compliance_settings(key, value) VALUES (?,?)", (k, str(v)))
    conn.commit()
    conn.close()


def get_settings():
    """إرجاع كل الحدود كقاموس (القيم الرقمية تُحوّل تلقائياً)."""
    conn = get_connection()
    try:
        rows = conn.execute("SELECT key, value FROM compliance_settings").fetchall()
    except sqlite3.OperationalError:
        init_compliance_tables()
        rows = conn.execute("SELECT key, value FROM compliance_settings").fetchall()
    conn.close()
    out = {}
    for r in rows:
        key, val = r[0], r[1]
        default = DEFAULT_SETTINGS.get(key)
        if isinstance(default, bool):
            out[key] = str(val).lower() in ("1", "true", "نعم")
        elif isinstance(default, float) or isinstance(default, int):
            try:
                out[key] = float(val)
            except (TypeError, ValueError):
                out[key] = default
        else:
            out[key] = val
    # ضمان وجود أي مفتاح ناقص
    for k, v in DEFAULT_SETTINGS.items():
        out.setdefault(k, v)
    return out


def save_settings(new_values):
    conn = get_connection()
    cur = conn.cursor()
    for k, v in new_values.items():
        cur.execute("INSERT INTO compliance_settings(key,value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
    conn.commit()
    conn.close()


def reset_settings():
    conn = get_connection()
    cur = conn.cursor()
    for k, v in DEFAULT_SETTINGS.items():
        cur.execute("INSERT INTO compliance_settings(key,value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# 3) أدوات مساعدة لتحليل القيد
# ---------------------------------------------------------------------------

BANK_KEYWORDS = ("بنك", "البنك")
CASH_KEYWORDS = ("خزينة", "الخزينة", "نقدية", "نقديه", "صندوق", "الصندوق")
SUPPLIER_KEYWORDS = ("مورد", "المورد", "موردين", "حساب مورد", "مستخلص", "دفعة مقدمة لمقاول")
# إيرادات اشتراكات الأعضاء (تُستثنى اشتراكات الجرائد والاشتراكات الأخرى)
SUBSCRIPTION_KEYWORDS = ("اشتراكات الأعضاء", "اشتراكات الاعضاء", "إشتراكات الأعضاء",
                         "اشتراك أعضاء", "اشتراكات العضوية", "حصيلة اشتراكات",
                         "اشتراكات اللجان", "اشتراكات العمال", "اشتراكات شركات",
                         "إيرادات من اشتراكات")
# كلمات تدل على تعاقد/شراء/أعمال (تُستخدم لتضييق نطاق القواعد)
CONTRACT_KEYWORDS = ("تعاقد", "تعاقدات", "مقاول", "مقاولات", "شراء أصل", "بيع أصل",
                     "مناقصة", "ممارسة", "توريد مهمات", "تنفيذ أعمال")
# تضارب المصالح (مادة 59) — كلمات تدل على طرف التعاقد
CONFLICT_PARTY_KEYWORDS = ("عضو مجلس الإدارة", "أعضاء مجلس الإدارة", "عضو مجلس ادارة",
                           "أقارب حتى الدرجة", "زوجة عضو", "نجل عضو", "ابن عضو مجلس",
                           "ابنة عضو مجلس")
# بنود صرف لا تخضع لحدود المشتريات (مادة 61) لأنها ليست شراء/تنفيذ أعمال
PROCUREMENT_EXCLUDE_KEYWORDS = (
    "مرتب", "مرتبات", "أجور", "اجور", "رواتب", "راتب", "مكافأة", "مكافاة", "مكافآت",
    "منحة", "منحه", "بدلات", "بدل ", "تأمينات", "تامينات", "ضريبة", "ضريبه", "قرض",
    "وديعة", "اذون خزانة", "أذون خزانة", "شهادات استثمار", "استثمار", "سلف", "سلفيات",
    "عهدة", "عهده", "اشتراكات", "إعانة", "اعانة", "إعانات", "دعم", "هبات", "تبرع",
    "القيد الافتتاحى", "رصيد افتتاحي", "رصيد مرحل", "تحويل بين", "تحويل داخلي",
)


def _amount(v):
    try:
        return round(float(v or 0), 2)
    except (TypeError, ValueError):
        return 0.0


def _entry_lines(conn, entry_id):
    return conn.execute("""
        SELECT a.name AS account, a.category, jl.debit, jl.credit, jl.line_description
        FROM journal_lines jl
        JOIN accounts a ON a.id = jl.account_id
        WHERE jl.entry_id = ?
        ORDER BY jl.id
    """, (entry_id,)).fetchall()


def _is_cash_account(name):
    return any(k in name for k in CASH_KEYWORDS)


def _is_bank_account(name):
    return any(k in name for k in BANK_KEYWORDS)


def _is_cash_or_bank(name):
    return _is_cash_account(name) or _is_bank_account(name)


def _entry_kind(lines):
    """
    تحديد طبيعة القيد:
      'payment'  -> صرف (خروج أموال من بنك/خزينة: الجانب الدائن لحساب نقدي/بنكي)
      'receipt'  -> تحصيل/إيراد (دخول أموال لبنك/خزينة: الجانب المدين لحساب نقدي/بنكي)
      'other'
    """
    cash_out = 0.0   # دائن على بنك/خزينة = صرف
    cash_in = 0.0    # مدين على بنك/خزينة = تحصيل
    for ln in lines:
        name = ln["account"] or ""
        if _is_cash_or_bank(name):
            cash_out += _amount(ln["credit"])
            cash_in += _amount(ln["debit"])
    if cash_out > 0 and cash_out >= cash_in:
        return "payment", cash_out
    if cash_in > 0:
        return "receipt", cash_in
    return "other", max(cash_in, cash_out)


def _payment_targets(lines):
    """الحسابات المستفيدة من الصرف (المدينة غير النقدية/البنكية)."""
    targets = []
    for ln in lines:
        name = ln["account"] or ""
        if not _is_cash_or_bank(name) and _amount(ln["debit"]) > 0:
            targets.append({"account": name, "amount": _amount(ln["debit"]),
                            "desc": ln.get("desc") or ln.get("line_description") or "",
                            "category": ln.get("category") or ""})
    return targets


def _payment_source_cash(lines):
    """المبلغ المخصوم من الخزينة نقداً (وليس من البنك)."""
    total = 0.0
    for ln in lines:
        if _is_cash_account(ln["account"] or ""):
            total += _amount(ln["credit"])
    return total


def _text_blob(description, lines):
    parts = [description or ""]
    for ln in lines:
        parts.append(ln.get("desc") or ln.get("line_description") or "")
        parts.append(ln.get("account") or "")
    return " ".join(parts)


# ---------------------------------------------------------------------------
# 4) فحص قيد صرف/إيراد واحد
# ---------------------------------------------------------------------------

def check_entry(entry_id=None, date=None, reference=None, description=None,
                draft_lines=None):
    """
    فحص قيد (قائم من قاعدة البيانات عبر entry_id، أو مسودة قيد تحت التحرير عبر
    draft_lines = [{"account":..., "debit":..., "credit":..., "desc":...}]).

    تُرجع قائمة نتائج:
        [{"severity": "violation|approval|info", "rule_code":..., "article":..., "message":..., "amount":...}]
    """
    s = get_settings()
    scope = s.get("entity_scope", "union")
    is_union = (scope != "committee")

    if entry_id is not None:
        conn = get_connection()
        head = conn.execute(
            "SELECT entry_date, reference, description FROM journal_entries WHERE id=?",
            (entry_id,)).fetchone()
        lines_raw = _entry_lines(conn, entry_id)
        conn.close()
        if not head:
            return []
        date = head["entry_date"]
        reference = head["reference"]
        description = head["description"]
        lines = [{"account": r["account"], "category": r["category"],
                  "debit": _amount(r["debit"]), "credit": _amount(r["credit"]),
                  "desc": r["line_description"] or ""} for r in lines_raw]
    else:
        lines = draft_lines or []

    results = []

    def add(severity, code, article, message, amount=0.0):
        results.append({"severity": severity, "rule_code": code, "article": article,
                        "message": message, "amount": _amount(amount)})

    if not lines:
        return results

    kind, amount = _entry_kind(lines)
    blob = _text_blob(description, lines)

    # ============================ الصرف ============================
    if kind == "payment":
        cash_paid = _payment_source_cash(lines)
        targets = _payment_targets(lines)
        per_purpose_cap = (s["cash_payment_per_purpose_union"] if is_union
                           else s["cash_payment_per_purpose_committee"])

        # (مادة 9): سقف الصرف النقدي في غرض واحد
        if cash_paid > per_purpose_cap:
            add("violation", "CASH_PURPOSE_CAP", "مادة (9)",
                f"الصرف النقدي في هذا الغرض بلغ {fmt(cash_paid)} جنيه ويتجاوز سقف اللائحة "
                f"({fmt(per_purpose_cap)} جنيه). يجوز تجاوزه بعد موافقة رئيس النقابة بناءً "
                f"على عرض أمين الصندوق مع بيان الأسباب.",
                cash_paid)
        elif cash_paid > per_purpose_cap * 0.8:
            add("approval", "CASH_PURPOSE_NEAR", "مادة (9)",
                f"الصرف النقدي {fmt(cash_paid)} جنيه يقترب من سقف اللائحة "
                f"({fmt(per_purpose_cap)} جنيه) — تأكد من توفر موافقة رئيس النقابة إذا تجاوز السقف.",
                cash_paid)

        # (مادة 8): المعاملات تتم بموجب شيكات — صرف نقدي كبير يجب أن يكون لعملية تستلزم طبيعتها ذلك
        if cash_paid > 0 and amount > per_purpose_cap:
            add("approval", "CHECKS_POLICY", "مادة (8)",
                "تنص اللائحة على أن جميع المعاملات تتم بموجب شيكات، ولا يجوز التوريد/الصرف نقداً "
                "إلا في المعاملات التي تستلزم طبيعتها ذلك وبموجب إيصالات مسلسلة مطبوعة بدفاتر قبض وصرف.")

        # (مادة 10): فواتير الموردين
        is_supplier = any(any(k in (t["account"] + t["desc"] + blob) for k in SUPPLIER_KEYWORDS)
                          for t in targets)
        if is_supplier:
            doc_ok = any(k in blob for k in ("فاتورة", "أصل الفاتورة", "إذن توريد", "محضر استلام",
                                              "مطابق", "خاتم صرف", "صرف"))
            add(("info" if doc_ok else "violation"), "SUPPLIER_DOCS", "مادة (10)",
                ("صرف لفواتير موردين: تأكد من إرفاق أصل الفاتورة وإذن التوريد/محضر الاستلام "
                 "وختم المستندات بخاتم (صرف) فور السداد.")
                if not doc_ok else
                "تم رصد مستندات مؤيدة لفاتورة المورد — تأكد من ختمها بخاتم (صرف).",
                amount)

        # (مادة 61): حدود الشراء وتنفيذ الأعمال
        proc_keywords = ("شراء", "توريد", "مقاول", "صيانة", "إصلاح", "اصلاح", "أعمال",
                         "تعاقد", "عملية", "مشتريات", "مشروع", "مباني", "مبانى",
                         "إنشاء", "انشاء", "ترميم", "تجهيزات", "معدات", "أثاث", "اثاث",
                         "سيارة", "سيارات", "مهمات", "خامات", "مستلزمات")
        excluded = any(k in blob for k in PROCUREMENT_EXCLUDE_KEYWORDS)
        is_procurement = (not excluded) and (any(k in blob for k in proc_keywords) or is_supplier)
        if is_procurement:
            if is_union:
                direct_cap = s["proc_direct_order_union"]
                practice_cap = s["proc_practice_union"]
                limited_cap = s["proc_limited_tender_union"]
            else:
                direct_cap = s["proc_direct_order_committee"]
                practice_cap = s["proc_practice_committee"]
                limited_cap = s["proc_limited_tender_committee"]
            if amount <= direct_cap:
                method = "الأمر المباشر"
            elif amount <= practice_cap:
                method = "الممارسة (التفاوض مع عدد مناسب من الموردين)"
            elif amount <= limited_cap:
                method = "المناقصة المحدودة بين موردين ذوي تخصص وسمعة طيبة"
            else:
                method = "المناقصة العامة (مع الإعلان بصحيفة يومية والتأمين الابتدائي)"
            # تنبيه إذا كانت القيمة تقتضي إجراءً أعلى من الأمر المباشر
            if amount > direct_cap:
                add("violation", "PROC_METHOD", "مادة (61)",
                    f"قيمة العملية {fmt(amount)} جنيه تتجاوز حد الأمر المباشر "
                    f"({fmt(direct_cap)} جنيه). الواجب وفق اللائحة: {method}. "
                    f"ويُستثنى من ذلك التعامل مع الجهات الحكومية وشركات القطاع العام والأعمال العام "
                    f"والجمعيات التعاونية المشهرة (أمر مباشر أياً كانت القيمة).",
                    amount)
            else:
                add("info", "PROC_METHOD", "مادة (61)",
                    f"قيمة العملية {fmt(amount)} جنيه في حدود {method}.", amount)

        # (مادة 72): الدفعة المقدمة للمقاول
        if "دفعة مقدمة" in blob or "دفعه مقدمه" in blob or "مقدم" in blob:
            # لا يمكن حساب النسبة من القيد وحده؛ تنبيه توجيهي بالسقف
            add("approval", "CONTRACT_ADVANCE", "مادة (72)",
                f"دفعة مقدمة لمقاول: لا يجوز أن تزيد على {s['contract_advance_pct_cap']:.0f}% "
                f"من قيمة التعاقد، وتُصرف مقابل خطاب ضمان بنكي غير مشروط بذات القيمة.", amount)

        # (مادة 37): بدل السفر الداخلي (لا يشمل بدل الانتقال الشهري الثابت مادة 39)
        travel_keywords = ("بدل سفر", "مصروفات سفر", "مصاريف سفر", "مأمورية", "انتداب", "بدل مبيت")
        if any(k in blob for k in travel_keywords):
            daily_cap = s["travel_allowance_daily_union"] if is_union else s["travel_allowance_daily_committee"]
            max_daily = daily_cap * (1 + s["travel_allowance_max_increase_pct"] / 100.0)
            # افتراض ليليتين كحد أقصى للمأموريات العادية (القاهرة-الأقاليم ذهاباً وإياباً)
            generous_trip_limit = max_daily * 2
            if amount > generous_trip_limit:
                add("approval", "TRAVEL_ALLOWANCE", "مادة (37)",
                    f"قيمة بدل السفر {fmt(amount)} جنيه مرتفعة: الحد الأدنى {fmt(daily_cap)} جنيه/ليلة "
                    f"وأقصى قيمة {fmt(max_daily)} جنيه/ليلة بقرار مجلس الإدارة. تأكد من عدد ليالي "
                    f"المأمورية، وتخفيض 25% إذا تحملت المنظمة المبيت، وأن السفر بالطائرة تم بموافقة "
                    f"رئيس مجلس الإدارة (مادة 36).",
                    amount)

        # (مادة 39): بدل الانتقال الشهري
        if "بدل انتقال" in blob and amount > s["monthly_transport_allowance_cap"]:
            add("violation", "MONTHLY_TRANSPORT", "مادة (39)",
                f"بدل الانتقال الشهري {fmt(amount)} جنيه يتجاوز الحد الأقصى "
                f"({fmt(s['monthly_transport_allowance_cap'])} جنيه شهرياً) بقرار من مجلس الإدارة.",
                amount)

        # (مادة 40): بدل الأعباء
        if "بدل أعباء" in blob and amount > s["monthly_burden_allowance_cap"]:
            add("violation", "MONTHLY_BURDEN", "مادة (40)",
                f"بدل الأعباء {fmt(amount)} جنيه يتجاوز الحد الأقصى "
                f"({fmt(s['monthly_burden_allowance_cap'])} جنيه شهرياً).", amount)

        # (مادة 50/51): الهدايا والعلاقات الدولية
        if "هدايا" in blob or "هدية" in blob or "إكرامي" in blob:
            if amount > s["delegation_gifts_exception_cap"]:
                add("violation", "DELEGATION_GIFTS", "مادة (50)",
                    f"قيمة الهدايا {fmt(amount)} جنيه تتجاوز حتى الحد الاستثنائي "
                    f"({fmt(s['delegation_gifts_exception_cap'])} جنيه بقرار رئيس المنظمة).", amount)
            elif amount > s["delegation_gifts_cap"]:
                add("approval", "DELEGATION_GIFTS", "مادة (50)",
                    f"قيمة الهدايا {fmt(amount)} جنيه تتجاوز الحد العادي "
                    f"({fmt(s['delegation_gifts_cap'])} جنيه) — يلزم قرار مسبق من رئيس المنظمة.", amount)

        # (مادة 59): حظر التعاقد مع أعضاء مجلس الإدارة والعاملين وأقاربهم حتى الدرجة الثانية
        is_contract_context = any(k in blob for k in CONTRACT_KEYWORDS)
        if is_contract_context and any(k in blob for k in CONFLICT_PARTY_KEYWORDS):
            add("violation", "CONFLICT_INTEREST", "مادة (59)",
                "لا يجوز التعاقد على بيع أو شراء الأصول أو تنفيذ الأعمال مع أعضاء مجلس "
                "الإدارة أو العاملين بالمنظمة أو أقاربهم حتى الدرجة الثانية.", amount)

    # ============================ الإيراد / التحصيل ============================
    elif kind == "receipt":
        # (مادة 2): الاشتراكات وتوزيعها
        if any(k in blob for k in SUBSCRIPTION_KEYWORDS):
            add("info", "SUBSCRIPTION_SHARE", "مادة (2)",
                f"تحصيل اشتراكات: توزع الحصيلة بواقع "
                f"{s['subscription_committee_share_pct']:.0f}% للجنة النقابية، "
                f"{s['subscription_general_share_pct']:.0f}% للنقابة العامة، "
                f"{s['subscription_union_share_pct']:.0f}% تورد للاتحاد النقابي إن وجد. "
                f"وعلى المنشأة التوريد في النصف الأول من كل شهر.", amount)

        # (مادة 13): إيداع الشيكات والحوالات في اليوم التالي على الأكثر
        if "شيك" in blob or "حواله" in blob or "حوالة" in blob:
            add("info", "CHECK_DEPOSIT", "مادة (13)",
                f"يجب تحويل الشيكات والحوالات النقدية إلى البنك خلال "
                f"{int(s['check_deposit_deadline_days'])} يوم على الأكثر من ورودها بموجب حافظة "
                f"يحتفظ بصورة معتمدة منها لدى أمين الصندوق.", amount)

        # (مادة 77/78): متحصلات المزادات
        if "مزاد" in blob:
            add("info", "AUCTION_RECEIPT", "مادة (77/78)",
                f"متحصلات مزاد: 30% من ثمن المنقولات تُسدد فور رسو المزاد، و10% من ثمن العقارات "
                f"فور الرسو مع استكمال الباقي خلال 3 أشهر، ولا تنتقل الملكية إلا بعد السداد الكامل.",
                amount)

    return results


def fmt(v):
    try:
        return f"{float(v or 0):,.0f}"
    except Exception:
        return "0"


# ---------------------------------------------------------------------------
# 5) فحص كل القيود (للتقرير / المسح الشامل)
# ---------------------------------------------------------------------------

def scan_all_entries():
    """فحص جميع قيود اليومية وإرجاع قائمة مخالفاتها (بدون حفظ)."""
    conn = get_connection()
    entries = conn.execute("SELECT id FROM journal_entries ORDER BY entry_date, id").fetchall()
    conn.close()
    all_rows = []
    for e in entries:
        for r in check_entry(entry_id=e["id"]):
            all_rows.append((e["id"], r))
    return all_rows


def save_violations_for_entry(entry_id, results):
    """حفظ نتائج فحص قيد في جدول المخالفات (يُستخدم بعد حفظ القيد)."""
    conn = get_connection()
    cur = conn.cursor()
    head = cur.execute("SELECT entry_date, reference FROM journal_entries WHERE id=?",
                       (entry_id,)).fetchone()
    entry_date = head["entry_date"] if head else ""
    entry_ref = head["reference"] if head else ""
    # إعادة فحص نظيفة: حذف نتائج الفحص السابق لهذا القيد غير المعتمدة
    cur.execute("DELETE FROM compliance_violations WHERE entry_id=? AND acknowledged=0", (entry_id,))
    for r in results:
        cur.execute("""
            INSERT INTO compliance_violations
                (entry_id, entry_date, entry_ref, rule_code, article, severity, message, amount, checked_at)
            VALUES (?,?,?,?,?,?,?,?,?)
        """, (entry_id, entry_date, entry_ref, r["rule_code"], r["article"],
              r["severity"], r["message"], r.get("amount", 0.0),
              datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()
    conn.close()


def list_violations(severity=None, include_ack=False):
    conn = get_connection()
    sql = """
        SELECT v.id, v.entry_id, v.entry_date, v.entry_ref, v.rule_code, v.article,
               v.severity, v.message, v.amount, v.checked_at, v.acknowledged, v.ack_note,
               e.description AS entry_desc
        FROM compliance_violations v
        LEFT JOIN journal_entries e ON e.id = v.entry_id
    """
    clauses = []
    params = []
    if severity:
        clauses.append("v.severity=?")
        params.append(severity)
    if not include_ack:
        clauses.append("v.acknowledged=0")
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " ORDER BY v.entry_date DESC, v.id DESC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


def acknowledge_violation(violation_id, note=""):
    conn = get_connection()
    conn.execute("UPDATE compliance_violations SET acknowledged=1, ack_note=? WHERE id=?",
                 (note, violation_id))
    conn.commit()
    conn.close()


def violations_summary():
    """أعداد المخالفات حسب المستوى للوحة التحكم."""
    conn = get_connection()
    try:
        rows = conn.execute("""
            SELECT severity, COUNT(*) AS c
            FROM compliance_violations
            WHERE acknowledged=0
            GROUP BY severity
        """).fetchall()
    except sqlite3.OperationalError:
        init_compliance_tables()
        rows = conn.execute("""
            SELECT severity, COUNT(*) AS c
            FROM compliance_violations
            WHERE acknowledged=0
            GROUP BY severity
        """).fetchall()
    conn.close()
    out = {"violation": 0, "approval": 0, "info": 0}
    for r in rows:
        out[r["severity"]] = r["c"]
    return out


# ---------------------------------------------------------------------------
# 6) فحص رصيد الخزينة (مادة 6) — مقابل الأرصدة الفعلية بالدفاتر
# ---------------------------------------------------------------------------

def check_cash_on_hand():
    """فحص رصيد حسابات الخزينة مقابل سقف السلفة المستديمة (مادة 6)."""
    s = get_settings()
    is_union = s.get("entity_scope", "union") != "committee"
    cap = s["cash_on_hand_union"] if is_union else s["cash_on_hand_committee"]
    conn = get_connection()
    where = " OR ".join(["a.name LIKE ?" for _ in CASH_KEYWORDS])
    params = tuple(f"%{k}%" for k in CASH_KEYWORDS)
    rows = conn.execute(f"""
        SELECT a.name AS account,
               COALESCE(SUM(jl.debit),0) - COALESCE(SUM(jl.credit),0) AS balance
        FROM accounts a
        JOIN journal_lines jl ON jl.account_id = a.id
        WHERE {where}
        GROUP BY a.id
        HAVING balance > 0
    """, params).fetchall()
    conn.close()
    results = []
    for r in rows:
        bal = _amount(r["balance"])
        if bal > cap:
            results.append({"severity": "violation", "rule_code": "CASH_ON_HAND",
                            "article": "مادة (6)", "account": r["account"],
                            "amount": bal,
                            "message": f"رصيد {r['account']} بلغ {fmt(bal)} جنيه متجاوزاً سقف "
                                       f"السلفة المستديمة ({fmt(cap)} جنيه). يلزم إيداع الزائد "
                                       f"بالبنك، ولا تجوز الزيادة إلا باعتماد رئيس النقابة."})
    return results
