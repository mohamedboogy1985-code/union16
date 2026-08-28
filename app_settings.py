# -*- coding: utf-8 -*-
"""
إعدادات النظام العامة + المستخدمون والصلاحيات + سجل التدقيق
===========================================================
يخزن هذا الموديول:
  1) إعدادات عامة للمنشأة (الاسم، الشعار، السنة المالية، التنبيهات).
  2) إعدادات الحسابات والضرائب (الفاتورة الإلكترونية، ضريبة القيمة المضافة،
     الخصم تحت حساب الضريبة).
  3) إعدادات الموارد البشرية (نسب التأمينات، شرائح ضريبة كسب العمل «استمارة 2»،
     بنود مسير الرواتب).
  4) المستخدمون وأدوار الصلاحيات.
  5) سجل التدقيق (Audit Log) لكل العمليات الحساسة.
"""

import os
import json
import hashlib
from datetime import datetime

from db import get_connection

# ---------------------------------------------------------------------------
# الإعدادات الافتراضية
# ---------------------------------------------------------------------------

DEFAULT_TAX_BRACKETS = [
    # شرائح صافي الدخل السنوي الخاضع للضريبة (جنيه) — قابلة للتعديل من الشاشة
    {"from": 0, "to": 30000, "rate": 0, "label": "الشريحة المعفاة"},
    {"from": 30000, "to": 45000, "rate": 10, "label": "الشريحة الثانية"},
    {"from": 45000, "to": 60000, "rate": 15, "label": "الشريحة الثالثة"},
    {"from": 60000, "to": 200000, "rate": 20, "label": "الشريحة الرابعة"},
    {"from": 200000, "to": 400000, "rate": 22.5, "label": "الشريحة الخامسة"},
    {"from": 400000, "to": 0, "rate": 25, "label": "أعلى من 400 ألف (0 = بدون حد)"},
]

DEFAULT_PAYROLL_ITEMS = [
    "الأجر الأساسي", "بدل طبيعة عمل", "بدل تمثيل", "بدل انتقال",
    "علاوة خاصة", "علاوات اجتماعية", "حوافز ومكافآت", "مناسبات موسمية",
    "تأمينات اجتماعية (حصة العامل 11%)", "ضريبة كسب العمل",
    "قسط سلفة شخصية", "قرض بنك التنمية", "تأمين تكافلي",
]

DEFAULT_SETTINGS = {
    # ---- عامة ----
    "org_name": "النقابة العامة للعاملين بصناعات البناء والأخشاب وصنع مواد البناء",
    "org_logo": "",
    "fiscal_year": str(datetime.now().year),
    "fiscal_year_start": "01-01",
    "currency": "جنيه مصري",
    "notify_cash_cap": "1",          # تنبيه تجاوز سقف الخزينة
    "notify_violations": "1",        # تنبيه مخالفات اللائحة
    "notify_monthly_close": "1",     # تنبيه إقفال الشهر (جرد الخزينة مادة 16)

    # ---- الحسابات والضرائب / الفاتورة الإلكترونية ----
    "tax_registration": "",                       # رقم التسجيل الضريبي
    "einvoice_enabled": "0",                      # تفعيل الربط مع الفاتورة الإلكترونية
    "einvoice_env": "test",                       # test | production
    "einvoice_client_id": "",
    "einvoice_client_secret": "",
    "einvoice_activity_code": "",
    "vat_rate": "14",                             # ضريبة القيمة المضافة %
    "withholding_fees_pct": "10",                 # خصم تحت حساب الضريبة: أتعاب مهنية
    "withholding_commissions_pct": "5",           # عمولات وسمسرة
    "withholding_supply_pct": "2",                # مقاولات وتوريدات
    "withholding_min_amount": "300",              # حد تطبيق الخصم

    # ---- الموارد البشرية والمرتبات ----
    "ins_employee_pct": "11",                     # حصة العامل في التأمينات
    "ins_employer_pct": "18.75",                  # حصة المنشأة
    "ins_min_wage": "2300",                       # الحد الأدنى للأجر التأميني
    "ins_max_wage": "14500",                      # الحد الأقصى للأجر التأميني
    "tax_annual_relief": "15000",                 # التخفيض الضريبي السنوي (استمارة 2)
    "tax_brackets": json.dumps(DEFAULT_TAX_BRACKETS, ensure_ascii=False),
    "payroll_items": json.dumps(DEFAULT_PAYROLL_ITEMS, ensure_ascii=False),
}

# ---------------------------------------------------------------------------
# الأدوار والصلاحيات
# ---------------------------------------------------------------------------

ROLES = {
    "admin": {
        "label": "مدير النظام",
        "permissions": {"entries_edit", "entries_delete", "import", "export",
                        "backup", "restore", "compliance_ack", "compliance_rescan",
                        "settings", "users_manage", "payroll_edit", "reports"},
    },
    "accountant": {
        "label": "محاسب",
        "permissions": {"entries_edit", "entries_delete", "import", "export",
                        "payroll_edit", "reports"},
    },
    "reviewer": {
        "label": "مراجع / رقابة",
        "permissions": {"export", "reports", "compliance_ack", "compliance_rescan"},
    },
    "viewer": {
        "label": "مشاهد فقط",
        "permissions": {"export", "reports"},
    },
}

PERMISSION_LABELS = {
    "entries_edit": "إدخال وتعديل القيود",
    "entries_delete": "حذف القيود",
    "import": "استيراد ملفات Excel",
    "export": "تصدير التقارير",
    "backup": "نسخة احتياطية",
    "restore": "استرجاع نسخة احتياطية",
    "compliance_ack": "اعتماد مخالفات اللائحة كاستثناء",
    "compliance_rescan": "إعادة الفحص الشامل للائحة",
    "settings": "تعديل إعدادات النظام واللائحة",
    "users_manage": "إدارة المستخدمين",
    "payroll_edit": "تعديل بيانات المرتبات",
    "reports": "عرض التقارير",
}


def init_settings_tables():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            full_name TEXT,
            password_hash TEXT,
            role TEXT DEFAULT 'viewer',
            active INTEGER DEFAULT 1,
            created_at TEXT
        )""")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT,
            username TEXT,
            action TEXT,
            details TEXT
        )""")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_audit_ts ON audit_log(ts)")

    existing = {r[0] for r in cur.execute("SELECT key FROM app_settings").fetchall()}
    for k, v in DEFAULT_SETTINGS.items():
        if k not in existing:
            cur.execute("INSERT INTO app_settings(key,value) VALUES(?,?)", (k, v))

    # مستخدم المدير الافتراضي (admin / admin)
    admin = cur.execute("SELECT id FROM users WHERE username='admin'").fetchone()
    if not admin:
        cur.execute(
            "INSERT INTO users(username, full_name, password_hash, role, active, created_at) "
            "VALUES(?,?,?,?,1,?)",
            ("admin", "مدير النظام", hash_password("admin"), "admin",
             datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# الإعدادات
# ---------------------------------------------------------------------------

def get_all_settings():
    conn = get_connection()
    try:
        rows = conn.execute("SELECT key, value FROM app_settings").fetchall()
    except Exception:
        init_settings_tables()
        rows = conn.execute("SELECT key, value FROM app_settings").fetchall()
    conn.close()
    out = dict(DEFAULT_SETTINGS)
    for r in rows:
        out[r[0]] = r[1]
    return out


def get_setting(key, default=""):
    return get_all_settings().get(key, default)


def set_setting(key, value):
    conn = get_connection()
    conn.execute("INSERT INTO app_settings(key,value) VALUES(?,?) "
                 "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, str(value)))
    conn.commit()
    conn.close()


def set_settings(mapping):
    conn = get_connection()
    for k, v in mapping.items():
        conn.execute("INSERT INTO app_settings(key,value) VALUES(?,?) "
                     "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
    conn.commit()
    conn.close()


def get_json_setting(key, default):
    import json as _json
    try:
        return _json.loads(get_setting(key, ""))
    except Exception:
        return default


# ---------------------------------------------------------------------------
# المستخدمون
# ---------------------------------------------------------------------------

def hash_password(password):
    return hashlib.sha256(("SA:" + str(password)).encode("utf-8")).hexdigest()


def verify_user(username, password):
    conn = get_connection()
    row = conn.execute(
        "SELECT * FROM users WHERE username=? AND active=1", (username,)).fetchone()
    conn.close()
    if not row:
        return None
    stored = row["password_hash"] or ""
    # توافق مع المستخدمين القدامى (صيغة salt$digest من auth.py السابقة)
    if stored and "$" in stored:
        salt, digest = stored.split("$", 1)
        if hashlib.sha256((salt + password).encode("utf-8")).hexdigest() != digest:
            return None
    elif stored != hash_password(password):
        return None
    return {"id": row["id"], "username": row["username"],
            "full_name": row["full_name"], "role": row["role"]}


def list_users():
    conn = get_connection()
    rows = conn.execute("SELECT id, username, full_name, role, active, created_at "
                        "FROM users ORDER BY id").fetchall()
    conn.close()
    return rows


def add_user(username, full_name, password, role):
    if role not in ROLES:
        raise ValueError("دور غير معروف")
    conn = get_connection()
    try:
        conn.execute(
            "INSERT INTO users(username, full_name, password_hash, role, active, created_at) "
            "VALUES(?,?,?,?,1,?)",
            (username.strip(), full_name.strip(), hash_password(password), role,
             datetime.now().strftime("%Y-%m-%d %H:%M")))
        conn.commit()
    except Exception as e:
        conn.close()
        raise ValueError("اسم المستخدم موجود بالفالسابق أو البيانات غير صحيحة") from e
    conn.close()


def update_user(user_id, full_name=None, role=None, password=None, active=None):
    conn = get_connection()
    if full_name is not None:
        conn.execute("UPDATE users SET full_name=? WHERE id=?", (full_name, user_id))
    if role is not None:
        if role not in ROLES:
            conn.close()
            raise ValueError("دور غير معروف")
        conn.execute("UPDATE users SET role=? WHERE id=?", (role, user_id))
    if password:
        conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                     (hash_password(password), user_id))
    if active is not None:
        conn.execute("UPDATE users SET active=? WHERE id=?", (1 if active else 0, user_id))
    conn.commit()
    conn.close()


def can(role, permission):
    return permission in ROLES.get(role, {}).get("permissions", set())


# ---------------------------------------------------------------------------
# سجل التدقيق
# ---------------------------------------------------------------------------

def log_action(username, action, details=""):
    try:
        conn = get_connection()
        conn.execute(
            "INSERT INTO audit_log(ts, username, action, details) VALUES(?,?,?,?)",
            (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), username or "غير معروف",
             action, details))
        conn.commit()
        conn.close()
    except Exception:
        pass


def list_audit(limit=500):
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT ts, username, action, details FROM audit_log "
            "ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    except Exception:
        rows = []
    conn.close()
    return rows
