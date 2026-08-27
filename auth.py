"""
وحدة المستخدمين والصلاحيات.

- مستخدمون بثلاث صلاحيات: مدير (admin) / محاسب (accountant) / مشاهد (viewer).
- المدير يملك كل الصلاحيات، المحاسب يمكنه الإضافة والتعديل والحذف والتصدير،
  المشاهد يقرأ فقط (لا استيراد ولا تعديل ولا طباعة مستندات؟ بل مشاهدة فقط).
- كلمة المرور تُخزَّن كـ hash (sha256 + salt) وليست نصاً صريحاً.

المستخدم الافتراضي عند أول تشغيل: admin / admin
"""

import hashlib
import os
import sqlite3
from db import get_connection, init_db

ROLE_ADMIN = "admin"
ROLE_ACCOUNTANT = "accountant"
ROLE_VIEWER = "viewer"

ROLE_LABELS = {
    ROLE_ADMIN: "مدير النظام",
    ROLE_ACCOUNTANT: "محاسب",
    ROLE_VIEWER: "مشاهد",
}

# صلاحيات التنفيذ داخل التطبيق
PERM_IMPORT = "import"          # استيراد اليومية والمرتبات
PERM_EDIT_ENTRIES = "edit"      # إضافة/تعديل/حذف القيود
PERM_SETTLEMENTS = "settle"     # حفظ وترحيل التسويات البنكية
PERM_EXPORT = "export"          # تصدير Excel وطباعة PDF
PERM_USERS = "users"            # إدارة المستخدمين (للمدير فقط)

ROLE_PERMISSIONS = {
    ROLE_ADMIN: {PERM_IMPORT, PERM_EDIT_ENTRIES, PERM_SETTLEMENTS, PERM_EXPORT, PERM_USERS},
    ROLE_ACCOUNTANT: {PERM_IMPORT, PERM_EDIT_ENTRIES, PERM_SETTLEMENTS, PERM_EXPORT},
    ROLE_VIEWER: set(),
}


def _hash_password(password, salt=None):
    if salt is None:
        salt = os.urandom(16).hex()
    digest = hashlib.sha256((salt + password).encode("utf-8")).hexdigest()
    return f"{salt}${digest}"


def _verify_password(password, stored):
    if not stored or "$" not in stored:
        return False
    salt = stored.split("$")[0]
    return _hash_password(password, salt) == stored


def _ensure_default_admin():
    """ينشئ مستخدم admin/admin إذا لم يوجد أي مستخدم في النظام."""
    conn = get_connection()
    try:
        count = conn.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if count == 0:
            conn.execute(
                "INSERT INTO users(username, password_hash, full_name, role, active) VALUES (?, ?, ?, ?, 1)",
                ("admin", _hash_password("admin"), "مدير النظام", ROLE_ADMIN),
            )
            conn.commit()
    finally:
        conn.close()


def authenticate(username, password):
    """يستعلم عن المستخدم ويعيد قاموسه أو None عند الخطأ."""
    init_db()
    _ensure_default_admin()
    conn = get_connection()
    try:
        row = conn.execute(
            "SELECT id, username, password_hash, full_name, role, active FROM users WHERE username = ?",
            (username.strip(),),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    if not row["active"]:
        return None
    if not _verify_password(password, row["password_hash"]):
        return None
    return {
        "id": row["id"],
        "username": row["username"],
        "full_name": row["full_name"] or "",
        "role": row["role"],
        "role_label": ROLE_LABELS.get(row["role"], row["role"]),
    }


def has_permission(user, permission):
    return user is not None and permission in ROLE_PERMISSIONS.get(user.get("role"), set())


def users_list():
    conn = get_connection()
    try:
        rows = conn.execute(
            "SELECT id, username, full_name, role, active FROM users ORDER BY id"
        ).fetchall()
    finally:
        conn.close()
    return [dict(r) for r in rows]


def add_user(username, password, full_name="", role=ROLE_VIEWER):
    username = (username or "").strip()
    if not username or not password:
        return "أدخل اسم المستخدم وكلمة المرور"
    if role not in ROLE_LABELS:
        return "صلاحية غير معروفة"
    conn = get_connection()
    try:
        exists = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
        if exists:
            return f"المستخدم '{username}' موجود مسبقاً"
        conn.execute(
            "INSERT INTO users(username, password_hash, full_name, role, active) VALUES (?, ?, ?, ?, 1)",
            (username, _hash_password(password), (full_name or "").strip(), role),
        )
        conn.commit()
    finally:
        conn.close()
    return f"تمت إضافة المستخدم '{username}'"


def update_user(user_id, full_name=None, role=None, active=None, password=None):
    conn = get_connection()
    try:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return "المستخدم غير موجود"
        if role is not None:
            if role not in ROLE_LABELS:
                return "صلاحية غير معروفة"
            conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
        if full_name is not None:
            conn.execute("UPDATE users SET full_name = ? WHERE id = ?", (full_name.strip(), user_id))
        if active is not None:
            conn.execute("UPDATE users SET active = ? WHERE id = ?", (1 if active else 0, user_id))
        if password:
            conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (_hash_password(password), user_id))
        conn.commit()
    finally:
        conn.close()
    return "تم تحديث بيانات المستخدم"


def delete_user(user_id, current_user_id):
    conn = get_connection()
    try:
        row = conn.execute("SELECT username FROM users WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return "المستخدم غير موجود"
        if row["username"] == "admin":
            return "لا يمكن حذف مستخدم admin الأساسي"
        if int(user_id) == int(current_user_id):
            return "لا يمكنك حذف حسابك الحالي"
        conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
    finally:
        conn.close()
    return f"تم حذف المستخدم '{row['username']}'"
