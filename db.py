
import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "accounting.db")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS imports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT,
        imported_at TEXT,
        entries_count INTEGER,
        lines_count INTEGER
    )""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS accounts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT,
        name TEXT UNIQUE,
        category TEXT,
        normal_side TEXT
    )""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS journal_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_date TEXT NOT NULL,
        reference TEXT,
        description TEXT,
        total_debit REAL DEFAULT 0,
        total_credit REAL DEFAULT 0,
        source_row INTEGER,
        import_id INTEGER
    )""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS journal_lines (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER NOT NULL,
        account_id INTEGER NOT NULL,
        debit REAL DEFAULT 0,
        credit REAL DEFAULT 0,
        line_description TEXT,
        FOREIGN KEY(entry_id) REFERENCES journal_entries(id),
        FOREIGN KEY(account_id) REFERENCES accounts(id)
    )""")

    cur.execute("CREATE INDEX IF NOT EXISTS idx_entries_date ON journal_entries(entry_date)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_lines_entry ON journal_lines(entry_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_lines_account ON journal_lines(account_id)")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS payroll_imports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        source_file TEXT,
        imported_at TEXT,
        payroll_month TEXT,
        employees_count INTEGER DEFAULT 0,
        gross_total REAL DEFAULT 0,
        deductions_total REAL DEFAULT 0,
        net_total REAL DEFAULT 0,
        salary_advance_total REAL DEFAULT 0,
        bank_loan_total REAL DEFAULT 0
    )""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS payroll_rows (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_id INTEGER,
        employee_no INTEGER,
        employee_name TEXT,
        payroll_month TEXT,
        gross_total REAL DEFAULT 0,
        total_deductions REAL DEFAULT 0,
        net_pay REAL DEFAULT 0,
        salary_advance_installment REAL DEFAULT 0,
        bank_loan_installment REAL DEFAULT 0,
        insurance_employee REAL DEFAULT 0,
        tax_amount REAL DEFAULT 0,
        FOREIGN KEY(import_id) REFERENCES payroll_imports(id)
    )""")

    cur.execute("CREATE INDEX IF NOT EXISTS idx_payroll_rows_name ON payroll_rows(employee_name)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_payroll_rows_import ON payroll_rows(import_id)")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bank_settlements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bank_key TEXT NOT NULL,
        bank_name TEXT NOT NULL,
        settlement_year INTEGER NOT NULL,
        checks_under_collection REAL DEFAULT 0,
        uncashed_check1 REAL DEFAULT 0,
        uncashed_check2 REAL DEFAULT 0,
        bank_statement_balance REAL DEFAULT 0,
        notes TEXT,
        updated_at TEXT,
        UNIQUE(bank_key, settlement_year)
    )""")

    cur.execute("CREATE INDEX IF NOT EXISTS idx_bank_settlement_key_year ON bank_settlements(bank_key, settlement_year)")

    conn.commit()
    conn.close()

def clear_accounting_data():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("DELETE FROM journal_lines")
    cur.execute("DELETE FROM journal_entries")
    cur.execute("DELETE FROM accounts")
    conn.commit()
    conn.close()
