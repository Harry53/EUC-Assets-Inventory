from __future__ import annotations

import csv
import io
import json
import os
import shutil
import smtplib
import sqlite3
from datetime import date, datetime
from functools import wraps
from pathlib import Path
from typing import Any

from flask import Flask, Response, flash, redirect, render_template, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "inventory.db"
LOG_PATH = BASE_DIR / "logs" / "activity.log"
ASSET_TYPES = ["Laptop", "Desktop", "Workstation", "Monitor", "G-Drive", "G-Raid"]
ROLES = ("Admin", "Editor", "Viewer")

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-change-me")
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

ASSET_COLUMNS = [
    "device_tag", "location", "asset_type", "serial_number", "ip_address", "device_model", "ram_gb",
    "storage_type", "storage_capacity", "action", "status", "username", "system_name", "previous_user_history",
    "purchase_year", "invoice_number", "finance_code", "supplier_name", "asset_value", "warranty_amc",
    "eol", "insurance_status", "insurance_type", "e_code"
]
EMPLOYEE_COLUMNS = ["e_code", "employee_name", "doj", "vertical", "sub_vertical", "content_category", "reporting_manager", "rfds", "employee_status"]
SOFTWARE_COLUMNS = ["software_name", "version", "license_type", "license_key", "number_of_licenses", "sub_vertical", "purchase_date", "registration_email", "status"]
SOFTWARE_ASSIGN_COLUMNS = [
    "desktop_laptop_tag", "e_code", "employee_name", "doj", "vertical", "sub_vertical", "content_category",
    "reporting_manager", "rfds", "location", "software_name", "version", "license_type", "license_key",
    "number_of_licenses", "installation_date", "expiry_date", "vendor_supplier", "purchase_year", "invoice_number",
    "cost", "status", "support_amc", "last_updated_date", "remarks"
]
UNPROCESSED_ASSET_COLUMNS = [
    "device_tag", "e_code", "employee_name", "doj", "vertical", "sub_vertical", "content_category", "reporting_manager",
    "rfds", "location", "asset_type", "action", "status", "username", "system_name", "previous_user_history",
    "serial_number", "ip_address", "device_model", "ram_gb", "storage_type", "storage_capacity", "purchase_year",
    "invoice_number", "finance_code", "supplier_name", "asset_value", "warranty_amc", "eol", "insurance_status", "insurance_type"
]
VENDOR_COLUMNS = ["vendor_name", "contact_person", "email", "phone", "address", "services", "status", "remarks"]
CHECKLIST_COLUMNS = ["checklist_type", "employee_name", "e_code", "vertical", "sub_vertical", "rfds", "payload", "asset_updates", "software_updates", "employee_status", "status"]
NOTIFICATION_TYPES = ["Approval", "Change", "Asset Removal", "Password Expiry", "License Expiry"]
DATE_FIELDS = {"doj", "eol", "purchase_date", "installation_date", "expiry_date", "last_updated_date", "email_created_on", "date_of_exit", "form_submission_date"}
SAMPLE_MAP = {
    "unprocessed-assets": UNPROCESSED_ASSET_COLUMNS,
    "unprocessed-software": SOFTWARE_ASSIGN_COLUMNS,
    "unprocessed-employees": EMPLOYEE_COLUMNS,
    "vendors": VENDOR_COLUMNS,
    "checklist-joining": ["employee_name", "e_code", "vertical", "sub_vertical", "rfds", "machine_asset_tag", "system_serial_number", "model", "ip_address", "operating_system", "office_365_license"],
    "checklist-replacement": ["employee_name", "e_code", "old_assets", "old_softwares", "machine_asset_tag", "system_serial_number", "model", "remarks"],
    "checklist-rebuild": ["employee_name", "e_code", "old_assets", "old_softwares", "machine_asset_tag", "system_serial_number", "model", "remarks"],
    "checklist-exit": ["employee_name", "e_code", "date_of_exit", "form_submission_date", "machine", "data_card", "external_hdd", "host_name", "ad_login_id", "data_handover"],
}

LABELS = {
    "device_tag": "Device Tag", "desktop_laptop_tag": "Desktop/Laptop Tag", "location": "Location", "asset_type": "Asset Type",
    "serial_number": "Serial Number", "ip_address": "IP Address", "device_model": "Device Model", "ram_gb": "RAM GB",
    "storage_type": "Storage Type", "storage_capacity": "Storage Capacity", "action": "Action", "status": "Status",
    "username": "Username", "system_name": "System Name", "previous_user_history": "Previous User History", "purchase_year": "Purchase Year",
    "invoice_number": "Invoice Number", "finance_code": "Finance Code", "supplier_name": "Supplier Name", "asset_value": "Asset Value",
    "warranty_amc": "Warranty/AMC", "eol": "End of Life", "insurance_status": "Insurance Status", "insurance_type": "Insurance Type",
    "e_code": "E-Code", "employee_name": "Employee Name", "doj": "Date of Joining", "vertical": "Vertical",
    "sub_vertical": "Sub-Vertical", "content_category": "Content Category", "reporting_manager": "Reporting Manager", "rfds": "RFDs",
    "employee_status": "Employee Status", "software_name": "Software Name", "version": "Version", "license_type": "License Type",
    "license_key": "License Key", "number_of_licenses": "Number of Licenses", "purchase_date": "Purchase Date", "registration_email": "Registration email id",
    "installation_date": "Installation Date", "expiry_date": "Expiry Date", "vendor_supplier": "Vendor/Supplier",
    "cost": "Cost", "support_amc": "Support/AMC", "last_updated_date": "Last Updated Date", "remarks": "Remarks", "vendor_name": "Vendor Name", "contact_person": "Contact Person", "services": "Services", "checklist_type": "Checklist Type", "payload": "Checklist Data", "asset_updates": "Asset Updates", "software_updates": "Software Updates"
}


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def execute(sql: str, params: tuple[Any, ...] = ()) -> sqlite3.Cursor:
    conn = db()
    cur = conn.execute(sql, params)
    conn.commit()
    conn.close()
    return cur


def init_db() -> None:
    schema = f"""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'Viewer', status TEXT NOT NULL DEFAULT 'Active', disabled INTEGER NOT NULL DEFAULT 0,
        last_login TEXT, last_ip TEXT, full_name TEXT, email TEXT, phone TEXT, department TEXT, theme TEXT DEFAULT 'dark', notification_preferences TEXT DEFAULT '[]', password_changed_at TEXT, created_at TEXT NOT NULL
    );
    CREATE TABLE IF NOT EXISTS employees (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(EMPLOYEE_COLUMNS)}, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS assets (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(ASSET_COLUMNS)}, approved INTEGER DEFAULT 1, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS software (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(SOFTWARE_COLUMNS)}, used_count INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS software_assignments (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(SOFTWARE_ASSIGN_COLUMNS)}, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS unprocessed_assets (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(UNPROCESSED_ASSET_COLUMNS)}, maker TEXT, checker TEXT, approval_status TEXT DEFAULT 'Pending', processed INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS unprocessed_software (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(SOFTWARE_ASSIGN_COLUMNS)}, maker TEXT, checker TEXT, approval_status TEXT DEFAULT 'Pending', processed INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS unprocessed_employees (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(EMPLOYEE_COLUMNS)}, maker TEXT, checker TEXT, approval_status TEXT DEFAULT 'Pending', processed INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS audit_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, actor TEXT, action TEXT, table_name TEXT, record_id INTEGER, details TEXT, ip TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS login_history (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, ip TEXT, status TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS notifications (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, type TEXT, message TEXT, is_read INTEGER DEFAULT 0, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS email_config (id INTEGER PRIMARY KEY CHECK (id=1), smtp_host TEXT, smtp_port TEXT, smtp_username TEXT, smtp_password TEXT, sender_email TEXT, use_tls INTEGER DEFAULT 1, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS email_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, recipient TEXT, subject TEXT, body TEXT, status TEXT, created_at TEXT NOT NULL);
    CREATE TABLE IF NOT EXISTS vendors (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(VENDOR_COLUMNS)}, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS checklists (id INTEGER PRIMARY KEY AUTOINCREMENT, {cols(CHECKLIST_COLUMNS)}, maker TEXT, checker TEXT, processed INTEGER DEFAULT 0, created_at TEXT, updated_at TEXT);
    CREATE TABLE IF NOT EXISTS checklist_images (id INTEGER PRIMARY KEY AUTOINCREMENT, checklist_id INTEGER, file_path TEXT, original_name TEXT, created_at TEXT NOT NULL);
    """
    conn = db()
    conn.executescript(schema)
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assets_device_tag ON assets(device_tag)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assets_serial ON assets(serial_number) WHERE serial_number IS NOT NULL AND serial_number != ''")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_ecode ON employees(e_code)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_software_key ON software(license_key) WHERE license_key IS NOT NULL AND license_key != ''")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assignment_device_license ON software_assignments(desktop_laptop_tag, license_key)")
    ensure_columns(conn)
    if not conn.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
        conn.execute("INSERT INTO users (username,password_hash,role,status,disabled,created_at) VALUES (?,?,?,?,?,?)",
                     ("admin", generate_password_hash("admin123"), "Admin", "Active", 0, now()))
    conn.commit()
    conn.close()


def cols(names: list[str]) -> str:
    return ", ".join(f"{name} TEXT" for name in names)


def ensure_columns(conn: sqlite3.Connection) -> None:
    additions = {
        "users": {"full_name": "TEXT", "email": "TEXT", "phone": "TEXT", "department": "TEXT", "theme": "TEXT DEFAULT 'dark'", "notification_preferences": "TEXT DEFAULT '[]'", "password_changed_at": "TEXT"},
    }
    for table, columns in additions.items():
        existing = {row[1] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}
        for column, definition in columns.items():
            if column not in existing:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")


def now() -> str:
    return datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")


def current_user() -> sqlite3.Row | None:
    uid = session.get("user_id")
    if not uid:
        return None
    conn = db()
    user = conn.execute("SELECT * FROM users WHERE id=? AND disabled=0", (uid,)).fetchone()
    conn.close()
    return user


def unread_notifications() -> list[sqlite3.Row]:
    username = session.get("username")
    if not username:
        return []
    conn = db()
    rows = conn.execute("SELECT * FROM notifications WHERE username=? AND is_read=0 ORDER BY id DESC LIMIT 10", (username,)).fetchall()
    conn.close()
    return rows


def create_notification(username: str, note_type: str, message: str) -> None:
    execute("INSERT INTO notifications (username,type,message,is_read,created_at) VALUES (?,?,?,?,?)", (username, note_type, message, 0, now()))
    conn = db()
    target = conn.execute("SELECT email, notification_preferences FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    prefs = json.loads(target["notification_preferences"] or "[]") if target else []
    if target and target["email"] and note_type in prefs:
        send_email(target["email"], f"EUC Inventory {note_type} Notification", message)


@app.context_processor
def inject_globals() -> dict[str, Any]:
    return {"user": current_user(), "labels": LABELS, "asset_types": ASSET_TYPES, "notification_types": NOTIFICATION_TYPES, "unread_notifications": unread_notifications()}


def login_required(view):
    @wraps(view)
    def wrapper(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapper


def roles_required(*roles: str):
    def decorator(view):
        @wraps(view)
        def wrapper(*args, **kwargs):
            user = current_user()
            if not user or user["role"] not in roles:
                flash("You do not have permission for this action.", "danger")
                return redirect(url_for("dashboard"))
            return view(*args, **kwargs)
        return wrapper
    return decorator


def log_action(action: str, table: str = "", record_id: int | None = None, details: Any = None) -> None:
    actor = session.get("username", "system")
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "") if request else ""
    detail_text = details if isinstance(details, str) else json.dumps(details or {}, default=str)
    LOG_PATH.parent.mkdir(exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        handle.write(f"{now()} | {actor} | {ip} | {action} | {table} | {record_id or ''} | {detail_text}\n")
    conn = db()
    conn.execute("INSERT INTO audit_logs (actor,action,table_name,record_id,details,ip,created_at) VALUES (?,?,?,?,?,?,?)",
                 (actor, action, table, record_id, detail_text, ip, now()))
    conn.commit()
    conn.close()


def row_dict(columns: list[str], source: dict[str, Any]) -> dict[str, str]:
    return {col: (source.get(col) or source.get(LABELS.get(col, "")) or "").strip() for col in columns}


def parse_upload(file) -> list[dict[str, str]]:
    filename = (file.filename or "").lower()
    data = file.read()
    rows: list[dict[str, str]] = []
    if filename.endswith(".xlsx"):
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws = wb.active
        headers = [str(c.value or "").strip() for c in next(ws.iter_rows(max_row=1))]
        for record in ws.iter_rows(min_row=2, values_only=True):
            rows.append({headers[i]: str(record[i] or "").strip() for i in range(len(headers))})
    else:
        text = data.decode("utf-8-sig")
        rows = [dict(r) for r in csv.DictReader(io.StringIO(text))]
    return rows


def insert_record(table: str, columns: list[str], data: dict[str, str], extra: dict[str, str] | None = None) -> int:
    payload = row_dict(columns, data)
    payload.update(extra or {})
    payload.setdefault("created_at", now())
    payload.setdefault("updated_at", now())
    keys = list(payload.keys())
    placeholders = ",".join("?" for _ in keys)
    conn = db()
    cur = conn.execute(f"INSERT INTO {table} ({','.join(keys)}) VALUES ({placeholders})", tuple(payload[k] for k in keys))
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    log_action("insert", table, rid, payload)
    return rid


def update_record(table: str, columns: list[str], record_id: int, data: dict[str, str]) -> None:
    payload = row_dict(columns, data)
    payload["updated_at"] = now()
    assignments = ",".join(f"{k}=?" for k in payload)
    execute(f"UPDATE {table} SET {assignments} WHERE id=?", tuple(payload.values()) + (record_id,))
    log_action("update", table, record_id, payload)


def approve_asset(record_id: int) -> None:
    conn = db()
    row = conn.execute("SELECT * FROM unprocessed_assets WHERE id=?", (record_id,)).fetchone()
    if not row:
        raise ValueError("Asset request not found")
    data = dict(row)
    employee = row_dict(EMPLOYEE_COLUMNS, data)
    if employee["e_code"]:
        upsert(conn, "employees", EMPLOYEE_COLUMNS, employee, "e_code")
    asset = row_dict(ASSET_COLUMNS, data)
    upsert(conn, "assets", ASSET_COLUMNS, asset, "device_tag")
    conn.execute("UPDATE unprocessed_assets SET approval_status='Approved', processed=1, checker=?, updated_at=? WHERE id=?", (session.get("username"), now(), record_id))
    conn.commit(); conn.close()
    log_action("approve", "unprocessed_assets", record_id, {"device_tag": data.get("device_tag")})


def approve_software(record_id: int) -> None:
    conn = db()
    row = conn.execute("SELECT * FROM unprocessed_software WHERE id=?", (record_id,)).fetchone()
    if not row:
        raise ValueError("Software request not found")
    data = dict(row)
    software = row_dict(SOFTWARE_COLUMNS, data)
    software["purchase_date"] = data.get("installation_date", "")
    upsert(conn, "software", SOFTWARE_COLUMNS, software, "license_key")
    assignment = row_dict(SOFTWARE_ASSIGN_COLUMNS, data)
    upsert(conn, "software_assignments", SOFTWARE_ASSIGN_COLUMNS, assignment, "desktop_laptop_tag", extra_unique="license_key")
    conn.execute("UPDATE software SET used_count=(SELECT COUNT(*) FROM software_assignments WHERE software_assignments.license_key=software.license_key)")
    conn.execute("UPDATE unprocessed_software SET approval_status='Approved', processed=1, checker=?, updated_at=? WHERE id=?", (session.get("username"), now(), record_id))
    conn.commit(); conn.close()
    log_action("approve", "unprocessed_software", record_id, {"license_key": data.get("license_key")})


def approve_employee(record_id: int) -> None:
    conn = db()
    row = conn.execute("SELECT * FROM unprocessed_employees WHERE id=?", (record_id,)).fetchone()
    if not row:
        raise ValueError("Employee request not found")
    upsert(conn, "employees", EMPLOYEE_COLUMNS, row_dict(EMPLOYEE_COLUMNS, dict(row)), "e_code")
    conn.execute("UPDATE unprocessed_employees SET approval_status='Approved', processed=1, checker=?, updated_at=? WHERE id=?", (session.get("username"), now(), record_id))
    conn.commit(); conn.close()
    log_action("approve", "unprocessed_employees", record_id, {"e_code": row["e_code"]})


def upsert(conn: sqlite3.Connection, table: str, columns: list[str], payload: dict[str, str], unique: str, extra_unique: str | None = None) -> None:
    payload["updated_at"] = now()
    where = f"{unique}=?"
    params = [payload.get(unique, "")]
    if extra_unique:
        where += f" AND {extra_unique}=?"
        params.append(payload.get(extra_unique, ""))
    existing = conn.execute(f"SELECT id FROM {table} WHERE {where}", params).fetchone()
    if existing:
        assignments = ",".join(f"{k}=?" for k in payload)
        conn.execute(f"UPDATE {table} SET {assignments} WHERE id=?", tuple(payload.values()) + (existing["id"],))
    else:
        payload["created_at"] = now()
        keys = list(payload.keys())
        conn.execute(f"INSERT INTO {table} ({','.join(keys)}) VALUES ({','.join('?' for _ in keys)})", tuple(payload.values()))


def rows_for(table: str, order: str = "id DESC", limit: int | None = None, where: str = "1=1", params: tuple[Any, ...] = ()) -> list[sqlite3.Row]:
    conn = db()
    sql = f"SELECT * FROM {table} WHERE {where} ORDER BY {order}"
    if limit:
        sql += f" LIMIT {limit}"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"]
        conn = db()
        user = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        ok = user and not user["disabled"] and user["status"] == "Active" and check_password_hash(user["password_hash"], password)
        ip = request.remote_addr or ""
        conn.execute("INSERT INTO login_history (username,ip,status,created_at) VALUES (?,?,?,?)", (username, ip, "Success" if ok else "Failed", now()))
        if ok:
            conn.execute("UPDATE users SET last_login=?, last_ip=? WHERE id=?", (now(), ip, user["id"]))
            session.update(user_id=user["id"], username=user["username"], role=user["role"])
            conn.commit(); conn.close(); log_action("login", "users", user["id"])
            return redirect(url_for("dashboard"))
        conn.commit(); conn.close(); flash("Invalid credentials or disabled user.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    log_action("logout")
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
@login_required
def dashboard():
    conn = db()
    asset_total = conn.execute("SELECT COUNT(*) c FROM assets").fetchone()["c"]
    employee_total = conn.execute("SELECT COUNT(*) c FROM employees").fetchone()["c"]
    software_total = conn.execute("SELECT COUNT(*) c FROM software").fetchone()["c"]
    active_users = conn.execute("SELECT COUNT(*) c FROM users WHERE disabled=0 AND status='Active'").fetchone()["c"]
    total_users = conn.execute("SELECT COUNT(*) c FROM users").fetchone()["c"]
    counts = {t: conn.execute("SELECT COUNT(*) c FROM assets WHERE asset_type=?", (t,)).fetchone()["c"] for t in ASSET_TYPES}
    approved = conn.execute("SELECT COUNT(*) c FROM assets WHERE approved=1").fetchone()["c"]
    pending = conn.execute("SELECT COUNT(*) c FROM unprocessed_assets WHERE approval_status='Pending'").fetchone()["c"]
    software = conn.execute("SELECT *, MAX(CAST(number_of_licenses AS INTEGER)-used_count,0) free_count FROM software ORDER BY software_name").fetchall()
    unapproved = conn.execute("SELECT asset_type, COUNT(*) c FROM unprocessed_assets WHERE approval_status='Pending' GROUP BY asset_type").fetchall()
    logs = conn.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT 200").fetchall()
    logins = conn.execute("SELECT * FROM login_history ORDER BY id DESC LIMIT 200").fetchall()
    conn.close()
    return render_template("dashboard.html", asset_total=asset_total, employee_total=employee_total, software_total=software_total,
                           active_users=active_users, total_users=total_users, counts=counts, approved=approved, pending=pending,
                           software=software, unapproved=unapproved, logs=logs, logins=logins)


@app.route("/inventory/<kind>", methods=["GET", "POST"])
@login_required
def inventory(kind: str):
    config = {
        "assets": ("Assets Inventory", "assets", ASSET_COLUMNS, ["Admin"]),
        "software": ("Licensed Software", "software", SOFTWARE_COLUMNS, ["Admin"]),
        "employees": ("Employee List", "employees", EMPLOYEE_COLUMNS, ["Admin"]),
        "unprocessed-assets": ("Unprocessed Assets", "unprocessed_assets", UNPROCESSED_ASSET_COLUMNS, ["Admin", "Editor"]),
        "unprocessed-software": ("Unprocessed Software", "unprocessed_software", SOFTWARE_ASSIGN_COLUMNS, ["Admin", "Editor"]),
        "unprocessed-employees": ("Unprocessed Employees", "unprocessed_employees", EMPLOYEE_COLUMNS, ["Admin", "Editor"]),
    }.get(kind)
    if not config:
        return redirect(url_for("dashboard"))
    title, table, columns, writers = config
    user = current_user()
    if request.method == "POST" and user["role"] in writers:
        extra = {}
        if table.startswith("unprocessed"):
            extra = {"maker": user["username"], "approval_status": "Pending", "processed": "0"}
        try:
            rid = insert_record(table, columns, request.form, extra)
            flash(f"Saved record #{rid}.", "success")
        except sqlite3.IntegrityError as exc:
            flash(f"Duplicate entry blocked: {exc}", "danger")
        return redirect(url_for("inventory", kind=kind))
    search = request.args.get("q", "").strip()
    where, params = "1=1", ()
    if search:
        like = f"%{search}%"
        where = " OR ".join(f"{c} LIKE ?" for c in columns)
        params = tuple([like] * len(columns))
    rows = rows_for(table, where=where, params=params)
    return render_template("inventory.html", kind=kind, title=title, table=table, columns=columns, rows=rows, can_write=user["role"] in writers)


def validate_upload(table: str, columns: list[str], records: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    headers = set(records[0].keys()) if records else set()
    missing = [LABELS.get(c, c) for c in columns if c not in headers and LABELS.get(c, "") not in headers]
    if missing:
        errors.append("Missing columns: " + ", ".join(missing[:12]))
    conn = db()
    for idx, record in enumerate(records, start=2):
        normalized = row_dict(columns, record)
        if table == "unprocessed_assets":
            tag, serial = normalized.get("device_tag", ""), normalized.get("serial_number", "")
            if tag and conn.execute("SELECT 1 FROM assets WHERE device_tag=? UNION SELECT 1 FROM unprocessed_assets WHERE device_tag=? AND approval_status='Pending'", (tag, tag)).fetchone():
                errors.append(f"Row {idx}: duplicate pending/approved device tag {tag}")
            if serial and conn.execute("SELECT 1 FROM assets WHERE serial_number=? UNION SELECT 1 FROM unprocessed_assets WHERE serial_number=? AND approval_status='Pending'", (serial, serial)).fetchone():
                errors.append(f"Row {idx}: duplicate pending/approved serial number {serial}")
        elif table == "unprocessed_software":
            key, tag = normalized.get("license_key", ""), normalized.get("desktop_laptop_tag", "")
            if key and tag and conn.execute("SELECT 1 FROM software_assignments WHERE desktop_laptop_tag=? AND license_key=? UNION SELECT 1 FROM unprocessed_software WHERE desktop_laptop_tag=? AND license_key=? AND approval_status='Pending'", (tag, key, tag, key)).fetchone():
                errors.append(f"Row {idx}: duplicate software allocation for {tag}/{key}")
        elif table == "unprocessed_employees":
            ecode = normalized.get("e_code", "")
            if ecode and conn.execute("SELECT 1 FROM employees WHERE e_code=? UNION SELECT 1 FROM unprocessed_employees WHERE e_code=? AND approval_status='Pending'", (ecode, ecode)).fetchone():
                errors.append(f"Row {idx}: duplicate pending/approved employee {ecode}")
    conn.close()
    return errors


@app.route("/inventory/<kind>/upload", methods=["POST"])
@login_required
@roles_required("Admin", "Editor")
def upload(kind: str):
    maps = {"unprocessed-assets": ("unprocessed_assets", UNPROCESSED_ASSET_COLUMNS), "unprocessed-software": ("unprocessed_software", SOFTWARE_ASSIGN_COLUMNS), "unprocessed-employees": ("unprocessed_employees", EMPLOYEE_COLUMNS)}
    if kind not in maps:
        return redirect(url_for("inventory", kind=kind))
    table, columns = maps[kind]
    try:
        records = parse_upload(request.files["file"])
        errors = validate_upload(table, columns, records)
        if errors:
            for error in errors[:10]:
                flash(error, "danger")
            if len(errors) > 10:
                flash(f"{len(errors) - 10} more validation errors hidden. Correct the file and upload again.", "danger")
            log_action("upload_validation_failed", table, details={"errors": errors[:25]})
            return redirect(url_for("inventory", kind=kind))
        for record in records:
            insert_record(table, columns, record, {"maker": session["username"], "approval_status": "Pending", "processed": "0"})
        create_notification(session["username"], "Approval", f"{len(records)} {kind} records submitted for approval")
        flash(f"Validated and uploaded {len(records)} pending records.", "success")
    except Exception as exc:
        flash(f"Upload failed: {exc}", "danger")
    return redirect(url_for("inventory", kind=kind))


@app.route("/record/<table>/<int:record_id>/edit", methods=["GET", "POST"])
@login_required
def edit_record(table: str, record_id: int):
    allowed = {"assets": ASSET_COLUMNS, "software": SOFTWARE_COLUMNS, "employees": EMPLOYEE_COLUMNS, "unprocessed_assets": UNPROCESSED_ASSET_COLUMNS, "unprocessed_software": SOFTWARE_ASSIGN_COLUMNS, "unprocessed_employees": EMPLOYEE_COLUMNS}
    if table not in allowed:
        return redirect(url_for("dashboard"))
    if current_user()["role"] == "Viewer":
        flash("Viewers cannot edit records.", "danger"); return redirect(url_for("dashboard"))
    columns = allowed[table]
    if request.method == "POST":
        try:
            update_record(table, columns, record_id, request.form)
            flash("Record updated.", "success")
        except sqlite3.IntegrityError as exc:
            flash(f"Duplicate entry blocked: {exc}", "danger")
    conn = db(); row = conn.execute(f"SELECT * FROM {table} WHERE id=?", (record_id,)).fetchone(); conn.close()
    return render_template("edit.html", table=table, row=row, columns=columns)


@app.route("/approve/<kind>/<int:record_id>", methods=["POST"])
@login_required
@roles_required("Admin")
def approve(kind: str, record_id: int):
    try:
        {"asset": approve_asset, "software": approve_software, "employee": approve_employee}[kind](record_id)
        create_notification(session["username"], "Change", "Approval completed and final inventory updated")
        flash("Request approved and posted to final inventory.", "success")
    except (KeyError, sqlite3.IntegrityError, ValueError) as exc:
        flash(f"Approval failed: {exc}", "danger")
    target = {"asset": "unprocessed-assets", "software": "unprocessed-software", "employee": "unprocessed-employees"}.get(kind, "assets")
    return redirect(url_for("inventory", kind=target))


@app.route("/reject/<table>/<int:record_id>", methods=["POST"])
@login_required
@roles_required("Admin")
def reject(table: str, record_id: int):
    if table not in ("unprocessed_assets", "unprocessed_software", "unprocessed_employees"):
        return redirect(url_for("dashboard"))
    execute(f"UPDATE {table} SET approval_status='Rejected', checker=?, updated_at=? WHERE id=?", (session["username"], now(), record_id))
    log_action("reject", table, record_id)
    create_notification(session["username"], "Approval", f"{table} request #{record_id} rejected")
    flash("Request rejected.", "warning")
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/reports/<name>")
@login_required
def report(name: str):
    start, end, asset_filter, software_filter = request.args.get("start", ""), request.args.get("end", ""), request.args.get("asset", ""), request.args.get("software", "")
    conn = db()
    if name == "inventory":
        columns = ["device_tag"] + EMPLOYEE_COLUMNS + [c for c in ASSET_COLUMNS if c != "e_code"]
        sql = "SELECT a.device_tag,e.*,a.location,a.asset_type,a.action,a.status,a.username,a.system_name,a.previous_user_history,a.serial_number,a.ip_address,a.device_model,a.ram_gb,a.storage_type,a.storage_capacity,a.purchase_year,a.invoice_number,a.finance_code,a.supplier_name,a.asset_value,a.warranty_amc,a.eol,a.insurance_status,a.insurance_type FROM assets a LEFT JOIN employees e ON a.e_code=e.e_code WHERE 1=1"
        params=[]
        if asset_filter: sql += " AND a.asset_type=?"; params.append(asset_filter)
        if start: sql += " AND date(a.created_at)>=date(?)"; params.append(start)
        if end: sql += " AND date(a.created_at)<=date(?)"; params.append(end)
        rows = conn.execute(sql, params).fetchall(); title = "Inventory Report"
    elif name == "software":
        columns = SOFTWARE_ASSIGN_COLUMNS
        sql = "SELECT * FROM software_assignments WHERE 1=1"; params=[]
        if software_filter: sql += " AND software_name LIKE ?"; params.append(f"%{software_filter}%")
        if start: sql += " AND date(created_at)>=date(?)"; params.append(start)
        if end: sql += " AND date(created_at)<=date(?)"; params.append(end)
        rows = conn.execute(sql, params).fetchall(); title = "Software Report"
    elif name == "allotted":
        columns = ["employee_name", "e_code", "device_tag", "asset_type", "created_at", "previous_user_history"]
        sql = "SELECT e.employee_name,a.e_code,a.device_tag,a.asset_type,a.created_at,a.previous_user_history FROM assets a LEFT JOIN employees e ON a.e_code=e.e_code WHERE a.e_code IS NOT NULL AND a.e_code!=''"
        params=[]
        if asset_filter: sql += " AND a.asset_type=?"; params.append(asset_filter)
        if start: sql += " AND date(a.created_at)>=date(?)"; params.append(start)
        if end: sql += " AND date(a.created_at)<=date(?)"; params.append(end)
        rows = conn.execute(sql, params).fetchall(); title = "Employee Allotted Assets Report"
    else:
        columns = ASSET_COLUMNS
        sql = "SELECT * FROM assets WHERE (e_code IS NULL OR e_code='' OR username IS NULL OR username='')"
        params=[]
        if asset_filter: sql += " AND asset_type=?"; params.append(asset_filter)
        if start: sql += " AND date(created_at)>=date(?)"; params.append(start)
        if end: sql += " AND date(created_at)<=date(?)"; params.append(end)
        rows = conn.execute(sql, params).fetchall(); title = "Unallocated Assets"
    conn.close()
    if request.args.get("download"):
        return csv_response(title, columns, rows)
    return render_template("report.html", name=name, title=title, columns=columns, rows=rows, start=start, end=end, asset_filter=asset_filter, software_filter=software_filter)


def csv_response(title: str, columns: list[str], rows: list[sqlite3.Row]) -> Response:
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow([LABELS.get(c, c.replace("_", " ").title()) for c in columns])
    for row in rows:
        writer.writerow([row[c] if c in row.keys() else "" for c in columns])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename={title.lower().replace(' ', '_')}.csv"})


@app.route("/admin", methods=["GET", "POST"])
@login_required
@roles_required("Admin")
def admin():
    if request.method == "POST":
        action = request.form.get("action")
        try:
            if action == "create_user":
                execute("INSERT INTO users (username,password_hash,role,status,disabled,created_at) VALUES (?,?,?,?,?,?)",
                        (request.form["username"].strip(), generate_password_hash(request.form["password"]), request.form["role"], "Active", 0, now()))
                log_action("create_user", "users", details={"username": request.form["username"], "role": request.form["role"]})
            elif action == "update_user":
                disabled = 1 if request.form.get("disabled") == "yes" else 0
                params: list[Any] = [request.form["role"], disabled, "Disabled" if disabled else "Active"]
                sql = "UPDATE users SET role=?, disabled=?, status=?"
                if request.form.get("new_password"):
                    sql += ", password_hash=?"; params.append(generate_password_hash(request.form["new_password"]))
                sql += " WHERE id=?"; params.append(request.form["user_id"])
                execute(sql, tuple(params)); log_action("update_user", "users", int(request.form["user_id"]))
            flash("User control updated.", "success")
        except sqlite3.IntegrityError as exc:
            flash(f"User update failed: {exc}", "danger")
        return redirect(url_for("admin"))
    conn = db()
    users = conn.execute("SELECT id,username,role,status,disabled,last_login,last_ip,created_at FROM users ORDER BY username").fetchall()
    logs = conn.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT 500").fetchall()
    assets = conn.execute("SELECT * FROM unprocessed_assets ORDER BY id DESC").fetchall()
    conn.close()
    return render_template("admin.html", users=users, roles=ROLES, logs=logs, assets=assets)


@app.route("/sample/<kind>")
@login_required
def sample_csv(kind: str):
    columns = SAMPLE_MAP.get(kind)
    if not columns:
        return redirect(url_for("dashboard"))
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([LABELS.get(c, c.replace("_", " ").title()) for c in columns])
    writer.writerow(["sample" if c not in DATE_FIELDS else date.today().isoformat() for c in columns])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename={kind}_sample.csv"})


@app.route("/delete/<table>/<int:record_id>", methods=["POST"])
@login_required
@roles_required("Admin", "Editor")
def delete_record(table: str, record_id: int):
    allowed = {"unprocessed_assets", "unprocessed_software", "unprocessed_employees", "vendors"}
    if table not in allowed:
        return redirect(url_for("dashboard"))
    if table.startswith("unprocessed"):
        execute(f"DELETE FROM {table} WHERE id=? AND approval_status!='Approved'", (record_id,))
    else:
        execute(f"DELETE FROM {table} WHERE id=?", (record_id,))
    log_action("delete", table, record_id)
    flash("Row deleted.", "warning")
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    user = current_user()
    if request.method == "POST":
        preferences = request.form.getlist("notifications")
        theme = request.form.get("theme", "dark")
        params: list[Any] = [request.form.get("full_name", ""), request.form.get("email", ""), request.form.get("phone", ""), request.form.get("department", ""), theme, json.dumps(preferences)]
        sql = "UPDATE users SET full_name=?, email=?, phone=?, department=?, theme=?, notification_preferences=?"
        if request.form.get("new_password"):
            if not check_password_hash(user["password_hash"], request.form.get("current_password", "")):
                flash("Current password is incorrect.", "danger")
                return redirect(url_for("profile"))
            sql += ", password_hash=?, password_changed_at=?"
            params.extend([generate_password_hash(request.form["new_password"]), now()])
        sql += " WHERE id=?"
        params.append(user["id"])
        execute(sql, tuple(params))
        session["theme"] = theme
        log_action("profile_update", "users", user["id"], {"theme": theme, "notifications": preferences})
        flash("Profile updated.", "success")
        return redirect(url_for("profile"))
    prefs = json.loads(user["notification_preferences"] or "[]") if user else []
    conn = db()
    notes = conn.execute("SELECT * FROM notifications WHERE username=? ORDER BY id DESC LIMIT 100", (user["username"],)).fetchall()
    conn.close()
    return render_template("profile.html", prefs=prefs, notes=notes)


@app.route("/notifications/read", methods=["POST"])
@login_required
def read_notifications():
    execute("UPDATE notifications SET is_read=1 WHERE username=?", (session["username"],))
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/admin/logs")
@login_required
@roles_required("Admin")
def admin_logs():
    q = request.args.get("q", "")
    conn = db()
    like = f"%{q}%"
    logs = conn.execute("SELECT * FROM audit_logs WHERE actor LIKE ? OR action LIKE ? OR table_name LIKE ? OR details LIKE ? ORDER BY id DESC LIMIT 500", (like, like, like, like)).fetchall()
    assets = conn.execute("SELECT id,device_tag,asset_type,approval_status,maker,checker,updated_at FROM unprocessed_assets ORDER BY id DESC LIMIT 300").fetchall()
    emails = conn.execute("SELECT * FROM email_logs ORDER BY id DESC LIMIT 300").fetchall()
    conn.close()
    return render_template("admin_logs.html", logs=logs, assets=assets, emails=emails, q=q)


@app.route("/admin/db", methods=["GET", "POST"])
@login_required
@roles_required("Admin")
def admin_db():
    result, columns = [], []
    predefined = {
        "pending_assets": "SELECT id, device_tag, asset_type, maker, created_at FROM unprocessed_assets WHERE approval_status='Pending' ORDER BY id DESC LIMIT 100",
        "expiring_licenses": "SELECT software_name, license_key, expiry_date, status FROM software_assignments WHERE expiry_date!='' ORDER BY expiry_date LIMIT 100",
        "inactive_employees": "SELECT e_code, employee_name, employee_status FROM employees WHERE employee_status!='Active' LIMIT 100",
        "unallocated_assets": "SELECT device_tag, asset_type, status, location FROM assets WHERE e_code IS NULL OR e_code='' LIMIT 100",
    }
    if request.method == "POST":
        action = request.form.get("action")
        if action == "backup":
            backup_dir = BASE_DIR / "backups"; backup_dir.mkdir(exist_ok=True)
            target = backup_dir / f"inventory_backup_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy2(DB_PATH, target)
            log_action("db_backup", "database", details=str(target))
            flash(f"Backup created: {target.name}", "success")
        elif action == "restore" and request.files.get("db_file"):
            request.files["db_file"].save(DB_PATH)
            log_action("db_restore", "database")
            flash("Database restored. Please restart the app if needed.", "warning")
        else:
            query = predefined.get(request.form.get("predefined"), request.form.get("query", "").strip())
            if not query.lower().startswith("select"):
                flash("Only SELECT queries are allowed from the web console.", "danger")
            else:
                conn = db(); cur = conn.execute(query); rows = cur.fetchall(); conn.close()
                columns = [d[0] for d in cur.description or []]
                result = rows
                log_action("db_query", "database", details=query)
    return render_template("admin_db.html", predefined=predefined, result=result, columns=columns)


@app.route("/admin/email", methods=["GET", "POST"])
@login_required
@roles_required("Admin")
def admin_email():
    if request.method == "POST":
        if request.form.get("action") == "save":
            execute("INSERT OR REPLACE INTO email_config (id,smtp_host,smtp_port,smtp_username,smtp_password,sender_email,use_tls,updated_at) VALUES (1?,?,?,?,?,?,?)",
                    (request.form.get("smtp_host", ""), request.form.get("smtp_port", "587"), request.form.get("smtp_username", ""), request.form.get("smtp_password", ""), request.form.get("sender_email", ""), 1 if request.form.get("use_tls") else 0, now()))
            flash("Email server configuration saved.", "success")
        elif request.form.get("action") == "test":
            status = send_email(request.form.get("test_recipient", ""), "EUC Inventory test email", "This is a test notification from EUC Inventory.")
            flash(status, "success" if status.startswith("Sent") else "danger")
        return redirect(url_for("admin_email"))
    conn = db()
    config = conn.execute("SELECT * FROM email_config WHERE id=1").fetchone()
    logs = conn.execute("SELECT * FROM email_logs ORDER BY id DESC LIMIT 200").fetchall()
    conn.close()
    return render_template("admin_email.html", config=config, logs=logs)


def send_email(recipient: str, subject: str, body: str) -> str:
    conn = db(); config = conn.execute("SELECT * FROM email_config WHERE id=1").fetchone(); conn.close()
    if not config or not recipient:
        status = "Email configuration or recipient missing"
    else:
        message = f"From: {config['sender_email']}\r\nTo: {recipient}\r\nSubject: {subject}\r\n\r\n{body}"
        try:
            server = smtplib.SMTP(config["smtp_host"], int(config["smtp_port"] or 587), timeout=10)
            if config["use_tls"]:
                server.starttls()
            if config["smtp_username"]:
                server.login(config["smtp_username"], config["smtp_password"])
            server.sendmail(config["sender_email"], [recipient], message)
            server.quit()
            status = "Sent test email"
        except Exception as exc:
            status = f"Email failed: {exc}"
    execute("INSERT INTO email_logs (recipient,subject,body,status,created_at) VALUES (?,?,?,?,?)", (recipient, subject, body, status, now()))
    return status


@app.route("/inventory/vendors", methods=["GET", "POST"])
@login_required
def vendors():
    if request.method == "POST" and current_user()["role"] in ("Admin", "Editor"):
        if request.files.get("file") and request.files["file"].filename:
            records = parse_upload(request.files["file"])
            for record in records:
                insert_record("vendors", VENDOR_COLUMNS, record)
            flash(f"Uploaded {len(records)} vendors.", "success")
        else:
            insert_record("vendors", VENDOR_COLUMNS, request.form)
            flash("Vendor saved.", "success")
        return redirect(url_for("vendors"))
    search = request.args.get("q", "")
    where, params = "1=1", ()
    if search:
        like = f"%{search}%"; where = " OR ".join(f"{c} LIKE ?" for c in VENDOR_COLUMNS); params = tuple([like] * len(VENDOR_COLUMNS))
    conn = db()
    rows = conn.execute(f"SELECT * FROM vendors WHERE {where} ORDER BY id DESC", params).fetchall()
    assets = conn.execute("SELECT supplier_name, COUNT(*) c FROM assets WHERE supplier_name!='' GROUP BY supplier_name").fetchall()
    software = conn.execute("SELECT vendor_supplier, COUNT(*) c FROM software_assignments WHERE vendor_supplier!='' GROUP BY vendor_supplier").fetchall()
    conn.close()
    return render_template("vendors.html", rows=rows, columns=VENDOR_COLUMNS, assets=assets, software=software)


@app.route("/checklist", methods=["GET", "POST"])
@login_required
def checklist():
    if request.method == "POST" and current_user()["role"] in ("Admin", "Editor"):
        ctype = request.form.get("checklist_type", "Joining")
        payload = {k: v for k, v in request.form.items() if k not in CHECKLIST_COLUMNS}
        record = {
            "checklist_type": ctype, "employee_name": request.form.get("employee_name", ""), "e_code": request.form.get("e_code", ""),
            "vertical": request.form.get("vertical", ""), "sub_vertical": request.form.get("sub_vertical", ""), "rfds": request.form.get("rfds", ""),
            "payload": json.dumps(payload), "asset_updates": request.form.get("asset_updates", ""), "software_updates": request.form.get("software_updates", ""),
            "employee_status": request.form.get("employee_status", "Active" if ctype != "Exit" else "Exited"), "status": "Pending"
        }
        rid = insert_record("checklists", CHECKLIST_COLUMNS, record, {"maker": session["username"], "processed": "0"})
        upload_dir = BASE_DIR / "uploads" / "checklists" / str(rid); upload_dir.mkdir(parents=True, exist_ok=True)
        for image in request.files.getlist("images"):
            if image and image.filename:
                target = upload_dir / image.filename
                image.save(target)
                execute("INSERT INTO checklist_images (checklist_id,file_path,original_name,created_at) VALUES (?,?,?,?)", (rid, str(target.relative_to(BASE_DIR)), image.filename, now()))
        create_notification(session["username"], "Approval", f"Checklist #{rid} submitted for approval")
        flash("Checklist submitted for approval.", "success")
        return redirect(url_for("checklist"))
    conn = db()
    pending = conn.execute("SELECT * FROM checklists WHERE status='Pending' ORDER BY id DESC LIMIT 200").fetchall()
    processed = conn.execute("SELECT * FROM checklists WHERE status='Approved' ORDER BY id DESC LIMIT 200").fetchall()
    employees = conn.execute("SELECT e_code, employee_name, vertical, sub_vertical, rfds FROM employees ORDER BY employee_name LIMIT 500").fetchall()
    conn.close()
    return render_template("checklist.html", pending=pending, processed=processed, employees=employees)


@app.route("/checklist/upload", methods=["POST"])
@login_required
@roles_required("Admin", "Editor")
def checklist_upload():
    ctype = request.form.get("checklist_type", "Joining")
    records = parse_upload(request.files["file"])
    for record in records:
        payload = row_dict(SAMPLE_MAP.get(f"checklist-{ctype.lower()}", []), record)
        insert_record("checklists", CHECKLIST_COLUMNS, {"checklist_type": ctype, "employee_name": payload.get("employee_name", ""), "e_code": payload.get("e_code", ""), "vertical": payload.get("vertical", ""), "sub_vertical": payload.get("sub_vertical", ""), "rfds": payload.get("rfds", ""), "payload": json.dumps(payload), "employee_status": "Active" if ctype != "Exit" else "Exited", "status": "Pending"}, {"maker": session["username"], "processed": "0"})
    flash(f"Uploaded {len(records)} checklist rows.", "success")
    return redirect(url_for("checklist"))


@app.route("/checklist/approve/<int:record_id>", methods=["POST"])
@login_required
@roles_required("Admin")
def approve_checklist(record_id: int):
    conn = db(); row = conn.execute("SELECT * FROM checklists WHERE id=?", (record_id,)).fetchone()
    if row:
        if row["e_code"]:
            conn.execute("UPDATE employees SET employee_status=?, updated_at=? WHERE e_code=?", (row["employee_status"] or "Active", now(), row["e_code"]))
            if row["checklist_type"] == "Exit":
                conn.execute("UPDATE assets SET status='Unallocated', e_code='', username='', updated_at=? WHERE e_code=?", (now(), row["e_code"]))
                conn.execute("UPDATE software_assignments SET status='Available', updated_at=? WHERE e_code=?", (now(), row["e_code"]))
        conn.execute("UPDATE checklists SET status='Approved', processed=1, checker=?, updated_at=? WHERE id=?", (session["username"], now(), record_id))
        conn.commit()
    conn.close(); log_action("approve_checklist", "checklists", record_id)
    flash("Checklist approved and inventory references updated.", "success")
    return redirect(url_for("checklist"))


@app.route("/checklist/reject/<int:record_id>", methods=["POST"])
@login_required
@roles_required("Admin")
def reject_checklist(record_id: int):
    execute("UPDATE checklists SET status='Rejected', checker=?, updated_at=? WHERE id=?", (session["username"], now(), record_id))
    flash("Checklist rejected.", "warning")
    return redirect(url_for("checklist"))


init_db()

if __name__ == "__main__":
    #app.run(debug=True)
    app.run(host="0.0.0.0", port=5000, debug=True)

