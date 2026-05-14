from __future__ import annotations

import csv
import io
import json
import os
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

LABELS = {
    "device_tag": "Device Tag", "desktop_laptop_tag": "Desktop/Laptop Tag", "location": "Location", "asset_type": "Asset Type",
    "serial_number": "Serial Number", "ip_address": "IP Address", "device_model": "Device Model", "ram_gb": "RAM (GB)",
    "storage_type": "Storage Type (HDD/SSD)", "storage_capacity": "Storage Capacity", "action": "Action", "status": "Status",
    "username": "Username", "system_name": "System Name", "previous_user_history": "Previous User History", "purchase_year": "Purchase Year",
    "invoice_number": "Invoice Number", "finance_code": "Finance Code", "supplier_name": "Supplier Name", "asset_value": "Asset Value",
    "warranty_amc": "Warranty/AMC", "eol": "End of Life (EOL)", "insurance_status": "Insurance Status", "insurance_type": "Insurance Type",
    "e_code": "E-Code", "employee_name": "Employee Name", "doj": "Date of Joining (YYYY-MM-DD)", "vertical": "Vertical",
    "sub_vertical": "Sub-Vertical", "content_category": "Content Category", "reporting_manager": "Reporting Manager (RM)", "rfds": "RFDs",
    "employee_status": "Employee Status", "software_name": "Software Name", "version": "Version", "license_type": "License Type",
    "license_key": "License Key", "number_of_licenses": "Number of Licenses", "purchase_date": "Purchase Date", "registration_email": "Registration email id",
    "installation_date": "Installation Date (YYYY-MM-DD)", "expiry_date": "Expiry Date (YYYY-MM-DD)", "vendor_supplier": "Vendor/Supplier",
    "cost": "Cost", "support_amc": "Support/AMC", "last_updated_date": "Last Updated Date (YYYY-MM-DD)", "remarks": "Remarks"
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
        last_login TEXT, last_ip TEXT, created_at TEXT NOT NULL
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
    """
    conn = db()
    conn.executescript(schema)
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assets_device_tag ON assets(device_tag)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assets_serial ON assets(serial_number) WHERE serial_number IS NOT NULL AND serial_number != ''")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_ecode ON employees(e_code)")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_software_key ON software(license_key) WHERE license_key IS NOT NULL AND license_key != ''")
    conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_assignment_device_license ON software_assignments(desktop_laptop_tag, license_key)")
    if not conn.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
        conn.execute("INSERT INTO users (username,password_hash,role,status,disabled,created_at) VALUES (?,?,?,?,?,?)",
                     ("admin", generate_password_hash("admin123"), "Admin", "Active", 0, now()))
    conn.commit()
    conn.close()


def cols(names: list[str]) -> str:
    return ", ".join(f"{name} TEXT" for name in names)


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


@app.context_processor
def inject_globals() -> dict[str, Any]:
    return {"user": current_user(), "labels": LABELS, "asset_types": ASSET_TYPES}


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
        for record in records:
            insert_record(table, columns, record, {"maker": session["username"], "approval_status": "Pending", "processed": "0"})
        flash(f"Uploaded {len(records)} pending records.", "success")
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


init_db()

if __name__ == "__main__":
    #app.run(debug=True)
    app.run(host="0.0.0.0", port=5000, debug=True)
