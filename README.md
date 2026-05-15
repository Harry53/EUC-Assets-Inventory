# EUC Assets Inventory

A Flask and SQLite web application for EUC inventory management with an administration-style UI, dark/light themes, dashboards, maker/checker approvals, CSV/XLSX ingestion, reports, notifications, vendor management, checklist workflows, audit logging, database administration, and email server configuration.

## Run locally

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python app.py
```

The first run creates `inventory.db` and an admin account:

- Username: `admin`
- Password: `admin123`

Change the default password immediately in **Administration → Active User Management**.

## Major features

- Dashboard with KPI cards, asset/software/employee graphs, compact audit tables, login history, and approval summaries.
- Dark/light theme support from the profile page, plus an instant header theme toggle.
- Profile page for user details, password reset, notification preferences, and personal notification history.
- Notification tray for approval, change, asset removal, password expiry, and license expiry notifications.
- Final inventory pages for assets, licensed software, employees, and vendor management.
- Unprocessed staging pages with maker/checker approvals, duplicate validation before upload, sample CSV downloads, row delete, and edit only while pending/rejected.
- Checklist module for Joining, Replacement, Rebuild, and Exit workflows, including multi-image upload to the server filesystem and approval updates to employee/assets/software statuses.
- Reports with filters and CSV downloads for inventory, software, allotted assets, and unallocated assets.
- Administration sub-pages for users, logs, database backup/restore/query/predefined queries, and SMTP email server configuration with test email and mail logs.
- Every auditable action is stored in `audit_logs` and appended to `logs/activity.log`.

## Upload notes

Use the **Sample CSV** buttons on unprocessed inventory, vendor, and checklist pages to download correct headers before uploading. Uploads validate required columns and duplicate approved/pending records before inserting into unprocessed tables.

