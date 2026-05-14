# EUC Assets Inventory

A Flask and SQLite web application for EUC inventory management with dashboards, maker/checker approvals, CSV/XLSX uploads, reports, audit logging, and role-based administration.

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

## Features

- Dashboard for total assets, employees, approved/pending assets, users, asset type counts, software used/free counts, audit logs, and login history.
- Inventory pages for final assets, licensed software, employees, and unprocessed staging tables.
- CSV/XLSX uploads for unprocessed assets, software allocations, and employees.
- Maker/checker approval flow that posts approved records into final tables.
- Duplicate protection for asset tags, serial numbers, employee codes, software license keys, and per-device software allocations.
- Reports with filters and CSV downloads for inventory, software, allotted assets, and unallocated assets.
- Admin user control with hashed passwords, roles, disable/update controls, and generated password helper.
- Every auditable action is stored in `audit_logs` and appended to `logs/activity.log`.

