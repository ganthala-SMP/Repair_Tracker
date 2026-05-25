# Repair Control - Multi Company

## Default super admin login
- Email: `ganthala@starmobileparts.be`
- Password: `SMP@2000`

Change this immediately after first login.

## What changed in this hardening build
- Environment variable based configuration
- Secret key loaded from `SECRET_KEY`
- PostgreSQL support through `DATABASE_URL`
- Debug mode off by default
- Basic upload validation for logos
- Basic 403 / 404 / 500 error pages
- Session cookie hardening

## Local setup
```bash
python -m venv .venv
source .venv/bin/activate  # Windows PowerShell: .\.venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

Open http://127.0.0.1:5000

## PostgreSQL instead of SQLite
The app now reads `DATABASE_URL`.

Example for local PostgreSQL:
```env
DATABASE_URL=postgresql://repair_user:strongpassword@localhost:5432/repair_tracker
```

If `DATABASE_URL` is not set, the app falls back to SQLite for local testing.

### Windows local PostgreSQL workflow
1. Install PostgreSQL
2. Create a database and user
3. Set environment variables in PowerShell:
```powershell
$env:SECRET_KEY = "replace-with-a-long-random-secret"
$env:DATABASE_URL = "postgresql://repair_user:strongpassword@localhost:5432/repair_tracker"
$env:FLASK_DEBUG = "false"
```
4. Start the app:
```powershell
python run.py
```

## Production checklist
- Use PostgreSQL
- Set a strong `SECRET_KEY`
- Keep `FLASK_DEBUG=false`
- Put the app behind Gunicorn + Nginx
- Enable HTTPS
- Use the Data Manager to export backups regularly
- Test restore by importing a backup into a staging copy before going live

## Backup strategy
Use **Data Manager** to export:
- Full backup (JSON)
- Customers / Parts / Repairs (CSV)

Recommended:
- take a full JSON backup daily
- store backups outside the app folder
- test restore weekly on a copy of the app

## Upload safety
Logo uploads are limited by:
- file extension
- content type
- max size: 4 MB

## Notes
This build is suitable for pilot production once you move it to PostgreSQL and deploy it behind a proper server setup.
