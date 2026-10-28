# Customer Management and Payment Tracking

This repository contains a Flask-based application for managing customers and tracking payments.

Files included:
- `app.py` — main Flask app
- `create_db.py`, `database/` — database setup and pending statements
- `templates/`, `static/` — frontend templates and styles

To run locally:

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## Deploy to Render

This Flask application needs a Python web-service host; uploading the repository
to Netlify as a static site will not run the Flask server. The included
`render.yaml` configures a Render web service with a persistent disk for the
SQLite database and pending statement files.

1. Commit and push the project to a Git provider connected to your Render
   account. Uncommitted local changes are not included in a Git-based deploy.
2. In Render, create a new **Blueprint** and select the repository. Review the
   service and persistent-disk costs before confirming.
3. Set the prompted `ADMIN_USERNAME` and `ADMIN_PASSWORD` values to strong,
   unique credentials. Render generates `SECRET_KEY` for the service.
4. Deploy and verify that the login page loads.

The production database starts empty. Local database files are excluded from
Git and are not migrated by this configuration. The SQLite disk is intended for
a single running service instance; use a managed database before scaling to
multiple instances.

This repo was generated and pushed via an automated assistant.
