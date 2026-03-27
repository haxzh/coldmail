# Cold Email Automation Dashboard

GitHub-ready Python project for sending personalized cold emails from Excel using Gmail SMTP, with Streamlit dashboard and status-aware processing.

## Features

- Reads Excel with pandas
- Required columns: Name, Email, Company Name, Title, Status
- Processes only rows where Status is empty or Not Sent
- Skips already sent rows automatically
- Personalized subject/body placeholders: {name}, {company}, {title}
- PDF resume attachment support
- Per-row Excel update with Sent or Failed status
- Optional Date column update
- Retry support for temporary failures
- Random delay between emails
- Per-run send limit in dashboard
- Streamlit upload support for deploy use (Excel, Resume, Template)

## Project Files

- email_automation.py: automation engine (CLI + optional Tkinter)
- streamlit_app.py: Streamlit dashboard
- email_template.txt: editable mail body template
- contacts_template.csv: sample data format
- requirements.txt: dependencies

## Quick Start

1. Install dependencies

```bash
pip install -r requirements.txt
```

2. Start dashboard

```bash
streamlit run streamlit_app.py
```

3. In dashboard sidebar

- Enter Sender Gmail and Gmail App Password
- Upload Excel, Resume PDF, and Template TXT in Quick File Select
- Set Max Emails Per Run if needed
- Click Start Automation

4. Download updated Excel after run

- Use Download Updated Excel button shown after successful run

## Excel Format

Use these columns exactly:

- Name
- Email
- Company Name
- Title
- Status
- Date (optional)

Reference sample: contacts_template.csv

## Status Rules

- Empty or Not Sent: will be processed
- Sent: skipped
- Failed: skipped by default (unless you reset status manually)

## Security Notes

- Use Gmail App Password only (not your normal Gmail login password)
- Do not commit personal files or secrets
- .gitignore already excludes:
  - mail_settings.json
  - .runtime_uploads/
  - Excel files
  - resume.pdf

## GitHub Push Checklist

1. Confirm no secret data in files
2. Keep only template/sample data
3. Run local check

```bash
python -m py_compile email_automation.py streamlit_app.py
```

4. Commit and push

```bash
git init
git add .
git commit -m "Initial commit: cold email automation dashboard"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```
