# Employee Ledger — Employee Management System

A Flask web app for a college/placement project. It demonstrates three things
together: **user authentication**, **Excel (.xlsx) as a working database**,
and an **automated email sent on successful login**.

## Features

- Sign up / log in (passwords are hashed with Werkzeug, never stored as plain text)
- On every successful login, an email is sent **to the email address tied to
  that account**, confirming the login (with timestamp)
- Employee records are stored in `data/employees.xlsx` — open it directly in
  Excel any time, outside the app
- Full CRUD: add, edit, delete employee records from the dashboard
- "Export .xlsx" button downloads the current Excel file
- User accounts themselves are stored in `data/users.xlsx`

## Tech stack

| Layer       | Choice                                   |
|-------------|-------------------------------------------|
| Backend     | Python, Flask                             |
| Data store  | Excel files via `openpyxl` (no SQL database) |
| Auth        | Flask session + Werkzeug password hashing |
| Email       | `smtplib` + Gmail SMTP (App Password)     |
| Frontend    | Server-rendered HTML/CSS (Jinja2 templates) |

## Project structure

```
ems/
├── app.py                 # routes, Excel read/write helpers, auth
├── mailer.py               # sends the "successful login" email
├── requirements.txt
├── .env.example             # copy to .env and fill in your values
├── data/
│   ├── users.xlsx           # created automatically on first run
│   └── employees.xlsx       # created automatically, seeded with sample rows
├── templates/               # login, signup, dashboard, employee form
└── static/style.css
```

## Setup

1. **Create a virtual environment and install dependencies**
   ```bash
   cd ems
   python3 -m venv venv
   source venv/bin/activate      # Windows: venv\Scripts\activate
   pip install -r requirements.txt
   ```

2. **Create your `.env` file**
   ```bash
   cp .env.example .env
   ```
   Open `.env` and set:
   - `SECRET_KEY` — any random string
   - `EMAIL_ADDRESS` — the Gmail account that will *send* the alert
   - `EMAIL_PASSWORD` — a Gmail **App Password** (see below — this is not your normal password)

3. **Run the app**
   ```bash
   python app.py
   ```
   Open `http://127.0.0.1:5000` in your browser.

4. **Try it**: Sign up with your own real email → log in → check that inbox
   for "Successful Login Alert - Employee Management System".

If you skip step 2, the app still works fully — login just shows a warning
that the email couldn't be sent, instead of crashing.

## Setting up the login email (Gmail App Password)

Gmail blocks plain-password SMTP logins, so you need an **App Password**:

1. Go to your Google Account → **Security**
2. Turn on **2-Step Verification** (required before App Passwords are available)
3. Go to **Security → App passwords**
4. Create one for "Mail" / "Other", name it e.g. "Employee Ledger"
5. Google gives you a 16-character password — copy it into `.env` as `EMAIL_PASSWORD`
   (no spaces)

Use that Gmail address as `EMAIL_ADDRESS` too — Gmail's SMTP server only lets
an account send mail as itself.

## How the pieces fit together (useful for your viva/interview)

- **Why Excel instead of a database?** It keeps the data store transparent
  and human-readable — anyone can open `employees.xlsx` directly — and
  demonstrates file I/O (`openpyxl`) instead of hiding it behind an ORM.
  Each "table" (users, employees) is just a sheet with a header row.
- **Why hash passwords?** So that even someone who opens `users.xlsx`
  directly can't read the actual passwords — `generate_password_hash` /
  `check_password_hash` from Werkzeug.
- **Why does login email use the account's *own* email, not a fixed one?**
  Each row in `users.xlsx` stores the email the user signed up with;
  `mailer.send_login_alert()` is called with that value at login time, so the
  confirmation always goes to the right inbox.
- **What happens if email fails?** `mailer.py` catches the SMTP exception and
  returns `(False, error_message)` instead of raising — login still succeeds,
  and the dashboard shows a warning instead of a server error. This is worth
  mentioning if asked about error handling.

## Notes

- This uses the Flask **development server** — fine for a project demo, not
  for production deployment.
- `data/*.xlsx` and `.env` are gitignored; don't commit real credentials.
