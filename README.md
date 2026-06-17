# Gmail Bulk Mailer

[![Python](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

A small production-ready Python utility to read recipients from Excel and send personalized emails via Gmail SMTP with an attachment (resume). Designed for safe, repeatable runs with testing, dry-run, retries, randomized delays, and HTML email support.

**Highlights**
- Supports `--test` and `--prod` datasets.
- Dry-run mode for safe previews.
- Retry and per-message randomized delay to reduce throttling.
- Logs all activity to `logs.txt`.

**Repository files**
- `main.py` — CLI entrypoint for sending emails.
- `mailer.py` — SMTP sending logic and retry handling.
- `config.py` — environment and settings loader.
- `utils.py` — helper utilities.
- [requirements.txt](requirements.txt) — Python dependencies.

## Quickstart

Prerequisites: Python 3.8+ and an App Password for your Gmail account (enable 2FA and create an App Password).

1. Clone the repository and create a virtual environment:

```bash
python -m venv .venv
source .venv/bin/activate    # macOS / Linux
.venv\\Scripts\\Activate.ps1  # Windows PowerShell
```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Create a `.env` file (copy from `.env.example` if present) and set:

```
EMAIL=your_email@gmail.com
APP_PASSWORD=your_app_password
```

4. Provide an Excel file with the required columns (see Excel Format below) and a resume PDF.

## Excel Format
The Excel sheet must include the following exact column headers (whitespace is stripped):

- "Contact Person Name"
- "Email_ID"
- "Company Name"

Example filenames used by the CLI: `testing.xlsx` (small dataset) and `500 HR's Email ID.xlsx` (main dataset).

## Usage
Run the script with one of the modes and provide a resume to attach:

```bash
python main.py --test --resume resume.pdf
python main.py --prod --resume resume.pdf --subject "Application for SDE Role"
python main.py --prod --resume resume.pdf --dry-run
```

Options:
- `--test` : use `testing.xlsx`
- `--prod` : use `500 HR's Email ID.xlsx`
- `--resume` : path to resume PDF to attach (required)
- `--subject` : email subject (default: "Application")
- `--dry-run` : prints messages instead of sending
- `--html` : sends an HTML email body

## Logs
All activity is appended to `logs.txt` with timestamps, recipient email, and status (success/fail/skipped). Use this file to audit runs.

## Gmail Limits & Safety
- Gmail daily sending limits vary (≈500/day). For safety target 100–150/day for regular accounts.
- Use `--dry-run` to verify content before sending.

## Troubleshooting
- Authentication errors: verify `EMAIL` and `APP_PASSWORD` in `.env` and confirm App Password is active.
- Attachment errors: verify the path passed to `--resume` is readable.
- Excel errors: ensure required headers are present exactly as specified.

## Contributing
Feel free to open issues or PRs. For quick local testing, use `--test` with a small Excel file and `--dry-run`.

## License
This project is provided under the MIT License — add a `LICENSE` file to the repository.

---

If you'd like, I can also add a `LICENSE` file and a `.gitignore` suitable for Python projects, and prepare a short commit message for pushing to GitHub.

