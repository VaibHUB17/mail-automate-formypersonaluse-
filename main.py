#!/usr/bin/env python3
import argparse
import math
from pathlib import Path
import sys
import time

import pandas as pd
import json

from mailer import Mailer
import utils
import config


def parse_args():
    p = argparse.ArgumentParser(description="Send personalized emails from Excel lists or export draft files")

    # Sending modes (keeps backward compatibility)
    group = p.add_mutually_exclusive_group(required=False)
    group.add_argument('--test', action='store_true', help='Use testing.xlsx')
    group.add_argument('--prod', action='store_true', help='Use 500 HR\'s Email ID.xlsx')

    # Export/draft mode
    p.add_argument('--input', type=str, help='Path to the HR Excel workbook (.xlsx)')
    p.add_argument('--output', default='outbox', help='Directory where draft emails will be written')
    p.add_argument('--batch-size', type=int, default=100, help='How many drafts to generate per batch')
    p.add_argument('--batch-number', type=int, default=1, help='1-based batch number to export')

    # Sender / profile info used for drafts and signature
    p.add_argument('--sender-name', required=False, help="Your name to place in the signature")
    p.add_argument('--sender-email', required=False, help="Your email to place in the signature")
    p.add_argument('--linkedin-url', default='https://www.linkedin.com/in/vaibhav-shivhare17', help='LinkedIn profile URL')
    p.add_argument('--github-url', default='https://github.com/VaibHUB17', help='GitHub profile URL')

    # Resume / sending options
    p.add_argument('--resume', type=str, required=False, help='Optional path to the resume file to attach')
    p.add_argument('--subject', type=str, default=None, help='Email subject')
    p.add_argument('--daily-limit', type=int, default=100, help='Maximum successful emails to send per run/day')
    p.add_argument('--dry-run', action='store_true', help='Print messages instead of sending')
    p.add_argument('--html', action='store_true', help='Send HTML formatted email (bonus)')
    p.add_argument('--reset-checkpoint', action='store_true', help='Clear saved progress checkpoint for the selected Excel file')
    return p.parse_args()


def load_dataframe(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {path}")
    df = pd.read_excel(path, engine='openpyxl')
    # Strip spaces from column names
    df.columns = df.columns.str.strip()
    required = {"Contact Person Name", "Email_ID", "Company Name"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in Excel: {missing}")
    return df


def clean_cell(value) -> str:
    if pd.isna(value):
        return ''
    return str(value).strip()


def build_body(greeting_name: str, company: str, sender_name: str, sender_email: str,
               linkedin: str, github: str, phone: str) -> tuple:
    """
    Return (plain_text_body, html_body_or_None) using the standardized template.
    """
    company_line = f"at {company}" if company else ''
    plain = (
        f"Hi {greeting_name},\n\n"
        f"I'm reaching out to inquire about Software Engineer or Full-Stack Developer opportunities at {company}.\n\n"
        f"I recently graduated with a B.Tech in Information Technology from MITS Gwalior (CGPA: 8.3) and have worked as a Software Developer Intern at DataVinci Analytics. Additionally, I have developed AI-powered applications, an MCP server, automation tools, and full-stack SaaS platforms.\n\n"
        f"Tech stack: C++, Python, JavaScript/TypeScript, React, Next.js, Node.js, SQL, MongoDB, REST APIs, Docker, CI/CD, Data Structures & Algorithms, System Design \n\n"
        f"I am available immediately for full-time opportunities and have attached my resume for your reference. I would appreciate the opportunity to discuss how my skills and experience can contribute to your team.\n\n"
        f"Best regards,\n{sender_name}\n{sender_email}\nPhone: {phone}\n"
        f"LinkedIn: {linkedin}\n"
        f"GitHub: {github}\n\n"
    )
    # Minimal HTML alternative (optional); keep simple formatting
    html = None
    return plain, html


CHECKPOINT_FILE = Path('.checkpoint.json')


def load_checkpoint(excel_name: str) -> int:
    """Return last processed 0-based index for excel_name, or -1 if none."""
    try:
        if not CHECKPOINT_FILE.exists():
            return -1
        data = json.loads(CHECKPOINT_FILE.read_text(encoding='utf-8'))
        return int(data.get(excel_name, -1))
    except Exception:
        return -1


def save_checkpoint(excel_name: str, last_index: int) -> None:
    try:
        data = {}
        if CHECKPOINT_FILE.exists():
            data = json.loads(CHECKPOINT_FILE.read_text(encoding='utf-8'))
        data[excel_name] = int(last_index)
        CHECKPOINT_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')
    except Exception:
        pass


def clear_checkpoint(excel_name: str) -> None:
    try:
        if not CHECKPOINT_FILE.exists():
            return
        data = json.loads(CHECKPOINT_FILE.read_text(encoding='utf-8'))
        if excel_name in data:
            del data[excel_name]
            CHECKPOINT_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')
    except Exception:
        pass


def generate_drafts(args):
    """Generate draft emails from an input Excel and write them to the output directory.

    Batch numbering is 1-based. Files are written as <index>_<safe_email>.txt
    """
    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Excel file not found: {input_path}")

    df = load_dataframe(input_path)
    total = len(df)
    batch_size = int(args.batch_size)
    batch_number = int(args.batch_number)
    start = (batch_number - 1) * batch_size
    end = start + batch_size
    batch = df.iloc[start:end]

    outdir = Path(args.output)
    outdir.mkdir(parents=True, exist_ok=True)

    sender_name = args.sender_name or 'Vaibhav Shivhare'
    sender_email = args.sender_email or ''
    linkedin = args.linkedin_url
    github = args.github_url
    subject = args.subject or 'Application'

    written = 0
    for i, row in enumerate(batch.itertuples(index=False), start=start + 1):
        name = clean_cell(row[df.columns.get_loc('Contact Person Name')] if 'Contact Person Name' in df.columns else '')
        email = clean_cell(row[df.columns.get_loc('Email_ID')] if 'Email_ID' in df.columns else '')
        company = clean_cell(row[df.columns.get_loc('Company Name')] if 'Company Name' in df.columns else '')

        greeting_name = name.strip() or 'Hiring Team'
        company_line = f"at {company}" if company.strip() else ''
        body, _ = build_body(greeting_name, company, sender_name, sender_email, linkedin, github, phone='')

        safe_email = email.replace('@', '_at_').replace(' ', '_') if email else f'no-email-{i}'
        out_file = outdir / f"{i:04d}_{safe_email}.txt"

        header = f"To: {email}\nSubject: {subject}\nFrom: {sender_name} <{sender_email}>\n\n"
        out_file.write_text(header + body, encoding='utf-8')
        written += 1

    print(f"Exported {written} drafts (batch {batch_number}) to {outdir} ({total} total rows in sheet)")


def main():
    args = parse_args()
    # If --input provided, run export/draft generation mode and exit
    if args.input:
        try:
            generate_drafts(args)
        except Exception as e:
            print(f"Failed to export drafts: {e}")
            sys.exit(1)
        return

    # Otherwise, proceed with sending mode (original flow)
    excel_file = Path('testing.xlsx') if args.test else Path("500 HR's Email ID.xlsx")
    resume_path = Path(args.resume) if args.resume else None
    dry_run = args.dry_run

    if resume_path is None or not resume_path.exists():
        print(f"Resume file not found: {resume_path}")
        sys.exit(1)

    try:
        df = load_dataframe(excel_file)
    except Exception as e:
        print(f"Failed to load Excel: {e}")
        sys.exit(1)

    total = len(df)
    sent_count = 0
    daily_limit = max(1, int(args.daily_limit))

    mailer = Mailer()
    # Sender/profile defaults
    sender_name = args.sender_name or 'Vaibhav Shivhare'
    sender_email = args.sender_email or 'vaibhavshivhare1709@gmail.com'
    linkedin = args.linkedin_url
    github = args.github_url
    phone = '+91 98269 12535'

    # Note: subject may be customized per-recipient below (test mode includes company name)

    # Normalize dataframe index to zero-based positions for checkpointing
    df = df.reset_index(drop=True)

    excel_name = excel_file.name
    if args.reset_checkpoint:
        clear_checkpoint(excel_name)
        print(f"Checkpoint cleared for {excel_name}")

    start_idx = load_checkpoint(excel_name) + 1
    if start_idx > 0:
        print(f"Resuming from row {start_idx} (0-based) for {excel_name}")

    remaining_rows = max(0, total - start_idx)
    estimated_days = math.ceil(remaining_rows / daily_limit) if remaining_rows else 0
    print(f"Daily limit: {daily_limit} emails/day")
    if remaining_rows:
        print(f"Remaining rows: {remaining_rows} | Estimated days at this limit: {estimated_days}")

    for idx in range(start_idx, len(df)):
        row = df.iloc[idx]
        name = clean_cell(row.get('Contact Person Name', ''))
        email = clean_cell(row.get('Email_ID', ''))
        company = clean_cell(row.get('Company Name', ''))

        if not name and not email and not company:
            continue

        if not email or pd.isna(email):
            utils.log_status(email, 'skipped', 'Missing email')
            continue

        # Use the shared body template for both test and prod
        greeting_name = name or 'Hiring Team'
        body_plain, body_html = build_body(greeting_name, company, sender_name, sender_email, linkedin, github, phone)

        # Both test and prod use the company-specific subject when a company is available.
        if company:
            subject_to_use = f"Application for FTE Opportunity | Freshers at {company}"
        else:
            subject_to_use = "Application for FTE Opportunity | Freshers"

        # Retry logic: attempt up to 2 times
        attempt = 0
        success = False
        while attempt < 2 and not success:
            attempt += 1
            try:
                if dry_run:
                    print(f"[DRY-RUN] To: {email} | Subject: {subject_to_use}\n{body_plain}\n---")
                    success = True
                else:
                    # Get credentials only when actually sending
                    email_addr, app_password = config.get_email_credentials()
                    success, error = mailer.send_email(
                        to_email=email,
                        to_name=name,
                        company=company,
                        subject=subject_to_use,
                        body_plain=body_plain,
                        body_html=body_html if args.html else None,
                        resume_path=resume_path,
                        from_email=email_addr,
                        app_password=app_password,
                    )
                    if not success:
                        raise RuntimeError(error)

                if success:
                    sent_count += 1
                    print(f"Sent {sent_count}/{total} to {name}")
                    utils.log_status(email, 'success')
                    # Persist progress so we can resume after interruptions
                    save_checkpoint(excel_name, idx)
                    if sent_count >= daily_limit:
                        print(f"Daily limit of {daily_limit} reached. Stop here and rerun tomorrow to continue from the checkpoint.")
                        break
                else:
                    utils.log_status(email, 'fail', 'Unknown failure')

            except Exception as e:
                utils.log_status(email, 'fail', str(e))
                if attempt < 2:
                    print(f"Error sending to {email}, retrying (attempt {attempt + 1})...")
                    time.sleep(2)
                else:
                    print(f"Failed to send to {email} after retry: {e}")

        # Anti-spam delay only on actual sends (skip for dry-run)
        if not dry_run:
            utils.random_delay(20, 40)

        if sent_count >= daily_limit:
            break


if __name__ == '__main__':
    main()
