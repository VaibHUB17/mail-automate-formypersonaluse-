import smtplib
from pathlib import Path
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional, Tuple


class Mailer:
    def __init__(self, smtp_server: str = 'smtp.gmail.com', port: int = 587):
        self.smtp_server = smtp_server
        self.port = port

    def send_email(self, *, to_email: str, to_name: str, company: str, subject: str,
                   body_plain: str, body_html: Optional[str], resume_path: Path,
                   from_email: str, app_password: str) -> Tuple[bool, Optional[str]]:
        """
        Send an email with optional HTML and a PDF resume attachment.
        Returns (True, None) on success or (False, error_message) on failure.
        """
        try:
            if not Path(resume_path).exists():
                return False, f"Resume file not found: {resume_path}"

            # Build message
            msg = MIMEMultipart('mixed')
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = subject

            # Alternative part for plain and html
            alt = MIMEMultipart('alternative')
            alt.attach(MIMEText(body_plain, 'plain'))
            if body_html:
                alt.attach(MIMEText(body_html, 'html'))

            msg.attach(alt)

            # Attach resume
            with open(resume_path, 'rb') as f:
                part = MIMEBase('application', 'pdf')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            filename = Path(resume_path).name
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)

            # Send via SMTP
            server = smtplib.SMTP(self.smtp_server, self.port, timeout=60)
            server.ehlo()
            server.starttls()
            server.login(from_email, app_password)
            server.sendmail(from_email, to_email, msg.as_string())
            server.quit()
            return True, None

        except Exception as e:
            return False, str(e)
