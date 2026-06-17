from dotenv import load_dotenv
import os
from typing import Optional, Tuple

load_dotenv()


def get_email_credentials() -> Tuple[Optional[str], Optional[str]]:
    """Return (EMAIL, APP_PASSWORD) from environment or .env file."""
    email = os.getenv('EMAIL')
    app_password = os.getenv('APP_PASSWORD')
    return email, app_password
