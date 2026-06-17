import random
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

LOG_FILE = Path('logs.txt')


def random_delay(min_seconds: int = 20, max_seconds: int = 40) -> None:
    """Sleep for a random time between min_seconds and max_seconds."""
    delay = random.randint(min_seconds, max_seconds)
    print(f"Waiting {delay} seconds to avoid spam throttling...")
    time.sleep(delay)


def log_status(email: str, status: str, message: Optional[str] = None) -> None:
    """Append a log line to logs.txt with timestamp, email, status, and optional message."""
    try:
        LOG_FILE.touch(exist_ok=True)
        ts = datetime.now().isoformat()
        with LOG_FILE.open('a', encoding='utf-8') as f:
            line = f"{ts}\t{email}\t{status}"
            if message:
                line += f"\t{message}"
            f.write(line + "\n")
    except Exception:
        # Best-effort logging; do not raise errors that stop the sending loop
        pass
