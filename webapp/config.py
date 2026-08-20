import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("DATA_DIR", BASE_DIR / "data")).resolve()
BAK_DIR = DATA_DIR / ".bak"
AUDIT_LOG = DATA_DIR / "audit.log"

ALLOWED_EXTENSION = ".xlsx"
MAX_UPLOAD_MB = 5
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

SESSION_IDLE_MINUTES = 60
SESSION_ABSOLUTE_DAYS = 30
SESSION_COOKIE_NAME = "ct_session"

LOGIN_RATE_LIMIT_BURST = os.environ.get("LOGIN_RATE_LIMIT_BURST", "10 per 15 minutes")
LOGIN_RATE_LIMIT_SUSTAINED = os.environ.get("LOGIN_RATE_LIMIT_SUSTAINED", "20 per hour")
LOGIN_RATE_LIMIT = f"{LOGIN_RATE_LIMIT_BURST};{LOGIN_RATE_LIMIT_SUSTAINED}"
DEFAULT_RATE_LIMIT = os.environ.get("DEFAULT_RATE_LIMIT", "60 per minute")

BACKUP_KEEP = 5
AUDIT_MAX_BYTES = 5 * 1024 * 1024

DEFAULT_APPEND_DIRECTION = "bottom"

RESERVED_COLUMN = "FLAGGED"

BANK_ACCOUNT_KEYWORDS = ("BANK", "ACCOUNT")

TIMEZONE = os.environ.get("TIMEZONE", "Asia/Kolkata")


def env(name: str, default: str) -> str:
    return os.environ.get(name, default)


def get_timezone_name() -> str:
    return TIMEZONE


def get_secret_key() -> str:
    return env("SECRET_KEY", "")


def get_password_hash() -> str:
    return env("PASSWORD_HASH", "")


def get_max_rows() -> int:
    try:
        return int(env("MAX_ROWS", "20000"))
    except ValueError:
        return 20000