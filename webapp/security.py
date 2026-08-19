import secrets

from flask import request, session

import config


def new_csrf_token() -> str:
    token = secrets.token_urlsafe(32)
    session["csrf_token"] = token
    return token


def get_csrf_token() -> str:
    token = session.get("csrf_token")
    if not token:
        token = new_csrf_token()
    return token


def validate_csrf() -> bool:
    token = session.get("csrf_token", "")
    if not token:
        return False
    provided = request.headers.get("X-CSRF-Token", "")
    return secrets.compare_digest(provided, token)


def safe_workbook_name(name) -> str | None:
    if not isinstance(name, str):
        return None
    name = name.strip()
    if not name.endswith(".xlsx"):
        return None
    if "/" in name or "\\" in name or name.startswith("."):
        return None
    path = (config.DATA_DIR / name).resolve()
    if not str(path).startswith(str(config.DATA_DIR.resolve())):
        return None
    return name


def safe_sheet_name(name) -> str | None:
    if not isinstance(name, str):
        return None
    name = name.strip()
    if not name:
        return None
    if len(name) > 31:
        return None
    if any(ch in name for ch in "[]:*?/\\"):
        return None
    return name


def security_headers(response):
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Cache-Control"] = "no-store"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; style-src 'self'; script-src 'self'; "
        "img-src 'self' data:; connect-src 'self'; font-src 'self'; "
        "frame-ancestors 'none'"
    )
    return response