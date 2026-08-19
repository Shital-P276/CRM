import time

from flask import session
from werkzeug.security import check_password_hash

import config


def password_configured() -> bool:
    return bool(config.get_password_hash())


def verify_password(password: str) -> bool:
    if not password_configured():
        return False
    return check_password_hash(config.get_password_hash(), password)


def login(password: str) -> bool:
    if not verify_password(password):
        return False
    now = time.time()
    session.clear()
    session["authed"] = True
    session["created_at"] = now
    session["last_active"] = now
    return True


def logout() -> None:
    session.clear()


def is_authenticated() -> bool:
    if not session.get("authed"):
        return False
    now = time.time()
    if now - float(session.get("last_active", 0)) > config.SESSION_IDLE_MINUTES * 60:
        session.clear()
        return False
    if now - float(session.get("created_at", now)) > config.SESSION_ABSOLUTE_DAYS * 86400:
        session.clear()
        return False
    session["last_active"] = now
    return True