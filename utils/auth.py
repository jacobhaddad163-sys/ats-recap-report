"""
Authentication module with security hardening.
- Rate limiting on login attempts
- Session timeout
- Input sanitization
- Server-side logging (no secrets exposed to client)
"""

import logging
import re
from datetime import datetime, timedelta

import streamlit as st

logger = logging.getLogger(__name__)

# ─── Constants ───────────────────────────────────────────────────────────────
MAX_LOGIN_ATTEMPTS = 5
LOCKOUT_MINUTES = 15
SESSION_TIMEOUT_HOURS = 8
MAX_NAME_LENGTH = 50


def _sanitize_name(name: str) -> str:
    """Sanitize user name: strip, limit length, remove dangerous chars."""
    name = str(name).strip()[:MAX_NAME_LENGTH]
    # Allow only alphanumeric, spaces, hyphens, periods
    name = re.sub(r'[^a-zA-Z0-9\s\-.]', '', name)
    return name.strip()


def _is_rate_limited() -> bool:
    """Check if login attempts are rate-limited."""
    attempts = st.session_state.get("_login_attempts", [])
    now = datetime.now()
    # Keep only recent attempts within the lockout window
    recent = [t for t in attempts if now - t < timedelta(minutes=LOCKOUT_MINUTES)]
    st.session_state["_login_attempts"] = recent
    return len(recent) >= MAX_LOGIN_ATTEMPTS


def _record_attempt():
    """Record a failed login attempt."""
    attempts = st.session_state.get("_login_attempts", [])
    attempts.append(datetime.now())
    st.session_state["_login_attempts"] = attempts


def _is_session_expired() -> bool:
    """Check if the current session has expired."""
    login_time = st.session_state.get("_login_time")
    if login_time is None:
        return True
    return datetime.now() - login_time > timedelta(hours=SESSION_TIMEOUT_HOURS)


def check_password() -> bool:
    """Returns True if the user is authenticated and session is valid."""
    if not st.session_state.get("authenticated", False):
        return False
    # Check session timeout
    if _is_session_expired():
        logger.info(f"Session expired for user: {st.session_state.get('user_name', 'unknown')}")
        logout()
        return False
    return True


def login(name: str, password: str) -> bool:
    """Validate credentials with rate limiting and logging."""
    # Rate limiting
    if _is_rate_limited():
        logger.warning(f"Login rate limit exceeded")
        return False

    # Sanitize name
    clean_name = _sanitize_name(name)
    if not clean_name:
        return False

    # Validate password
    correct_password = st.secrets.get("APP_PASSWORD", "")
    if not correct_password:
        logger.error("APP_PASSWORD not configured in secrets")
        return False

    if password != correct_password:
        _record_attempt()
        logger.warning(f"Failed login attempt for: {clean_name}")
        return False

    # Success
    st.session_state["authenticated"] = True
    st.session_state["user_name"] = clean_name.title()
    st.session_state["_login_time"] = datetime.now()
    logger.info(f"Successful login: {clean_name}")
    return True


def logout():
    """Clear all authentication and session state."""
    keys_to_clear = [
        "authenticated", "user_name", "_login_time",
        "_login_attempts", "output_excel", "output_filename",
        "category_configs",
    ]
    for key in keys_to_clear:
        st.session_state.pop(key, None)


def require_auth():
    """Redirect to login if not authenticated."""
    if not check_password():
        st.switch_page("app.py")
