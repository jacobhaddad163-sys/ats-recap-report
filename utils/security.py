"""
Security utilities for ATS Recap Report Builder.
- Input validation and sanitization
- File validation (MIME type, size, zip bomb protection)
- Excel formula injection prevention
- Rate limiting
"""

import io
import logging
import re
import zipfile
from datetime import datetime, timedelta

import streamlit as st

logger = logging.getLogger(__name__)

# ─── Constants ───────────────────────────────────────────────────────────────
MAX_FILE_SIZE = 100 * 1024 * 1024       # 100 MB
MAX_UNCOMPRESSED_SIZE = 500 * 1024 * 1024  # 500 MB
MAX_FILES_IN_ZIP = 10_000
MAX_TEXT_LENGTH = 500
MAX_TITLE_LENGTH = 200
MAX_CATEGORY_LINES = 200

# Excel formula injection characters
FORMULA_STARTERS = ('=', '+', '-', '@', '\t', '\r', '\n')

# XLSX magic number (ZIP PK header)
XLSX_MAGIC = b'PK\x03\x04'


# ─── Input Validation ────────────────────────────────────────────────────────

def sanitize_text(value: str, max_length: int = MAX_TEXT_LENGTH,
                  allow_special: bool = False) -> str:
    """
    Sanitize user text input.
    Strips whitespace, limits length, removes dangerous characters.
    """
    if not isinstance(value, str):
        value = str(value)

    value = value.strip()[:max_length]

    # Remove null bytes and control characters (except basic whitespace)
    value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', value)

    if not allow_special:
        # Remove characters that could cause injection
        # Allow: alphanumeric, spaces, hyphens, underscores, periods, commas,
        #        parentheses, slashes, ampersands, dollar signs, apostrophes
        value = re.sub(r"[^a-zA-Z0-9\s\-_.,()/'&$#:+]", '', value)

    return value.strip()


def sanitize_for_excel(value: str) -> str:
    """
    Prevent Excel formula injection.
    Prefixes dangerous characters so Excel treats them as text.
    """
    if not value:
        return value

    value = str(value).strip()

    # Remove control characters
    value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', value)

    # If starts with formula character, prefix with single quote to neutralize
    if value and value[0] in FORMULA_STARTERS:
        value = "'" + value

    return value


def validate_category_mapping(text: str) -> tuple:
    """
    Validate and parse category mapping text.
    Returns (is_valid: bool, result: dict or error_message: str)

    Expected format: CATEGORY NAME: REF1, REF2, REF3
    """
    if not text or not text.strip():
        return False, "Category mapping cannot be empty"

    lines = text.strip().split("\n")
    if len(lines) > MAX_CATEGORY_LINES:
        return False, f"Too many categories (max {MAX_CATEGORY_LINES})"

    category_map = {}
    for line_num, line in enumerate(lines, 1):
        line = line.strip()
        if not line:
            continue

        if ":" not in line:
            return False, f"Line {line_num}: Missing ':' separator. Format: CATEGORY NAME: REF1, REF2"

        parts = line.split(":", 1)
        cat_name = sanitize_text(parts[0].strip().upper(), max_length=100)
        refs_str = parts[1].strip()

        if not cat_name:
            return False, f"Line {line_num}: Empty category name"

        refs = []
        for ref in refs_str.split(","):
            ref = sanitize_text(ref.strip().upper(), max_length=10)
            if ref:
                # Refs should be alphanumeric only
                if re.match(r'^[A-Z0-9]+$', ref):
                    refs.append(ref)
                else:
                    return False, f"Line {line_num}: Invalid ref# '{ref}' (alphanumeric only)"

        if not refs:
            return False, f"Line {line_num}: No valid ref#s for category '{cat_name}'"

        if cat_name in category_map:
            # Merge refs for duplicate category names
            category_map[cat_name].extend(refs)
        else:
            category_map[cat_name] = refs

    if not category_map:
        return False, "No valid categories found"

    return True, category_map


# ─── File Validation ─────────────────────────────────────────────────────────

def validate_xlsx_file(file_bytes: bytes, filename: str) -> tuple:
    """
    Validate uploaded XLSX file for security.
    Returns (is_valid: bool, error_message: str or None)
    """
    # Check file size
    if len(file_bytes) > MAX_FILE_SIZE:
        size_mb = len(file_bytes) / (1024 * 1024)
        return False, f"File too large ({size_mb:.1f} MB). Maximum is {MAX_FILE_SIZE // (1024*1024)} MB."

    if len(file_bytes) < 100:
        return False, "File is too small to be a valid Excel file."

    # Check extension
    if not filename.lower().endswith(('.xlsx', '.xls')):
        return False, "Invalid file type. Please upload an .xlsx file."

    # Check magic number (XLSX files are ZIP archives)
    if filename.lower().endswith('.xlsx'):
        if not file_bytes[:4] == XLSX_MAGIC:
            return False, "File does not appear to be a valid XLSX file."

        # Check for zip bomb
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                # Check number of files in archive
                file_count = len(zf.namelist())
                if file_count > MAX_FILES_IN_ZIP:
                    return False, f"File contains too many internal files ({file_count})."

                # Check uncompressed size
                total_uncompressed = sum(info.file_size for info in zf.infolist())
                if total_uncompressed > MAX_UNCOMPRESSED_SIZE:
                    size_mb = total_uncompressed / (1024 * 1024)
                    return False, f"File uncompressed size too large ({size_mb:.0f} MB)."

                # Verify it has workbook.xml (basic XLSX structure check)
                if 'xl/workbook.xml' not in zf.namelist():
                    return False, "File is missing required Excel structure."

        except zipfile.BadZipFile:
            return False, "File is corrupted or not a valid ZIP/XLSX file."
        except Exception as e:
            logger.error(f"File validation error: {e}")
            return False, "Could not validate file. Please try again."

    return True, None


# ─── Rate Limiting ───────────────────────────────────────────────────────────

def check_rate_limit(action: str, max_requests: int = 20,
                     window_seconds: int = 3600) -> bool:
    """
    Check if action is rate-limited.
    Returns True if action is ALLOWED, False if rate-limited.
    """
    key = f"_rate_limit_{action}"
    now = datetime.now()

    timestamps = st.session_state.get(key, [])
    # Keep only timestamps within the window
    timestamps = [t for t in timestamps if now - t < timedelta(seconds=window_seconds)]

    if len(timestamps) >= max_requests:
        logger.warning(f"Rate limit exceeded for action: {action}")
        return False

    timestamps.append(now)
    st.session_state[key] = timestamps
    return True
