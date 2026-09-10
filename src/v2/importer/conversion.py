"""
Conversion functions for ImportSpec column processing — FS section 5.4.1.

These functions parse raw text values into native Python date/float values,
used exclusively by the Import Module (importer/) to convert CSV/external
file columns at import time. This is the ONE place in the framework where
text-to-date and text-to-number parsing happens — validation types
(DATE, NUMERIC, INTEGER) require native values and do not parse text
themselves (FS sections 6.9-6.11).
"""
import math
from datetime import datetime


def convert_date(text: str, date_formats: list[str]):
    """
    Parse text against date_formats in order. Returns a datetime on the
    first match, or None if no format matches.
    """
    if text is None:
        return None
    text = str(text).strip()
    if not text:
        return None

    for fmt in date_formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    return None


def convert_numeric(text: str, decimal_sep: str = ".", thousands_sep: str | None = None):
    """
    Parse text as a float, validating thousands-grouping structure rather
    than blindly stripping separators. Returns None if not parseable or
    if the separators are not used in a structurally valid way.

    Structural rules enforced when thousands_sep is configured:
    - thousands_sep groups must contain exactly 3 digits each, except
      the leftmost group which may have 1-3 digits
    - thousands_sep may not appear after the decimal point
    - decimal_sep may appear at most once, and only after all
      thousands groups
    """
    if text is None:
        return None

    raw = str(text).strip()
    if not raw:
        return None

    sign = ""
    if raw and raw[0] in "+-":
        sign, raw = raw[0], raw[1:]

    if decimal_sep and decimal_sep in raw:
        parts = raw.split(decimal_sep)
        if len(parts) != 2:
            return None  # more than one decimal separator
        int_part, frac_part = parts
        if not frac_part.isdigit():
            return None
    else:
        int_part, frac_part = raw, ""

    if not int_part:
        return None

    if thousands_sep:
        if thousands_sep in frac_part:
            return None
        if thousands_sep in int_part:
            groups = int_part.split(thousands_sep)
            if not groups[0].isdigit() or not (1 <= len(groups[0]) <= 3):
                return None
            for g in groups[1:]:
                if not g.isdigit() or len(g) != 3:
                    return None
            int_part = "".join(groups)
        else:
            if not int_part.isdigit():
                return None
    else:
        if not int_part.isdigit():
            return None

    normalized = f"{sign}{int_part}.{frac_part}" if frac_part else f"{sign}{int_part}"

    try:
        parsed = float(normalized)
        if math.isnan(parsed) or math.isinf(parsed):
            return None
        return parsed
    except ValueError:
        return None


def convert_integer(text: str, decimal_sep: str = ".", thousands_sep: str | None = None):
    """
    Parse text as a number (per convert_numeric) and require the result
    to have a zero fractional part. Per FS section 6.11 ERP migration
    context, "123.0" and "123.000" are valid integers — only a genuinely
    non-zero fractional part is rejected.
    Returns the parsed float (not int — callers decide native cell type)
    or None if not parseable or fractional part is non-zero.
    """
    parsed = convert_numeric(text, decimal_sep, thousands_sep)
    if parsed is None:
        return None
    if parsed != int(parsed):
        return None
    return parsed
