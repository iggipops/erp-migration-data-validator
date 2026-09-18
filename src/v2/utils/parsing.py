"""
Small shared value-parsing helpers used across config loading and
function/validator parameter handling.
"""
from typing import Any

_TRUE_STRINGS  = {"true", "yes", "1"}
_FALSE_STRINGS = {"false", "no", "0"}


def parse_bool(value: Any, *, field_name: str = "value") -> bool:
    """
    Coerce a config/parameter value to bool.

    A plain Python bool is returned as-is. A string is matched
    case-insensitively (after stripping whitespace) against
    true/false, yes/no, 1/0 — this covers values that arrive as text,
    e.g. a quoted YAML scalar ("false") or a FunctionArguments cell
    ("trim=false"), where Python's own bool("false") would otherwise
    evaluate to True. Any other value falls back to bool(value).

    Raises ValueError with `field_name` in the message if a string
    value doesn't match any recognized boolean spelling.
    """
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in _TRUE_STRINGS:
            return True
        if normalized in _FALSE_STRINGS:
            return False
        raise ValueError(
            f"{field_name}: cannot interpret '{value}' as a boolean "
            f"(expected one of: true/false, yes/no, 1/0)"
        )
    return bool(value)
