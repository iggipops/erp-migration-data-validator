"""
Shared helpers for built-in validation type modules.
Avoids duplicating common patterns (issue recording, cell iteration)
across individual validator files.
"""


def iter_data_rows(sheet):
    """Yield row numbers for all data rows (excludes header row 1)."""
    return range(2, sheet.max_row + 1)
