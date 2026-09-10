import math
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# ---------------------------------------------------------------------------
# Workbook open / copy
# ---------------------------------------------------------------------------

def copy_workbook(source_path: str, output_path: str) -> None:
    """Copy input file to output path, creating parent dirs if needed."""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(source_path, output_path)


def open_workbook(path: str) -> Workbook:
    """Open an xlsx workbook with openpyxl (keep_vba=False)."""
    return load_workbook(path)


def save_workbook(workbook: Workbook, path: str) -> None:
    """Save the workbook to the given path."""
    workbook.save(path)


# ---------------------------------------------------------------------------
# Sheet helpers
# ---------------------------------------------------------------------------

def find_sheet_name(workbook: Workbook, sheet_name: str) -> str | None:
    """
    Case-insensitive worksheet name lookup (FS section 9: worksheet names
    are case-insensitive framework-wide). Returns the sheet's actual name
    as stored in the workbook (original case preserved) if a case-
    insensitive match exists, else None.
    """
    target = str(sheet_name).strip().lower()
    for name in workbook.sheetnames:
        if name.strip().lower() == target:
            return name
    return None


def sheet_exists(workbook: Workbook, sheet_name: str) -> bool:
    return find_sheet_name(workbook, sheet_name) is not None


def get_sheet(workbook: Workbook, sheet_name: str) -> Worksheet:
    """
    Case-insensitive worksheet lookup. Raises KeyError (matching the
    dict-access contract callers previously relied on from workbook[name])
    if no case-insensitive match exists.
    """
    actual = find_sheet_name(workbook, sheet_name)
    if actual is None:
        raise KeyError(f'Worksheet "{sheet_name}" not found')
    return workbook[actual]


def get_or_create_sheet(workbook: Workbook, sheet_name: str) -> Worksheet:
    """Return sheet if it exists (case-insensitive), create it if not."""
    actual = find_sheet_name(workbook, sheet_name)
    if actual is not None:
        return workbook[actual]
    return workbook.create_sheet(sheet_name)


def remove_sheet_if_exists(workbook: Workbook, sheet_name: str) -> None:
    actual = find_sheet_name(workbook, sheet_name)
    if actual is not None:
        workbook.remove(workbook[actual])


# ---------------------------------------------------------------------------
# Header / row helpers
# ---------------------------------------------------------------------------

def get_headers(sheet: Worksheet) -> dict[str, int]:
    """
    Return {column_name: column_index} from the first row.
    Column names are stripped but NOT lowercased — matching is done
    by callers that need case-insensitive behaviour.
    """
    headers: dict[str, int] = {}
    for cell in sheet[1]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column
    return headers


def get_headers_lower(sheet: Worksheet) -> dict[str, int]:
    """Same as get_headers but keys are lowercased for case-insensitive lookup."""
    return {k.lower(): v for k, v in get_headers(sheet).items()}


def find_column_index(headers: dict[str, int], col_name: str) -> int | None:
    """
    Case-insensitive column name lookup within a headers dict returned
    by get_headers() (FS section 9: column names are case-insensitive
    framework-wide). Returns the column index if a case-insensitive
    match exists, else None.
    """
    target = str(col_name).strip().lower()
    for name, idx in headers.items():
        if name.strip().lower() == target:
            return idx
    return None


def sheet_to_dataframe(sheet: Worksheet) -> pd.DataFrame:
    """
    Read an openpyxl sheet into a pandas DataFrame.
    First row is used as column headers.
    All values are kept as-is (no dtype coercion).
    """
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()

    headers = [str(c).strip() if c is not None else f"_col{i}"
               for i, c in enumerate(rows[0])]
    data = rows[1:]
    return pd.DataFrame(data, columns=headers)


# ---------------------------------------------------------------------------
# Value helpers
# ---------------------------------------------------------------------------

def is_empty(value) -> bool:
    """
    True if value is None, a whitespace-only string, or the text "nan"
    (case-insensitive). The "nan" case covers artifacts produced when
    pandas reads empty Excel formula cells (e.g. VLOOKUP results that
    evaluate to nothing) — per FS section 2.5.2, these are treated as
    empty business data, not as the literal text "nan".
    """
    if value is None:
        return True
    s = str(value).strip()
    if s == "":
        return True
    return s.lower() == "nan"


def clean_field(value) -> str:
    """
    Returns '' if is_empty(value) is True (None, blank, or the literal
    text "nan"), otherwise the stripped string form of value.

    Use this in place of str(value).strip() whenever the result will be
    passed to is_empty(): stringifying None first turns it into the
    literal text "None", which is_empty() does not recognize as empty,
    so the empty case silently falls through as if a real value were set.
    """
    return "" if is_empty(value) else str(value).strip()


def normalize_string(value) -> str:
    """Strip and lowercase. Returns '' for None."""
    if value is None:
        return ""
    return str(value).strip().lower()


def is_numeric_null(value) -> bool:
    """
    True if value represents a numeric null:
    empty/None, or a numeric zero in any decimal format
    (0, 0.0, 0,0, 00, 0.00, etc.).
    Used by the NULL validation type.
    """
    if is_empty(value):
        return True
    normalized = str(value).strip().replace(",", ".")
    try:
        parsed = float(normalized)
        if math.isnan(parsed) or math.isinf(parsed):
            return False
        return parsed == 0.0
    except ValueError:
        return False
