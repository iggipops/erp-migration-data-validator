from datetime import datetime, date
from workbook.excel_utils import is_empty


def validate(sheet, col_idx, col_name, rule, context):
    """
    DATE validation per FS section 6.9 (V2.3).

    Requires the cell to be a native Excel date/datetime value.
    Text or plain numeric values fail outright — DATE does not parse text.
    Checks: native-value-required, null_dates, min_date, max_date.
    """
    issues_writer = context["issues_writer"]
    config        = context["config"]

    dv            = config.date_validation
    min_date_str  = dv.get("min_date")
    max_date_str  = dv.get("max_date")
    null_dates    = set(dv.get("null_dates", []) or [])
    treat_null_as = str(dv.get("treat_null_dates_as", "empty")).strip().lower()

    min_date = datetime.strptime(min_date_str, "%Y-%m-%d") if min_date_str else None
    max_date = datetime.strptime(max_date_str, "%Y-%m-%d") if max_date_str else None

    for row in range(2, sheet.max_row + 1):
        cell  = sheet.cell(row=row, column=col_idx)
        value = cell.value

        if is_empty(value):
            continue

        # Require native date/datetime — text and numbers fail outright
        parsed = _as_datetime(value)
        if parsed is None:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=(
                    f"Value is not a native Excel date — "
                    f"text and plain numeric values are not accepted by DATE validation. "
                    f"Ensure the cell is formatted as a date in Excel."
                ),
            )
            continue

        parsed_str = parsed.strftime("%Y-%m-%d")

        # Null date check
        if parsed_str in null_dates:
            if treat_null_as == "invalid":
                issues_writer.record(
                    cell=cell, rule=rule,
                    sheet_name=sheet.title, column_name=col_name,
                    row_number=row, cell_content=value,
                    comment=f"Date {parsed_str} is a configured null date",
                )
            continue  # "empty" or "valid" → pass

        # Range checks
        if min_date and parsed < min_date:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=f"Date {parsed_str} is before min_date {min_date_str}",
            )
            continue

        if max_date and parsed > max_date:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=f"Date {parsed_str} is after max_date {max_date_str}",
            )


def _as_datetime(value):
    """
    Returns a datetime if value is a native Excel date/datetime, else None.
    Handles both datetime.datetime and datetime.date objects from openpyxl.
    """
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    return None
