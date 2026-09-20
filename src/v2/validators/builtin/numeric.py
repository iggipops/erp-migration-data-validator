import math
from workbook.excel_utils import is_empty, cell_value, last_data_row


def validate(sheet, col_idx, col_name, rule, context):
    """
    NUMERIC validation per FS section 6.10 (V2.3).

    Requires the cell to be a native Excel numeric value (int or float).
    Text values fail outright — NUMERIC does not parse text.
    Checks: native-value-required, min_value, max_value.
    """
    _validate_numeric(sheet, col_idx, col_name, rule, context, integer_only=False)


def _validate_numeric(sheet, col_idx, col_name, rule, context, integer_only: bool):
    issues_writer = context["issues_writer"]
    config        = context["config"]

    nv        = config.numeric_validation
    min_value = nv.get("min_value")
    max_value = nv.get("max_value")

    for row in range(2, last_data_row(sheet) + 1):
        cell  = sheet.cell(row=row, column=col_idx)
        value = cell_value(cell)

        if is_empty(value):
            continue

        # Require native numeric — text fails outright
        parsed = _as_float(value)
        if parsed is None:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=(
                    f"Value is not a native Excel number — "
                    f"text values are not accepted by NUMERIC/INTEGER validation. "
                    f"Ensure the cell is formatted as a number in Excel."
                ),
            )
            continue

        # INTEGER: fractional part must be zero
        if integer_only and parsed != int(parsed):
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=(
                    f"Value {parsed} has a non-zero fractional part "
                    f"and cannot be imported into an integer field"
                ),
            )
            continue

        if min_value is not None and parsed < min_value:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=f"Value {parsed} is below min_value {min_value}",
            )
            continue

        if max_value is not None and parsed > max_value:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=value,
                comment=f"Value {parsed} is above max_value {max_value}",
            )


def _as_float(value):
    """
    Returns float if value is a native Excel numeric (int or float), else None.
    Excludes bool explicitly — bool is a subclass of int in Python, and
    openpyxl reads Excel TRUE/FALSE cells as native bool, which must NOT
    pass NUMERIC/INTEGER validation as 1.0/0.0.
    Guards against float('nan') and float('inf').
    """
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        f = float(value)
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    return None
