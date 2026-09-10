from collections import Counter
from workbook.excel_utils import is_empty, normalize_string


def validate(sheet, col_idx, col_name, rule, context):
    """
    Flags all occurrences of values that appear more than once.
    Empty cells are excluded (empty-exclusion rule, FS section 2.5.4).
    Comparison is case-insensitive (FS section 9).
    """
    issues_writer = context["issues_writer"]

    values = [
        normalize_string(sheet.cell(row=row, column=col_idx).value)
        for row in range(2, sheet.max_row + 1)
        if not is_empty(sheet.cell(row=row, column=col_idx).value)
    ]

    duplicates = {v for v, count in Counter(values).items() if count > 1}

    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=col_idx)
        if is_empty(cell.value):
            continue
        if normalize_string(cell.value) in duplicates:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell.value,
            )
