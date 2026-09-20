import re
from workbook.excel_utils import is_empty, cell_value, last_data_row


def validate(sheet, col_idx, col_name, rule, context):
    """
    Flags values with leading/trailing whitespace, consecutive internal
    spaces, or characters listed in format_special_chars config.
    Empty cells are excluded (FS section 2.5.4).
    """
    issues_writer  = context["issues_writer"]
    config         = context["config"]
    special_chars  = config.format_special_chars or ""

    for row in range(2, last_data_row(sheet) + 1):
        cell = sheet.cell(row=row, column=col_idx)
        if is_empty(cell_value(cell)):
            continue
        value   = str(cell_value(cell))
        invalid = (
            value != value.strip()
            or bool(re.search(r"\s{2,}", value))
            or any(ch in value for ch in special_chars)
        )
        if invalid:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell_value(cell),
            )
