from workbook.excel_utils import is_numeric_null, cell_value, last_data_row


def validate(sheet, col_idx, col_name, rule, context):
    """
    NULL flags empty values AND numeric zero (sole exception to
    the empty-exclusion principle — FS section 2.5.4).
    """
    issues_writer = context["issues_writer"]
    for row in range(2, last_data_row(sheet) + 1):
        cell = sheet.cell(row=row, column=col_idx)
        if is_numeric_null(cell_value(cell)):
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell_value(cell),
            )
