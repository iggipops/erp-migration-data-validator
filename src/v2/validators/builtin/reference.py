from workbook.excel_utils import is_empty, normalize_string, get_headers, get_sheet, find_column_index


def validate(sheet, col_idx, col_name, rule, context):
    """
    Flags values not present in ReferenceDataSheet.ReferenceDataColumn.
    Empty cells are excluded (FS section 2.5.4).
    Comparison is case-insensitive (FS section 9).
    """
    issues_writer   = context["issues_writer"]
    workbook        = context["workbook"]
    ref_sheet_name  = str(rule.get("ReferenceDataSheet", "")).strip()
    ref_column_name = str(rule.get("ReferenceDataColumn", "")).strip()

    ref_sheet   = get_sheet(workbook, ref_sheet_name)
    ref_headers = get_headers(ref_sheet)
    ref_col_idx = find_column_index(ref_headers, ref_column_name)

    reference_values = {
        normalize_string(ref_sheet.cell(row=r, column=ref_col_idx).value)
        for r in range(2, ref_sheet.max_row + 1)
        if not is_empty(ref_sheet.cell(row=r, column=ref_col_idx).value)
    }

    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=col_idx)
        if is_empty(cell.value):
            continue
        if normalize_string(cell.value) not in reference_values:
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell.value,
            )
