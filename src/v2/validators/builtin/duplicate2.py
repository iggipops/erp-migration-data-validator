from collections import Counter
from workbook.excel_utils import is_empty, normalize_string, get_headers, find_column_index, cell_value, last_data_row


def validate(sheet, col_idx, col_name, rule, context):
    """
    Flags all rows where the combination of ColumnName + ReferenceDataColumn
    appears more than once. Empty ColumnName cells are excluded entirely.
    Comparison is case-insensitive (FS section 9).
    """
    issues_writer = context["issues_writer"]
    ref_col_name  = str(rule.get("ReferenceDataColumn", "")).strip()
    headers       = get_headers(sheet)

    ref_col_idx = find_column_index(headers, ref_col_name)
    if ref_col_idx is None:
        raise ValueError(
            f"DUPLICATE2: ReferenceDataColumn '{ref_col_name}' "
            f"not found in sheet '{sheet.title}'"
        )

    pairs = {}
    for row in range(2, last_data_row(sheet) + 1):
        v1 = cell_value(sheet.cell(row=row, column=col_idx))
        if is_empty(v1):
            continue
        v2 = cell_value(sheet.cell(row=row, column=ref_col_idx))
        pairs[row] = (normalize_string(v1), normalize_string(v2))

    duplicates = {pair for pair, count in Counter(pairs.values()).items() if count > 1}

    for row, pair in pairs.items():
        if pair in duplicates:
            cell = sheet.cell(row=row, column=col_idx)
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell_value(cell),
            )
