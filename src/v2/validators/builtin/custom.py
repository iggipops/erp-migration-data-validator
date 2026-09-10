from workbook.excel_utils import sheet_to_dataframe


def validate(sheet, col_idx, col_name, rule, context):
    """
    Delegates to a registered custom function (FS section 2.6.2).
    CustomFunctionName must reference a function in the Function Registry.
    Custom functions are exempt from the empty-exclusion principle and
    handle their own empty-value logic internally (FS section 2.5.4).
    """
    issues_writer = context["issues_writer"]
    fn_registry   = context["fn_registry"]
    config        = context["config"]
    workbook      = context["workbook"]

    fn_name = str(rule.get("CustomFunctionName", "")).strip()
    func    = fn_registry.get(fn_name)
    if func is None:
        raise ValueError(f"Function '{fn_name}' not found in function registry")

    df = sheet_to_dataframe(sheet)

    fn_name_lower    = fn_name.lower()
    technical_config = config.custom_functions.get(fn_name_lower, {})

    named_columns = {col: df[col].tolist() for col in df.columns}

    fn_context = {
        "worksheet_data":   df,
        "named_columns":    named_columns,
        "technical_config": technical_config,
        "workbook":         workbook,
        "config":           config,
        "row_ids":          list(range(2, len(df) + 2)),
    }

    results = func(fn_context)

    for i, result in enumerate(results):
        row = i + 2
        if isinstance(result, (list, tuple)) and len(result) == 2:
            passed, comment = result
        else:
            passed  = bool(result)
            comment = ""

        if not passed:
            cell = sheet.cell(row=row, column=col_idx)
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell.value,
                comment=comment,
            )
