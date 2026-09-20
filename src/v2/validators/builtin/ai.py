"""
AI ValidationType dispatcher (FS section 6.8) — mechanically identical to
CUSTOM (section 6.7): CustomFunctionName is looked up in the Function
Registry and called with the same context shape and the same positional
row-order mapping back onto SheetName's data (see validators/builtin/custom.py
and FS Appendix C.3). The framework-level differences from CUSTOM are:

- Soft-skip when ai_enabled is false/omitted (same treatment as Active=No).
  The engine already skips before dispatch (FS section 3.4/6.8), but this
  module guards independently too since validate() can be called directly.

- The registered function may return, per row, one of:
    True / (True, "")   — passed
    False / (False, c)  — failed, with optional comment
    None                — the row belonged to an AI batch that failed;
                          no verdict is available (FS section 6.8)
  If every row was verdicted, this behaves exactly like CUSTOM (issues
  recorded for failing rows, no special handling). If some rows are
  unresolved and at least one row was verdicted, AIRulePartialError is
  raised after recording issues for every row that DID get a verdict —
  validation_engine.py catches this specifically to set the rule's status
  to 'partial' rather than 'failed' (FS section 6.8/8.8.2). If every row
  is unresolved (every batch failed), a plain exception is raised instead,
  so the rule is reported as 'failed' with no issues recorded.
"""
from utils.ai_client import AIRulePartialError
from utils.logger import get_logger
from workbook.excel_utils import sheet_to_dataframe, cell_value


def validate(sheet, col_idx, col_name, rule, context):
    logger = get_logger()
    config = context["config"]

    if not getattr(config, "ai_enabled", False):
        logger.info(
            f"AI rule [{rule.get('RuleCode')}] skipped (AI disabled)"
        )
        return

    issues_writer = context["issues_writer"]
    fn_registry   = context["fn_registry"]
    workbook      = context["workbook"]

    fn_name = str(rule.get("CustomFunctionName", "")).strip()
    func    = fn_registry.get(fn_name)
    if func is None:
        raise ValueError(f"Function '{fn_name}' not found in function registry")

    df = sheet_to_dataframe(sheet)

    fn_name_lower    = fn_name.lower()
    technical_config = config.custom_functions.get(fn_name_lower, {})
    named_columns    = {col: df[col].tolist() for col in df.columns}

    fn_context = {
        "worksheet_data":   df,
        "named_columns":    named_columns,
        "technical_config": technical_config,
        "workbook":         workbook,
        "config":           config,
        "row_ids":          list(range(2, len(df) + 2)),
    }

    results = func(fn_context)

    total_rows      = len(results)
    unresolved_rows = 0

    for i, result in enumerate(results):
        row = i + 2

        if result is None:
            unresolved_rows += 1
            continue

        if isinstance(result, (list, tuple)) and len(result) == 2:
            passed, comment = result
        else:
            passed, comment = bool(result), ""

        if not passed:
            cell = sheet.cell(row=row, column=col_idx)
            issues_writer.record(
                cell=cell, rule=rule,
                sheet_name=sheet.title, column_name=col_name,
                row_number=row, cell_content=cell_value(cell),
                comment=comment,
            )

    if unresolved_rows == 0:
        return

    if unresolved_rows == total_rows:
        raise RuntimeError(
            f"AI rule [{rule.get('RuleCode')}]: all {total_rows} row(s) "
            f"unresolved — every AI batch failed"
        )

    partial_exc = AIRulePartialError(
        f"AI rule [{rule.get('RuleCode')}]: {unresolved_rows} of {total_rows} "
        f"row(s) unresolved — one or more AI batches failed"
    )
    partial_exc.rows_unchecked = unresolved_rows
    raise partial_exc
