"""
ai_itemcategoryfit — AI Item/Category Fit Validation (FS Appendix C).

Judges whether an item's name is semantically relevant to the category
it's assigned to, using an AI model. Reference example for AI-flavored
validation functions (ValidationType = AI).

Logic chain (FS section C.4), for each row on SheetName (source_sheet,
which must equal SheetName exactly — FS section C.3):
1. If source_key_column's value is empty, the row is skipped (passes,
   no issue).
2. The key is looked up in reference_sheet via reference_key_column. If
   not found, the row is skipped (passes, no issue) — this function's
   only concern is relevance, not existence.
3. The text in source_text_column (primary row) and reference_text_column
   (matched reference row) are read.
4. Matched pairs are batched (batch_size rows per call, default 20) and
   sent to the AI service via the shared AI client module.
5. Alongside each batch, the full reference_text_column list across every
   row of reference_sheet is also sent (never the key/ID column) — once
   per batch, not once per row — so the AI can recommend a genuinely
   better-fitting category.
6. Rows verdicted F fail; the AI's comment text is written verbatim to
   the Comment column, no verification against the real category list.
7. If the AI call fails for a batch, the rows in that batch get no
   verdict (None) — see validators/builtin/ai.py for how that becomes
   the rule's 'partial'/'failed' status (FS section 6.8/8.8.2).

Returns a list of results, one per row of context["worksheet_data"]:
True (pass), (False, comment) (fail), or None (unresolved — batch failed).
"""
import math

import pandas as pd

from workbook.excel_utils import sheet_exists
import utils.ai_client as ai_client


REQUIRED_PARAMS = [
    "prompt",
    "source_sheet",
    "source_key_column",
    "source_text_column",
    "reference_sheet",
    "reference_key_column",
    "reference_text_column",
]

DEFAULT_BATCH_SIZE = 20


def check_metadata(rule: dict, workbook, config) -> list[str]:
    """
    Optional metadata-check hook (FS section 2.6.2), called during
    metadata validation (step 8.6) for any active AI ValidationRules row
    referencing ai_itemcategoryfit.
    """
    errors: list[str] = []

    fn_name = "ai_itemcategoryfit"
    cfg = config.custom_functions.get(fn_name, {})

    rule_sheet_name        = str(rule.get("SheetName", "")).strip()
    source_sheet           = str(cfg.get("source_sheet", "")).strip()
    source_key_column      = str(cfg.get("source_key_column", "")).strip()
    source_text_column     = str(cfg.get("source_text_column", "")).strip()
    reference_sheet        = str(cfg.get("reference_sheet", "")).strip()
    reference_key_column   = str(cfg.get("reference_key_column", "")).strip()
    reference_text_column  = str(cfg.get("reference_text_column", "")).strip()
    batch_size             = cfg.get("batch_size", DEFAULT_BATCH_SIZE)

    # --- Hard constraint (FS section C.3): source_sheet must equal SheetName ---
    if source_sheet and rule_sheet_name and \
            source_sheet.strip().lower() != rule_sheet_name.strip().lower():
        errors.append(
            f"ai_itemcategoryfit: source_sheet '{source_sheet}' must match "
            f"the ValidationRules row's SheetName '{rule_sheet_name}' exactly "
            f"(FS Appendix C.3 — required for the positional row mapping "
            f"back onto SheetName's data)"
        )

    if rule_sheet_name and sheet_exists(workbook, rule_sheet_name):
        source_cols = _columns_of(workbook, rule_sheet_name)
        if source_key_column and _resolve_column_name(source_cols, source_key_column) is None:
            errors.append(
                f"ai_itemcategoryfit: source_key_column '{source_key_column}' "
                f"not found in sheet '{rule_sheet_name}'"
            )
        if source_text_column and _resolve_column_name(source_cols, source_text_column) is None:
            errors.append(
                f"ai_itemcategoryfit: source_text_column '{source_text_column}' "
                f"not found in sheet '{rule_sheet_name}'"
            )

    if reference_sheet and not sheet_exists(workbook, reference_sheet):
        errors.append(
            f"ai_itemcategoryfit: reference_sheet '{reference_sheet}' not "
            f"found in workbook"
        )
    elif reference_sheet:
        ref_cols = _columns_of(workbook, reference_sheet)
        if reference_key_column and _resolve_column_name(ref_cols, reference_key_column) is None:
            errors.append(
                f"ai_itemcategoryfit: reference_key_column "
                f"'{reference_key_column}' not found in sheet '{reference_sheet}'"
            )
        if reference_text_column and _resolve_column_name(ref_cols, reference_text_column) is None:
            errors.append(
                f"ai_itemcategoryfit: reference_text_column "
                f"'{reference_text_column}' not found in sheet '{reference_sheet}'"
            )

    if batch_size is not None:
        try:
            if int(batch_size) <= 0:
                errors.append("ai_itemcategoryfit: batch_size must be a positive integer")
        except (ValueError, TypeError):
            errors.append(
                f"ai_itemcategoryfit: batch_size '{batch_size}' is not a valid integer"
            )

    return errors


def ai_itemcategoryfit(context: dict) -> list:
    from utils.logger import get_logger
    logger   = get_logger()
    cfg      = context.get("technical_config", {})
    workbook = context.get("workbook")
    config   = context.get("config")
    df       = context.get("worksheet_data")  # SheetName's data == source_sheet (C.3)

    if df is None or workbook is None or config is None:
        raise ValueError(
            "ai_itemcategoryfit: worksheet_data, workbook and config are required"
        )

    prompt                 = cfg.get("prompt", "")
    source_key_column      = cfg.get("source_key_column", "")
    source_text_column     = cfg.get("source_text_column", "")
    reference_sheet        = cfg.get("reference_sheet", "")
    reference_key_column   = cfg.get("reference_key_column", "")
    reference_text_column  = cfg.get("reference_text_column", "")
    batch_size             = cfg.get("batch_size", DEFAULT_BATCH_SIZE)

    _require(prompt, "prompt")
    _require(source_key_column, "source_key_column")
    _require(source_text_column, "source_text_column")
    _require(reference_sheet, "reference_sheet")
    _require(reference_key_column, "reference_key_column")
    _require(reference_text_column, "reference_text_column")

    try:
        batch_size = int(batch_size) if batch_size else DEFAULT_BATCH_SIZE
    except (ValueError, TypeError):
        raise ValueError(f"ai_itemcategoryfit: batch_size '{batch_size}' is not a valid integer")
    if batch_size <= 0:
        raise ValueError("ai_itemcategoryfit: batch_size must be a positive integer")

    resolved_source_key_col = _resolve_column_name(df.columns, source_key_column)
    if resolved_source_key_col is None:
        raise ValueError(
            f"ai_itemcategoryfit: source_key_column '{source_key_column}' "
            f"does not exist on the source worksheet"
        )
    resolved_source_text_col = _resolve_column_name(df.columns, source_text_column)
    if resolved_source_text_col is None:
        raise ValueError(
            f"ai_itemcategoryfit: source_text_column '{source_text_column}' "
            f"does not exist on the source worksheet"
        )

    ref_df = _sheet_to_df(workbook, reference_sheet)

    resolved_ref_key_col = _resolve_column_name(ref_df.columns, reference_key_column)
    if resolved_ref_key_col is None:
        raise ValueError(
            f"ai_itemcategoryfit: reference_key_column '{reference_key_column}' "
            f"does not exist on sheet '{reference_sheet}'"
        )
    resolved_ref_text_col = _resolve_column_name(ref_df.columns, reference_text_column)
    if resolved_ref_text_col is None:
        raise ValueError(
            f"ai_itemcategoryfit: reference_text_column '{reference_text_column}' "
            f"does not exist on sheet '{reference_sheet}'"
        )

    reference_lookup = _build_lookup(ref_df, resolved_ref_key_col)

    # Step 5: full category-description list across every reference row —
    # Category Description only, never the key/ID column — sent once per batch.
    full_reference_text_list = [
        _str(v) for v in ref_df[resolved_ref_text_col].tolist() if not _is_empty(v)
    ]

    # --- Steps 1-2: determine which rows actually need an AI verdict ---
    row_positions: list[int]  = []
    row_payloads:  list[dict] = []

    for pos, (_, row_data) in enumerate(df.iterrows()):
        key_value = _str(row_data.get(resolved_source_key_col)).strip()

        if _is_empty(key_value):
            continue  # step 1: empty key -> pass, no issue

        ref_row = reference_lookup.get(_normalize_key(key_value))
        if ref_row is None:
            continue  # step 2: key not found in reference -> pass, no issue

        row_positions.append(pos)
        row_payloads.append({
            "item_text":              _str(row_data.get(resolved_source_text_col)),
            "assigned_category_text": _str(ref_row.get(resolved_ref_text_col)),
        })

    # Default: rows never sent to the AI (steps 1-2) pass without an issue.
    results: list = [True] * len(df)

    # --- Steps 3-5: send matched rows to the AI service in batches ---
    for batch_start in range(0, len(row_payloads), batch_size):
        batch_positions = row_positions[batch_start:batch_start + batch_size]
        batch_payloads  = row_payloads[batch_start:batch_start + batch_size]

        try:
            verdicts = ai_client.send_batch(
                config, prompt, batch_payloads,
                reference_context=full_reference_text_list,
            )
        except ai_client.AIClientError as exc:
            logger.warning(f"ai_itemcategoryfit: batch failed — {exc}")
            for pos in batch_positions:
                results[pos] = None  # step 7: unresolved — batch failed
            continue

        # Step 6: F rows fail with the AI's verbatim comment; P rows pass.
        for pos, (passed, comment) in zip(batch_positions, verdicts):
            results[pos] = True if passed else (False, comment)

    return results


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sheet_to_df(workbook, sheet_name: str) -> pd.DataFrame:
    if not sheet_exists(workbook, sheet_name):
        raise ValueError(f"ai_itemcategoryfit: sheet '{sheet_name}' not found in workbook")
    from workbook.excel_utils import sheet_to_dataframe, get_sheet
    df = sheet_to_dataframe(get_sheet(workbook, sheet_name))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _columns_of(workbook, sheet_name: str) -> set[str]:
    if not sheet_exists(workbook, sheet_name):
        return set()
    from workbook.excel_utils import get_headers, get_sheet
    return set(get_headers(get_sheet(workbook, sheet_name)).keys())


def _build_lookup(df: pd.DataFrame, key_column: str) -> dict[str, dict]:
    """Case-insensitive lookup: normalized key value -> row dict (FS section 9)."""
    lookup = {}
    for _, row in df.iterrows():
        key = _normalize_key(row.get(key_column, ""))
        if key:
            lookup[key] = row.to_dict()
    return lookup


def _normalize_key(value) -> str:
    return _str(value).strip().lower()


def _str(value) -> str:
    if value is None:
        return ""
    return str(value)


def _is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    s = str(value).strip()
    return s == "" or s.lower() == "nan"


def _resolve_column_name(available_columns, wanted_name: str):
    """Case-insensitive column name resolution (FS section 9)."""
    target = str(wanted_name).strip().lower()
    for actual in available_columns:
        if str(actual).strip().lower() == target:
            return actual
    return None


def _require(value, param_name: str) -> None:
    if not value or str(value).strip() == "":
        raise ValueError(f"ai_itemcategoryfit: '{param_name}' is required in config")
