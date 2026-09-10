"""
dimpolicyvalidation — Stock Balance Dimension Validation (FS Appendix B).

Validates stock balance records against configurable dimension policies —
the required, forbidden, and quantity-restricted attributes each stock
record must carry before ERP import.

The dimension policy lookup key is composed from one or more *lookup
blocks* (`policy_key_lookups`, FS B.3/B.4). Each block names a Stock
Balance column whose value drives a lookup, the reference sheet and key
column to match it against, and one or more attribute columns to pull
from the matched reference row. All attribute values from all blocks —
in block order, then within-block column order — are concatenated with
`dim_policy_key_delimiter` to form the composed policy key. This covers
an item-only policy (one block, one or more columns — the AX/D365
shape), an item+warehouse policy (two blocks — the SAP / Infor LN
shape), or any other combination of business entities.

Logic chain (FS B.7), for each row on the Stock Balance sheet (the sheet
named in the ValidationRules row's SheetName):

1. Build the composed key. For each block, in list order:
   a. Read source_column on the current row. If empty, the row is
      skipped entirely — it passes without producing an issue (use a
      separate EMPTY rule to flag empty values in their own right).
   b. Look the value up in reference_sheet via reference_key_column. If
      no matching row is found, the row fails (terminal — steps 4-5 are
      not evaluated).
   c. Otherwise append each policy_key_columns value from the matched
      row, in list order, to the growing key components.
2. Compose the final key by joining all components with
   dim_policy_key_delimiter.
3. Look the composed key up in dim_policy_sheet (dim_policy_key_column).
   If not found, the row fails (terminal — steps 4-5 are not evaluated).
4. For each stock_to_dim_policy_columns_mapping pair, read the Yes/No
   value on the matched policy row: Yes -> the mapped Stock Balance cell
   must be not empty; No -> it must be empty. Every mismatch adds a
   message; the check continues across all mapped columns.
5. Read the Quantity Restriction value on the matched policy row. Empty
   means no restriction and this step is skipped; otherwise the
   stock_quantity_column cell must equal it, and a mismatch adds a
   message.
6. Report. A step 1b / step 3 failure is reported as that single
   message. Otherwise all step 4 and step 5 messages are reported
   together, semicolon-separated.

Config validation (FS B.5) — schema (A.1-A.7), cross-reference
(B.1-B.8), and data integrity (C.1-C.3) — is implemented in
check_metadata() below. Following the framework's standard
collect-and-report-together behavior (FS 8.1/8.6), every error found is
returned in one list, the metadata validator collects them alongside
every other rule's errors, and processing terminates before any
enrichment or validation runs. The same checks also run at the top of
dimpolicyvalidation() itself, so a direct call that bypassed metadata
validation still reports every config error together rather than
misbehaving row by row.

Returns list of (bool, str): (passed, comment_message).
"""
import math

import pandas as pd

# Imported as a module, not as individual names: the function registry scans
# every module-level function in this file and would otherwise register the
# imported helpers as if they were custom functions (FS section 8.4).
import workbook.excel_utils as excel_utils


FUNCTION_NAME = "dimpolicyvalidation"

# The Quantity Restriction column is named literally in FS B.6 — it is a
# fixed part of the Dim Policy sheet structure, not a configurable
# parameter. Matched case-insensitively, per FS section 9.
QUANTITY_RESTRICTION_COLUMN = "Quantity Restriction"

BLOCK_REQUIRED_KEYS = (
    "source_column",
    "reference_sheet",
    "reference_key_column",
    "policy_key_columns",
)

REQUIRED_PARAMS = [
    "policy_key_lookups",
    "dim_policy_sheet",
    "dim_policy_key_column",
    "stock_to_dim_policy_columns_mapping",
    "stock_quantity_column",
]


# ---------------------------------------------------------------------------
# Metadata-check hook (FS section 2.6.2) — Config Validation, FS B.5
# ---------------------------------------------------------------------------

def check_metadata(rule: dict, workbook, config) -> list[str]:
    """
    Optional metadata-check hook (FS section 2.6.2). Called during metadata
    validation (step 8.6) for every active CUSTOM ValidationRules row that
    references dimpolicyvalidation.

    Runs the full FS B.5 Config Validation — schema (B.5.1/A.1-A.7),
    cross-reference (B.5.2/B.1-B.8), and data integrity (B.5.3/C.1-C.3) —
    and returns every error found as one list, empty if the configuration
    is sound. The caller collects these alongside all other metadata
    errors and terminates before any row is processed (FS 8.6).
    """
    cfg = _function_config(config)

    stock_sheet_name = str(rule.get("SheetName", "")).strip()
    stock_columns    = _columns_of(workbook, stock_sheet_name)

    return _collect_config_errors(
        cfg, workbook,
        stock_columns=stock_columns,
        stock_sheet_label=stock_sheet_name or "the Stock Balance sheet",
    )


# ---------------------------------------------------------------------------
# Validation function
# ---------------------------------------------------------------------------

def dimpolicyvalidation(context: dict) -> list[tuple[bool, str]]:

    cfg      = context.get("technical_config") or {}
    workbook = context.get("workbook")
    df       = context.get("worksheet_data")  # Stock Balance sheet

    if df is None or workbook is None:
        raise ValueError(
            f"{FUNCTION_NAME}: worksheet_data and workbook are required"
        )

    # --- FS B.5 Config Validation — every error reported together, then stop.
    # Normally already run at step 8.6 via check_metadata(); repeated here so
    # a direct call is held to the same contract.
    errors = _collect_config_errors(
        cfg, workbook,
        stock_columns=list(df.columns),
        stock_sheet_label="the Stock Balance sheet",
    )
    if errors:
        raise ValueError(
            f"{FUNCTION_NAME}: configuration validation failed:\n"
            + "\n".join(f"  - {e}" for e in errors)
        )

    delimiter = _delimiter(cfg)

    # --- Resolve every configured name to its actual on-sheet casing once,
    # up front (FS section 9: sheet and column names are case-insensitive),
    # so the row loop never has to handle casing again.
    blocks = []
    for raw_block in cfg["policy_key_lookups"]:
        ref_sheet_name = excel_utils.find_sheet_name(workbook, raw_block["reference_sheet"])
        ref_df         = _sheet_to_df(workbook, ref_sheet_name)
        ref_key_col    = _resolve_column_name(ref_df.columns, raw_block["reference_key_column"])
        blocks.append({
            "source_column":      _resolve_column_name(df.columns, raw_block["source_column"]),
            "reference_sheet":    ref_sheet_name,
            "lookup":             _build_lookup(ref_df, ref_key_col),
            "policy_key_columns": [
                _resolve_column_name(ref_df.columns, c)
                for c in raw_block["policy_key_columns"]
            ],
        })

    dim_policy_df  = _sheet_to_df(workbook, cfg["dim_policy_sheet"])
    dim_policy_key = _resolve_column_name(dim_policy_df.columns, cfg["dim_policy_key_column"])
    policy_lookup  = _build_lookup(dim_policy_df, dim_policy_key)

    # Mapping order follows config order, so accumulated step 4 messages
    # come out in the order the consultant declared the mapping.
    mapping = {
        _resolve_column_name(df.columns, stock_col):
            _resolve_column_name(dim_policy_df.columns, policy_col)
        for stock_col, policy_col in cfg["stock_to_dim_policy_columns_mapping"].items()
    }

    stock_qty_col = _resolve_column_name(df.columns, cfg["stock_quantity_column"])

    # Optional per FS B.6 — absent means no row ever carries a restriction.
    qty_restriction_col = _resolve_column_name(
        dim_policy_df.columns, QUANTITY_RESTRICTION_COLUMN
    )

    # --- Process each stock row ---
    results: list[tuple[bool, str]] = []

    for _, stock_row in df.iterrows():
        results.append(_validate_row(
            stock_row=stock_row,
            blocks=blocks,
            delimiter=delimiter,
            policy_lookup=policy_lookup,
            mapping=mapping,
            stock_qty_col=stock_qty_col,
            qty_restriction_col=qty_restriction_col,
        ))

    return results


# ---------------------------------------------------------------------------
# Row-level validation (FS B.7) — messages per FS B.8
# ---------------------------------------------------------------------------

def _validate_row(
    stock_row,
    blocks,
    delimiter,
    policy_lookup,
    mapping,
    stock_qty_col,
    qty_restriction_col,
) -> tuple[bool, str]:

    # --- Step 1: build the composed key, block by block ---
    key_parts: list[str] = []

    for block in blocks:
        source_value = stock_row.get(block["source_column"])

        # Step 1a — empty source value: the row is skipped entirely and
        # passes without an issue. Flagging emptiness is an EMPTY rule's job.
        if _is_empty(source_value):
            return True, ""

        # Step 1b — no matching reference row: terminal failure for this row,
        # steps 4-5 are not evaluated.
        ref_row = block["lookup"].get(_normalize_key(source_value))
        if ref_row is None:
            return False, (
                f"{block['source_column']} value {_display(source_value)} "
                f"not found in {block['reference_sheet']} — "
                f"dimension policy cannot be determined"
            )

        # Step 1c — pull this block's key components, in list order.
        for col in block["policy_key_columns"]:
            key_parts.append(_display(ref_row.get(col)))

    # --- Step 2: compose the final key ---
    policy_key = delimiter.join(key_parts)

    # --- Step 3: look the composed key up in the Dim Policy sheet.
    # Terminal failure — steps 4-5 are not evaluated. ---
    policy_row = policy_lookup.get(_normalize_key(policy_key))
    if policy_row is None:
        return False, f"Dimension policy key {policy_key} not found in Dim Policy sheet"

    # --- Steps 4-5 accumulate: every mismatch is reported, not just the first ---
    failures: list[str] = []

    # --- Step 4: mapped dimension columns ---
    for stock_col, policy_col in mapping.items():
        policy_flag = _normalize_key(policy_row.get(policy_col))
        stock_value = stock_row.get(stock_col)

        if policy_flag == "yes":
            if _is_empty(stock_value):
                failures.append(
                    f"Dimension column {stock_col} is required by policy "
                    f"{policy_key} but is empty"
                )
        elif policy_flag == "no":
            if not _is_empty(stock_value):
                failures.append(
                    f"Dimension column {stock_col} must be absent for policy "
                    f"{policy_key} but value {_display(stock_value)} is present"
                )
        else:
            # Unreachable through the normal pipeline: FS B.5.3/C.2 makes a
            # non-Yes/No policy cell a config-validation hard stop. Kept so a
            # corrupt policy table can never silently pass if that check was
            # bypassed. (No FS B.8 message covers this case.)
            failures.append(
                f"Dimension column {stock_col} has unparseable policy value "
                f"{_display(policy_row.get(policy_col)) or '(empty)'} for policy "
                f"{policy_key} — must be Yes or No"
            )

    # --- Step 5: quantity restriction ---
    if qty_restriction_col is not None:
        required_qty = policy_row.get(qty_restriction_col)

        # Empty restriction: no restriction applies, step skipped (FS B.7).
        if not _is_empty(required_qty):
            stock_qty = stock_row.get(stock_qty_col)

            # An empty stock quantity is out of scope for the quantity
            # restriction check — leave it to a dedicated EMPTY/NULL rule
            # on that column rather than reporting it as a mismatch here.
            if not _is_empty(stock_qty) and not _values_equal(required_qty, stock_qty):
                failures.append(
                    f"Quantity restriction for policy {policy_key} requires "
                    f"{_display(required_qty)} but found {_display(stock_qty)} "
                    f"in column {stock_qty_col}"
                )

    # --- Step 6: report accumulated step 4-5 messages together ---
    if failures:
        return False, "; ".join(failures)

    return True, ""


# ---------------------------------------------------------------------------
# Config Validation (FS B.5)
# ---------------------------------------------------------------------------

def _collect_config_errors(
    cfg: dict, workbook, stock_columns, stock_sheet_label: str
) -> list[str]:
    """
    Run all three FS B.5 layers and return every error found, in order:
    schema (B.5.1/A.1-A.7), cross-reference (B.5.2/B.1-B.8), and data
    integrity (B.5.3/C.1-C.3).

    Each layer's checks are guarded so that a failure in an earlier layer
    only suppresses the specific later checks that cannot be evaluated
    without it — the aim is to report as much as possible in one pass
    rather than to stop at the first problem.

    `stock_columns` is None when the Stock Balance sheet itself is not
    available, in which case stock-side column checks are skipped.
    """
    errors: list[str] = []

    blocks = _check_schema(cfg, errors)
    _check_cross_references(cfg, workbook, stock_columns, stock_sheet_label, blocks, errors)
    _check_data_integrity(cfg, workbook, errors)

    return errors


def _check_schema(cfg: dict, errors: list[str]) -> list[dict]:
    """
    FS B.5.1 — structural checks A.1-A.7. Returns the subset of
    policy_key_lookups blocks that are structurally sound enough for the
    cross-reference layer to evaluate.
    """
    usable_blocks: list[dict] = []

    # --- A.1: policy_key_lookups present and a non-empty list ---
    raw_blocks = cfg.get("policy_key_lookups")
    if not isinstance(raw_blocks, list) or not raw_blocks:
        errors.append(
            f"{FUNCTION_NAME}: policy_key_lookups is required and must be a "
            f"non-empty list of lookup blocks (B.5.1/A.1)"
        )
        raw_blocks = []

    for position, block in enumerate(raw_blocks, start=1):

        if not isinstance(block, dict):
            errors.append(
                f"{FUNCTION_NAME}: policy_key_lookups block {position} must be a "
                f"mapping with the keys {', '.join(BLOCK_REQUIRED_KEYS)} (B.5.1/A.2)"
            )
            continue

        # --- A.2: all four required keys present.
        # For policy_key_columns this is a presence check only — whether the
        # list it holds is actually usable is A.3's concern, so an empty list
        # is reported there rather than as a missing key here. ---
        missing = []
        for key in BLOCK_REQUIRED_KEYS:
            value = block.get(key)
            if key == "policy_key_columns":
                if value is None:
                    missing.append(key)
            elif _is_blank(value):
                missing.append(key)

        if missing:
            errors.append(
                f"{FUNCTION_NAME}: policy_key_lookups block {position} is missing "
                f"required key(s): {', '.join(missing)} (B.5.1/A.2)"
            )

        # --- A.3: policy_key_columns is a non-empty list ---
        key_columns_usable = False
        if "policy_key_columns" not in missing:
            key_columns = block.get("policy_key_columns")
            if (isinstance(key_columns, (list, tuple)) and key_columns
                    and not any(_is_blank(c) for c in key_columns)):
                key_columns_usable = True
            else:
                errors.append(
                    f"{FUNCTION_NAME}: policy_key_lookups block {position}: "
                    f"policy_key_columns must be a non-empty list of column "
                    f"names (B.5.1/A.3)"
                )

        if not missing and key_columns_usable:
            usable_blocks.append(block)

    # --- A.4: dim_policy_sheet / dim_policy_key_column present, non-empty ---
    for param in ("dim_policy_sheet", "dim_policy_key_column"):
        if _is_blank(cfg.get(param)):
            errors.append(
                f"{FUNCTION_NAME}: {param} is required and must be a non-empty "
                f"string (B.5.1/A.4)"
            )

    # --- A.5: stock_to_dim_policy_columns_mapping present, non-empty mapping ---
    mapping = cfg.get("stock_to_dim_policy_columns_mapping")
    if not isinstance(mapping, dict) or not mapping:
        errors.append(
            f"{FUNCTION_NAME}: stock_to_dim_policy_columns_mapping is required "
            f"and must be a non-empty mapping of Stock Balance column names to "
            f"Dim Policy column names (B.5.1/A.5)"
        )
    else:
        for stock_col, policy_col in mapping.items():
            if _is_blank(stock_col) or _is_blank(policy_col):
                errors.append(
                    f"{FUNCTION_NAME}: stock_to_dim_policy_columns_mapping "
                    f"contains an entry with an empty column name "
                    f"({stock_col!r}: {policy_col!r}) (B.5.1/A.5)"
                )

    # --- A.6: stock_quantity_column present, non-empty ---
    if _is_blank(cfg.get("stock_quantity_column")):
        errors.append(
            f"{FUNCTION_NAME}: stock_quantity_column is required and must be "
            f"non-empty (B.5.1/A.6)"
        )

    # --- A.7: dim_policy_key_delimiter optional; defaults to empty string.
    # Only its type is checked — a list or mapping cannot be a delimiter. ---
    delimiter = cfg.get("dim_policy_key_delimiter")
    if delimiter is not None and isinstance(delimiter, (list, tuple, dict, set)):
        errors.append(
            f"{FUNCTION_NAME}: dim_policy_key_delimiter must be a string "
            f"(optional; defaults to an empty string) (B.5.1/A.7)"
        )

    return usable_blocks


def _check_cross_references(
    cfg: dict, workbook, stock_columns, stock_sheet_label: str,
    blocks: list[dict], errors: list[str],
) -> None:
    """FS B.5.2 — cross-reference checks B.1-B.8, against the actual workbook."""

    for position, block in enumerate(blocks, start=1):

        # --- B.1: source_column exists on the Stock Balance sheet ---
        if stock_columns is not None and \
                _resolve_column_name(stock_columns, block["source_column"]) is None:
            errors.append(
                f"{FUNCTION_NAME}: policy_key_lookups block {position}: "
                f"source_column '{block['source_column']}' does not exist on "
                f"sheet '{stock_sheet_label}' (B.5.2/B.1)"
            )

        # --- B.2: reference_sheet exists ---
        ref_sheet = str(block["reference_sheet"]).strip()
        if not excel_utils.sheet_exists(workbook, ref_sheet):
            errors.append(
                f"{FUNCTION_NAME}: policy_key_lookups block {position}: "
                f"reference_sheet '{ref_sheet}' not found in workbook (B.5.2/B.2)"
            )
            continue

        ref_columns = _columns_of(workbook, ref_sheet)

        # --- B.3: reference_key_column exists on reference_sheet ---
        if _resolve_column_name(ref_columns, block["reference_key_column"]) is None:
            errors.append(
                f"{FUNCTION_NAME}: policy_key_lookups block {position}: "
                f"reference_key_column '{block['reference_key_column']}' does not "
                f"exist on sheet '{ref_sheet}' (B.5.2/B.3)"
            )

        # --- B.4: every policy_key_columns entry exists on reference_sheet ---
        for col in block["policy_key_columns"]:
            if _resolve_column_name(ref_columns, col) is None:
                errors.append(
                    f"{FUNCTION_NAME}: policy_key_lookups block {position}: "
                    f"policy_key_columns references column '{col}' which does not "
                    f"exist on sheet '{ref_sheet}' (B.5.2/B.4)"
                )

    # --- B.5: dim_policy_sheet exists ---
    dim_policy_sheet  = str(cfg.get("dim_policy_sheet", "")).strip()
    dim_policy_exists = bool(dim_policy_sheet) and excel_utils.sheet_exists(workbook, dim_policy_sheet)
    if dim_policy_sheet and not dim_policy_exists:
        errors.append(
            f"{FUNCTION_NAME}: dim_policy_sheet '{dim_policy_sheet}' not found "
            f"in workbook (B.5.2/B.5)"
        )

    dim_policy_columns = _columns_of(workbook, dim_policy_sheet) if dim_policy_exists else None

    # --- B.6: dim_policy_key_column exists on dim_policy_sheet ---
    dim_policy_key = cfg.get("dim_policy_key_column")
    if dim_policy_columns is not None and not _is_blank(dim_policy_key) and \
            _resolve_column_name(dim_policy_columns, dim_policy_key) is None:
        errors.append(
            f"{FUNCTION_NAME}: dim_policy_key_column '{dim_policy_key}' does not "
            f"exist on sheet '{dim_policy_sheet}' (B.5.2/B.6)"
        )

    # --- B.7: every mapping key exists on Stock Balance, every mapping value
    # exists on dim_policy_sheet ---
    mapping = cfg.get("stock_to_dim_policy_columns_mapping")
    if isinstance(mapping, dict):
        for stock_col, policy_col in mapping.items():
            if _is_blank(stock_col) or _is_blank(policy_col):
                continue  # already reported by A.5
            if stock_columns is not None and \
                    _resolve_column_name(stock_columns, stock_col) is None:
                errors.append(
                    f"{FUNCTION_NAME}: stock_to_dim_policy_columns_mapping "
                    f"references column '{stock_col}' which does not exist on "
                    f"sheet '{stock_sheet_label}' (B.5.2/B.7)"
                )
            if dim_policy_columns is not None and \
                    _resolve_column_name(dim_policy_columns, policy_col) is None:
                errors.append(
                    f"{FUNCTION_NAME}: stock_to_dim_policy_columns_mapping "
                    f"references column '{policy_col}' which does not exist on "
                    f"sheet '{dim_policy_sheet}' (B.5.2/B.7)"
                )

    # --- B.8: stock_quantity_column exists on the Stock Balance sheet ---
    stock_qty_col = cfg.get("stock_quantity_column")
    if stock_columns is not None and not _is_blank(stock_qty_col) and \
            _resolve_column_name(stock_columns, stock_qty_col) is None:
        errors.append(
            f"{FUNCTION_NAME}: stock_quantity_column '{stock_qty_col}' does not "
            f"exist on sheet '{stock_sheet_label}' (B.5.2/B.8)"
        )


def _check_data_integrity(cfg: dict, workbook, errors: list[str]) -> None:
    """
    FS B.5.3 — data integrity checks C.1-C.3, against the loaded Dim Policy
    reference data. Checked once here, not per stock row.
    """
    dim_policy_sheet = str(cfg.get("dim_policy_sheet", "")).strip()
    if not dim_policy_sheet or not excel_utils.sheet_exists(workbook, dim_policy_sheet):
        return  # already reported by B.5

    dim_policy_df = _sheet_to_df(workbook, dim_policy_sheet)

    # --- C.1: no duplicate values in dim_policy_key_column ---
    # A duplicate composed key makes policy lookup ambiguous — hard stop.
    key_column = _resolve_column_name(dim_policy_df.columns, cfg.get("dim_policy_key_column"))
    if key_column is not None:
        seen, duplicates = set(), []
        for value in dim_policy_df[key_column].tolist():
            normalized = _normalize_key(value)
            if not normalized:
                continue
            if normalized in seen and _display(value) not in duplicates:
                duplicates.append(_display(value))
            seen.add(normalized)
        if duplicates:
            errors.append(
                f"{FUNCTION_NAME}: sheet '{dim_policy_sheet}' column "
                f"'{key_column}' contains duplicate policy key(s): "
                f"{', '.join(duplicates)} — policy lookup would be ambiguous "
                f"(B.5.3/C.1)"
            )

    # --- C.2: every cell in the mapped dimension columns is Yes or No ---
    mapping = cfg.get("stock_to_dim_policy_columns_mapping")
    if isinstance(mapping, dict):
        for policy_col in mapping.values():
            actual = _resolve_column_name(dim_policy_df.columns, policy_col)
            if actual is None:
                continue  # already reported by B.7
            bad_values = []
            for value in dim_policy_df[actual].tolist():
                normalized = _normalize_key(value)
                if normalized in ("yes", "no"):
                    continue
                shown = _display(value) or "(empty)"
                if shown not in bad_values:
                    bad_values.append(shown)
            if bad_values:
                errors.append(
                    f"{FUNCTION_NAME}: sheet '{dim_policy_sheet}' column "
                    f"'{actual}' contains value(s) other than Yes/No: "
                    f"{', '.join(bad_values)} (B.5.3/C.2)"
                )

    # --- C.3: Quantity Restriction values, where non-empty, are numeric ---
    qty_col = _resolve_column_name(dim_policy_df.columns, QUANTITY_RESTRICTION_COLUMN)
    if qty_col is not None:
        bad_values = []
        for value in dim_policy_df[qty_col].tolist():
            if _is_empty(value):
                continue
            if _to_number(value) is None and _display(value) not in bad_values:
                bad_values.append(_display(value))
        if bad_values:
            errors.append(
                f"{FUNCTION_NAME}: sheet '{dim_policy_sheet}' column "
                f"'{qty_col}' contains non-numeric value(s): "
                f"{', '.join(bad_values)} (B.5.3/C.3)"
            )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _function_config(config) -> dict:
    """This function's own config.yaml block (custom_functions.dimpolicyvalidation)."""
    return getattr(config, "custom_functions", {}).get(FUNCTION_NAME, {}) or {}


def _delimiter(cfg: dict) -> str:
    """FS B.5.1/A.7 — optional, defaults to an empty string."""
    value = cfg.get("dim_policy_key_delimiter")
    return "" if value is None else str(value)


def _sheet_to_df(workbook, sheet_name: str) -> pd.DataFrame:
    if not excel_utils.sheet_exists(workbook, sheet_name):
        raise ValueError(
            f"{FUNCTION_NAME}: sheet '{sheet_name}' not found in workbook"
        )
    df = excel_utils.sheet_to_dataframe(excel_utils.get_sheet(workbook, sheet_name))
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _columns_of(workbook, sheet_name: str) -> list[str] | None:
    """
    Column names of sheet_name, or None if the sheet is not in the workbook —
    None means "cannot be checked" rather than "has no columns".
    """
    if not sheet_name or not excel_utils.sheet_exists(workbook, sheet_name):
        return None
    return list(excel_utils.get_headers(excel_utils.get_sheet(workbook, sheet_name)).keys())


def _build_lookup(df: pd.DataFrame, key_column: str) -> dict[str, dict]:
    """
    Build a dict keyed by the normalized key_column value (FS section 9:
    comparison is case-insensitive and whitespace-trimmed). The first row
    for a given key wins.
    """
    lookup: dict[str, dict] = {}
    for _, row in df.iterrows():
        key = _normalize_key(row.get(key_column, ""))
        if key and key not in lookup:
            lookup[key] = row.to_dict()
    return lookup


def _normalize_key(value) -> str:
    """Normalized form used for all case-insensitive matching (FS section 9)."""
    return _display(value).lower()


def _display(value) -> str:
    """
    The value as it should read in a message or key component: stripped, with
    whole floats rendered without a trailing '.0' (Excel stores 1 as 1.0, and
    'requires 1' reads better than 'requires 1.0'). Empty values render as ''.
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def _is_empty(value) -> bool:
    """
    True if the value is empty business data (FS sections 2.5.2 / 6.1): None,
    NaN, whitespace-only, or the text 'nan' left behind by formula/import
    artifacts.
    """
    return _display(value) == ""


def _is_blank(value) -> bool:
    """True if a config value is absent or an empty/whitespace-only string."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    if isinstance(value, (list, tuple, dict, set)):
        return len(value) == 0
    return False


def _to_number(value) -> float | None:
    """The value as a float, or None if it is not numeric."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if (isinstance(value, float) and math.isnan(value)) else float(value)
    try:
        return float(str(value).strip())
    except (ValueError, TypeError):
        return None


def _values_equal(expected, actual) -> bool:
    """
    Compare a Dim Policy value against a Stock Balance value: numerically when
    both sides are numeric (so text '1' equals the number 1), otherwise as
    normalized strings (FS section 9).
    """
    expected_num, actual_num = _to_number(expected), _to_number(actual)
    if expected_num is not None and actual_num is not None:
        return expected_num == actual_num
    return _normalize_key(expected) == _normalize_key(actual)


def _resolve_column_name(available_columns, wanted_name):
    """
    Case-insensitive column name resolution (FS section 9). Returns the column
    name as it actually appears in available_columns, or None if absent.
    """
    if wanted_name is None:
        return None
    target = str(wanted_name).strip().lower()
    for actual in available_columns:
        if str(actual).strip().lower() == target:
            return actual
    return None
