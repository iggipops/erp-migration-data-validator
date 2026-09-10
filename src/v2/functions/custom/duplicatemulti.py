"""
DUPLICATEMULTI — Custom validation function for duplicate detection
across three or more columns.

Flags rows where the combination of N columns appears more than once
in the same worksheet. All occurrences of a duplicated combination are
flagged — not only the second occurrence. Comparison is case-insensitive
per FS section 9.

config.yaml parameters:
    columns: list[str]  — three or more column names to check together
"""
from collections import Counter


REQUIRED_PARAMS = ["columns"]


def duplicatemulti(context: dict) -> list[bool]:

    params  = context.get("technical_config", {})
    columns = params.get("columns", [])

    if isinstance(columns, str):
        columns = [columns]

    if len(columns) < 3:
        raise ValueError(
            "DUPLICATEMULTI: 'columns' must contain at least 3 column names "
            "(use DUPLICATE for 1 column or DUPLICATE2 for 2 columns)"
        )

    named_columns = context.get("named_columns", {})

    for col in columns:
        if col not in named_columns:
            raise ValueError(f"DUPLICATEMULTI: column '{col}' not found in context")

    value_lists = [named_columns[col] for col in columns]
    row_count   = len(value_lists[0])

    for col, values in zip(columns, value_lists):
        if len(values) != row_count:
            raise ValueError(
                f"DUPLICATEMULTI: column '{col}' has a different number of rows"
            )

    # Build normalized tuples — case-insensitive, stripped, per section 9
    combos = []
    for i in range(row_count):
        combo = tuple(
            _normalize(value_lists[c][i]) for c in range(len(columns))
        )
        combos.append(combo)

    counts     = Counter(combos)
    duplicates = {combo for combo, count in counts.items() if count > 1}

    # True = valid (not a duplicate), False = invalid (duplicate found)
    return [combo not in duplicates for combo in combos]


def _normalize(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() == "nan":
        return ""
    return s.lower()
