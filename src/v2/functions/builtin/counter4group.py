import pandas as pd


REQUIRED_PARAMS = ["grouping_columns"]


def counter4group(context: dict) -> list:
    """
    Generate a counter that resets within each group, sorted within the group.

    Expected context keys:
        worksheet_data  : pd.DataFrame
        technical_config: dict
            grouping_columns : list[str]  — columns that define group boundaries
            sort_columns     : list[str]  — columns to sort by within each group
                                            (ascending, applied before counting)
    """
    params           = context.get("technical_config", {})
    grouping_columns = params.get("grouping_columns", [])
    sort_columns     = params.get("sort_columns", [])

    # Handle both string (single value) and list (multiple values)
    if isinstance(grouping_columns, str):
        grouping_columns = [grouping_columns] if grouping_columns else []
    if isinstance(sort_columns, str):
        sort_columns = [sort_columns] if sort_columns else []

    df = context.get("worksheet_data")
    if df is None:
        raise ValueError("COUNTER4GROUP: worksheet_data not available in context")

    if not grouping_columns:
        raise ValueError("COUNTER4GROUP: 'grouping_columns' parameter is required")

    # Validate columns exist
    for col in grouping_columns + sort_columns:
        if col not in df.columns:
            raise ValueError(f"COUNTER4GROUP: column '{col}' not found in worksheet")

    # Work on a copy with original index to restore order later
    working = df.copy()
    working["_original_index"] = range(len(working))

    if sort_columns:
        working = working.sort_values(
            by=grouping_columns + sort_columns,
            kind="stable"
        )

    # Generate cumcount within each group (0-based → +1 for 1-based)
    working["_group_counter"] = (
        working.groupby(grouping_columns, sort=False).cumcount() + 1
    )

    # Restore original row order
    working = working.sort_values("_original_index")

    return working["_group_counter"].astype("int64").tolist()
