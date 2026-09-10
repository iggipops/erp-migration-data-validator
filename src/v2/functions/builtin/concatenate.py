REQUIRED_PARAMS = ["columns"]


def concatenate(context: dict) -> list:
    """
    Concatenate values from two or more columns into a single string per row.

    Expected context keys:
        named_columns   : dict[str, list]  — {column_name: [values...]}
        technical_config: dict             — function params from YAML rule
            columns   : list[str]  — ordered column names to concatenate
            separator : str        — string placed between values (default "")
            trim      : bool       — strip whitespace from each value (default True)
    """
    params    = context.get("technical_config", {})
    columns   = params.get("columns", [])
    separator = str(params.get("separator", ""))
    trim      = bool(params.get("trim", True))

    # Handle both string (single value) and list (multiple values)
    if isinstance(columns, str):
        columns = [columns] if columns else []

    named_columns = context.get("named_columns", {})

    if not columns:
        raise ValueError("CONCATENATE: 'columns' parameter is required")

    # Verify all requested columns are available
    for col in columns:
        if col not in named_columns:
            raise ValueError(f"CONCATENATE: column '{col}' not found in context")

    row_count = len(named_columns[columns[0]])
    result    = []

    for i in range(row_count):
        parts = []
        for col in columns:
            value = named_columns[col][i]
            # Treat None and pandas NaN as empty string
            import math
            if value is None or (isinstance(value, float) and math.isnan(value)):
                text = ""
            else:
                text = str(value)
            if trim:
                text = text.strip()
            parts.append(text)
        result.append(separator.join(parts))

    return result
