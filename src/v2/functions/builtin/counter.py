REQUIRED_PARAMS = []


def counter(context: dict) -> list:
    """
    Generate a sequential counter [1, 2, 3, ..., n] for each data row.

    No parameters required.
    """
    worksheet_data = context.get("worksheet_data")

    if worksheet_data is None:
        raise ValueError("COUNTER: worksheet_data not available in context")

    row_count = len(worksheet_data)
    return list(range(1, row_count + 1))
