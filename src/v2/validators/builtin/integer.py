from validators.builtin.numeric import _validate_numeric


def validate(sheet, col_idx, col_name, rule, context):
    """
    INTEGER validation per FS section 6.11 (V2.3).

    Same as NUMERIC but also requires zero fractional part.
    Native values like 123.0 or 123.0000 are valid (ERP migration context).
    """
    _validate_numeric(sheet, col_idx, col_name, rule, context, integer_only=True)
