from datetime import datetime
from importer.conversion import convert_date, convert_numeric, convert_integer


# --- convert_date ---

def test_convert_date_matches_first_format():
    result = convert_date("2025-01-15", ["%Y-%m-%d", "%d/%m/%Y"])
    assert result == datetime(2025, 1, 15)


def test_convert_date_tries_formats_in_order():
    result = convert_date("15/01/2025", ["%Y-%m-%d", "%d/%m/%Y"])
    assert result == datetime(2025, 1, 15)


def test_convert_date_no_match_returns_none():
    assert convert_date("not-a-date", ["%Y-%m-%d"]) is None


def test_convert_date_none_returns_none():
    assert convert_date(None, ["%Y-%m-%d"]) is None


def test_convert_date_empty_string_returns_none():
    assert convert_date("", ["%Y-%m-%d"]) is None


# --- convert_numeric ---

def test_convert_numeric_simple():
    assert convert_numeric("123.45") == 123.45


def test_convert_numeric_custom_decimal_separator():
    assert convert_numeric("123,45", decimal_sep=",") == 123.45


def test_convert_numeric_with_thousands_separator():
    assert convert_numeric("1.234.567,89", decimal_sep=",", thousands_sep=".") == 1234567.89


def test_convert_numeric_thousands_group_must_be_three_digits():
    # "12.34" with thousands_sep="." is not a valid group (34 is only 2 digits after a full group)
    assert convert_numeric("12.34", decimal_sep=",", thousands_sep=".") is None


def test_convert_numeric_thousands_not_allowed_after_decimal():
    assert convert_numeric("1,234.56", decimal_sep=",", thousands_sep=".") is None


def test_convert_numeric_negative_sign():
    assert convert_numeric("-123.45") == -123.45


def test_convert_numeric_none_returns_none():
    assert convert_numeric(None) is None


def test_convert_numeric_garbage_returns_none():
    assert convert_numeric("abc") is None


def test_convert_numeric_multiple_decimal_separators_invalid():
    assert convert_numeric("1.2.3") is None


# --- convert_integer ---

def test_convert_integer_whole_number():
    assert convert_integer("123") == 123.0


def test_convert_integer_zero_fractional_accepted():
    assert convert_integer("123.000") == 123.0


def test_convert_integer_nonzero_fractional_rejected():
    assert convert_integer("123.5") is None


def test_convert_integer_none_returns_none():
    assert convert_integer(None) is None
