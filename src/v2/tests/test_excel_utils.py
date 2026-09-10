import pytest
from workbook.excel_utils import (
    is_empty, is_numeric_null, normalize_string, get_headers, get_headers_lower,
    find_sheet_name, sheet_exists, get_sheet, get_or_create_sheet,
    remove_sheet_if_exists, find_column_index,
)


# --- is_empty ---

def test_is_empty_none():
    assert is_empty(None) is True


def test_is_empty_whitespace():
    assert is_empty("   ") is True


def test_is_empty_nan_text():
    assert is_empty("NaN") is True
    assert is_empty("nan") is True


def test_is_empty_zero_is_not_empty():
    assert is_empty(0) is False


def test_is_empty_value_present():
    assert is_empty("x") is False


# --- is_numeric_null ---

def test_is_numeric_null_zero_variants():
    assert is_numeric_null(0) is True
    assert is_numeric_null("0.00") is True
    assert is_numeric_null("0,00") is True


def test_is_numeric_null_empty():
    assert is_numeric_null(None) is True
    assert is_numeric_null("") is True


def test_is_numeric_null_nonzero():
    assert is_numeric_null(5) is False
    assert is_numeric_null("abc") is False


# --- normalize_string ---

def test_normalize_string_case_and_whitespace():
    assert normalize_string("  Hello  ") == "hello"


def test_normalize_string_none():
    assert normalize_string(None) == ""


# --- sheet lookups: case-insensitive per FS section 9 ---

def test_find_sheet_name_exact_match(make_workbook):
    wb = make_workbook()
    wb.create_sheet("Items")
    assert find_sheet_name(wb, "Items") == "Items"


def test_find_sheet_name_case_insensitive(make_workbook):
    wb = make_workbook()
    wb.create_sheet("Items")
    assert find_sheet_name(wb, "items") == "Items"
    assert find_sheet_name(wb, "ITEMS") == "Items"


def test_find_sheet_name_not_found(make_workbook):
    wb = make_workbook()
    assert find_sheet_name(wb, "NoSuchSheet") is None


def test_sheet_exists_case_insensitive(make_workbook):
    wb = make_workbook()
    wb.create_sheet("Items")
    assert sheet_exists(wb, "items") is True
    assert sheet_exists(wb, "Stock") is False


def test_get_sheet_case_insensitive(make_workbook):
    wb = make_workbook()
    ws = wb.create_sheet("Items")
    assert get_sheet(wb, "items") is ws


def test_get_sheet_raises_keyerror_when_missing(make_workbook):
    wb = make_workbook()
    with pytest.raises(KeyError):
        get_sheet(wb, "NoSuchSheet")


def test_get_or_create_sheet_returns_existing_case_insensitively(make_workbook):
    wb = make_workbook()
    ws = wb.create_sheet("Items")
    assert get_or_create_sheet(wb, "items") is ws
    assert wb.sheetnames == ["Items"]  # no duplicate created


def test_get_or_create_sheet_creates_when_missing(make_workbook):
    wb = make_workbook()
    ws = get_or_create_sheet(wb, "NewSheet")
    assert ws.title == "NewSheet"
    assert "NewSheet" in wb.sheetnames


def test_remove_sheet_if_exists_case_insensitive(make_workbook):
    wb = make_workbook()
    wb.create_sheet("Items")
    remove_sheet_if_exists(wb, "items")
    assert wb.sheetnames == []


def test_remove_sheet_if_exists_noop_when_missing(make_workbook):
    wb = make_workbook()
    remove_sheet_if_exists(wb, "NoSuchSheet")  # should not raise
    assert wb.sheetnames == []


# --- header lookups: case-insensitive per FS section 9 ---

def test_get_headers(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId", "Qty"]])
    headers = get_headers(wb["Items"])
    assert headers == {"ItemId": 1, "Qty": 2}


def test_get_headers_lower(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId", "Qty"]])
    headers = get_headers_lower(wb["Items"])
    assert headers == {"itemid": 1, "qty": 2}


def test_find_column_index_exact_and_case_insensitive(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId", "Qty"]])
    headers = get_headers(wb["Items"])
    assert find_column_index(headers, "ItemId") == 1
    assert find_column_index(headers, "itemid") == 1
    assert find_column_index(headers, "ITEMID") == 1


def test_find_column_index_not_found(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"]])
    headers = get_headers(wb["Items"])
    assert find_column_index(headers, "NoSuchColumn") is None
