import pytest
from validators.builtin import duplicate2


def run(wb, sheet_factory, config, issues_writer_factory, rows, ref_col="Warehouse"):
    sheet_factory(wb, "Stock", [["ItemId", "Warehouse"]] + rows)
    ws = wb["Stock"]
    iw = issues_writer_factory(wb, config)
    rule = {
        "RuleCode": "R1", "RuleName": "Dup2 check", "Severity": "ERROR", "Color": "RED",
        "ReferenceDataColumn": ref_col,
    }
    duplicate2.validate(ws, 1, "ItemId", rule, {"issues_writer": iw})
    return iw


def test_flags_duplicate_combination(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [
        ["A1", "WH1"], ["A1", "WH2"], ["A1", "WH1"],  # A1+WH1 duplicated
    ])
    assert iw.issue_counter == 2


def test_same_item_different_warehouse_is_not_duplicate(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [
        ["A1", "WH1"], ["A1", "WH2"],
    ])
    assert iw.issue_counter == 0


def test_empty_primary_column_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [
        [None, "WH1"], [None, "WH1"],
    ])
    assert iw.issue_counter == 0


def test_case_insensitive_combination(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [
        ["A1", "WH1"], ["a1", "wh1"],
    ])
    assert iw.issue_counter == 2


def test_missing_reference_column_raises(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    with pytest.raises(ValueError):
        run(wb, sheet_factory, fake_config, issues_writer_factory, [["A1", "WH1"]], ref_col="NoSuchColumn")


def test_reference_column_matched_case_insensitively(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    """FS section 9: ReferenceDataColumn should resolve even if its case
    differs from the actual header text."""
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
             [["A1", "WH1"], ["A1", "WH1"]], ref_col="warehouse")
    assert iw.issue_counter == 2
