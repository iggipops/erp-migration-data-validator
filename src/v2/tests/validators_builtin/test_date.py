from datetime import datetime
from validators.builtin import date as date_validator


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["ShipDate"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Date check", "Severity": "ERROR", "Color": "RED"}
    date_validator.validate(ws, 1, "ShipDate", rule, {"issues_writer": iw, "config": config})
    return iw


def test_native_date_passes(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[datetime(2025, 1, 1)]])
    assert iw.issue_counter == 0


def test_text_value_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["2025-01-01"]])
    assert iw.issue_counter == 1


def test_plain_number_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[45658]])
    assert iw.issue_counter == 1


def test_empty_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[None]])
    assert iw.issue_counter == 0


def test_min_max_range(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    fake_config.date_validation = {"min_date": "2025-01-01", "max_date": "2025-12-31"}
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [
        [datetime(2024, 12, 31)],   # before min
        [datetime(2025, 6, 1)],     # in range
        [datetime(2026, 1, 1)],     # after max
    ])
    assert iw.issue_counter == 2


def test_null_dates_treated_as_empty_by_default(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    fake_config.date_validation = {"null_dates": ["1900-01-01"]}
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[datetime(1900, 1, 1)]])
    assert iw.issue_counter == 0


def test_null_dates_treated_as_invalid_when_configured(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    fake_config.date_validation = {
        "null_dates": ["1900-01-01"],
        "treat_null_dates_as": "invalid",
    }
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[datetime(1900, 1, 1)]])
    assert iw.issue_counter == 1
