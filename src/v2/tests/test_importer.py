from datetime import datetime
from importer.importer import Importer


def test_imports_csv_as_new_sheet(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId,Qty\nA1,5\nA2,10\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    result = imp.import_external_files(rules, [])
    assert result["loaded"] == ["Items"]
    ws = wb["Items"]
    assert [c.value for c in ws[1]] == ["ItemId", "Qty"]
    assert ws.cell(row=2, column=1).value == "A1"


def test_replaces_existing_sheet_with_same_name(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId\nA1\n")
    wb = make_workbook()
    wb.create_sheet("Items").append(["OldData"])
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    imp.import_external_files(rules, [])
    assert wb["Items"].cell(row=1, column=1).value == "ItemId"


def test_replaces_existing_sheet_case_insensitively(make_workbook, fake_config, tmp_path):
    """Regression: sheet removal must resolve case-insensitively —
    otherwise this raises KeyError instead of replacing the sheet."""
    f = tmp_path / "items.csv"
    f.write_text("ItemId\nA1\n")
    wb = make_workbook()
    wb.create_sheet("items")  # lowercase, rule says "Items"
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    result = imp.import_external_files(rules, [])
    assert result["loaded"] == ["Items"]
    assert wb["Items"].cell(row=1, column=1).value == "ItemId"


def test_inactive_rule_skipped(make_workbook, fake_config, tmp_path):
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": "/does/not/exist.csv",
              "Format": "csv", "Active": "No"}]
    result = imp.import_external_files(rules, [])
    assert result["loaded"] == []
    assert result["failed"] == []


def test_missing_file_recorded_as_failed(make_workbook, fake_config):
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": "/does/not/exist.csv",
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ","}]
    result = imp.import_external_files(rules, [])
    assert result["loaded"] == []
    assert len(result["failed"]) == 1
    assert result["failed_sheets"] == {"Items"}


def test_successful_import_leaves_failed_sheets_empty(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId\nA1\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    result = imp.import_external_files(rules, [])
    assert result["failed_sheets"] == set()


def test_one_failed_one_ok_only_failed_sheet_recorded(make_workbook, fake_config, tmp_path):
    ok_file = tmp_path / "items.csv"
    ok_file.write_text("ItemId\nA1\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [
        {"SequenceNum": 1, "SheetName": "Stock", "FilePath": "/does/not/exist.csv",
         "Format": "csv", "Active": "Yes", "CSVDelimiter": ","},
        {"SequenceNum": 2, "SheetName": "Items", "FilePath": str(ok_file),
         "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""},
    ]
    result = imp.import_external_files(rules, [])
    assert result["loaded"] == ["Items"]
    assert result["failed_sheets"] == {"Stock"}


def test_import_spec_numeric_conversion_applied(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId,Qty\nA1,\"1.234,56\"\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    ef_rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
                 "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    isp_rules = [{"ExternalFileSheetName": "Items", "ColumnName": "Qty", "TargetType": "numeric",
                  "Active": "Yes", "DecimalSeparator": ",", "ThousandsSeparator": "."}]
    result = imp.import_external_files(ef_rules, isp_rules)
    assert result["loaded"] == ["Items"]
    assert wb["Items"].cell(row=2, column=2).value == 1234.56


def test_leading_equals_value_not_written_as_formula(make_workbook, fake_config, tmp_path):
    """A text value like '=HYPERLINK(...)' copied verbatim from an external
    CSV must be stored as literal text, not a live formula (formula/CSV
    injection guard)."""
    f = tmp_path / "items.csv"
    f.write_text('ItemId,Note\nA1,"=HYPERLINK(""http://evil"")"\n')
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    imp.import_external_files(rules, [])
    cell = wb["Items"].cell(row=2, column=2)
    assert cell.data_type == "s"
    assert cell.value == '=HYPERLINK("http://evil")'


def test_import_spec_column_matched_case_insensitively(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId,qty\nA1,5\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    ef_rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
                 "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "OriginalSheetName": ""}]
    # ImportSpec references "Qty" (capital Q) but actual CSV header is "qty"
    isp_rules = [{"ExternalFileSheetName": "Items", "ColumnName": "Qty", "TargetType": "integer",
                  "Active": "Yes"}]
    result = imp.import_external_files(ef_rules, isp_rules)
    assert result["conversion_warnings"]["Items"] == 0
    assert wb["Items"].cell(row=2, column=2).value == 5.0
