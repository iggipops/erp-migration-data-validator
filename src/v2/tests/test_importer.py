from datetime import datetime
from importer.importer import Importer


def test_imports_csv_as_new_sheet(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId,Qty\nA1,5\nA2,10\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [{"SequenceNum": 1, "SheetName": "Items", "FilePath": str(f),
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
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
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
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
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
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
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8"}]
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
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
    result = imp.import_external_files(rules, [])
    assert result["failed_sheets"] == set()


def test_one_failed_one_ok_only_failed_sheet_recorded(make_workbook, fake_config, tmp_path):
    ok_file = tmp_path / "items.csv"
    ok_file.write_text("ItemId\nA1\n")
    wb = make_workbook()
    imp = Importer(wb, fake_config)
    rules = [
        {"SequenceNum": 1, "SheetName": "Stock", "FilePath": "/does/not/exist.csv",
         "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8"},
        {"SequenceNum": 2, "SheetName": "Items", "FilePath": str(ok_file),
         "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""},
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
                 "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
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
              "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
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
                 "Format": "csv", "Active": "Yes", "CSVDelimiter": ",", "CSVEncoding": "utf-8", "OriginalSheetName": ""}]
    # ImportSpec references "Qty" (capital Q) but actual CSV header is "qty"
    isp_rules = [{"ExternalFileSheetName": "Items", "ColumnName": "Qty", "TargetType": "integer",
                  "Active": "Yes"}]
    result = imp.import_external_files(ef_rules, isp_rules)
    assert result["conversion_warnings"]["Items"] == 0
    assert wb["Items"].cell(row=2, column=2).value == 5.0


# --- F-04: NA-like text and csv encoding (FS 8.6.1, 8.6.4) ---

def csv_rule(path, encoding="utf-8", sheet="Items", seq=1):
    return {"SequenceNum": seq, "SheetName": sheet, "FilePath": str(path),
            "Format": "csv", "Active": "Yes", "CSVDelimiter": ",",
            "CSVEncoding": encoding, "OriginalSheetName": ""}


def test_na_like_text_survives_and_only_empty_cell_is_missing(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("Code,Note\n"
                 "NULL,a\nNA,b\nN/A,c\nNone,d\n#N/A,e\nnan,f\nnull,g\n,h\n")
    wb = make_workbook()
    result = Importer(wb, fake_config).import_external_files([csv_rule(f)], [])
    assert result["loaded"] == ["Items"]
    ws = wb["Items"]
    assert [ws.cell(row=r, column=1).value for r in range(2, 10)] == [
        "NULL", "NA", "N/A", "None", "#N/A", "nan", "null", None]
    assert [ws.cell(row=r, column=2).value for r in range(2, 10)] == list("abcdefgh")


def test_na_like_text_survives_when_import_spec_leaves_column_as_text(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("Code\nNULL\nNA\n")
    wb = make_workbook()
    isp = [{"ExternalFileSheetName": "Items", "ColumnName": "Code", "TargetType": "text",
            "Active": "Yes"}]
    Importer(wb, fake_config).import_external_files([csv_rule(f)], isp)
    assert [wb["Items"].cell(row=r, column=1).value for r in (2, 3)] == ["NULL", "NA"]


def test_utf8_sig_strips_bom_from_first_header(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_bytes("ItemId,Qty\nA1,5\n".encode("utf-8-sig"))
    wb = make_workbook()
    result = Importer(wb, fake_config).import_external_files([csv_rule(f, "utf-8-sig")], [])
    assert result["loaded"] == ["Items"]
    assert wb["Items"].cell(row=1, column=1).value == "ItemId"


def test_cp1251_decodes_cyrillic(make_workbook, fake_config, tmp_path):
    f = tmp_path / "units.csv"
    f.write_bytes("Код,Название\nA1,Штука\n".encode("cp1251"))
    wb = make_workbook()
    result = Importer(wb, fake_config).import_external_files([csv_rule(f, "cp1251")], [])
    assert result["loaded"] == ["Items"]
    ws = wb["Items"]
    assert [c.value for c in ws[1]] == ["Код", "Название"]
    assert ws.cell(row=2, column=2).value == "Штука"


def test_cp1252_decodes_western_characters(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_bytes("Name\nCafé €5\n".encode("cp1252"))
    wb = make_workbook()
    Importer(wb, fake_config).import_external_files([csv_rule(f, "cp1252")], [])
    assert wb["Items"].cell(row=2, column=1).value == "Café €5"


def test_encoding_value_is_case_insensitive(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_bytes("ItemId\nA1\n".encode("utf-8-sig"))
    wb = make_workbook()
    result = Importer(wb, fake_config).import_external_files([csv_rule(f, "UTF-8-SIG")], [])
    assert result["loaded"] == ["Items"]
    assert wb["Items"].cell(row=1, column=1).value == "ItemId"


def test_undecodable_file_is_named_import_failure(make_workbook, fake_config, tmp_path):
    """cp1251 bytes declared as utf-8: recorded like any other failed import
    (FS 8.6.4) - names the file and the declared encoding, no raw pandas error."""
    bad = tmp_path / "units.csv"
    bad.write_bytes("Код\nШтука\n".encode("cp1251"))
    ok = tmp_path / "items.csv"
    ok.write_text("ItemId\nA1\n")
    wb = make_workbook()
    result = Importer(wb, fake_config).import_external_files(
        [csv_rule(bad, "utf-8", sheet="Units", seq=1), csv_rule(ok, sheet="Items", seq=2)], [])
    assert result["loaded"] == ["Items"]              # processing continued
    assert result["failed_sheets"] == {"Units"}
    assert len(result["failed"]) == 1
    msg = result["failed"][0]
    assert "units.csv" in msg and "'utf-8'" in msg and "CSVEncoding" in msg
    assert "UnicodeDecodeError" not in msg
    assert "Units" not in wb.sheetnames


def test_missing_or_unsupported_encoding_is_named_import_failure(make_workbook, fake_config, tmp_path):
    f = tmp_path / "items.csv"
    f.write_text("ItemId\nA1\n")
    wb = make_workbook()
    for enc in ("", "latin-1"):
        result = Importer(wb, fake_config).import_external_files([csv_rule(f, enc)], [])
        assert result["loaded"] == []
        assert result["failed_sheets"] == {"Items"}
        assert "CSVEncoding" in result["failed"][0]
