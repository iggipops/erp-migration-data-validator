from metadata.metadata_reader import MetadataReader


def test_reads_validation_rules(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "ValidationRules", [
        ["SequenceNum", "RuleCode", "RuleName", "SheetName", "ColumnName",
         "ValidationType", "Severity", "Color", "Active",
         "ReferenceDataSheet", "ReferenceDataColumn", "CustomFunctionName"],
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    rows = MetadataReader(wb).read_validation_rules()
    assert len(rows) == 1
    assert rows[0]["SequenceNum"] == 1
    assert rows[0]["RuleCode"] == "R1"
    assert rows[0]["Active"] == "Yes"


def test_missing_sheet_returns_empty_list(make_workbook):
    wb = make_workbook()
    assert MetadataReader(wb).read_validation_rules() == []
    assert MetadataReader(wb).read_enrichment_rules() == []
    assert MetadataReader(wb).read_external_files() == []
    assert MetadataReader(wb).read_import_spec() == []


def test_blank_trailing_rows_skipped(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "ValidationRules", [
        ["SequenceNum", "RuleCode", "RuleName", "SheetName", "ColumnName",
         "ValidationType", "Severity", "Color", "Active",
         "ReferenceDataSheet", "ReferenceDataColumn", "CustomFunctionName"],
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
        [None, None, None, None, None, None, None, None, None, None, None, None],
    ])
    rows = MetadataReader(wb).read_validation_rules()
    assert len(rows) == 1


def test_canonical_keys_regardless_of_header_case(make_workbook, sheet_factory):
    """Regression: this is the critical fix — row dicts must be keyed by
    the CANONICAL column name even when the actual header text differs
    in case, since every downstream consumer (engines, importer) accesses
    rows via rule["SequenceNum"] / rule.get("Active") etc."""
    wb = make_workbook()
    sheet_factory(wb, "ValidationRules", [
        ["sequencenum", "RULECODE", "RuleName", "SHEETNAME", "columnname",
         "validationtype", "SEVERITY", "Color", "active",
         "ReferenceDataSheet", "ReferenceDataColumn", "CustomFunctionName"],
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    rows = MetadataReader(wb).read_validation_rules()
    assert rows[0]["SequenceNum"] == 1
    assert rows[0]["RuleCode"] == "R1"
    assert rows[0]["SheetName"] == "Items"
    assert rows[0]["ColumnName"] == "ItemId"
    assert rows[0]["ValidationType"] == "EMPTY"
    assert rows[0]["Severity"] == "ERROR"
    assert rows[0]["Active"] == "Yes"
    # direct bracket access (used throughout engines) must not KeyError
    assert rows[0]["SequenceNum"] == 1


def test_extra_unknown_columns_preserved_under_literal_name(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "ValidationRules", [
        ["SequenceNum", "RuleCode", "RuleName", "SheetName", "ColumnName",
         "ValidationType", "Severity", "Color", "Active",
         "ReferenceDataSheet", "ReferenceDataColumn", "CustomFunctionName", "Notes"],
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", "", "extra info"],
    ])
    rows = MetadataReader(wb).read_validation_rules()
    assert rows[0]["Notes"] == "extra info"


def test_reads_external_files_and_import_spec(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "ExternalFiles", [
        ["SequenceNum", "SheetName", "FilePath", "Format", "Active",
         "OriginalSheetName", "CSVDelimiter"],
        [1, "Items", "/tmp/x.csv", "csv", "Yes", "", ","],
    ])
    sheet_factory(wb, "ImportSpec", [
        ["ExternalFileSheetName", "ColumnName", "TargetType", "Active",
         "DateFormats", "DecimalSeparator", "ThousandsSeparator"],
        ["Items", "Qty", "numeric", "Yes", "", ".", ","],
    ])
    reader = MetadataReader(wb)
    ef = reader.read_external_files()
    isp = reader.read_import_spec()
    assert ef[0]["FilePath"] == "/tmp/x.csv"
    assert isp[0]["ColumnName"] == "Qty"
