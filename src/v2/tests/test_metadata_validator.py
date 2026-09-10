import pytest
from metadata.metadata_validator import MetadataValidator


class FakeRegistry:
    def __init__(self, known_fns=None, required_params=None):
        self.known_fns = known_fns or set()
        self.required_params = required_params or {}

    def has(self, name):
        return name in self.known_fns

    def get_required_params(self, name):
        return self.required_params.get(name)

    def get_metadata_check_fn(self, name):
        return None


class FakeVType:
    def __init__(self, known_types=("EMPTY", "NULL", "DUPLICATE", "DUPLICATE2",
                                     "FORMAT", "REFERENCE", "CUSTOM", "AI",
                                     "DATE", "NUMERIC", "INTEGER")):
        self.known_types = set(known_types)

    def has(self, t):
        return t in self.known_types

    def get_metadata_check_fn(self, t):
        return None


VR_HEADERS = ["SequenceNum", "RuleCode", "RuleName", "SheetName", "ColumnName",
              "ValidationType", "Severity", "Color", "Active",
              "ReferenceDataSheet", "ReferenceDataColumn", "CustomFunctionName"]

ER_HEADERS = ["SequenceNum", "RuleCode", "RuleName", "SheetName", "TargetColumn",
              "CustomFunctionName", "FunctionArguments", "Active", "Color"]


def build_wb(make_workbook, sheet_factory, items_rows=None, vr_rows=None, er_rows=None,
             vr_headers=None, er_headers=None):
    wb = make_workbook()
    sheet_factory(wb, "Items", items_rows or [["ItemId"], ["A1"]])
    if vr_rows is not None:
        sheet_factory(wb, "ValidationRules", [vr_headers or VR_HEADERS] + vr_rows)
    if er_rows is not None:
        sheet_factory(wb, "EnrichmentRules", [er_headers or ER_HEADERS] + er_rows)
    return wb


def mv(wb, config=None, failed_sheets=None):
    return MetadataValidator(
        wb, config=config or FakeConfig(), registry=FakeRegistry(),
        vtype_registry=FakeVType(), failed_sheets=failed_sheets,
    )


class FakeConfig:
    color_map = {"RED": "FF0000"}
    ai_enabled = False
    custom_functions = {}


def test_missing_validationrules_sheet_is_error(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    errors = mv(wb).validate()
    assert any("ValidationRules" in e and "not found" in e for e in errors)


def test_valid_rule_passes(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_unknown_validation_type(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "BOGUS", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("ValidationType" in e for e in errors)


def test_unknown_severity(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "CRITICAL", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("Severity" in e for e in errors)


def test_invalid_color(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "NOTACOLOR", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("Color" in e for e in errors)


def test_empty_color_is_valid(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_color_hex_with_hash_prefix_valid(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "#00FF00", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_color_hex_3_digit_shorthand_valid(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "0F0", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_color_hex_3_digit_shorthand_with_hash_valid(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "#0F0", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_sheetname_not_in_workbook(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "NoSuchSheet", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("NoSuchSheet" in e for e in errors)


def test_columnname_not_in_sheet(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "NoSuchColumn", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("NoSuchColumn" in e for e in errors)


def test_column_matched_case_insensitively(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "items", "itemid", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert errors == []


def test_custom_type_requires_known_function(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "CUSTOM", "ERROR", "RED", "Yes", "", "", "unknown_fn"],
    ])
    errors = mv(wb).validate()
    assert any("unknown_fn" in e for e in errors)


def test_reference_type_requires_ref_sheet_and_column(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("ReferenceDataSheet" in e or "ReferenceDataColumn" in e for e in errors)


def test_reference_type_requires_ref_sheet_and_column_blank_cell(make_workbook, sheet_factory):
    """Regression: a truly blank ReferenceDataSheet/ReferenceDataColumn
    cell reads as None from openpyxl (not ""). It must still produce the
    "required for REFERENCE type" message, not get stringified into the
    literal text "None" and reported as "not found in workbook"."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes", None, None, None],
    ])
    errors = mv(wb).validate()
    assert any("ReferenceDataSheet is required" in e for e in errors)
    assert any("ReferenceDataColumn is required" in e for e in errors)
    assert not any('"None"' in e for e in errors)


def test_reference_data_sheet_must_differ_from_sheet_name(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes",
         "Items", "ItemId", ""],
    ])
    errors = mv(wb).validate()
    assert any("ReferenceDataSheet" in e and "differ" in e for e in errors)


def test_reference_data_sheet_different_from_sheet_name_passes(make_workbook, sheet_factory):
    wb = build_wb(
        make_workbook, sheet_factory,
        items_rows=[["ItemId"], ["A1"]],
        vr_rows=[
            [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes",
             "Other", "OtherId", ""],
        ],
    )
    sheet_factory(wb, "Other", [["OtherId"], ["A1"]])
    errors = mv(wb).validate()
    assert errors == []


def test_duplicate_rulecode(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule A", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
        [2, "R1", "Rule B", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb).validate()
    assert any("RuleCode" in e for e in errors)


def test_mandatory_header_missing(make_workbook, sheet_factory):
    headers = [h for h in VR_HEADERS if h != "Active"]
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[[1, "R1", "Rule"]], vr_headers=headers)
    errors = mv(wb).validate()
    assert any('mandatory column "Active" not found' in e for e in errors)


# --- Import failure isolation (FS 8.5.4) ---

def test_row_depending_on_failed_sheet_is_skipped_not_errored(make_workbook, sheet_factory):
    """A row referencing a SheetName that failed to import must not raise
    a hard metadata error (which used to kill the entire run) — it's
    skipped, and the unrelated valid row still validates cleanly."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule A", "Items", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
        [2, "R2", "Rule B", "Stock", "Qty", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert errors == []


def test_row_depending_on_failed_sheet_matched_case_insensitively(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "stock", "Qty", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert errors == []


def test_row_not_depending_on_failed_sheet_still_errors(make_workbook, sheet_factory):
    """Only rows tied to the failed SheetName are spared — an unrelated
    row with its own genuine problem must still be reported."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "NoSuchSheet", "ItemId", "EMPTY", "ERROR", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert any("NoSuchSheet" in e for e in errors)


def test_reference_row_depending_on_failed_reference_sheet_is_skipped(make_workbook, sheet_factory):
    """A REFERENCE-type row whose ReferenceDataSheet (not SheetName) failed
    to import must not raise a hard 'ReferenceDataSheet not found' error —
    it's skipped, same as the SheetName case."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes",
         "Stock", "Qty", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert errors == []


def test_reference_row_depending_on_failed_reference_sheet_matched_case_insensitively(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes",
         "stock", "Qty", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert errors == []


def test_reference_row_with_unrelated_reference_sheet_still_errors(make_workbook, sheet_factory):
    """Only rows tied to the failed ReferenceDataSheet are spared — a
    REFERENCE row whose ReferenceDataSheet is a genuinely missing sheet
    (unrelated to any failed import) must still be reported."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "REFERENCE", "ERROR", "RED", "Yes",
         "NoSuchSheet", "Qty", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert any("NoSuchSheet" in e for e in errors)


def test_enrichment_row_depending_on_failed_sheet_is_skipped(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[], er_rows=[
        [1, "E1", "Enrich", "Stock", "NewCol", "counter", "", "Yes", ""],
    ])
    errors = mv(wb, failed_sheets={"Stock"}).validate()
    assert errors == []


# --- EnrichmentRules ---

def test_enrichment_target_column_must_not_already_exist(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, er_rows=[
        [1, "E1", "Enrich", "Items", "ItemId", "counter", "", "Yes", ""],
    ])
    errors = mv(wb).validate()
    assert any("ItemId" in e for e in errors)


def test_enrichment_valid_new_column(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, er_rows=[
        [1, "E1", "Enrich", "Items", "NewCol", "counter", "", "Yes", ""],
    ])
    errors = mv(wb, config=FakeConfig()).validate()
    # counter is not in FakeRegistry's known_fns by default -> CustomFunctionName error expected
    assert any("counter" in e for e in errors)


def test_enrichment_valid_new_column_known_function(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[], er_rows=[
        [1, "E1", "Enrich", "Items", "NewCol", "counter", "", "Yes", ""],
    ])
    m = MetadataValidator(wb, config=FakeConfig(), registry=FakeRegistry(known_fns={"counter"}), vtype_registry=FakeVType())
    errors = m.validate()
    assert errors == []


# --- AI type (FS section 6.8/8.6.3) ---

def ai_config(ai_enabled=False, custom_functions=None):
    cfg = FakeConfig()
    cfg.ai_enabled = ai_enabled
    cfg.custom_functions = custom_functions or {}
    return cfg


def test_ai_type_with_no_reference_fields_and_ai_disabled_passes(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, config=ai_config(ai_enabled=False)).validate()
    assert errors == []


def test_ai_type_customfunctionname_not_required_when_ai_disabled(make_workbook, sheet_factory):
    """FS 8.6.3: CustomFunctionName is only required when ai_enabled=true —
    a soft-skipped (ai_enabled=false) row isn't held to this check."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, config=ai_config(ai_enabled=False)).validate()
    assert errors == []


def test_ai_type_customfunctionname_required_when_ai_enabled(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, config=ai_config(ai_enabled=True)).validate()
    assert any("CustomFunctionName" in e for e in errors)


def test_ai_type_customfunctionname_must_be_known_function_when_ai_enabled(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", "unknown_fn"],
    ])
    errors = mv(wb, config=ai_config(ai_enabled=True)).validate()
    assert any("unknown_fn" in e for e in errors)


def test_ai_type_known_function_with_ai_enabled_passes(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", "ai_fn"],
    ])
    m = MetadataValidator(
        wb, config=ai_config(ai_enabled=True), registry=FakeRegistry(known_fns={"ai_fn"}),
        vtype_registry=FakeVType(),
    )
    assert m.validate() == []


def test_ai_type_reference_data_sheet_optional_when_absent(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", "", "", ""],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert errors == []


def test_ai_type_reference_data_sheet_optional_when_blank_cell(make_workbook, sheet_factory):
    """Regression: a truly blank Excel cell reads as None (not ""), which
    must still be recognized as empty rather than stringified into the
    literal text "None"."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes", None, None, None],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert errors == []


def test_ai_type_reference_data_sheet_must_exist_if_provided(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "NoSuchSheet", "", ""],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert any("NoSuchSheet" in e for e in errors)


def test_ai_type_reference_data_sheet_must_differ_from_sheet_name(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "Items", "", ""],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert any("ReferenceDataSheet" in e and "differ" in e for e in errors)


def test_ai_type_reference_data_sheet_different_from_sheet_name_passes(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "Other", "", ""],
    ])
    sheet_factory(wb, "Other", [["OtherId"], ["A1"]])
    errors = mv(wb, config=ai_config()).validate()
    assert errors == []


def test_ai_type_reference_data_column_requires_reference_data_sheet(make_workbook, sheet_factory):
    """FS section 6.8/8.6.3 consistency rule: ReferenceDataColumn must be
    empty when ReferenceDataSheet is empty for AI type."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "", "SomeColumn", ""],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert any("ReferenceDataColumn" in e and "ReferenceDataSheet" in e for e in errors)


def test_ai_type_reference_data_column_requires_reference_data_sheet_blank_cell(make_workbook, sheet_factory):
    """Same consistency rule as above, but with a truly blank
    ReferenceDataSheet cell (None) rather than an explicit empty string —
    must still be flagged, not silently accepted."""
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         None, "SomeColumn", ""],
    ])
    errors = mv(wb, config=ai_config()).validate()
    assert any("ReferenceDataColumn" in e and "ReferenceDataSheet" in e for e in errors)


def test_ai_type_reference_data_column_must_exist_on_reference_data_sheet(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "Other", "NoSuchColumn", ""],
    ])
    sheet_factory(wb, "Other", [["OtherId"], ["A1"]])
    errors = mv(wb, config=ai_config()).validate()
    assert any("NoSuchColumn" in e for e in errors)


def test_ai_type_reference_data_column_existing_on_reference_data_sheet_passes(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "Other", "OtherId", ""],
    ])
    sheet_factory(wb, "Other", [["OtherId"], ["A1"]])
    errors = mv(wb, config=ai_config()).validate()
    assert errors == []


def test_ai_type_reference_data_column_pending_enrichment_target_passes(make_workbook, sheet_factory):
    """FS section 2.5.3 forward-reference exception: ReferenceDataColumn
    may name a column not yet created, pending an active EnrichmentRules
    row targeting ReferenceDataSheet."""
    wb = build_wb(
        make_workbook, sheet_factory,
        vr_rows=[
            [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
             "Other", "PendingCol", ""],
        ],
        er_rows=[
            [1, "E1", "Enrich", "Other", "PendingCol", "counter", "", "Yes", ""],
        ],
    )
    sheet_factory(wb, "Other", [["OtherId"], ["A1"]])
    m = MetadataValidator(
        wb, config=ai_config(), registry=FakeRegistry(known_fns={"counter"}),
        vtype_registry=FakeVType(),
    )
    assert m.validate() == []


def test_ai_row_depending_on_failed_reference_sheet_is_skipped(make_workbook, sheet_factory):
    wb = build_wb(make_workbook, sheet_factory, vr_rows=[
        [1, "R1", "Rule", "Items", "ItemId", "AI", "WARNING", "RED", "Yes",
         "Stock", "Qty", ""],
    ])
    errors = mv(wb, config=ai_config(), failed_sheets={"Stock"}).validate()
    assert errors == []


def test_shared_error_list_regression(make_workbook, sheet_factory):
    """Regression: a ValidationRules row-level error must not suppress
    EnrichmentRules' own errors (previously both checked `if self.errors`
    against a list shared across both sheets)."""
    wb = build_wb(
        make_workbook, sheet_factory,
        vr_rows=[[1, "R1", "Rule", "Items", "ItemId", "EMPTY", "ERROR", "NOTACOLOR", "Yes", "", "", ""]],
        er_rows=[[1, "E1", "Enrich", "Items", "NewCol", "unknown_fn", "", "Yes", ""]],
    )
    errors = mv(wb).validate()
    assert any("Color" in e for e in errors)
    assert any("unknown_fn" in e for e in errors)
