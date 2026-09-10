from validators.builtin import reference


def run(wb, sheet_factory, config, issues_writer_factory, item_rows, ref_rows,
        ref_sheet="ItemMaster", ref_col="ItemId"):
    sheet_factory(wb, "Stock", [["ItemId"]] + item_rows)
    sheet_factory(wb, ref_sheet, [["ItemId"]] + ref_rows)
    ws = wb["Stock"]
    iw = issues_writer_factory(wb, config)
    rule = {
        "RuleCode": "R1", "RuleName": "Ref check", "Severity": "ERROR", "Color": "RED",
        "ReferenceDataSheet": ref_sheet, "ReferenceDataColumn": ref_col,
    }
    reference.validate(ws, 1, "ItemId", rule, {"issues_writer": iw, "workbook": wb})
    return iw


def test_flags_value_not_in_reference(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
              [["A1"], ["A2"]], [["A1"]])
    assert iw.issue_counter == 1


def test_no_issue_when_all_present(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
              [["A1"], ["A2"]], [["A1"], ["A2"], ["A3"]])
    assert iw.issue_counter == 0


def test_empty_cells_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
              [[None], ["A1"]], [["A1"]])
    assert iw.issue_counter == 0


def test_case_insensitive_match(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
              [["a1"]], [["A1"]])
    assert iw.issue_counter == 0


def test_reference_sheet_and_column_matched_case_insensitively(
    make_workbook, sheet_factory, fake_config, issues_writer_factory
):
    """FS section 9: ReferenceDataSheet/ReferenceDataColumn should resolve
    even if their case differs from the actual sheet/header."""
    wb = make_workbook()
    # Sheet is created as "ItemMaster" with header "ItemId" (normal case),
    # but the rule references it in a completely different case.
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
              [["A1"]], [["A1"]], ref_sheet="ItemMaster", ref_col="ItemId")
    assert iw.issue_counter == 0

    # Now re-run with the rule using different case than the actual sheet/header
    wb2 = make_workbook()
    sheet_factory(wb2, "Stock", [["ItemId"], ["A1"]])
    sheet_factory(wb2, "ItemMaster", [["ItemId"], ["A1"]])
    ws2 = wb2["Stock"]
    iw2 = issues_writer_factory(wb2, fake_config)
    rule = {
        "RuleCode": "R1", "RuleName": "Ref check", "Severity": "ERROR", "Color": "RED",
        "ReferenceDataSheet": "itemmaster", "ReferenceDataColumn": "itemid",
    }
    reference.validate(ws2, 1, "ItemId", rule, {"issues_writer": iw2, "workbook": wb2})
    assert iw2.issue_counter == 0
