from validators.builtin import duplicate


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["ItemId"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Dup check", "Severity": "ERROR", "Color": "RED"}
    duplicate.validate(ws, 1, "ItemId", rule, {"issues_writer": iw})
    return iw


def test_flags_all_occurrences_of_duplicate(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
             [["A1"], ["A2"], ["A1"], ["A3"]])
    # Both occurrences of A1 flagged, not just the second
    assert iw.issue_counter == 2


def test_case_insensitive_duplicate(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
             [["A1"], ["a1"], ["ITEM"]])
    assert iw.issue_counter == 2


def test_empty_cells_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
             [[None], [""], ["A1"]])
    assert iw.issue_counter == 0


def test_no_duplicates(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory,
             [["A1"], ["A2"], ["A3"]])
    assert iw.issue_counter == 0
