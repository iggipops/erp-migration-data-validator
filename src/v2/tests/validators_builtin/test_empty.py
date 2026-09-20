from validators.builtin import empty


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    # A second, always-filled column keeps blank cells inside the data range —
    # rows with no value in any column are not data rows (F-02).
    sheet_factory(wb, "Items", [["ItemId", "Other"]] + [[*r, "x"] for r in rows])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Empty check", "Severity": "ERROR", "Color": "RED"}
    empty.validate(ws, 1, "ItemId", rule, {"issues_writer": iw})
    return iw


def test_flags_empty_cells(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["A1"], [None], [""]])
    assert iw.issue_counter == 2


def test_no_issues_when_all_filled(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["A1"], ["A2"]])
    assert iw.issue_counter == 0
