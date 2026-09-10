from validators.builtin import null


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["Qty"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Null check", "Severity": "ERROR", "Color": "RED"}
    null.validate(ws, 1, "Qty", rule, {"issues_writer": iw})
    return iw


def test_flags_empty_and_zero(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[5], [0], [None], [""]])
    # NULL flags empty AND numeric zero (sole exception to empty-exclusion)
    assert iw.issue_counter == 3


def test_no_issues_for_nonzero(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[5], [-1], [0.5]])
    assert iw.issue_counter == 0
