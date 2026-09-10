from validators.builtin import integer


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["Qty"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Integer check", "Severity": "ERROR", "Color": "RED"}
    integer.validate(ws, 1, "Qty", rule, {"issues_writer": iw, "config": config})
    return iw


def test_whole_numbers_pass(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[5], [5.0], [5.000]])
    assert iw.issue_counter == 0


def test_fractional_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[5.5]])
    assert iw.issue_counter == 1


def test_bool_fails_regression(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[True]])
    assert iw.issue_counter == 1


def test_text_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["5"]])
    assert iw.issue_counter == 1
