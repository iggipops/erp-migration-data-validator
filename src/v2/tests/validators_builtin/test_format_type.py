from validators.builtin import format_type


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["Name"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Format check", "Severity": "WARNING", "Color": "YELLOW"}
    format_type.validate(ws, 1, "Name", rule, {"issues_writer": iw, "config": config})
    return iw


def test_flags_leading_trailing_whitespace(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[" Widget"], ["Widget "]])
    assert iw.issue_counter == 2


def test_flags_double_internal_spaces(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["Big  Widget"]])
    assert iw.issue_counter == 1


def test_flags_configured_special_chars(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    fake_config.format_special_chars = "#@"
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["Widget#1"], ["Widget"]])
    assert iw.issue_counter == 1


def test_clean_values_pass(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["Widget"], ["Big Widget"]])
    assert iw.issue_counter == 0


def test_empty_cells_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[None], [""]])
    assert iw.issue_counter == 0
