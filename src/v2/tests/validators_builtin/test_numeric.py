from validators.builtin import numeric


def run(wb, sheet_factory, config, issues_writer_factory, rows):
    sheet_factory(wb, "Items", [["Price"]] + rows)
    ws = wb["Items"]
    iw = issues_writer_factory(wb, config)
    rule = {"RuleCode": "R1", "RuleName": "Numeric check", "Severity": "ERROR", "Color": "RED"}
    numeric.validate(ws, 1, "Price", rule, {"issues_writer": iw, "config": config})
    return iw


def test_native_numbers_pass(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[5], [5.5], [-3]])
    assert iw.issue_counter == 0


def test_text_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [["5"]])
    assert iw.issue_counter == 1


def test_bool_fails_regression(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    """Regression test: Excel TRUE/FALSE cells (native Python bool via
    openpyxl) must NOT pass NUMERIC as 1.0/0.0 — bool is a subclass of int."""
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[True], [False]])
    assert iw.issue_counter == 2


def test_empty_excluded(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[None]])
    assert iw.issue_counter == 0


def test_min_max_range(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    fake_config.numeric_validation = {"min_value": 0, "max_value": 100}
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[-1], [50], [101]])
    assert iw.issue_counter == 2


def test_nan_treated_as_empty(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    """is_empty() treats float('nan') as empty per FS 2.5.2 (pandas artifact
    convention), so it's excluded before NUMERIC ever inspects it."""
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[float("nan")]])
    assert iw.issue_counter == 0


def test_inf_fails(make_workbook, sheet_factory, fake_config, issues_writer_factory):
    wb = make_workbook()
    iw = run(wb, sheet_factory, fake_config, issues_writer_factory, [[float("inf")]])
    assert iw.issue_counter == 1
