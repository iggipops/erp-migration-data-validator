from output.issues_writer import IssuesWriter, ISSUES_HEADERS


def test_initializes_issues_sheet(make_workbook, fake_config):
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    ws = wb["Issues"]
    assert [c.value for c in ws[1]] == ISSUES_HEADERS


def test_replaces_existing_issues_sheet(make_workbook, fake_config):
    wb = make_workbook()
    wb.create_sheet("Issues").append(["Stale"])
    IssuesWriter(wb, fake_config)
    ws = wb["Issues"]
    assert ws.cell(row=1, column=1).value == "IssueNum"


def test_record_appends_row_and_increments_counters(make_workbook, fake_config):
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    iw = IssuesWriter(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "Rule", "Severity": "ERROR", "Color": "RED"}
    cell = sheet.cell(row=2, column=1)
    iw.record(cell=cell, rule=rule, sheet_name="Items", column_name="ItemId",
               row_number=2, cell_content="A1")
    assert iw.issue_counter == 1
    assert iw.error_counter == 1
    assert iw.warning_counter == 0
    issues_ws = wb["Issues"]
    assert issues_ws.cell(row=2, column=1).value == 1
    assert issues_ws.cell(row=2, column=6).value == "R1"


def test_warning_severity_increments_warning_counter(make_workbook, fake_config):
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    iw = IssuesWriter(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "Rule", "Severity": "WARNING", "Color": "YELLOW"}
    cell = sheet.cell(row=2, column=1)
    iw.record(cell=cell, rule=rule, sheet_name="Items", column_name="ItemId",
               row_number=2, cell_content="A1")
    assert iw.warning_counter == 1
    assert iw.error_counter == 0


def test_record_highlights_cell(make_workbook, fake_config):
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    iw = IssuesWriter(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "Rule", "Severity": "ERROR", "Color": "RED"}
    cell = sheet.cell(row=2, column=1)
    iw.record(cell=cell, rule=rule, sheet_name="Items", column_name="ItemId",
               row_number=2, cell_content="A1")
    assert cell.fill.start_color.rgb.endswith("FF0000")
