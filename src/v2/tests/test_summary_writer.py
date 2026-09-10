from datetime import datetime
from output.issues_writer import IssuesWriter
from output.summary_writer import SummaryWriter


def test_summary_sheet_created_first(make_workbook, fake_config):
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write(
        enrichment_stats=[],
        validation_stats=[],
        external_files_result={"loaded": [], "failed": [], "conversion_warnings": {}},
        external_files_rules=[],
        start_time=datetime(2026, 1, 1, 10, 0, 0),
    )
    assert "Summary" in wb.sheetnames
    assert wb.sheetnames[0] == "Summary"


def test_summary_replaces_existing_summary_sheet(make_workbook, fake_config):
    wb = make_workbook()
    wb.create_sheet("Summary").append(["Stale"])
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write([], [], {"loaded": [], "failed": [], "conversion_warnings": {}}, [],
              start_time=datetime(2026, 1, 1))
    ws = wb["Summary"]
    # first data cell should not be the stale content
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "Stale" not in values


def test_summary_contains_run_info(make_workbook, fake_config):
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write([], [], {"loaded": [], "failed": [], "conversion_warnings": {}}, [],
              start_time=datetime(2026, 1, 1, 10, 0, 0))
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert fake_config.input_file in values
    assert fake_config.output_file in values


def test_summary_counts_issues_by_rule(make_workbook, fake_config):
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    sheet.append(["A2"])
    iw = IssuesWriter(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "Rule", "Severity": "ERROR", "Color": "RED"}
    iw.record(cell=sheet.cell(row=2, column=1), rule=rule, sheet_name="Items",
               column_name="ItemId", row_number=2, cell_content="A1")
    iw.record(cell=sheet.cell(row=3, column=1), rule=rule, sheet_name="Items",
               column_name="ItemId", row_number=3, cell_content="A2")

    validation_stats = [{
        "SequenceNum": 1, "RuleCode": "R1", "RuleName": "Rule",
        "SheetName": "Items", "ColumnName": "ItemId", "Severity": "ERROR",
        "status": "executed", "issue_count": 2,
    }]
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write([], validation_stats, {"loaded": [], "failed": [], "conversion_warnings": {}}, [],
              start_time=datetime(2026, 1, 1))
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "R1" in values
    assert 2 in values


def test_write_metadata_errors_creates_summary_with_errors_listed(make_workbook, fake_config):
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    errors = [
        'ValidationRules row 2: Severity "CRITICAL" must be ERROR or WARNING',
        'EnrichmentRules row 3: CustomFunctionName is required',
    ]
    sw.write_metadata_errors(errors, start_time=datetime(2026, 1, 1, 10, 0, 0))

    assert "Summary" in wb.sheetnames
    assert wb.sheetnames[0] == "Summary"
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert errors[0] in values
    assert errors[1] in values
    assert any("TERMINATED" in str(v) for v in values)


def test_write_metadata_errors_includes_external_files_section(make_workbook, fake_config):
    """FS 8.6: Step 5 (external file import) already ran before Step 6
    terminates — its results must not be silently dropped from the
    Summary sheet written on termination."""
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    ef_result = {
        "loaded": [], "failed": ["'Bad/Sheet' from 'stock.csv': bad sheet name"],
        "failed_sheets": {"Bad/Sheet"}, "conversion_warnings": {},
    }
    ef_rules = [{"SequenceNum": 1, "Active": "Yes", "FilePath": "stock.csv", "SheetName": "Bad/Sheet"}]
    sw.write_metadata_errors(
        ["some metadata error"], start_time=datetime(2026, 1, 1, 10, 0, 0),
        external_files_result=ef_result, external_files_rules=ef_rules,
    )
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "External Files Statistics" in values
    assert "stock.csv" in values
    assert any("failed" in str(v) for v in values)
    assert "some metadata error" in values


def test_write_metadata_errors_omits_external_files_section_when_not_provided(make_workbook, fake_config):
    """Back-compat: callers that don't pass external_files_result (e.g. a
    Step 6 failure encountered before Step 5 ran) get the same Summary as
    before — no External Files section, no crash."""
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write_metadata_errors(["some error"], start_time=datetime(2026, 1, 1))
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "External Files Statistics" not in values


def test_write_metadata_errors_replaces_existing_summary_sheet(make_workbook, fake_config):
    wb = make_workbook()
    wb.create_sheet("Summary").append(["Stale"])
    iw = IssuesWriter(wb, fake_config)
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write_metadata_errors(["some error"], start_time=datetime(2026, 1, 1))
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "Stale" not in values


def test_summary_reports_partial_status_and_rows_unchecked(make_workbook, fake_config):
    """FS section 6.8/8.10: a 'partial' AI rule must be counted distinctly
    from executed/skipped/failed, and the count of never-checked rows
    must be reported alongside the issue count from successful batches."""
    wb = make_workbook()
    sheet = wb.create_sheet("Items")
    sheet.append(["ItemId"])
    sheet.append(["A1"])
    iw = IssuesWriter(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "AI Rule", "Severity": "WARNING", "Color": "YELLOW"}
    iw.record(cell=sheet.cell(row=2, column=1), rule=rule, sheet_name="Items",
              column_name="ItemId", row_number=2, cell_content="A1")

    validation_stats = [{
        "SequenceNum": 1, "RuleCode": "R1", "RuleName": "AI Rule",
        "SheetName": "Items", "ColumnName": "ItemId", "Severity": "WARNING",
        "status": "partial", "issue_count": 1, "rows_unchecked": 3,
    }]
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write([], validation_stats, {"loaded": [], "failed": [], "conversion_warnings": {}}, [],
              start_time=datetime(2026, 1, 1))
    ws = wb["Summary"]
    values = [c.value for row in ws.iter_rows() for c in row if c.value is not None]
    assert "partial" in values
    assert 3 in values  # rows_unchecked
    assert any("partial: 1" in str(v) for v in values)
    # partial rule's issues (from its successful batch) must still be reported
    assert 1 in values


def test_write_does_not_crash_with_failed_rules(make_workbook, fake_config):
    wb = make_workbook()
    iw = IssuesWriter(wb, fake_config)
    enrichment_stats = [{
        "SequenceNum": 1, "RuleCode": "E1", "RuleName": "Enrich",
        "SheetName": "Items", "ColumnName": "New", "status": "failed",
    }]
    sw = SummaryWriter(wb, fake_config, iw)
    sw.write(enrichment_stats, [], {"loaded": [], "failed": ["'X' from 'y.csv': error"], "conversion_warnings": {}},
              external_files_rules=[{"SequenceNum": 1, "Active": "Yes", "FilePath": "y.csv"}],
              start_time=datetime(2026, 1, 1))
    assert "Summary" in wb.sheetnames
