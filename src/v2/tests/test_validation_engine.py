from engine.validation_engine import ValidationEngine


def make_rule(**overrides):
    rule = {
        "SequenceNum": 1, "RuleCode": "R1", "RuleName": "Rule",
        "SheetName": "Items", "ColumnName": "ItemId",
        "ValidationType": "EMPTY", "Severity": "ERROR", "Color": "RED",
        "Active": "Yes", "ReferenceDataSheet": "", "ReferenceDataColumn": "",
        "CustomFunctionName": "",
    }
    rule.update(overrides)
    return rule


def test_executes_active_rule_and_records_issues(make_workbook, sheet_factory, fake_config,
                                                   issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], [None]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    stats = engine.execute([make_rule()])
    assert stats[0]["status"] == "executed"
    assert stats[0]["issue_count"] == 1


def test_inactive_rule_skipped(make_workbook, sheet_factory, fake_config,
                                issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], [None]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    stats = engine.execute([make_rule(Active="No")])
    assert stats == []


def test_rules_executed_in_sequencenum_order(make_workbook, sheet_factory, fake_config,
                                              issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rules = [make_rule(SequenceNum=2, RuleCode="R2"), make_rule(SequenceNum=1, RuleCode="R1")]
    stats = engine.execute(rules)
    assert [s["RuleCode"] for s in stats] == ["R1", "R2"]


def test_sheet_and_column_resolved_case_insensitively(make_workbook, sheet_factory, fake_config,
                                                        issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], [None]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(SheetName="items", ColumnName="itemid")
    stats = engine.execute([rule])
    assert stats[0]["status"] == "executed"
    assert stats[0]["issue_count"] == 1


def test_unknown_sheet_marks_rule_failed_not_crash(make_workbook, sheet_factory, fake_config,
                                                     issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    stats = engine.execute([make_rule(SheetName="NoSuchSheet")])
    assert stats[0]["status"] == "failed"


def test_rule_depending_on_failed_import_is_skipped_not_failed(make_workbook, sheet_factory, fake_config,
                                                                 issues_writer_factory, fn_registry, vtype_registry):
    """A rule whose SheetName failed to import must be reported distinctly
    as 'skipped' — not folded into 'failed', which would look like the
    rule itself has a real problem (FS 8.5.4)."""
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    stats = engine.execute([make_rule(SheetName="Stock")], failed_sheets={"Stock"})
    assert stats[0]["status"] == "skipped"


def test_reference_rule_depending_on_failed_reference_sheet_is_skipped_not_failed(
        make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry, vtype_registry):
    """A REFERENCE rule whose own SheetName imported fine, but whose
    ReferenceDataSheet failed to import, must be reported as 'skipped' —
    not dispatched to the REFERENCE validator (which would error trying to
    open the missing sheet) and not folded into a generic 'failed' status."""
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(
        ValidationType="REFERENCE", ReferenceDataSheet="Stock", ReferenceDataColumn="Qty",
    )
    stats = engine.execute([rule], failed_sheets={"Stock"})
    assert stats[0]["status"] == "skipped"


def test_ai_rule_skipped_when_ai_disabled(make_workbook, sheet_factory, fake_config,
                                            issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = False
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(ValidationType="AI", CustomFunctionName="unused_fn")
    stats = engine.execute([rule])
    assert stats[0]["status"] == "skipped"


def test_ai_rule_partial_status_when_some_batches_fail(make_workbook, sheet_factory, fake_config,
                                                          issues_writer_factory, fn_registry, vtype_registry):
    """FS section 6.8: some rows verdicted (issues recorded), some rows
    unresolved (batch failed) -> overall rule status is 'partial', not
    'failed', and issues from the successful rows are kept."""
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"], ["A3"]])
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_ai_fn(context):
        return [True, (False, "mismatch"), None]

    fn_registry.register("fake_ai_fn", fake_ai_fn, source="test")
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(ValidationType="AI", CustomFunctionName="fake_ai_fn")
    stats = engine.execute([rule])
    assert stats[0]["status"] == "partial"
    assert stats[0]["issue_count"] == 1
    assert stats[0]["rows_unchecked"] == 1


def test_ai_rule_failed_status_when_all_batches_fail(make_workbook, sheet_factory, fake_config,
                                                       issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"]])
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_ai_fn(context):
        return [None, None]

    fn_registry.register("fake_ai_fn", fake_ai_fn, source="test")
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(ValidationType="AI", CustomFunctionName="fake_ai_fn")
    stats = engine.execute([rule])
    assert stats[0]["status"] == "failed"
    assert stats[0]["issue_count"] == 0


def test_ai_rule_executed_status_when_all_batches_succeed(make_workbook, sheet_factory, fake_config,
                                                            issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"]])
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_ai_fn(context):
        return [True, (False, "bad")]

    fn_registry.register("fake_ai_fn", fake_ai_fn, source="test")
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rule = make_rule(ValidationType="AI", CustomFunctionName="fake_ai_fn")
    stats = engine.execute([rule])
    assert stats[0]["status"] == "executed"
    assert stats[0]["issue_count"] == 1
    assert stats[0]["rows_unchecked"] == 0


def test_unrelated_rule_still_executes_when_other_sheet_failed(make_workbook, sheet_factory, fake_config,
                                                                 issues_writer_factory, fn_registry, vtype_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], [None]])
    iw = issues_writer_factory(wb, fake_config)
    engine = ValidationEngine(wb, fake_config, fn_registry, vtype_registry, iw)
    rules = [
        make_rule(RuleCode="R1", SheetName="Stock", ColumnName="Qty"),
        make_rule(RuleCode="R2", SheetName="Items", ColumnName="ItemId"),
    ]
    stats = engine.execute(rules, failed_sheets={"Stock"})
    by_code = {s["RuleCode"]: s for s in stats}
    assert by_code["R1"]["status"] == "skipped"
    assert by_code["R2"]["status"] == "executed"
    assert by_code["R2"]["issue_count"] == 1
