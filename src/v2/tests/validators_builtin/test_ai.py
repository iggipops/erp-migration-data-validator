import pytest
from utils.ai_client import AIRulePartialError
from validators.builtin import ai


def base_rule(**overrides):
    rule = {"RuleCode": "R1", "RuleName": "AI check", "Severity": "WARNING", "Color": "YELLOW",
            "CustomFunctionName": "fake_ai_fn"}
    rule.update(overrides)
    return rule


def test_soft_skipped_when_ai_disabled(make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = False

    def should_not_be_called(context):
        raise AssertionError("function must not be called when ai_enabled is false")

    fn_registry.register("fake_ai_fn", should_not_be_called, source="test")
    ai.validate(ws, 1, "ItemId", base_rule(), {
        "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
    })
    assert iw.issue_counter == 0


def test_delegates_to_registered_function_like_custom(make_workbook, sheet_factory, fake_config,
                                                         issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_fn(context):
        return [True, (False, "bad fit")]

    fn_registry.register("fake_ai_fn", fake_fn, source="test")
    ai.validate(ws, 1, "ItemId", base_rule(), {
        "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
    })
    assert iw.issue_counter == 1


def test_unregistered_function_raises(make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True
    with pytest.raises(ValueError):
        ai.validate(ws, 1, "ItemId", base_rule(CustomFunctionName="does_not_exist"), {
            "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
        })


def test_all_rows_unresolved_raises_plain_exception_no_issues(make_workbook, sheet_factory, fake_config,
                                                                 issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_fn(context):
        return [None, None]  # every batch failed

    fn_registry.register("fake_ai_fn", fake_fn, source="test")
    with pytest.raises(Exception) as exc_info:
        ai.validate(ws, 1, "ItemId", base_rule(), {
            "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
        })
    assert not isinstance(exc_info.value, AIRulePartialError)
    assert iw.issue_counter == 0


def test_some_rows_unresolved_raises_partial_error_after_recording_issues(
        make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"], ["A3"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.ai_enabled = True

    def fake_fn(context):
        # row 1 passes, row 2 fails (recorded), row 3 unresolved (batch failed)
        return [True, (False, "mismatch"), None]

    fn_registry.register("fake_ai_fn", fake_fn, source="test")
    with pytest.raises(AIRulePartialError) as exc_info:
        ai.validate(ws, 1, "ItemId", base_rule(), {
            "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
        })
    assert iw.issue_counter == 1  # the failing verdicted row was recorded
    assert exc_info.value.rows_unchecked == 1
