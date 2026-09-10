import pytest
from validators.builtin import custom


def test_delegates_to_registered_function(make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"], ["A1"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    fake_config.custom_functions = {"duplicatemulti": {}}

    # Register a trivial custom function directly for this test
    def fake_fn(context):
        # fail every other row
        return [(i % 2 == 0, "bad" if i % 2 else "") for i in range(len(context["worksheet_data"]))]

    fn_registry.register("fake_fn", fake_fn, source="test")

    rule = {"RuleCode": "R1", "RuleName": "Custom", "Severity": "ERROR", "Color": "RED",
            "CustomFunctionName": "fake_fn"}
    custom.validate(ws, 1, "ItemId", rule, {
        "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
    })
    assert iw.issue_counter == 1  # only row index 1 (second row) fails


def test_unregistered_function_raises(make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)
    rule = {"RuleCode": "R1", "RuleName": "Custom", "Severity": "ERROR", "Color": "RED",
            "CustomFunctionName": "does_not_exist"}
    with pytest.raises(ValueError):
        custom.validate(ws, 1, "ItemId", rule, {
            "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
        })


def test_bool_return_treated_as_pass_fail(make_workbook, sheet_factory, fake_config, issues_writer_factory, fn_registry):
    wb = make_workbook()
    sheet_factory(wb, "Items", [["ItemId"], ["A1"], ["A2"]])
    ws = wb["Items"]
    iw = issues_writer_factory(wb, fake_config)

    def fake_fn(context):
        return [True, False]

    fn_registry.register("fake_fn2", fake_fn, source="test")
    rule = {"RuleCode": "R1", "RuleName": "Custom", "Severity": "ERROR", "Color": "RED",
            "CustomFunctionName": "fake_fn2"}
    custom.validate(ws, 1, "ItemId", rule, {
        "issues_writer": iw, "fn_registry": fn_registry, "config": fake_config, "workbook": wb,
    })
    assert iw.issue_counter == 1
