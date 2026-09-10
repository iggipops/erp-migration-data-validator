import pytest

from workbook.excel_utils import sheet_to_dataframe
from utils.ai_client import AIClientError
from functions.custom import ai_itemcategoryfit as aif
from functions.custom.ai_itemcategoryfit import ai_itemcategoryfit, check_metadata


def build_workbook(make_workbook, sheet_factory):
    wb = make_workbook()
    sheet_factory(wb, "ItemData", [
        ["CategoryId", "Item Name"],
        ["C1", "Steel Bolt"],       # matched -> sent to AI
        ["C9", "Mystery Widget"],   # key not found in CategoryMaster -> skipped
        [None, "No Category Item"],  # empty key -> skipped
    ])
    sheet_factory(wb, "CategoryMaster", [
        ["CategoryId", "Category Description"],
        ["C1", "Widgets"],
        ["C2", "Fasteners"],
    ])
    return wb


BASE_CFG = {
    "prompt": "Evaluate whether the item name is relevant to the category.",
    "source_sheet": "ItemData",
    "source_key_column": "CategoryId",
    "source_text_column": "Item Name",
    "reference_sheet": "CategoryMaster",
    "reference_key_column": "CategoryId",
    "reference_text_column": "Category Description",
    "batch_size": 20,
}


def run(make_workbook, sheet_factory, cfg=None, config=None):
    wb = build_workbook(make_workbook, sheet_factory)
    df = sheet_to_dataframe(wb["ItemData"])
    context = {
        "worksheet_data": df,
        "workbook": wb,
        "technical_config": cfg if cfg is not None else BASE_CFG,
        "config": config if config is not None else object(),
    }
    return ai_itemcategoryfit(context)


def test_empty_key_row_skipped_no_issue(make_workbook, sheet_factory, monkeypatch):
    def fake_send_batch(config, prompt, payloads, reference_context=None):
        return [(True, "") for _ in payloads]
    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)

    results = run(make_workbook, sheet_factory)
    assert results[2] is True  # empty CategoryId row


def test_unmatched_key_row_skipped_no_issue(make_workbook, sheet_factory, monkeypatch):
    def fake_send_batch(config, prompt, payloads, reference_context=None):
        return [(True, "") for _ in payloads]
    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)

    results = run(make_workbook, sheet_factory)
    assert results[1] is True  # CategoryId C9 not in CategoryMaster


def test_matched_row_passes_when_ai_verdict_is_pass(make_workbook, sheet_factory, monkeypatch):
    def fake_send_batch(config, prompt, payloads, reference_context=None):
        return [(True, "")]
    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)

    results = run(make_workbook, sheet_factory)
    assert results[0] is True


def test_matched_row_fails_with_comment_passthrough(make_workbook, sheet_factory, monkeypatch):
    def fake_send_batch(config, prompt, payloads, reference_context=None):
        return [(False, "Item name doesn't match Widgets — closer to Fasteners.")]
    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)

    results = run(make_workbook, sheet_factory)
    assert results[0] == (False, "Item name doesn't match Widgets — closer to Fasteners.")


def test_full_reference_text_list_sent_once_per_batch_never_the_key_column(
        make_workbook, sheet_factory, monkeypatch):
    captured = {}

    def fake_send_batch(config, prompt, payloads, reference_context=None):
        captured["payloads"] = payloads
        captured["reference_context"] = reference_context
        return [(True, "") for _ in payloads]

    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)
    run(make_workbook, sheet_factory)

    assert set(captured["reference_context"]) == {"Widgets", "Fasteners"}
    assert "C1" not in captured["reference_context"]
    assert "C2" not in captured["reference_context"]
    assert captured["payloads"] == [{"item_text": "Steel Bolt", "assigned_category_text": "Widgets"}]


def test_batch_failure_leaves_rows_unresolved(make_workbook, sheet_factory, monkeypatch):
    def fake_send_batch(config, prompt, payloads, reference_context=None):
        raise AIClientError("service unavailable")

    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)
    results = run(make_workbook, sheet_factory)
    assert results[0] is None  # matched row's batch failed -> unresolved
    assert results[1] is True  # unmatched key row still passes (never sent to AI)
    assert results[2] is True  # empty key row still passes (never sent to AI)


def test_batch_size_splits_matched_rows_across_multiple_calls(make_workbook, sheet_factory, monkeypatch):
    wb = make_workbook()
    sheet_factory(wb, "ItemData", [
        ["CategoryId", "Item Name"],
        ["C1", "Item A"],
        ["C1", "Item B"],
        ["C1", "Item C"],
    ])
    sheet_factory(wb, "CategoryMaster", [["CategoryId", "Category Description"], ["C1", "Widgets"]])
    df = sheet_to_dataframe(wb["ItemData"])

    calls = []

    def fake_send_batch(config, prompt, payloads, reference_context=None):
        calls.append(len(payloads))
        return [(True, "") for _ in payloads]

    monkeypatch.setattr(aif.ai_client, "send_batch", fake_send_batch)
    cfg = dict(BASE_CFG)
    cfg["batch_size"] = 2
    context = {"worksheet_data": df, "workbook": wb, "technical_config": cfg, "config": object()}
    results = ai_itemcategoryfit(context)
    assert calls == [2, 1]
    assert results == [True, True, True]


def test_missing_required_param_raises(make_workbook, sheet_factory):
    cfg = dict(BASE_CFG)
    del cfg["prompt"]
    with pytest.raises(ValueError):
        run(make_workbook, sheet_factory, cfg=cfg)


def test_unknown_reference_sheet_raises(make_workbook, sheet_factory):
    cfg = dict(BASE_CFG)
    cfg["reference_sheet"] = "NoSuchSheet"
    with pytest.raises(ValueError):
        run(make_workbook, sheet_factory, cfg=cfg)


# --- check_metadata (FS section 2.6.2 / Appendix C.3) ---

def make_config(custom_functions):
    class FakeConfig:
        pass
    c = FakeConfig()
    c.custom_functions = {"ai_itemcategoryfit": custom_functions}
    return c


def test_check_metadata_source_sheet_must_match_sheetname(make_workbook, sheet_factory):
    wb = build_workbook(make_workbook, sheet_factory)
    rule = {"SheetName": "ItemData"}
    cfg = dict(BASE_CFG)
    cfg["source_sheet"] = "SomeOtherSheet"
    errors = check_metadata(rule, wb, make_config(cfg))
    assert any("source_sheet" in e and "SheetName" in e for e in errors)


def test_check_metadata_matching_source_sheet_passes(make_workbook, sheet_factory):
    wb = build_workbook(make_workbook, sheet_factory)
    rule = {"SheetName": "ItemData"}
    errors = check_metadata(rule, wb, make_config(BASE_CFG))
    assert errors == []


def test_check_metadata_missing_reference_sheet(make_workbook, sheet_factory):
    wb = build_workbook(make_workbook, sheet_factory)
    rule = {"SheetName": "ItemData"}
    cfg = dict(BASE_CFG)
    cfg["reference_sheet"] = "NoSuchSheet"
    errors = check_metadata(rule, wb, make_config(cfg))
    assert any("NoSuchSheet" in e for e in errors)


def test_check_metadata_missing_reference_text_column(make_workbook, sheet_factory):
    wb = build_workbook(make_workbook, sheet_factory)
    rule = {"SheetName": "ItemData"}
    cfg = dict(BASE_CFG)
    cfg["reference_text_column"] = "NoSuchColumn"
    errors = check_metadata(rule, wb, make_config(cfg))
    assert any("NoSuchColumn" in e for e in errors)


def test_check_metadata_invalid_batch_size(make_workbook, sheet_factory):
    wb = build_workbook(make_workbook, sheet_factory)
    rule = {"SheetName": "ItemData"}
    cfg = dict(BASE_CFG)
    cfg["batch_size"] = 0
    errors = check_metadata(rule, wb, make_config(cfg))
    assert any("batch_size" in e for e in errors)
