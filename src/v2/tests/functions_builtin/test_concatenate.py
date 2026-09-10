import pytest
from functions.builtin.concatenate import concatenate


def test_basic_concatenation():
    ctx = {
        "named_columns": {"First": ["A", "B"], "Last": ["Smith", "Jones"]},
        "technical_config": {"columns": ["First", "Last"], "separator": " "},
    }
    assert concatenate(ctx) == ["A Smith", "B Jones"]


def test_default_separator_is_empty():
    ctx = {
        "named_columns": {"A": ["x"], "B": ["y"]},
        "technical_config": {"columns": ["A", "B"]},
    }
    assert concatenate(ctx) == ["xy"]


def test_trim_strips_whitespace_by_default():
    ctx = {
        "named_columns": {"A": [" x "], "B": [" y "]},
        "technical_config": {"columns": ["A", "B"], "separator": "-"},
    }
    assert concatenate(ctx) == ["x-y"]


def test_trim_false_preserves_whitespace():
    ctx = {
        "named_columns": {"A": [" x "]},
        "technical_config": {"columns": ["A"], "trim": False},
    }
    assert concatenate(ctx) == [" x "]


def test_none_and_nan_become_empty_string():
    ctx = {
        "named_columns": {"A": [None, float("nan")], "B": ["x", "y"]},
        "technical_config": {"columns": ["A", "B"], "separator": "-"},
    }
    assert concatenate(ctx) == ["-x", "-y"]


def test_single_column_as_string():
    ctx = {
        "named_columns": {"A": ["x", "y"]},
        "technical_config": {"columns": "A"},
    }
    assert concatenate(ctx) == ["x", "y"]


def test_missing_columns_param_raises():
    ctx = {"named_columns": {}, "technical_config": {}}
    with pytest.raises(ValueError):
        concatenate(ctx)


def test_unknown_column_raises():
    ctx = {
        "named_columns": {"A": ["x"]},
        "technical_config": {"columns": ["A", "NoSuchColumn"]},
    }
    with pytest.raises(ValueError):
        concatenate(ctx)
