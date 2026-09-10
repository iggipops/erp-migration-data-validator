import pytest
from functions.custom.duplicatemulti import duplicatemulti


def test_flags_all_occurrences_of_duplicate_combo():
    ctx = {
        "named_columns": {
            "A": ["x", "x", "y"],
            "B": ["1", "1", "1"],
            "C": ["a", "a", "a"],
        },
        "technical_config": {"columns": ["A", "B", "C"]},
    }
    # rows 0,1 have identical (x,1,a); row 2 differs (y,1,a)
    assert duplicatemulti(ctx) == [False, False, True]


def test_no_duplicates():
    ctx = {
        "named_columns": {"A": ["x", "y"], "B": ["1", "2"], "C": ["a", "b"]},
        "technical_config": {"columns": ["A", "B", "C"]},
    }
    assert duplicatemulti(ctx) == [True, True]


def test_case_insensitive_and_whitespace_normalized():
    ctx = {
        "named_columns": {"A": ["X", " x "], "B": ["1", "1"], "C": ["a", "A"]},
        "technical_config": {"columns": ["A", "B", "C"]},
    }
    assert duplicatemulti(ctx) == [False, False]


def test_requires_at_least_three_columns():
    ctx = {
        "named_columns": {"A": ["x"], "B": ["1"]},
        "technical_config": {"columns": ["A", "B"]},
    }
    with pytest.raises(ValueError):
        duplicatemulti(ctx)


def test_unknown_column_raises():
    ctx = {
        "named_columns": {"A": ["x"], "B": ["1"], "C": ["a"]},
        "technical_config": {"columns": ["A", "B", "NoSuchColumn"]},
    }
    with pytest.raises(ValueError):
        duplicatemulti(ctx)


def test_mismatched_row_counts_raises():
    ctx = {
        "named_columns": {"A": ["x", "y"], "B": ["1"], "C": ["a", "b"]},
        "technical_config": {"columns": ["A", "B", "C"]},
    }
    with pytest.raises(ValueError):
        duplicatemulti(ctx)
