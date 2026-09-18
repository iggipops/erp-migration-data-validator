import pandas as pd
import pytest
from functions.builtin.counter4group import counter4group


def test_counter_resets_per_group():
    df = pd.DataFrame({"Warehouse": ["WH1", "WH1", "WH2", "WH1"], "Item": ["A", "B", "C", "D"]})
    ctx = {"worksheet_data": df, "technical_config": {"grouping_columns": ["Warehouse"]}}
    result = counter4group(ctx)
    assert result == [1, 2, 1, 3]


def test_sort_columns_applied_within_group():
    df = pd.DataFrame({
        "Warehouse": ["WH1", "WH1", "WH1"],
        "Priority":  [3, 1, 2],
    })
    ctx = {
        "worksheet_data": df,
        "technical_config": {"grouping_columns": ["Warehouse"], "sort_columns": ["Priority"]},
    }
    result = counter4group(ctx)
    # Original row order preserved in output, but counter reflects sorted-by-Priority rank
    # row0 (Priority=3) -> rank 3, row1 (Priority=1) -> rank 1, row2 (Priority=2) -> rank 2
    assert result == [3, 1, 2]


def test_missing_grouping_columns_raises():
    df = pd.DataFrame({"A": [1]})
    with pytest.raises(ValueError):
        counter4group({"worksheet_data": df, "technical_config": {}})


def test_unknown_column_raises():
    df = pd.DataFrame({"A": [1]})
    with pytest.raises(ValueError):
        counter4group({"worksheet_data": df, "technical_config": {"grouping_columns": ["NoSuchColumn"]}})


def test_string_param_normalized_to_list():
    df = pd.DataFrame({"Warehouse": ["WH1", "WH1"]})
    ctx = {"worksheet_data": df, "technical_config": {"grouping_columns": "Warehouse"}}
    assert counter4group(ctx) == [1, 2]


def test_counter_values_are_python_ints_not_floats():
    """Regression: cumcount() can promote the result to float64 (e.g. when
    pandas needs a common dtype); the returned counters must always be
    plain ints, matching the documented 1-based integer counter (FS 7.1.3)."""
    df = pd.DataFrame({"Warehouse": ["WH1", "WH1", "WH2", "WH1"], "Item": ["A", "B", "C", "D"]})
    ctx = {"worksheet_data": df, "technical_config": {"grouping_columns": ["Warehouse"]}}
    result = counter4group(ctx)
    assert all(isinstance(v, int) for v in result)
    assert not any(isinstance(v, float) for v in result)
