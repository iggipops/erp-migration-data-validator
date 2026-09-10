import pandas as pd
import pytest
from functions.builtin.counter import counter


def test_sequential_counter():
    ctx = {"worksheet_data": pd.DataFrame({"A": ["x", "y", "z"]})}
    assert counter(ctx) == [1, 2, 3]


def test_empty_dataframe():
    ctx = {"worksheet_data": pd.DataFrame({"A": []})}
    assert counter(ctx) == []


def test_missing_worksheet_data_raises():
    with pytest.raises(ValueError):
        counter({})
