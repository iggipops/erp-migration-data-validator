"""
Shared pytest fixtures for the ERP Migration Data Validator test suite.
"""
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

# Make the project importable when running `pytest` from the repo root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config.config_loader import AppConfig
from utils.logger import setup_logger
from output.issues_writer import IssuesWriter
from registry.function_registry import FunctionRegistry
from validators.registry import ValidationTypeRegistry


@pytest.fixture(autouse=True, scope="session")
def _logger(tmp_path_factory):
    """Ensure the app logger has handlers so no code path crashes on logging."""
    log_dir = tmp_path_factory.mktemp("logs")
    setup_logger(str(log_dir / "test.log"))


@pytest.fixture
def make_workbook():
    """Factory: returns a fresh empty Workbook with the default sheet removed."""
    def _make():
        wb = Workbook()
        wb.remove(wb.active)
        return wb
    return _make


def add_sheet(wb, name, rows):
    """Helper: create a sheet named `name` and append each row in `rows` (first row = headers)."""
    ws = wb.create_sheet(name)
    for row in rows:
        ws.append(row)
    return ws


@pytest.fixture
def sheet_factory():
    return add_sheet


@pytest.fixture
def fake_config(tmp_path):
    """A minimal but real AppConfig, with paths pointed at pytest's tmp_path."""
    return AppConfig(
        input_file=str(tmp_path / "input.xlsx"),
        output_file=str(tmp_path / "output.xlsx"),
        log_file=str(tmp_path / "run.log"),
        custom_functions_directory=str(tmp_path),
        format_special_chars="",
        predefined_colors={},
        ai_enabled=False,
        ai_model=None,
        ai_api_key=None,
        custom_functions={},
        date_validation={},
        numeric_validation={},
        color_map={
            "RED": "FF0000", "GREEN": "00FF00", "YELLOW": "FFFF00",
            "ORANGE": "FFA500", "BLUE": "0000FF", "GREY": "808080",
            "LIGHTRED": "FFCCCC", "LIGHTYELLOW": "FFFFCC",
        },
    )


@pytest.fixture
def issues_writer_factory():
    def _make(wb, config):
        return IssuesWriter(wb, config)
    return _make


@pytest.fixture
def fn_registry():
    reg = FunctionRegistry()
    reg.register_builtins()
    return reg


@pytest.fixture
def vtype_registry():
    reg = ValidationTypeRegistry()
    reg.register_builtins()
    return reg


def make_validation_context(workbook, config, issues_writer, fn_registry=None):
    return {
        "workbook":      workbook,
        "config":        config,
        "issues_writer": issues_writer,
        "fn_registry":   fn_registry,
    }


@pytest.fixture
def validation_context_factory():
    return make_validation_context
