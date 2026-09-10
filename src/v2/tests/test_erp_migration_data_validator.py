import sys
from pathlib import Path

import erp_migration_data_validator


def test_config_default_resolves_relative_to_script_dir(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "argv", ["erp_migration_data_validator.py"])
    monkeypatch.chdir(tmp_path)

    config_path = erp_migration_data_validator.parse_args()

    expected = Path(erp_migration_data_validator.__file__).resolve().parent / "config.yaml"
    assert Path(config_path) == expected


def test_config_explicit_arg_overrides_default(monkeypatch, tmp_path):
    monkeypatch.setattr(sys, "argv", ["erp_migration_data_validator.py", "--config", "custom.yaml"])

    config_path = erp_migration_data_validator.parse_args()

    assert config_path == "custom.yaml"
