from datetime import date

import yaml
import pytest
from openpyxl import Workbook
from config.config_loader import ConfigLoader
from registry.provider_registry import ProviderRegistry

_provider_registry = ProviderRegistry()
_provider_registry.register_builtins()


def write_config(tmp_path, data):
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(data))
    return str(path)


def make_input_xlsx(tmp_path):
    p = tmp_path / "input.xlsx"
    Workbook().save(p)
    return str(p)


def base_config_dict(tmp_path):
    (tmp_path / "custom_functions").mkdir(exist_ok=True)
    return {
        "input_file": make_input_xlsx(tmp_path),
        "output_file": str(tmp_path / "output.xlsx"),
        "log_file": str(tmp_path / "run.log"),
        "custom_functions_directory": str(tmp_path / "custom_functions"),
    }


def test_loads_minimal_valid_config(tmp_path):
    cfg_path = write_config(tmp_path, base_config_dict(tmp_path))
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.input_file.endswith("input.xlsx")
    assert config.color_map["RED"] == "FF0000"


def test_missing_config_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        ConfigLoader().load(str(tmp_path / "does_not_exist.yaml"), _provider_registry)


def test_missing_mandatory_field_raises(tmp_path):
    data = base_config_dict(tmp_path)
    del data["output_file"]
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_input_file_must_be_xlsx(tmp_path):
    data = base_config_dict(tmp_path)
    txt = tmp_path / "input.txt"
    txt.write_text("x")
    data["input_file"] = str(txt)
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_input_file_must_exist(tmp_path):
    data = base_config_dict(tmp_path)
    data["input_file"] = str(tmp_path / "nope.xlsx")
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_custom_functions_directory_must_exist(tmp_path):
    data = base_config_dict(tmp_path)
    data["custom_functions_directory"] = str(tmp_path / "nope")
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_predefined_colors_merged_with_builtins(tmp_path):
    data = base_config_dict(tmp_path)
    data["predefined_colors"] = {"purple": "800080"}
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.color_map["PURPLE"] == "800080"
    assert config.color_map["RED"] == "FF0000"  # builtins preserved


def test_invalid_hex_color_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["predefined_colors"] = {"bad": "ZZZZZZ"}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_custom_functions_keys_lowercased(tmp_path):
    data = base_config_dict(tmp_path)
    data["custom_functions"] = {"DimPolicyValidation": {"a": 1}}
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert "dimpolicyvalidation" in config.custom_functions


def test_date_validation_min_after_max_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["date_validation"] = {"min_date": "2025-12-31", "max_date": "2025-01-01"}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_numeric_validation_min_after_max_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["numeric_validation"] = {"min_value": 100, "max_value": 0}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_numeric_validation_non_numeric_min_value_raises_config_error(tmp_path):
    """A non-numeric min_value/max_value (e.g. a stray text string from
    YAML) must be reported as a normal config validation error, not crash
    the loader with an unhandled TypeError from the min_v > max_v comparison."""
    data = base_config_dict(tmp_path)
    data["numeric_validation"] = {"min_value": "not-a-number", "max_value": 100}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="must both be numeric"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_date_validation_unquoted_yaml_date_normalized_to_string(tmp_path):
    """An unquoted YAML date (e.g. min_date: 1900-01-02) is parsed by
    PyYAML into a native date object, not a string — the loader must
    normalize it to 'YYYY-MM-DD' so it doesn't crash the DATE validator
    later, which expects a string it can strptime()."""
    data = base_config_dict(tmp_path)
    # yaml.safe_dump serializes a real date object the same way PyYAML
    # would parse an unquoted date in a hand-written config.yaml.
    data["date_validation"] = {"min_date": date(1900, 1, 2), "max_date": date(2100, 1, 1)}
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.date_validation["min_date"] == "1900-01-02"
    assert config.date_validation["max_date"] == "2100-01-01"


def test_date_validation_null_dates_as_yaml_date_objects_normalized(tmp_path):
    data = base_config_dict(tmp_path)
    data["date_validation"] = {"null_dates": [date(1900, 1, 1), "1999-12-31"]}
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.date_validation["null_dates"] == ["1900-01-01", "1999-12-31"]


def test_date_validation_min_after_max_detected_with_date_objects(tmp_path):
    data = base_config_dict(tmp_path)
    data["date_validation"] = {"min_date": date(2025, 12, 31), "max_date": date(2025, 1, 1)}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="min_date"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_date_validation_invalid_string_format_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["date_validation"] = {"min_date": "02-01-1900"}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="min_date"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_numeric_validation_same_decimal_and_thousands_sep_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["numeric_validation"] = {"decimal_separator": ",", "thousands_separator": ","}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_malformed_yaml_raises_valueerror(tmp_path):
    cfg_path = tmp_path / "config.yaml"
    cfg_path.write_text("input_file: [unclosed\nbad: : yaml")
    with pytest.raises(ValueError):
        ConfigLoader().load(str(cfg_path), _provider_registry)


def test_yaml_not_a_mapping_raises_valueerror(tmp_path):
    cfg_path = tmp_path / "config.yaml"
    cfg_path.write_text("- item1\n- item2\n")
    with pytest.raises(ValueError):
        ConfigLoader().load(str(cfg_path), _provider_registry)


def test_unreadable_config_file_raises_valueerror(tmp_path, monkeypatch):
    cfg_path = write_config(tmp_path, base_config_dict(tmp_path))

    def fake_open(*args, **kwargs):
        raise PermissionError("Permission denied")

    monkeypatch.setattr("config.config_loader.open", fake_open, raising=False)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_output_file_directory_must_exist(tmp_path):
    data = base_config_dict(tmp_path)
    data["output_file"] = str(tmp_path / "nonexistent_subdir" / "output.xlsx")
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_output_file_directory_not_writable_raises(tmp_path, monkeypatch):
    data = base_config_dict(tmp_path)
    cfg_path = write_config(tmp_path, data)

    monkeypatch.setattr("config.config_loader.os.access", lambda path, mode: False)
    with pytest.raises(ValueError):
        ConfigLoader().load(cfg_path, _provider_registry)


# --- AI config (FS 3.4/3.8) ---

def test_ai_disabled_by_default(tmp_path):
    cfg_path = write_config(tmp_path, base_config_dict(tmp_path))
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_enabled is False
    assert config.ai_model is None
    assert config.ai_api_key is None
    assert config.ai_connection_test_prompt is None
    assert config.ai_provider is None
    assert config.ai_provider_settings == {}


def test_ai_enabled_requires_model(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_api_key"] = "secret"
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="ai_model"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_enabled_requires_api_key(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="ai_api_key"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_enabled_with_model_and_key_loads(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "anthropic"
    data["ai_provider_settings"] = {
        "anthropic": {"ai_endpoint": "https://example.test", "ai_api_version": "v1"}
    }
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_enabled is True
    assert config.ai_model == "gpt-4o"
    assert config.ai_api_key == "secret"


def test_ai_disabled_does_not_require_model_or_key(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = False
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_enabled is False


def test_ai_enabled_quoted_string_false_is_treated_as_false(tmp_path):
    """Regression: a quoted YAML string 'false' must not be coerced via
    Python's bool(), which treats any non-empty string (including "false")
    as truthy."""
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = "false"
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_enabled is False


def test_ai_enabled_quoted_string_true_is_treated_as_true(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = "true"
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "anthropic"
    data["ai_provider_settings"] = {
        "anthropic": {"ai_endpoint": "https://example.test", "ai_api_version": "v1"}
    }
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_enabled is True


def test_ai_enabled_unrecognized_string_raises_config_error(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = "enabled-ish"
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="ai_enabled"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_connection_test_prompt_loaded(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "anthropic"
    data["ai_provider_settings"] = {
        "anthropic": {"ai_endpoint": "https://example.test", "ai_api_version": "v1"}
    }
    data["ai_connection_test_prompt"] = "Respond with OK."
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_connection_test_prompt == "Respond with OK."


# --- AI provider (FS 3.4/3.8, section 2.6.3) ---

def test_ai_enabled_requires_provider(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="ai_provider"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_provider_must_be_registered(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "not-a-real-provider"
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="not a registered provider"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_provider_settings_missing_required_key_raises(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "anthropic"
    data["ai_provider_settings"] = {"anthropic": {"ai_endpoint": "https://example.test"}}
    cfg_path = write_config(tmp_path, data)
    with pytest.raises(ValueError, match="ai_api_version"):
        ConfigLoader().load(cfg_path, _provider_registry)


def test_ai_provider_settings_all_required_keys_present_loads(tmp_path):
    data = base_config_dict(tmp_path)
    data["ai_enabled"] = True
    data["ai_model"] = "gpt-4o"
    data["ai_api_key"] = "secret"
    data["ai_provider"] = "Anthropic"  # case-insensitive, like CustomFunctionName
    data["ai_provider_settings"] = {
        "anthropic": {"ai_endpoint": "https://example.test", "ai_api_version": "v1"}
    }
    cfg_path = write_config(tmp_path, data)
    config = ConfigLoader().load(cfg_path, _provider_registry)
    assert config.ai_provider == "anthropic"
    assert config.ai_provider_settings["anthropic"]["ai_endpoint"] == "https://example.test"
    assert config.provider_registry is _provider_registry
