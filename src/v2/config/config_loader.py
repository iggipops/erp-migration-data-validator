import os
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

from utils.logger import get_logger
from utils.parsing import parse_bool


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

SUPPORTED_COLOR_NAMES: dict[str, str] = {
    "RED":         "FF0000",
    "GREEN":       "00FF00",
    "YELLOW":      "FFFF00",
    "ORANGE":      "FFA500",
    "BLUE":        "0000FF",
    "GREY":        "808080",
    "LIGHTRED":    "FFCCCC",
    "LIGHTYELLOW": "FFFFCC",
}

SUPPORTED_VALIDATION_TYPES = [
    "EMPTY",
    "NULL",
    "DUPLICATE",
    "DUPLICATE2",
    "FORMAT",
    "REFERENCE",
    "CUSTOM",
    "AI",
    "DATE",
    "NUMERIC",
    "INTEGER",
]

SUPPORTED_ENRICHMENT_TYPES = [
    "CUSTOM",
]

SUPPORTED_SEVERITIES = [
    "ERROR",
    "WARNING",
]

VALIDATION_RULES_SHEET  = "ValidationRules"
ENRICHMENT_RULES_SHEET  = "EnrichmentRules"
EXTERNAL_FILES_SHEET    = "ExternalFiles"
IMPORT_SPEC_SHEET       = "ImportSpec"
ISSUES_SHEET            = "Issues"
SUMMARY_SHEET           = "Summary"

# FS 5.1 / 8.3.1 — the only encodings a csv external file may declare
# (ExternalFiles.CSVEncoding). Shared by preflight validation and the importer.
SUPPORTED_CSV_ENCODINGS = ("utf-8", "utf-8-sig", "cp1251", "cp1252")


# ---------------------------------------------------------------------------
# Config dataclass
# ---------------------------------------------------------------------------

@dataclass
class AppConfig:

    # --- Framework (mandatory) ---
    input_file:                   str
    output_file:                  str
    log_file:                     str
    custom_functions_directory:   str

    # --- Formatting (optional) ---
    format_special_chars:         str | None              = None
    predefined_colors:            dict[str, str]          = field(default_factory=dict)

    # --- AI (optional) ---
    ai_enabled:                   bool                    = False
    ai_model:                     str | None              = None
    ai_api_key:                   str | None              = None
    ai_connection_test_prompt:    str | None              = None
    ai_provider:                  str | None              = None
    ai_provider_settings:         dict[str, dict[str, Any]] = field(default_factory=dict)

    # --- Provider Registry (runtime, not loaded from YAML — section 2.6.3) ---
    provider_registry:            Any                     = None

    # --- Custom functions (optional) ---
    custom_functions:             dict[str, dict[str, Any]] = field(default_factory=dict)

    # --- Date Validation (optional unless ValidationType=DATE is used) ---
    date_validation:               dict[str, Any]          = field(default_factory=dict)

    # --- Numeric Validation (optional, used by NUMERIC and INTEGER types) ---
    numeric_validation:             dict[str, Any]          = field(default_factory=dict)

    # --- Resolved color map (built-in + predefined, uppercased) ---
    color_map:                    dict[str, str]          = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Loader
# ---------------------------------------------------------------------------

class ConfigLoader:

    def load(self, config_path: str, provider_registry) -> AppConfig:
        """
        Load and validate config.yaml.
        Raises ValueError with all validation errors if any field is invalid.
        Terminates immediately on missing mandatory fields (spec 3.6).

        `provider_registry` is built at Step 1 (FS 8.1), before this method
        is ever called — required here to validate ai_provider (section
        3.4/3.8) and stored on the returned AppConfig so every later AI
        call can reach it through config alone (section 2.6.3).
        """
        logger = get_logger()
        logger.info(f"Loading configuration from: {config_path}")

        path = Path(config_path)
        if not path.exists():
            raise FileNotFoundError(f"Configuration file not found: {config_path}")

        try:
            with open(path, encoding="utf-8") as f:
                raw = yaml.safe_load(f)
        except yaml.YAMLError as exc:
            raise ValueError(
                f"Configuration file contains malformed YAML: {exc}"
            ) from exc
        except OSError as exc:
            raise ValueError(
                f"Configuration file could not be read: {exc}"
            ) from exc

        if raw is None:
            raw = {}
        elif not isinstance(raw, dict):
            raise ValueError(
                "Configuration file must contain a YAML mapping (key: value "
                f"pairs) at the top level, got {type(raw).__name__}"
            )

        errors: list[str] = []

        # --- Framework (mandatory) ---
        input_file                 = self._require(raw, "input_file", errors)
        output_file                = self._require(raw, "output_file", errors)
        log_file                   = self._require(raw, "log_file", errors)
        custom_functions_directory = self._require(raw, "custom_functions_directory", errors)

        if errors:
            raise ValueError(
                "Configuration errors (mandatory fields missing):\n"
                + "\n".join(f"  - {e}" for e in errors)
            )

        # Validate input file exists and is a valid xlsx
        input_path = Path(input_file)
        if not input_path.exists():
            errors.append(f"input_file not found: {input_file}")
        elif input_path.suffix.lower() != ".xlsx":
            errors.append(
                f"input_file must be a valid Excel workbook "
                f"in .xlsx format (MS Excel 2010 - 365 Spreadsheet): {input_file}"
            )

        # Validate custom functions directory exists
        if not Path(custom_functions_directory).exists():
            errors.append(
                f"custom_functions_directory not found: {custom_functions_directory}"
            )

        # Validate output file's directory exists and is writable
        output_dir = Path(output_file).parent
        if not output_dir.exists():
            errors.append(
                f"output_file directory does not exist: {output_dir}"
            )
        elif not os.access(output_dir, os.W_OK):
            errors.append(
                f"output_file directory is not writable: {output_dir}"
            )

        # --- Formatting (optional) ---
        format_special_chars = raw.get("format_special_chars", None)
        if format_special_chars is not None:
            format_special_chars = str(format_special_chars)

        predefined_colors: dict[str, str] = {}
        raw_colors = raw.get("predefined_colors", {}) or {}
        for name, hex_val in raw_colors.items():
            upper_name = str(name).upper()
            upper_hex  = str(hex_val).lstrip("#").upper()
            if len(upper_hex) != 6 or not all(
                c in "0123456789ABCDEF" for c in upper_hex
            ):
                errors.append(
                    f"predefined_colors.{name}: invalid hex value '{hex_val}'"
                )
            else:
                predefined_colors[upper_name] = upper_hex

        # --- AI (optional) ---
        try:
            ai_enabled = parse_bool(raw.get("ai_enabled", False), field_name="ai_enabled")
        except ValueError as exc:
            errors.append(str(exc))
            ai_enabled = False
        ai_model    = raw.get("ai_model", None)
        ai_api_key  = raw.get("ai_api_key", None)

        ai_connection_test_prompt = raw.get("ai_connection_test_prompt", None)
        if ai_connection_test_prompt is not None:
            ai_connection_test_prompt = str(ai_connection_test_prompt)

        ai_provider = raw.get("ai_provider", None)
        if ai_provider is not None:
            ai_provider = str(ai_provider).strip().lower()

        ai_provider_settings: dict[str, dict[str, Any]] = {}
        raw_ps = raw.get("ai_provider_settings", {}) or {}
        for provider_name, settings in raw_ps.items():
            ai_provider_settings[str(provider_name).strip().lower()] = settings or {}

        if ai_enabled:
            if ai_model is None or str(ai_model).strip() == "":
                errors.append("ai_model is required when ai_enabled is true")
            if ai_api_key is None or str(ai_api_key).strip() == "":
                errors.append("ai_api_key is required when ai_enabled is true")
            if ai_provider is None or ai_provider == "":
                errors.append("ai_provider is required when ai_enabled is true")
            elif not provider_registry.has(ai_provider):
                errors.append(
                    f"ai_provider '{ai_provider}' is not a registered provider "
                    f"(available: {provider_registry.list_names()})"
                )
            else:
                required = provider_registry.get_required_settings(ai_provider)
                if required:
                    provided = set(ai_provider_settings.get(ai_provider, {}).keys())
                    missing = [s for s in required if s not in provided]
                    if missing:
                        errors.append(
                            f"ai_provider_settings.{ai_provider} is missing "
                            f"required setting(s): {', '.join(missing)}"
                        )

        # --- Custom functions (optional) ---
        custom_functions: dict[str, dict[str, Any]] = {}
        raw_cf = raw.get("custom_functions", {}) or {}
        for func_name, params in raw_cf.items():
            custom_functions[str(func_name).lower()] = params or {}

        # --- Date Validation (optional unless ValidationType=DATE is used) ---
        date_validation: dict[str, Any] = raw.get("date_validation", {}) or {}
        if date_validation:
            # min_date/max_date/null_dates may arrive from YAML as either a
            # native date/datetime object (unquoted "2025-01-01") or a plain
            # string — normalize all of them to "%Y-%m-%d" strings here so
            # every downstream consumer (this comparison, and the DATE
            # validator) sees one consistent representation.
            for date_field in ("min_date", "max_date"):
                if date_validation.get(date_field) is not None:
                    date_validation[date_field] = self._normalize_date_value(
                        date_validation[date_field], date_field, errors
                    )
            raw_null_dates = date_validation.get("null_dates") or []
            if raw_null_dates:
                date_validation["null_dates"] = [
                    self._normalize_date_value(nd, "null_dates", errors)
                    for nd in raw_null_dates
                ]

            min_d = date_validation.get("min_date")
            max_d = date_validation.get("max_date")
            if min_d and max_d:
                try:
                    if datetime.strptime(min_d, "%Y-%m-%d") > datetime.strptime(max_d, "%Y-%m-%d"):
                        errors.append(
                            f"date_validation: min_date ({min_d}) is after "
                            f"max_date ({max_d})"
                        )
                except ValueError as exc:
                    errors.append(f"date_validation: invalid min_date/max_date format: {exc}")

        # --- Numeric Validation (optional) ---
        numeric_validation: dict[str, Any] = raw.get("numeric_validation", {}) or {}
        if numeric_validation:
            min_v = numeric_validation.get("min_value")
            max_v = numeric_validation.get("max_value")
            if min_v is not None and max_v is not None:
                try:
                    min_v_num = float(min_v)
                    max_v_num = float(max_v)
                except (TypeError, ValueError):
                    errors.append(
                        f"numeric_validation: min_value ({min_v!r}) and max_value "
                        f"({max_v!r}) must both be numeric"
                    )
                else:
                    if min_v_num > max_v_num:
                        errors.append(
                            f"numeric_validation: min_value ({min_v}) is greater than "
                            f"max_value ({max_v})"
                        )
            dec_sep   = numeric_validation.get("decimal_separator", ".")
            thou_sep  = numeric_validation.get("thousands_separator")
            if thou_sep and dec_sep and str(thou_sep) == str(dec_sep):
                errors.append(
                    f"numeric_validation: decimal_separator and thousands_separator "
                    f"must not be the same character ('{dec_sep}')"
                )

        if errors:
            raise ValueError(
                "Configuration validation errors:\n"
                + "\n".join(f"  - {e}" for e in errors)
            )

        # Build unified color map: built-ins + predefined (predefined can override)
        color_map = {**SUPPORTED_COLOR_NAMES, **predefined_colors}

        config = AppConfig(
            input_file=input_file,
            output_file=output_file,
            log_file=log_file,
            custom_functions_directory=custom_functions_directory,
            format_special_chars=format_special_chars,
            predefined_colors=predefined_colors,
            ai_enabled=ai_enabled,
            ai_model=ai_model,
            ai_api_key=ai_api_key,
            ai_connection_test_prompt=ai_connection_test_prompt,
            ai_provider=ai_provider,
            ai_provider_settings=ai_provider_settings,
            provider_registry=provider_registry,
            custom_functions=custom_functions,
            date_validation=date_validation,
            numeric_validation=numeric_validation,
            color_map=color_map,
        )

        logger.info("Configuration loaded successfully")
        logger.debug(f"  input_file:                 {config.input_file}")
        logger.debug(f"  output_file:                {config.output_file}")
        logger.debug(f"  log_file:                   {config.log_file}")
        logger.debug(f"  custom_functions_directory: {config.custom_functions_directory}")
        logger.debug(f"  ai_enabled:                 {config.ai_enabled}")
        logger.debug(f"  ai_provider:                {config.ai_provider}")
        logger.debug(f"  predefined_colors:          {config.predefined_colors}")
        logger.debug(f"  custom_functions blocks:    {list(config.custom_functions.keys())}")

        return config

    # -----------------------------------------------------------------------
    # Helpers
    # -----------------------------------------------------------------------

    def _require(self, raw: dict, key: str, errors: list[str]) -> str | None:
        value = raw.get(key)
        if value is None or str(value).strip() == "":
            errors.append(f"'{key}' is required but missing or empty")
            return None
        return str(value).strip()

    def _normalize_date_value(self, value: Any, field_name: str, errors: list[str]) -> str | None:
        """
        Normalize a date_validation value (min_date/max_date/null_dates entry)
        to a "%Y-%m-%d" string. PyYAML parses an unquoted date like
        1900-01-02 into a native date/datetime object rather than a string;
        accept both so the validator always receives one consistent type.
        """
        if isinstance(value, datetime):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, str):
            try:
                datetime.strptime(value, "%Y-%m-%d")
            except ValueError:
                errors.append(
                    f"date_validation: {field_name} ({value!r}) is not a valid "
                    f"date in YYYY-MM-DD format"
                )
                return None
            return value
        errors.append(
            f"date_validation: {field_name} must be a date or a 'YYYY-MM-DD' "
            f"string, got {type(value).__name__}"
        )
        return None
