"""
Validation Type Registry — FS section 2.6.1.

Each built-in validation type is a self-contained module exposing:
    validate(sheet, col_idx, col_name, rule, context) -> None
        Performs the validation, calling context["issues_writer"].record(...)
        for each failing row.

    check_metadata(rule, workbook, config) -> list[str]   [optional]
        Returns a list of error message strings for type-specific metadata
        checks (section 8.6). If not declared, no extra checks are run
        for that type beyond the universal ones in metadata_validator.py.

The registry is fixed at the framework level (not scanned from a directory)
since validation types are a framework capability, not user-extensible,
in V2.3. Adding a new built-in type means adding a new module here and
registering it below — validation_engine.py itself does not change.
"""
from utils.logger import get_logger


class ValidationTypeRegistry:

    def __init__(self):
        self.logger = get_logger()
        self._validate_fns: dict[str, callable] = {}
        self._metadata_check_fns: dict[str, callable] = {}

    def register(self, type_name: str, validate_fn: callable, check_metadata_fn=None) -> None:
        key = type_name.strip().upper()
        self._validate_fns[key] = validate_fn
        if check_metadata_fn is not None:
            self._metadata_check_fns[key] = check_metadata_fn
        self.logger.debug(f"Registered validation type '{key}'")

    def register_builtins(self) -> None:
        from validators.builtin import (
            empty, null, duplicate, duplicate2, format_type,
            reference, custom, ai, date, numeric, integer,
        )

        modules = {
            "EMPTY":      empty,
            "NULL":       null,
            "DUPLICATE":  duplicate,
            "DUPLICATE2": duplicate2,
            "FORMAT":     format_type,
            "REFERENCE":  reference,
            "CUSTOM":     custom,
            "AI":         ai,
            "DATE":       date,
            "NUMERIC":    numeric,
            "INTEGER":    integer,
        }

        for type_name, module in modules.items():
            validate_fn       = getattr(module, "validate")
            check_metadata_fn = getattr(module, "check_metadata", None)
            self.register(type_name, validate_fn, check_metadata_fn)

        self.logger.info(
            f"Validation type registry ready — {len(self._validate_fns)} type(s): "
            f"{sorted(self._validate_fns.keys())}"
        )

    def get_validate_fn(self, type_name: str):
        return self._validate_fns.get(type_name.strip().upper())

    def get_metadata_check_fn(self, type_name: str):
        return self._metadata_check_fns.get(type_name.strip().upper())

    def has(self, type_name: str) -> bool:
        return type_name.strip().upper() in self._validate_fns

    def list_types(self) -> list[str]:
        return sorted(self._validate_fns.keys())
