from __future__ import annotations

import importlib.util
import inspect
from pathlib import Path

from utils.logger import get_logger


class FunctionRegistry:
    """
    Single registry for all functions — built-in and custom alike.
    Built-ins are registered first, then custom directory is scanned.
    Function names are stored and looked up case-insensitively.
    Each function may optionally declare REQUIRED_PARAMS at module level.
    """

    def __init__(self):
        self.logger   = get_logger()
        self._registry: dict[str, callable] = {}
        self._required_params: dict[str, list] = {}
        self._metadata_check_fns: dict[str, callable] = {}

    # -----------------------------------------------------------------------
    # Registration
    # -----------------------------------------------------------------------

    def register(self, name: str, func: callable, source: str = "", required_params: list = None) -> None:
        key = name.strip().lower()
        if key in self._registry:
            self.logger.warning(
                f"Function '{name}' already registered "
                f"(from {source}) — keeping existing, skipping duplicate"
            )
            return
        self._registry[key] = func
        self._required_params[key] = required_params or []
        self.logger.debug(f"Registered function '{name}' [{source}]")

    def register_builtins(self) -> None:
        """Register all built-in functions from functions/builtin/."""
        from functions.builtin.concatenate    import concatenate,   REQUIRED_PARAMS as RP_CONCAT
        from functions.builtin.counter        import counter,       REQUIRED_PARAMS as RP_COUNTER
        from functions.builtin.counter4group  import counter4group, REQUIRED_PARAMS as RP_C4G

        for func, rp in (
            (concatenate,   RP_CONCAT),
            (counter,       RP_COUNTER),
            (counter4group, RP_C4G),
        ):
            self.register(func.__name__, func, source="built-in", required_params=rp)

        self.logger.info(
            f"Built-in functions registered: "
            f"{[f.__name__ for f in (concatenate, counter, counter4group)]}"
        )

    def register_custom_directory(self, directory: str) -> dict:
        """
        Scan directory for .py files, import each, register all callable
        functions found. Returns {"loaded": [...], "failed": [...]} report.
        """
        result = {"loaded": [], "failed": []}
        dir_path = Path(directory)

        py_files = sorted(dir_path.glob("*.py"))

        if not py_files:
            self.logger.info(f"No custom function files found in: {directory}")
            return result

        for py_file in py_files:
            if py_file.name.startswith("_"):
                continue
            try:
                self._load_module(py_file, result)
            except Exception as exc:
                msg = f"{py_file.name}: {exc}"
                result["failed"].append(msg)
                self.logger.error(
                    f"Failed to load custom function file: {msg}",
                    exc_info=True
                )

        self.logger.info(
            f"Custom functions loaded: {result['loaded']}  "
            f"failed: {result['failed']}"
        )
        return result

    # -----------------------------------------------------------------------
    # Lookup
    # -----------------------------------------------------------------------

    def get(self, name: str) -> callable | None:
        return self._registry.get(name.strip().lower())

    def get_required_params(self, name: str) -> list:
        return self._required_params.get(name.strip().lower(), [])

    def get_metadata_check_fn(self, name: str) -> callable | None:
        """Return the optional check_metadata function for this function, if declared."""
        return self._metadata_check_fns.get(name.strip().lower())

    def has(self, name: str) -> bool:
        return name.strip().lower() in self._registry

    def list_names(self) -> list[str]:
        return list(self._registry.keys())

    # -----------------------------------------------------------------------
    # Private
    # -----------------------------------------------------------------------

    def _load_module(self, py_file: Path, result: dict) -> None:
        spec   = importlib.util.spec_from_file_location(py_file.stem, py_file)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)

        # Read optional REQUIRED_PARAMS from module level
        required_params = getattr(module, "REQUIRED_PARAMS", None)

        # Read optional check_metadata function (FS section 2.6.2)
        check_metadata_fn = getattr(module, "check_metadata", None)
        if check_metadata_fn and callable(check_metadata_fn):
            # Will be associated with the primary function name after registration
            pass

        found = []
        for attr_name in dir(module):
            if attr_name.startswith("_") or attr_name == "check_metadata":
                continue
            attr = getattr(module, attr_name)
            # Only register functions this module actually defines. Helpers it
            # imported at module level (e.g. `from workbook.excel_utils import
            # sheet_exists`) are also module-level callables, but they are not
            # custom functions and must never end up in the registry under
            # their own name — comparing __module__ tells the two apart
            # regardless of how the author wrote their imports.
            if (callable(attr) and inspect.isfunction(attr)
                    and attr.__module__ == module.__name__):
                self.register(
                    attr_name, attr,
                    source=py_file.name,
                    required_params=required_params
                )
                # Associate check_metadata with this function name
                if check_metadata_fn:
                    self._metadata_check_fns[attr_name.strip().lower()] = check_metadata_fn
                found.append(attr_name)

        if found:
            result["loaded"].extend(found)
