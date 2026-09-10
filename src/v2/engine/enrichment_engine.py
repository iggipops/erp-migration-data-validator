import pandas as pd
from openpyxl.workbook import Workbook

from config.config_loader import AppConfig
from registry.function_registry import FunctionRegistry
from workbook.excel_utils import get_headers, sheet_to_dataframe, get_sheet, find_column_index, normalize_string
from utils.logger import get_logger


class EnrichmentEngine:

    def __init__(
        self,
        workbook: Workbook,
        config: AppConfig,
        registry: FunctionRegistry,
    ):
        self.workbook = workbook
        self.config   = config
        self.registry = registry
        self.logger   = get_logger()

    def execute(self, rules: list[dict], failed_sheets: set[str] | None = None) -> list[dict]:
        """
        Execute active enrichment rules in SequenceNum order.
        Returns execution statistics per rule.
        """
        failed_sheets = {normalize_string(s) for s in (failed_sheets or set())}

        active_rules = [
            r for r in rules
            if str(r.get("Active", "")).strip().upper() == "YES"
        ]
        active_rules.sort(key=lambda r: int(r["SequenceNum"]))

        stats = []

        for rule in active_rules:
            rule_code = str(rule.get("RuleCode", "")).strip()
            rule_name = str(rule.get("RuleName", "")).strip()
            stat = {
                "SequenceNum": rule.get("SequenceNum", ""),
                "RuleCode":    rule_code,
                "RuleName":    rule_name,
                "SheetName":   str(rule.get("SheetName", "")).strip(),
                "ColumnName":  str(rule.get("TargetColumn", "")).strip(),
                "status":      None,
            }

            self.logger.info(
                f"Enrichment [{rule_code}] {rule_name} — "
                f"{rule.get('SheetName')}.{rule.get('TargetColumn')}"
            )

            # --- Import failure skip (FS section 8.5.4) ---
            if normalize_string(stat["SheetName"]) in failed_sheets:
                stat["status"] = "skipped"
                self.logger.warning(
                    f"  [{rule_code}] skipped — depends on SheetName "
                    f"'{stat['SheetName']}' which failed to import"
                )
                stats.append(stat)
                continue

            try:
                self._execute_rule(rule)
                stat["status"] = "executed"
                self.logger.info(f"  [{rule_code}] completed")

            except Exception as exc:
                stat["status"] = "failed"
                self.logger.error(
                    f"  [{rule_code}] FAILED: {exc}", exc_info=True
                )

            stats.append(stat)

        return stats

    # -----------------------------------------------------------------------

    def _execute_rule(self, rule: dict) -> None:

        sheet_name    = str(rule["SheetName"]).strip()
        target_column = str(rule["TargetColumn"]).strip()
        fn_name       = str(rule["CustomFunctionName"]).strip()

        func = self.registry.get(fn_name)
        if func is None:
            raise ValueError(f"Function '{fn_name}' not found in registry")

        sheet   = get_sheet(self.workbook, sheet_name)
        headers = get_headers(sheet)
        df      = sheet_to_dataframe(sheet)

        context = self._build_context(rule, sheet_name, df)
        result  = func(context)

        if not isinstance(result, (list, tuple)) or len(result) != len(df):
            raise ValueError(
                f"Function '{fn_name}' must return a list with "
                f"{len(df)} elements, got {len(result) if hasattr(result, '__len__') else type(result)}"
            )

        # Write result into target column
        # If column exists — overwrite; if not — create it
        col_idx = find_column_index(headers, target_column)
        if col_idx is None:
            # New column — append after last used column
            col_idx = sheet.max_column + 1
            sheet.cell(row=1, column=col_idx, value=target_column)

        for i, value in enumerate(result):
            sheet.cell(row=i + 2, column=col_idx, value=value)

        # Apply Color (if defined) to header and all data cells in the column
        color_value = str(rule.get("Color", "") or "").strip()
        if color_value:
            fill = self._get_fill(color_value)
            sheet.cell(row=1, column=col_idx).fill = fill
            for i in range(len(result)):
                sheet.cell(row=i + 2, column=col_idx).fill = fill

    def _get_fill(self, color_value: str):
        from openpyxl.styles import PatternFill
        upper = color_value.upper()
        rgb   = self.config.color_map.get(upper, upper)
        return PatternFill(start_color=rgb, end_color=rgb, fill_type="solid")

    def _build_context(
        self, rule: dict, sheet_name: str, df: pd.DataFrame
    ) -> dict:

        fn_name_lower = str(rule.get("CustomFunctionName", "")).strip().lower()

        # Built-in functions: read parameters from FunctionArguments column
        # Custom functions: read parameters from config.yaml
        fn_args_str = str(rule.get("FunctionArguments", "") or "").strip()
        if fn_args_str:
            technical_config = self._parse_function_arguments(fn_args_str)
        else:
            technical_config = self.config.custom_functions.get(fn_name_lower, {})

        # Build named_columns from all columns referenced in technical_config
        named_columns = self._extract_named_columns(technical_config, df)

        return {
            "worksheet_data":   df,
            "named_columns":    named_columns,
            "technical_config": technical_config,
            "workbook":         self.workbook,
            "config":           self.config,
            "row_ids":          list(range(2, len(df) + 2)),
        }

    def _parse_function_arguments(self, args_str: str) -> dict:
        """
        Parse FunctionArguments string into a dict.
        Format: key=value;key=value
        Lists:  key=val1,val2,val3
        Example: columns=ItemName,ItemId;separator=_

        A value wrapped in matching single or double quotes is taken
        literally — quotes protect leading/trailing whitespace (e.g. a
        one-character space separator: separator=" ") and disable comma
        splitting, so a quoted value is always a scalar, never a list.
        """
        result = {}
        if not args_str:
            return result

        for part in args_str.split(";"):
            part = part.strip()
            if "=" not in part:
                continue
            key, _, value = part.partition("=")
            key = key.strip()
            if not key:
                continue
            # Quoted value: take it literally (preserves whitespace, no list split)
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                result[key] = value[1:-1]
                continue
            value = value.strip()
            # If value contains comma — treat as list
            if "," in value:
                result[key] = [v.strip() for v in value.split(",")]
            else:
                result[key] = value

        self.logger.debug(f"  Parsed FunctionArguments: {result}")
        return result

    def _extract_named_columns(
        self, technical_config: dict, df: pd.DataFrame
    ) -> dict[str, list]:
        """
        Build a flat dict of {column_name: [values]} for all column name
        values found anywhere in technical_config.
        Case-insensitive matching: config may use ITEMNAME, actual column is ItemName.
        The key returned uses the ACTUAL column name from the dataframe.
        Also adds the config-specified name as an alias so functions can find it.
        """
        named = {}
        # Build a lowercase -> actual name map from df columns
        lower_to_actual = {col.lower(): col for col in df.columns}

        col_names = self._collect_column_names(technical_config)
        for col in col_names:
            actual = lower_to_actual.get(col.lower())
            if actual:
                # Register under both the actual name and the requested name
                named[actual] = df[actual].tolist()
                if col != actual:
                    named[col] = df[actual].tolist()
        return named

    def _collect_column_names(self, obj) -> list[str]:
        """Recursively collect string values from a config dict/list that match df columns."""
        names = []
        if isinstance(obj, str):
            names.append(obj)
        elif isinstance(obj, list):
            for item in obj:
                names.extend(self._collect_column_names(item))
        elif isinstance(obj, dict):
            for v in obj.values():
                names.extend(self._collect_column_names(v))
        return names
