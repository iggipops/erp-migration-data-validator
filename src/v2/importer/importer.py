"""
Import Module — FS section 8.5.

Loads active external files defined in ExternalFiles, optionally applying
per-column ImportSpec conversions (section 5.4), and writes the result into
the output workbook as new worksheets.

This module is dedicated to external file import — kept separate from
WorkbookManager (which handles copy/open/save of the main workbook) per the
architectural split agreed for V2.3.r2+.
"""
from pathlib import Path

import pandas as pd
from openpyxl.workbook import Workbook

from config.config_loader import AppConfig
from importer.conversion import convert_date, convert_numeric, convert_integer
from workbook.excel_utils import append_row_safe, remove_sheet_if_exists
from utils.logger import get_logger


class Importer:

    def __init__(self, workbook: Workbook, config: AppConfig):
        self.workbook = workbook
        self.config   = config
        self.logger   = get_logger()

    def import_external_files(
        self,
        external_files_rules: list[dict],
        import_spec_rules: list[dict],
    ) -> dict[str, list]:
        """
        Import all active external files into the workbook as new worksheets,
        applying ImportSpec conversions where defined (section 8.5.2).

        Returns:
            {
                "loaded": [sheet_name, ...],
                "failed": ["'<sheet>' from '<path>': <reason>", ...],
                "failed_sheets": {sheet_name, ...},
                "conversion_warnings": {sheet_name: count, ...}
            }
        """
        results = {"loaded": [], "failed": [], "failed_sheets": set(), "conversion_warnings": {}}

        active_rules = [
            r for r in external_files_rules
            if str(r.get("Active", "")).strip().upper() == "YES"
        ]
        active_rules.sort(key=lambda r: int(r["SequenceNum"]))

        # Index ImportSpec rows by (lowercased ExternalFileSheetName) for fast lookup
        spec_by_sheet: dict[str, list[dict]] = {}
        for spec_row in import_spec_rules:
            if str(spec_row.get("Active", "")).strip().upper() != "YES":
                continue
            key = str(spec_row.get("ExternalFileSheetName", "")).strip().lower()
            spec_by_sheet.setdefault(key, []).append(spec_row)

        for rule in active_rules:
            sheet_name = str(rule["SheetName"]).strip()
            file_path  = str(rule["FilePath"]).strip()
            file_type  = str(rule.get("Format", "xlsx")).strip().lower()

            self.logger.info(f"Importing external file: {file_path} -> sheet '{sheet_name}'")

            try:
                df = self._load_external_file(rule, file_path, file_type)

                applicable_specs = spec_by_sheet.get(sheet_name.lower(), [])
                warning_count = 0
                if applicable_specs:
                    warning_count = self._apply_import_spec(df, sheet_name, applicable_specs)

                self._write_dataframe_to_sheet(df, sheet_name)
                results["loaded"].append(sheet_name)
                results["conversion_warnings"][sheet_name] = warning_count
                self.logger.info(
                    f"  Imported {len(df)} rows into sheet '{sheet_name}'"
                    + (f" ({warning_count} conversion warning(s))" if warning_count else "")
                )

            except Exception as exc:
                msg = f"'{sheet_name}' from '{file_path}': {exc}"
                results["failed"].append(msg)
                results["failed_sheets"].add(sheet_name)
                self.logger.error(f"  Failed to import external file: {msg}", exc_info=True)

        return results

    # -----------------------------------------------------------------------
    # 8.5.1 Loading
    # -----------------------------------------------------------------------

    def _load_external_file(
        self, rule: dict, file_path: str, file_type: str
    ) -> pd.DataFrame:

        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        if file_type == "xlsx":
            original_sheet = str(rule.get("OriginalSheetName", "")).strip() or None
            df = pd.read_excel(
                file_path,
                sheet_name=original_sheet,
                dtype=str,
                engine="openpyxl",
            )

        elif file_type == "csv":
            raw_delimiter = str(rule.get("CSVDelimiter", ",") or ",").strip()
            delimiter = raw_delimiter.replace('"', '').replace("'", "").strip() or ","
            self.logger.debug(f"  CSV delimiter: '{delimiter}'")
            df = pd.read_csv(file_path, sep=delimiter, dtype=str, engine="python")

        else:
            raise ValueError(f"Unsupported Format: {file_type}")

        # Normalize column names — strip whitespace, preserve original case
        df.columns = [str(c).strip() for c in df.columns]
        return df

    # -----------------------------------------------------------------------
    # 8.5.2 ImportSpec Column Conversion
    # -----------------------------------------------------------------------

    def _apply_import_spec(
        self, df: pd.DataFrame, sheet_name: str, spec_rows: list[dict]
    ) -> int:
        """
        Apply per-column conversions to df in place, per spec_rows.
        Returns the total number of per-cell conversion warnings logged.
        """
        total_warnings = 0

        # Build case-insensitive column lookup
        lower_to_actual = {c.lower(): c for c in df.columns}

        for spec in spec_rows:
            col_name    = str(spec.get("ColumnName", "")).strip()
            target_type = str(spec.get("TargetType", "text")).strip().lower()

            if target_type == "text" or not target_type:
                continue  # no conversion needed

            actual_col = lower_to_actual.get(col_name.lower())
            if actual_col is None:
                self.logger.warning(
                    f"  ImportSpec: column '{col_name}' not found in sheet "
                    f"'{sheet_name}' — skipping conversion, column will "
                    f"remain as text"
                )
                continue

            warnings = self._convert_column(df, actual_col, target_type, spec, sheet_name)
            total_warnings += warnings

        return total_warnings

    def _convert_column(
        self, df: pd.DataFrame, col_name: str, target_type: str,
        spec: dict, sheet_name: str
    ) -> int:

        warning_count = 0

        if target_type == "date":
            formats_raw = str(spec.get("DateFormats", "") or "")
            date_formats = [f.strip() for f in formats_raw.split(";") if f.strip()]
            if not date_formats:
                self.logger.warning(
                    f"  ImportSpec: column '{col_name}' in sheet '{sheet_name}' "
                    f"has TargetType=date but no DateFormats defined — "
                    f"skipping conversion"
                )
                return 0

            new_values = []
            for i, raw_value in enumerate(df[col_name]):
                converted = convert_date(raw_value, date_formats)
                if converted is None and raw_value not in (None, ""):
                    warning_count += 1
                    self.logger.warning(
                        f"  ImportSpec: row {i + 2} column '{col_name}' "
                        f"in sheet '{sheet_name}': value '{raw_value}' does "
                        f"not match any DateFormats — kept as text"
                    )
                    new_values.append(raw_value)
                else:
                    new_values.append(converted if converted is not None else raw_value)
            df[col_name] = new_values

        elif target_type in ("numeric", "integer"):
            decimal_sep   = str(spec.get("DecimalSeparator", ".") or ".")
            thousands_sep = spec.get("ThousandsSeparator") or None

            new_values = []
            for i, raw_value in enumerate(df[col_name]):
                if target_type == "integer":
                    converted = convert_integer(raw_value, decimal_sep, thousands_sep)
                else:
                    converted = convert_numeric(raw_value, decimal_sep, thousands_sep)

                if converted is None and raw_value not in (None, ""):
                    warning_count += 1
                    self.logger.warning(
                        f"  ImportSpec: row {i + 2} column '{col_name}' "
                        f"in sheet '{sheet_name}': value '{raw_value}' could "
                        f"not be converted to {target_type} — kept as text"
                    )
                    new_values.append(raw_value)
                else:
                    new_values.append(converted if converted is not None else raw_value)
            df[col_name] = new_values

        else:
            self.logger.warning(
                f"  ImportSpec: column '{col_name}' in sheet '{sheet_name}' "
                f"has unsupported TargetType '{target_type}' — skipping conversion"
            )

        return warning_count

    # -----------------------------------------------------------------------
    # 8.5.3 Writing to Output Workbook
    # -----------------------------------------------------------------------

    def _write_dataframe_to_sheet(self, df: pd.DataFrame, sheet_name: str) -> None:
        """Write a DataFrame into the workbook as a new sheet (replace if exists)."""
        remove_sheet_if_exists(self.workbook, sheet_name)

        ws = self.workbook.create_sheet(sheet_name)

        # Header row
        ws.append(list(df.columns))

        # Data rows — native Python date/float values are written directly;
        # openpyxl will store them as native Excel date/number cells.
        # append_row_safe guards text values (e.g. "=HYPERLINK(...)" copied
        # verbatim from an external CSV) from being interpreted as formulas.
        for _, row in df.iterrows():
            append_row_safe(ws, [
                None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
                for v in row
            ])
