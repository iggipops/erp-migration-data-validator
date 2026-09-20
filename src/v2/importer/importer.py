"""
Import Module — FS section 8.6.

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
from workbook.excel_utils import (
    append_row_safe, remove_sheet_if_exists, get_sheet, get_values_twin,
    last_data_row, load_workbook_pair, values_sheet, warn_missing_cached_values,
)
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
                warning_count = 0
                if file_type == "xlsx":
                    # xlsx keeps native types and formulas; no ImportSpec (FS 8.6.1)
                    row_count = self._import_xlsx(rule, file_path, sheet_name)
                else:
                    df = self._load_external_file(rule, file_path, file_type)
                    applicable_specs = spec_by_sheet.get(sheet_name.lower(), [])
                    if applicable_specs:
                        warning_count = self._apply_import_spec(df, sheet_name, applicable_specs)
                    self._write_dataframe_to_sheet(df, sheet_name)
                    row_count = len(df)

                results["loaded"].append(sheet_name)
                results["conversion_warnings"][sheet_name] = warning_count
                self.logger.info(
                    f"  Imported {row_count} rows into sheet '{sheet_name}'"
                    + (f" ({warning_count} conversion warning(s))" if warning_count else "")
                )

            except Exception as exc:
                msg = f"'{sheet_name}' from '{file_path}': {exc}"
                results["failed"].append(msg)
                results["failed_sheets"].add(sheet_name)
                self.logger.error(f"  Failed to import external file: {msg}", exc_info=True)

        return results

    # -----------------------------------------------------------------------
    # 8.6.1 Loading
    # -----------------------------------------------------------------------

    def _import_xlsx(self, rule: dict, file_path: str, sheet_name: str) -> int:
        """
        Import one worksheet of an external xlsx file, read the same way as
        the primary workbook (F-01/F-02): the file is loaded twice —
        formulas for the output workbook, cached values (data_only=True)
        for the values twin — and rows stop at the last real value. Cells
        are copied with their native type, formulas included. Returns the
        number of data rows imported.
        """
        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        original_sheet = str(rule.get("OriginalSheetName", "") or "").strip()
        if not original_sheet:
            raise ValueError("OriginalSheetName is required when Format = xlsx")

        src_wb    = load_workbook_pair(file_path)
        src_sheet = get_sheet(src_wb, original_sheet)
        warn_missing_cached_values(
            src_wb, [src_sheet.title], source=f"External file '{file_path}': ")
        src_values = values_sheet(src_sheet)
        last_row   = last_data_row(src_sheet)

        remove_sheet_if_exists(self.workbook, sheet_name)
        ws = self.workbook.create_sheet(sheet_name)
        twin = get_values_twin(self.workbook)
        twin_ws = None
        if twin is not None:
            remove_sheet_if_exists(twin, sheet_name)
            twin_ws = twin.create_sheet(sheet_name)

        for r in range(1, last_row + 1):
            for c in range(1, src_sheet.max_column + 1):
                src = src_sheet.cell(row=r, column=c)
                value = src.value
                if r == 1 and value is not None:
                    value = str(value).strip()  # column names stripped at load (FS 8.6.1)
                if value is None:
                    continue
                dst = ws.cell(row=r, column=c, value=value)
                dst.data_type = src.data_type  # a text "=x" must not become a formula
                dst.number_format = src.number_format
                if twin_ws is not None:
                    cached = value if r == 1 else src_values.cell(row=r, column=c).value
                    if cached is not None:
                        tdst = twin_ws.cell(row=r, column=c, value=cached)
                        tdst.number_format = src.number_format
        return max(last_row - 1, 0)

    def _load_external_file(
        self, rule: dict, file_path: str, file_type: str
    ) -> pd.DataFrame:
        """Load a csv external file as text. (xlsx goes through _import_xlsx.)"""

        if not Path(file_path).exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        if file_type == "csv":
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
    # 8.6.2 ImportSpec Column Conversion (csv files only)
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
    # 8.6.3 Writing to Output Workbook
    # -----------------------------------------------------------------------

    def _write_dataframe_to_sheet(self, df: pd.DataFrame, sheet_name: str) -> None:
        """Write a DataFrame into the workbook as a new sheet (replace if exists)."""
        remove_sheet_if_exists(self.workbook, sheet_name)

        ws = self.workbook.create_sheet(sheet_name)
        twin = get_values_twin(self.workbook)
        twin_ws = None
        if twin is not None:
            remove_sheet_if_exists(twin, sheet_name)
            twin_ws = twin.create_sheet(sheet_name)
        targets = [t for t in (ws, twin_ws) if t is not None]

        # Header row
        for t in targets:
            t.append(list(df.columns))

        # Data rows — native Python date/float values are written directly;
        # openpyxl will store them as native Excel date/number cells.
        # append_row_safe guards text values (e.g. "=HYPERLINK(...)" copied
        # verbatim from an external CSV) from being interpreted as formulas.
        # Rows go to the values twin too, so both workbooks stay in step.
        for _, row in df.iterrows():
            values = [
                None if (v is None or (isinstance(v, float) and pd.isna(v))) else v
                for v in row
            ]
            for t in targets:
                append_row_safe(t, values)
