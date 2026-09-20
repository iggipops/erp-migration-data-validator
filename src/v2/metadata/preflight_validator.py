"""
Preflight Validator — FS section 8.2 (ExternalFiles and ImportSpec
structural validation).

This runs before the input workbook is copied to the output path
(FS 8.3) and before any external file is imported (FS 8.5). All errors
across both sheets are collected and returned together; the caller is
responsible for terminating before step 3 if this returns any errors —
per FS 8.2, processing must not copy or import anything once a
structural error has been found.

Mirrors the collect-everything pattern used by MetadataValidator
(metadata/metadata_validator.py) for the later, step-6 metadata pass.
"""
from pathlib import Path

from openpyxl.workbook import Workbook

from config.config_loader import AppConfig, EXTERNAL_FILES_SHEET, IMPORT_SPEC_SHEET
from workbook.excel_utils import (
    sheet_exists, is_empty, normalize_string,
    get_headers, get_sheet, find_column_index, cell_value, last_data_row,
)
from utils.logger import get_logger


# ---------------------------------------------------------------------------
# Required columns per sheet
# ---------------------------------------------------------------------------

MANDATORY_EXTERNAL_FILES_COLUMNS = [
    "SequenceNum", "SheetName", "FilePath", "Format", "Active",
]

# Columns that may be empty per-row, but whose HEADER must still exist —
# a missing header means every row silently falls back to a default value,
# which can hide a renamed/misspelled column.
CONDITIONAL_EXTERNAL_FILES_COLUMNS = [
    "OriginalSheetName", "CSVDelimiter",
]

MANDATORY_IMPORT_SPEC_COLUMNS = [
    "ExternalFileSheetName", "ColumnName", "TargetType", "Active",
]

CONDITIONAL_IMPORT_SPEC_COLUMNS = [
    "DateFormats", "DecimalSeparator", "ThousandsSeparator",
]

VALID_CSV_DELIMITERS = {",", ";", "|", ":", " "}
VALID_TARGET_TYPES = {"text", "date", "numeric", "integer"}


class PreflightValidator:

    def __init__(self, workbook: Workbook, config: AppConfig):
        self.workbook = workbook
        self.config = config
        self.errors: list[str] = []
        self.logger = get_logger()

    def validate(self) -> list[str]:
        """
        Validate ExternalFiles (8.2.1) and ImportSpec (8.2.2).
        Returns list of error strings (empty = all OK).
        """
        self.errors = []

        active_ef_by_sheet = self._validate_external_files()
        self._validate_import_spec(active_ef_by_sheet)

        if self.errors:
            for err in self.errors:
                self.logger.error(f"Preflight error: {err}")
        else:
            self.logger.info("  Preflight validation passed")

        return self.errors

    # -----------------------------------------------------------------------
    # 8.2.1  ExternalFiles
    # -----------------------------------------------------------------------

    def _validate_external_files(self) -> dict:
        """
        Returns {normalized SheetName: row_dict} for Active=Yes rows —
        used by _validate_import_spec for referential integrity, and by
        the caller to know which external sheets are legitimately active.
        """
        if not sheet_exists(self.workbook, EXTERNAL_FILES_SHEET):
            self.logger.debug(f'Sheet "{EXTERNAL_FILES_SHEET}" not found — skipping (no external files)')
            return {}

        sheet, headers = self._sheet_and_headers(EXTERNAL_FILES_SHEET)

        header_error_count = 0
        for col in MANDATORY_EXTERNAL_FILES_COLUMNS:
            if find_column_index(headers, col) is None:
                self.errors.append(f'{EXTERNAL_FILES_SHEET}: mandatory column "{col}" not found')
                header_error_count += 1
        for col in CONDITIONAL_EXTERNAL_FILES_COLUMNS:
            if find_column_index(headers, col) is None:
                self.errors.append(
                    f'{EXTERNAL_FILES_SHEET}: column "{col}" not found. This column may be '
                    f'left empty on individual rows, but its header must exist in the sheet.'
                )
                header_error_count += 1
        if header_error_count:
            return {}

        all_ef_columns = MANDATORY_EXTERNAL_FILES_COLUMNS + CONDITIONAL_EXTERNAL_FILES_COLUMNS
        existing_sheet_names = {normalize_string(s) for s in self.workbook.sheetnames}

        active_rows = []
        for row in range(2, last_data_row(sheet) + 1):
            r = self._row_dict(sheet, headers, row, all_ef_columns)
            if self._all_empty(r):
                continue

            active_val = r.get("Active")
            if is_empty(active_val):
                self.errors.append(f'{EXTERNAL_FILES_SHEET} row {row}: Active is empty')
                continue
            active = str(active_val).strip().upper()
            if active not in ("YES", "NO"):
                self.errors.append(
                    f'{EXTERNAL_FILES_SHEET} row {row}: Active must be YES or NO, got "{active}"'
                )
                continue
            if active == "NO":
                continue

            active_rows.append((row, r))

        seq_nums = []
        sheet_names_seen = {}
        filepath_key_seen = {}
        by_sheet_name = {}

        for row, r in active_rows:
            # --- SequenceNum ---
            seq_val = r.get("SequenceNum")
            if is_empty(seq_val):
                self.errors.append(f'{EXTERNAL_FILES_SHEET} row {row}: SequenceNum is empty')
            else:
                try:
                    seq_nums.append(int(seq_val))
                except (ValueError, TypeError):
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: SequenceNum must be numeric'
                    )

            # --- SheetName: not empty, unique across active rows, no collision
            #     with a worksheet already present in the input workbook ---
            sheet_name = str(r.get("SheetName", "") or "").strip()
            if is_empty(sheet_name):
                self.errors.append(f'{EXTERNAL_FILES_SHEET} row {row}: SheetName is empty')
            else:
                key = normalize_string(sheet_name)
                if key in sheet_names_seen:
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: SheetName "{sheet_name}" '
                        f'duplicates the SheetName in row {sheet_names_seen[key]} — each '
                        f'active external file must target a unique sheet'
                    )
                else:
                    sheet_names_seen[key] = row
                if key in existing_sheet_names:
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: SheetName "{sheet_name}" already '
                        f'exists in the input workbook and would be overwritten by external '
                        f'file import'
                    )
                by_sheet_name[key] = r

            # --- FilePath: not empty, must exist ---
            file_path = str(r.get("FilePath", "") or "").strip()
            file_path_ok = False
            if is_empty(file_path):
                self.errors.append(f'{EXTERNAL_FILES_SHEET} row {row}: FilePath is empty')
            elif not Path(file_path).exists():
                self.errors.append(
                    f'{EXTERNAL_FILES_SHEET} row {row}: FilePath not found: "{file_path}"'
                )
            else:
                file_path_ok = True

            # --- Format ---
            file_format = str(r.get("Format", "") or "").strip().lower()
            if is_empty(file_format):
                self.errors.append(f'{EXTERNAL_FILES_SHEET} row {row}: Format is empty')
            elif file_format not in ("xlsx", "csv"):
                self.errors.append(
                    f'{EXTERNAL_FILES_SHEET} row {row}: Format must be xlsx or csv, '
                    f'got "{file_format}"'
                )

            # --- OriginalSheetName: required + must exist in the external
            #     file when Format=xlsx; must be empty when Format=csv ---
            original_sheet = str(r.get("OriginalSheetName", "") or "").strip()
            if file_format == "xlsx":
                if is_empty(original_sheet):
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: OriginalSheetName is required '
                        f'when Format = xlsx'
                    )
                elif file_path_ok:
                    ext_sheet_names = self._read_external_xlsx_sheet_names(row, file_path)
                    if ext_sheet_names is not None and normalize_string(original_sheet) not in ext_sheet_names:
                        self.errors.append(
                            f'{EXTERNAL_FILES_SHEET} row {row}: OriginalSheetName '
                            f'"{original_sheet}" not found in external file "{file_path}"'
                        )
            elif file_format == "csv" and not is_empty(original_sheet):
                self.errors.append(
                    f'{EXTERNAL_FILES_SHEET} row {row}: OriginalSheetName must be empty '
                    f'when Format = csv'
                )

            # --- CSVDelimiter: required when Format=csv, must be a supported value ---
            csv_delim = r.get("CSVDelimiter")
            csv_delim_str = str(csv_delim).strip() if csv_delim is not None else ""
            if file_format == "csv":
                if is_empty(csv_delim_str):
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: CSVDelimiter is required '
                        f'when Format = csv'
                    )
                elif csv_delim_str not in VALID_CSV_DELIMITERS:
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: CSVDelimiter "{csv_delim_str}" '
                        f'is not one of the supported delimiters (comma, semicolon, pipe, '
                        f'colon, space)'
                    )

            # --- FilePath + OriginalSheetName uniqueness (FilePath alone for csv) ---
            if file_path:
                if file_format == "csv":
                    fp_key = ("csv", normalize_string(file_path))
                    dup_desc = "FilePath"
                else:
                    fp_key = ("xlsx", normalize_string(file_path), normalize_string(original_sheet))
                    dup_desc = "FilePath + OriginalSheetName"
                if fp_key in filepath_key_seen:
                    self.errors.append(
                        f'{EXTERNAL_FILES_SHEET} row {row}: duplicate {dup_desc} combination '
                        f'— already used in row {filepath_key_seen[fp_key]}'
                    )
                else:
                    filepath_key_seen[fp_key] = row

        self._check_unique(seq_nums, EXTERNAL_FILES_SHEET, "SequenceNum")

        return by_sheet_name

    def _read_external_xlsx_sheet_names(self, row: int, file_path: str):
        """
        Opens the external xlsx file (read-only, no data load) to verify
        OriginalSheetName exists. Returns a set of normalized sheet names,
        or None if the file couldn't be opened (error already recorded).
        """
        try:
            from openpyxl import load_workbook
            ext_wb = load_workbook(file_path, read_only=True)
            try:
                return {normalize_string(s) for s in ext_wb.sheetnames}
            finally:
                ext_wb.close()
        except Exception as exc:
            self.errors.append(
                f'{EXTERNAL_FILES_SHEET} row {row}: could not open external file '
                f'"{file_path}" to verify OriginalSheetName: {exc}'
            )
            return None

    # -----------------------------------------------------------------------
    # 8.2.2  ImportSpec
    # -----------------------------------------------------------------------

    def _validate_import_spec(self, active_ef_by_sheet: dict) -> None:
        if not sheet_exists(self.workbook, IMPORT_SPEC_SHEET):
            self.logger.debug(f'Sheet "{IMPORT_SPEC_SHEET}" not found — treated as empty')
            return

        sheet, headers = self._sheet_and_headers(IMPORT_SPEC_SHEET)

        header_error_count = 0
        for col in MANDATORY_IMPORT_SPEC_COLUMNS:
            if find_column_index(headers, col) is None:
                self.errors.append(f'{IMPORT_SPEC_SHEET}: mandatory column "{col}" not found')
                header_error_count += 1
        for col in CONDITIONAL_IMPORT_SPEC_COLUMNS:
            if find_column_index(headers, col) is None:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET}: column "{col}" not found. This column may be '
                    f'left empty on individual rows, but its header must exist in the sheet.'
                )
                header_error_count += 1
        if header_error_count:
            return

        all_isp_columns = MANDATORY_IMPORT_SPEC_COLUMNS + CONDITIONAL_IMPORT_SPEC_COLUMNS
        seen_keys = {}

        for row in range(2, last_data_row(sheet) + 1):
            r = self._row_dict(sheet, headers, row, all_isp_columns)
            if self._all_empty(r):
                continue

            active_val = r.get("Active")
            if is_empty(active_val):
                self.errors.append(f'{IMPORT_SPEC_SHEET} row {row}: Active is empty')
                continue
            active = str(active_val).strip().upper()
            if active not in ("YES", "NO"):
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: Active must be YES or NO, got "{active}"'
                )
                continue
            if active == "NO":
                continue

            ef_sheet = str(r.get("ExternalFileSheetName", "") or "").strip()
            col_name = str(r.get("ColumnName", "") or "").strip()
            target_type = str(r.get("TargetType", "") or "").strip().lower()

            # --- ExternalFileSheetName: referential integrity to ExternalFiles ---
            if is_empty(ef_sheet):
                self.errors.append(f'{IMPORT_SPEC_SHEET} row {row}: ExternalFileSheetName is empty')
            elif normalize_string(ef_sheet) not in active_ef_by_sheet:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: ExternalFileSheetName "{ef_sheet}" does '
                    f'not reference an Active=Yes row in {EXTERNAL_FILES_SHEET}'
                )
            elif str(active_ef_by_sheet[normalize_string(ef_sheet)].get("Format", "") or "").strip().lower() == "xlsx":
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: ExternalFileSheetName "{ef_sheet}" '
                    f'references an xlsx row in {EXTERNAL_FILES_SHEET} — ImportSpec applies '
                    f'to csv external files only (xlsx files already carry native types)'
                )

            # --- ColumnName: existence in the actual external file is
            #     validated at import time (8.5), not here ---
            if is_empty(col_name):
                self.errors.append(f'{IMPORT_SPEC_SHEET} row {row}: ColumnName is empty')

            # --- TargetType ---
            if is_empty(target_type):
                self.errors.append(f'{IMPORT_SPEC_SHEET} row {row}: TargetType is empty')
            elif target_type not in VALID_TARGET_TYPES:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: TargetType "{target_type}" is not '
                    f'supported (must be text, date, numeric, or integer)'
                )

            # --- DateFormats: required when TargetType = date ---
            if target_type == "date":
                date_formats = str(r.get("DateFormats", "") or "").strip()
                if is_empty(date_formats):
                    self.errors.append(
                        f'{IMPORT_SPEC_SHEET} row {row}: DateFormats is required when '
                        f'TargetType = date'
                    )

            # --- DecimalSeparator / ThousandsSeparator: single char, must differ ---
            decimal_sep = r.get("DecimalSeparator")
            decimal_sep_str = str(decimal_sep).strip() if decimal_sep is not None else ""
            thousands_sep = r.get("ThousandsSeparator")
            thousands_sep_str = str(thousands_sep).strip() if thousands_sep is not None else ""

            if decimal_sep_str and len(decimal_sep_str) != 1:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: DecimalSeparator must be a single '
                    f'character, got "{decimal_sep_str}"'
                )
            if thousands_sep_str and len(thousands_sep_str) != 1:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: ThousandsSeparator must be a single '
                    f'character, got "{thousands_sep_str}"'
                )
            if decimal_sep_str and thousands_sep_str and decimal_sep_str == thousands_sep_str:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: DecimalSeparator and ThousandsSeparator '
                    f'must differ (both are "{decimal_sep_str}")'
                )

            # --- Duplicate ExternalFileSheetName + ColumnName across active rows ---
            key = (normalize_string(ef_sheet), normalize_string(col_name))
            if key in seen_keys:
                self.errors.append(
                    f'{IMPORT_SPEC_SHEET} row {row}: duplicate ExternalFileSheetName '
                    f'"{ef_sheet}" + ColumnName "{col_name}" combination — already used '
                    f'in row {seen_keys[key]}'
                )
            else:
                seen_keys[key] = row

    # -----------------------------------------------------------------------
    # Helpers
    # -----------------------------------------------------------------------

    def _sheet_and_headers(self, sheet_name: str):
        sheet = get_sheet(self.workbook, sheet_name)
        headers = get_headers(sheet)
        return sheet, headers

    def _row_dict(self, sheet, headers: dict, row: int, canonical_columns: list) -> dict:
        """
        Builds the row dict keyed by CANONICAL column name, resolving each
        one to its actual column index case-insensitively (FS section 9).
        This is what makes every downstream r.get("SheetName") etc. work
        correctly regardless of the header's actual case in the file.
        """
        result = {}
        for canon in canonical_columns:
            idx = find_column_index(headers, canon)
            result[canon] = cell_value(sheet.cell(row=row, column=idx)) if idx is not None else None
        return result

    def _all_empty(self, row_data: dict) -> bool:
        return all(is_empty(v) for v in row_data.values())

    def _check_unique(self, values: list, sheet_name: str, col: str) -> None:
        if len(values) != len(set(values)):
            self.errors.append(f'{sheet_name}: duplicate {col} values detected')
