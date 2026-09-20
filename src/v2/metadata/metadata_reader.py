from openpyxl.workbook import Workbook

from workbook.excel_utils import get_headers, get_sheet, sheet_exists, is_empty, find_column_index, cell_value, last_data_row
from utils.logger import get_logger
from config.config_loader import (
    VALIDATION_RULES_SHEET,
    ENRICHMENT_RULES_SHEET,
    EXTERNAL_FILES_SHEET,
    IMPORT_SPEC_SHEET,
)
from metadata.metadata_validator import (
    MANDATORY_VALIDATION_COLUMNS,
    MANDATORY_ENRICHMENT_COLUMNS,
    CONDITIONAL_ENRICHMENT_COLUMNS,
)
from metadata.preflight_validator import (
    MANDATORY_EXTERNAL_FILES_COLUMNS,
    CONDITIONAL_EXTERNAL_FILES_COLUMNS,
    MANDATORY_IMPORT_SPEC_COLUMNS,
    CONDITIONAL_IMPORT_SPEC_COLUMNS,
)


class MetadataReader:

    def __init__(self, workbook: Workbook):
        self.workbook = workbook
        self.logger   = get_logger()

    def read_validation_rules(self) -> list[dict]:
        return self._read_sheet(VALIDATION_RULES_SHEET, MANDATORY_VALIDATION_COLUMNS)

    def read_enrichment_rules(self) -> list[dict]:
        return self._read_sheet(
            ENRICHMENT_RULES_SHEET,
            MANDATORY_ENRICHMENT_COLUMNS + CONDITIONAL_ENRICHMENT_COLUMNS,
        )

    def read_external_files(self) -> list[dict]:
        return self._read_sheet(
            EXTERNAL_FILES_SHEET,
            MANDATORY_EXTERNAL_FILES_COLUMNS + CONDITIONAL_EXTERNAL_FILES_COLUMNS,
        )

    def read_import_spec(self) -> list[dict]:
        return self._read_sheet(
            IMPORT_SPEC_SHEET,
            MANDATORY_IMPORT_SPEC_COLUMNS + CONDITIONAL_IMPORT_SPEC_COLUMNS,
        )

    # -----------------------------------------------------------------------

    def _read_sheet(self, sheet_name: str, canonical_columns: list[str]) -> list[dict]:
        """
        Read all rows from a metadata sheet into a list of dicts, keyed by
        CANONICAL column name (e.g. "SequenceNum") resolved case-
        insensitively against the sheet's actual headers (FS section 9:
        column names are case-insensitive framework-wide). This is what
        lets every downstream rule.get("SheetName") / rule["ColumnName"]
        work correctly regardless of the header's actual case in the file.

        Any header present in the sheet but not in canonical_columns is
        also kept, under its literal (as-written) name, so unanticipated
        or custom columns aren't silently dropped.

        Returns empty list if sheet does not exist.
        """
        if not sheet_exists(self.workbook, sheet_name):
            self.logger.debug(f"Metadata sheet not found (skipped): {sheet_name}")
            return []

        sheet   = get_sheet(self.workbook, sheet_name)
        headers = get_headers(sheet)

        if not headers:
            self.logger.debug(f"Metadata sheet is empty (skipped): {sheet_name}")
            return []

        canonical_idx = {}
        matched_literal_names = set()
        for canon in canonical_columns:
            idx = find_column_index(headers, canon)
            if idx is not None:
                canonical_idx[canon] = idx
                # remember which literal header text this canonical name matched,
                # so we don't also emit it a second time under its literal name below
                for literal, literal_idx in headers.items():
                    if literal_idx == idx:
                        matched_literal_names.add(literal)

        extra_headers = {
            literal: idx for literal, idx in headers.items()
            if literal not in matched_literal_names
        }

        rows: list[dict] = []

        for row_num in range(2, last_data_row(sheet) + 1):

            row_data = {}
            all_empty = True

            for canon, idx in canonical_idx.items():
                value = cell_value(sheet.cell(row=row_num, column=idx))
                row_data[canon] = value
                if not is_empty(value):
                    all_empty = False

            for literal, idx in extra_headers.items():
                value = cell_value(sheet.cell(row=row_num, column=idx))
                row_data[literal] = value
                if not is_empty(value):
                    all_empty = False

            if all_empty:
                continue  # skip blank trailing rows

            rows.append(row_data)

        self.logger.debug(
            f"Read {len(rows)} rows from sheet '{sheet_name}'"
        )
        return rows
