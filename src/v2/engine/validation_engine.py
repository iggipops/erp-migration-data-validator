"""
Validation Engine — thin dispatcher per FS section 8.8 (V2.3).

For each active ValidationRules row, looks up the ValidationType in the
Validation Type Registry and delegates entirely to that module's validate()
function. The engine itself contains no type-specific logic.
"""
from openpyxl.workbook import Workbook

from config.config_loader import AppConfig
from registry.function_registry import FunctionRegistry
from validators.registry import ValidationTypeRegistry
from output.issues_writer import IssuesWriter
from utils.ai_client import AIRulePartialError
from workbook.excel_utils import get_headers, get_sheet, find_column_index, normalize_string
from utils.logger import get_logger


class ValidationEngine:

    def __init__(
        self,
        workbook: Workbook,
        config: AppConfig,
        fn_registry: FunctionRegistry,
        type_registry: ValidationTypeRegistry,
        issues_writer: IssuesWriter,
    ):
        self.workbook      = workbook
        self.config        = config
        self.fn_registry   = fn_registry
        self.type_registry = type_registry
        self.issues_writer = issues_writer
        self.logger        = get_logger()

    def execute(
        self,
        rules: list[dict],
        failed_enrichment_columns: set[tuple[str, str]] | None = None,
        failed_sheets: set[str] | None = None,
    ) -> list[dict]:
        """
        Execute active validation rules in SequenceNum order.
        Returns execution statistics per rule.

        `failed_enrichment_columns` is a set of (normalized SheetName,
        normalized ColumnName) tuples — scoped per worksheet, matched
        case-insensitively (FS section 9), per the failed-enrichment
        dependency skip in FS section 2.5.3.
        """
        failed_enrichment_columns = failed_enrichment_columns or set()
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
            vtype     = str(rule.get("ValidationType", "")).strip().upper()

            stat = {
                "SequenceNum": rule.get("SequenceNum", ""),
                "RuleCode":    rule_code,
                "RuleName":    rule_name,
                "SheetName":   str(rule.get("SheetName", "")).strip(),
                "ColumnName":  str(rule.get("ColumnName", "")).strip(),
                "Severity":    str(rule.get("Severity", "")).strip(),
                "status":      None,
                "issue_count": 0,
                "rows_unchecked": 0,
            }

            self.logger.info(
                f"Validation [{rule_code}] {rule_name} — "
                f"{rule.get('SheetName')}.{rule.get('ColumnName')} "
                f"type={vtype}"
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

            ref_data_sheet = str(rule.get("ReferenceDataSheet", "")).strip()
            if vtype == "REFERENCE" and normalize_string(ref_data_sheet) in failed_sheets:
                stat["status"] = "skipped"
                self.logger.warning(
                    f"  [{rule_code}] skipped — depends on ReferenceDataSheet "
                    f"'{ref_data_sheet}' which failed to import"
                )
                stats.append(stat)
                continue

            # --- Dependency skip (FS section 2.5.3) ---
            # ReferenceDataColumn is scoped to ReferenceDataSheet for
            # REFERENCE (the lookup source) but to the rule's own SheetName
            # for DUPLICATE2 (FS 6.4: "second column on the same sheet").
            ref_column = str(rule.get("ReferenceDataColumn", "")).strip()
            dep_column = None
            if (normalize_string(stat["SheetName"]), normalize_string(stat["ColumnName"])) in failed_enrichment_columns:
                dep_column = stat["ColumnName"]
            elif vtype == "REFERENCE" and (normalize_string(ref_data_sheet), normalize_string(ref_column)) in failed_enrichment_columns:
                dep_column = ref_column
            elif vtype == "DUPLICATE2" and (normalize_string(stat["SheetName"]), normalize_string(ref_column)) in failed_enrichment_columns:
                dep_column = ref_column

            if dep_column:
                stat["status"] = "skipped"
                self.logger.warning(
                    f"  [{rule_code}] skipped — depends on column '{dep_column}' "
                    f"which was not created due to a failed enrichment rule"
                )
                stats.append(stat)
                continue

            # --- AI skip when disabled ---
            if vtype == "AI" and not self.config.ai_enabled:
                stat["status"] = "skipped"
                self.logger.info(f"  [{rule_code}] skipped (AI disabled)")
                stats.append(stat)
                continue

            try:
                issue_count    = self._dispatch(rule, vtype)
                stat["status"] = "executed"
                stat["issue_count"] = issue_count
                self.logger.info(f"  [{rule_code}] completed — {issue_count} issue(s)")

            except AIRulePartialError as exc:
                # Some AI batches succeeded, some failed (FS section 6.8) —
                # issues for successful-batch rows are already recorded.
                stat["status"]         = "partial"
                stat["issue_count"]    = exc.issue_count
                stat["rows_unchecked"] = exc.rows_unchecked
                self.logger.warning(f"  [{rule_code}] partial — {exc}")

            except Exception as exc:
                stat["status"] = "failed"
                self.logger.error(f"  [{rule_code}] FAILED: {exc}", exc_info=True)

            stats.append(stat)

        return stats

    def _dispatch(self, rule: dict, vtype: str) -> int:
        """
        Look up the type in the Validation Type Registry and call its
        validate() function. Returns the number of issues recorded.
        """
        validate_fn = self.type_registry.get_validate_fn(vtype)
        if validate_fn is None:
            raise ValueError(
                f"ValidationType '{vtype}' not found in Validation Type Registry"
            )

        sheet_name = str(rule["SheetName"]).strip()
        col_name   = str(rule["ColumnName"]).strip()
        sheet      = get_sheet(self.workbook, sheet_name)
        headers    = get_headers(sheet)
        col_idx    = find_column_index(headers, col_name)

        context = {
            "workbook":     self.workbook,
            "config":       self.config,
            "fn_registry":  self.fn_registry,
            "issues_writer": self.issues_writer,
        }

        before = self.issues_writer.issue_counter
        try:
            validate_fn(sheet, col_idx, col_name, rule, context)
        except AIRulePartialError as exc:
            exc.issue_count = self.issues_writer.issue_counter - before
            raise
        return self.issues_writer.issue_counter - before
