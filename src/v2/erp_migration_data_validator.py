"""
ERP Migration Data Validator V2
Entry point — executes the full processing sequence per spec section 8.
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

from __version__ import __version__


def parse_args() -> str:
    # FS 3.1: if --config is omitted, look for config.yaml next to the
    # executable itself, not in the current working directory.
    default_config = str(Path(__file__).resolve().parent / "config.yaml")
    parser = argparse.ArgumentParser(
        description="ERP Migration Data Validator V2"
    )
    parser.add_argument(
        "--config",
        default=default_config,
        help="Path to YAML configuration file (default: config.yaml in the executable's directory)"
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}"
    )
    return parser.parse_args().config


def main() -> int:
    """
    Full execution sequence. Step numbers below are plain sequential
    numbers matching actual execution order in this function — they are
    NOT the same as FS section 2.4's list or FS section 8's subsection
    numbers, which disagree with each other on step order (2.4 places
    external file import before registry build; section 8 does the
    opposite, which is what this code follows — see notes on step 5/6).
    Console and log output use these same plain numbers.

    1  Build AI Provider Registry                          (FS 8.1)
    2  Load and validate technical configuration            (FS 8.2)
    3  Validate ExternalFiles and ImportSpec metadata        (FS 8.3)
    4  Copy input workbook to output path                    (FS 8.4)
    5  Build function registry and validation type registry (FS 8.5)
    6  Import external files as worksheets                  (FS 8.6)
    7  Validate ValidationRules and EnrichmentRules metadata (FS 8.7)
    8  Execute enrichments in SequenceNum order              (FS 8.8)
    9  Execute validations in SequenceNum order              (FS 8.9)
    10 Format Issues worksheet                               (FS 8.10)
    11 Generate Summary worksheet                            (FS 8.11)
    12 Save workbook and log completion
    """

    start_time = datetime.now()
    config_path = parse_args()

    # -----------------------------------------------------------------------
    # 1  Build AI Provider Registry (FS 8.1)
    # -----------------------------------------------------------------------
    # No log file path is known yet — it lives inside the config we're
    # about to read. Console-only logging until then; everything logged
    # in the meantime is buffered and replayed into the real log file
    # once config.log_file is known (FS 8.12.2).
    from utils.logger import setup_logger, get_logger
    bootstrap_logger = setup_logger()
    bootstrap_logger.info("=" * 60)
    bootstrap_logger.info(f"ERP Migration Data Validator V2 — starting (version {__version__})")
    bootstrap_logger.info(f"Config: {config_path}")

    bootstrap_logger.info("Step 1 — Building AI Provider Registry")
    from registry.provider_registry import ProviderRegistry
    provider_registry = ProviderRegistry()
    provider_registry.register_builtins()

    # -----------------------------------------------------------------------
    # 2  Load and validate technical configuration (FS 8.2)
    # -----------------------------------------------------------------------
    bootstrap_logger.info("Step 2 — Loading and validating technical configuration")
    from config.config_loader import ConfigLoader
    try:
        config = ConfigLoader().load(config_path, provider_registry)
    except (FileNotFoundError, ValueError) as exc:
        bootstrap_logger.error(f"Configuration error: {exc}")
        print(f"\nConfiguration error:\n{exc}", file=sys.stderr)
        return 1

    # Re-setup logger with the correct log file from config — replays
    # everything buffered above into it first (FS 8.12.2).
    setup_logger(config.log_file)
    logger = get_logger()
    logger.info("  Configuration loaded successfully")

    # --- AI connectivity pre-flight check (FS 3.4/3.8/8.2.1) ---
    if config.ai_enabled and config.ai_connection_test_prompt:
        logger.info("  Testing AI connectivity (ai_connection_test_prompt)")
        from utils.ai_client import AIClientError, test_connection
        try:
            test_connection(config, config.ai_connection_test_prompt)
            logger.info("  AI connectivity check passed")
        except AIClientError as exc:
            logger.error(f"AI connectivity check failed: {exc}")
            print(f"\nConfiguration error: AI connectivity check failed:\n{exc}", file=sys.stderr)
            return 1

    # -----------------------------------------------------------------------
    # 3  ExternalFiles and ImportSpec metadata — structural checks (FS 8.3),
    #    before touching the filesystem (no copy, no import yet)
    # -----------------------------------------------------------------------
    from workbook.workbook_manager import WorkbookManager
    from metadata.metadata_reader  import MetadataReader

    logger.info("Step 3 — Validating ExternalFiles and ImportSpec metadata (structural)")
    try:
        # We need to open the source workbook briefly to read ExternalFiles/ImportSpec
        from workbook.excel_utils import open_workbook
        source_wb      = open_workbook(config.input_file)
        reader_source  = MetadataReader(source_wb)
        external_files = reader_source.read_external_files()
        import_spec    = reader_source.read_import_spec()
        logger.info(
            f"  Found {len(external_files)} external file rule(s), "
            f"{len(import_spec)} ImportSpec rule(s)"
        )

        from metadata.preflight_validator import PreflightValidator
        preflight = PreflightValidator(source_wb, config)
        preflight_errors = preflight.validate()

        if preflight_errors:
            print(
                "\nExternalFiles/ImportSpec metadata errors:\n"
                + "\n".join(f"  - {e}" for e in preflight_errors),
                file=sys.stderr,
            )
            return 1

    except Exception as exc:
        logger.error(f"ExternalFiles/ImportSpec metadata read failed: {exc}", exc_info=True)
        return 1

    # -----------------------------------------------------------------------
    # 4  Copy input workbook to output path (FS 8.4)
    # -----------------------------------------------------------------------
    logger.info("Step 4 — Copying input workbook to output path")
    wb_manager = WorkbookManager(config)
    try:
        workbook = wb_manager.copy_and_open()
    except Exception as exc:
        logger.error(f"Failed to copy/open workbook: {exc}", exc_info=True)
        return 1

    # -----------------------------------------------------------------------
    # 5  Build function registry and validation type registry (FS 8.5)
    # -----------------------------------------------------------------------
    logger.info("Step 5 — Building function registry and validation type registry")

    from registry.function_registry import FunctionRegistry
    registry = FunctionRegistry()
    registry.register_builtins()
    load_result = registry.register_custom_directory(
        config.custom_functions_directory
    )
    if load_result["failed"]:
        for msg in load_result["failed"]:
            logger.warning(f"  Custom function load failed: {msg}")
    logger.info(
        f"  Function registry ready — {len(registry.list_names())} function(s): "
        f"{registry.list_names()}"
    )

    from validators.registry import ValidationTypeRegistry
    vtype_registry = ValidationTypeRegistry()
    vtype_registry.register_builtins()
    logger.info(
        f"  Validation type registry ready — {len(vtype_registry.list_types())} type(s)"
    )

    # -----------------------------------------------------------------------
    # 6  Import external files as worksheets (FS 8.6, Import Module)
    # -----------------------------------------------------------------------
    logger.info("Step 6 — Importing external files")
    from importer.importer import Importer
    ef_result = {"loaded": [], "failed": [], "failed_sheets": set(), "conversion_warnings": {}}
    if external_files:
        importer_instance = Importer(workbook, config)
        ef_result = importer_instance.import_external_files(external_files, import_spec)
        if ef_result["failed"]:
            for msg in ef_result["failed"]:
                logger.error(f"  External file import failed: {msg}")
            # Failed imports are logged; we continue — validation/enrichment
            # rules that depend on a failed SheetName are skipped individually
            # (FS 8.6.4) rather than failing metadata validation for everyone
        total_warnings = sum(ef_result.get("conversion_warnings", {}).values())
        if total_warnings:
            logger.warning(
                f"  ImportSpec conversion warnings: {total_warnings} cell(s) "
                f"could not be converted and were kept as text"
            )
    else:
        logger.info("  No active external file rules — skipping")

    # -----------------------------------------------------------------------
    # 7  Validate ValidationRules and EnrichmentRules metadata (FS 8.7)
    # -----------------------------------------------------------------------
    logger.info("Step 7 — Validating rules metadata")
    reader = MetadataReader(workbook)
    validation_rules  = reader.read_validation_rules()
    enrichment_rules  = reader.read_enrichment_rules()

    from metadata.metadata_validator import MetadataValidator
    mv = MetadataValidator(
        workbook, config, registry, vtype_registry,
        failed_sheets=ef_result.get("failed_sheets", set()),
    )
    errors = mv.validate()

    if errors:
        logger.error("Metadata validation failed. Terminating.")
        print("\nMetadata validation errors:", file=sys.stderr)
        for err in errors:
            print(f"  - {err}", file=sys.stderr)

        # FS 8.7: errors are written to the Summary worksheet before
        # terminating, so the output file explains why it stopped here.
        from output.issues_writer import IssuesWriter
        from output.summary_writer import SummaryWriter
        issues_writer  = IssuesWriter(workbook, config)
        summary_writer = SummaryWriter(workbook, config, issues_writer)
        summary_writer.write_metadata_errors(
            errors, start_time=start_time,
            external_files_result=ef_result,
            external_files_rules=external_files,
        )
        try:
            wb_manager.save()
        except Exception as exc:
            logger.error(f"Failed to save workbook: {exc}", exc_info=True)

        return 1

    # -----------------------------------------------------------------------
    # 8  Execute enrichments in SequenceNum order (FS 8.8)
    # -----------------------------------------------------------------------
    from output.issues_writer   import IssuesWriter
    from engine.enrichment_engine import EnrichmentEngine

    logger.info("Step 8 — Executing enrichment rules")
    issues_writer     = IssuesWriter(workbook, config)
    enrichment_engine = EnrichmentEngine(workbook, config, registry)
    enrichment_stats  = enrichment_engine.execute(
        enrichment_rules, failed_sheets=ef_result.get("failed_sheets", set())
    )

    # -----------------------------------------------------------------------
    # 9  Execute validations in SequenceNum order (FS 8.9)
    # -----------------------------------------------------------------------
    from engine.validation_engine import ValidationEngine
    from workbook.excel_utils import normalize_string

    # Collect (SheetName, TargetColumn) of enrichment rules that failed at
    # runtime (FS 2.5.3), scoped per worksheet and matched case-insensitively
    # (FS section 9 — column/worksheet names are case-insensitive framework-
    # wide) so a column name that happens to collide with an unrelated
    # column of the same name on a different sheet isn't treated as a
    # dependency.
    failed_enrichment_columns = {
        (normalize_string(stat["SheetName"]), normalize_string(stat["ColumnName"]))
        for stat in enrichment_stats
        if stat["status"] == "failed" and stat["ColumnName"]
    }
    if failed_enrichment_columns:
        logger.warning(
            f"Failed enrichment target columns (dependent validations will be skipped): "
            f"{sorted(failed_enrichment_columns)}"
        )

    logger.info("Step 9 — Executing validation rules")
    validation_engine = ValidationEngine(
        workbook, config, registry, vtype_registry, issues_writer
    )
    validation_stats  = validation_engine.execute(
        validation_rules, failed_enrichment_columns,
        failed_sheets=ef_result.get("failed_sheets", set()),
    )

    # -----------------------------------------------------------------------
    # 10  Issues worksheet is already written by IssuesWriter during execution (FS 8.10)
    # -----------------------------------------------------------------------
    issues_writer.apply_column_widths()
    logger.info(
        f"Step 10 — Issues recorded: "
        f"{issues_writer.error_counter} error(s), "
        f"{issues_writer.warning_counter} warning(s)"
    )

    # -----------------------------------------------------------------------
    # 11  Generate Summary worksheet (FS 8.11)
    # -----------------------------------------------------------------------
    from output.summary_writer import SummaryWriter

    logger.info("Step 11 — Generating Summary worksheet")
    summary_writer = SummaryWriter(workbook, config, issues_writer)
    summary_writer.write(
        enrichment_stats=enrichment_stats,
        validation_stats=validation_stats,
        external_files_result=ef_result,
        external_files_rules=external_files,
        start_time=start_time,
    )

    # -----------------------------------------------------------------------
    # 12  Save workbook
    # -----------------------------------------------------------------------
    logger.info("Step 12 — Saving output workbook")
    try:
        wb_manager.save()
    except Exception as exc:
        logger.error(f"Failed to save workbook: {exc}", exc_info=True)
        return 1

    elapsed = (datetime.now() - start_time).total_seconds()
    logger.info("=" * 60)
    logger.info(
        f"Completed in {elapsed:.1f}s — "
        f"{issues_writer.error_counter} error(s), "
        f"{issues_writer.warning_counter} warning(s). "
        f"Output: {config.output_file}"
    )
    print(
        f"\nDone in {elapsed:.1f}s — "
        f"{issues_writer.error_counter} error(s) and "
        f"{issues_writer.warning_counter} warning(s). "
        f'See output file "{config.output_file}"'
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
