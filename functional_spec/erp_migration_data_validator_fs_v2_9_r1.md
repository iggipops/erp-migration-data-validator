# ERP Migration Data Validator — Functional Specification v2_9_r1

Document history: Not retained prior to v2_9_r0, which marked the starting point for public distribution.

# 1. Document Overview

## 1.1 Purpose

This Functional Specification defines the required behavior of ERP Migration Data Validator V2. It describes what the system must do, how it must behave, and what inputs and outputs are expected. It serves as the primary reference for implementation and testing.

## 1.4 Version History

|             |          |                                              |
|-------------|----------|----------------------------------------------|
| **Version** | **Date** | **Description**                              |
| v2_9_r0      | 9/8/2026  | AI backend generalized from a single hardwired provider to a pluggable Provider Registry (new section 2.6.3; 2.6.4 now compares three registries, not two) — the framework's third extensibility point, alongside the Validation Type Registry (2.6.1) and Function Registry (2.6.2). Provider Registry construction gets its own genuine pipeline step, new section 8.1 — every subsequent section-8 subsection (and every bare "Step N" reference throughout the document, including 8.12's logging behavior) shifts up by one accordingly; 2.4's high-level list gains a matching new first item. New mandatory-when-ai_enabled `ai_provider` field (3.4) selects the active provider by name, validated against the registry exactly as CustomFunctionName is validated against the Function Registry, at Step 2 (8.2) — immediately after the registry itself is built at Step 1 (8.1), since it has no config or workbook dependency and needs to exist before ai_provider can be checked against it. `ai_endpoint` is removed from the top-level AI block entirely (reverses v2_8_r1) — provider-specific settings, including ai_endpoint and now ai_api_version too (both mandatory), move to a new `ai_provider_settings.<provider>.*` block, mirroring the existing `custom_functions.<name>.*` pattern (3.5) rather than inventing a new shape. ai_api_version is no longer a code constant anywhere — the last hardcoded AI value is gone. ai_model/ai_api_key/ai_provider remain unconditionally mandatory at Step 2/8.2 when ai_enabled = true, and ai_provider_settings' required keys for the selected provider are checked at the same step. The AI connectivity pre-flight check (ai_connection_test_prompt) also moves back to Step 2, as new section 8.2.1 — reversing v2_8_r0's move of this check into ai_itemcategoryfit's own metadata-check hook at Step 6/8.7. That move traded away Step 1/2 fail-fast timing to avoid a wasted connectivity call on workbooks with no active AI rule; on reflection, config.yaml is inherently one-per-workbook already (input_file is mandatory), so that waste isn't a realistic scenario worth the architectural cost — and it no longer requires ai_itemcategoryfit-specific code or the once-per-run dedup flag that move needed, since the Provider Registry now makes the check genuinely framework-level and provider-agnostic (dispatched like any other AI call, just not tied to any rule). Appendix C.7 (that function-specific implementation) is removed accordingly; 2.6.2's dedup-pattern paragraph is removed with it, since it no longer has a concrete example. Ships with one provider module, anthropic, implementing the full contract; ai_itemcategoryfit and the rest of the framework are entirely provider-agnostic. Appendix C.1's "Anthropic Messages API only" caveat is softened accordingly — Appendix D's example config updated to the new shape. Functional change, no BS impact — this is an internal architecture generalization, not a change to what the tool does. Also, unrelated editorial cleanup: 2.4's step order corrected to match section 8's actual order (registry build before external-file import; Issues formatting before Summary generation) — the two had disagreed with each other since before this revision, a discrepancy the entry point's own code docstring already called out; section 8 is what the code follows, so 2.4 is what moved. |
| v2_9_r1      | 9/19/2026 | ImportSpec now applies to csv external files only (5.4, 5.4.2, 8.3, 8.6). xlsx external files already carry native Excel types, like the primary workbook, so they need no conversion. New check in 8.3.2: an ImportSpec row whose ExternalFileSheetName points to an xlsx row in ExternalFiles is a metadata validation error, reported before any file is imported. Functional change, no BS impact. |

# 2. System Overview

## 2.1 What the Tool Does

The tool reads an Excel workbook containing business data and metadata rules. It executes enrichment functions to generate new derived columns, then executes validation functions to check data quality. Results are written back to the workbook as an Issues worksheet and a Summary worksheet, with color highlighting applied to flagged cells.

## 2.2 What the Tool Does Not Do

- Replicate ERP business logic

- Connect to ERP systems directly

- Perform ETL operations

- Act as a workflow or BPM engine

- Provide a graphical user interface (V2 scope)

## 2.3 Execution Model

The tool executes from the command line or as a compiled executable. It reads a YAML configuration file specifying input/output paths and settings. All processing is sequential, deterministic, and worksheet-wide. Partial execution and row filtering are not supported. All operations after the initial configuration load work exclusively on the output file — the input file is never modified.

## 2.4 High-Level Processing Sequence

1.  Build AI Provider Registry (section 2.6.3) — no config or workbook dependency, so this happens first
2.  Load and validate technical configuration
3.  Validate ExternalFiles and ImportSpec metadata — structural checks before touching the filesystem
4.  Copy input workbook to output path — all further operations apply to the output file only
5.  Build function registry and validation type registry (section 2.6)
6.  Import external files as worksheets into the output workbook, applying ImportSpec column conversions where defined
7.  Validate ValidationRules and EnrichmentRules metadata — all worksheets and columns now exist
8.  Execute enrichments in SequenceNum order — generate new columns
9.  Execute validations in SequenceNum order — highlight cells and populate Issues worksheet
10. Apply formatting
11. Generate Summary worksheet
12. Save output workbook
13. Write logs

## 2.5 Core Processing Principles

These principles apply across the entire tool — to enrichment functions, validation functions, and built-in validation types alike. They take precedence over any single section's wording where ambiguity might otherwise arise.

### 2.5.1 Row Processing Philosophy

- Empty rows remain part of the processing scope — they are not skipped or removed.

- Row positions remain stable throughout processing. The framework preserves the original worksheet row alignment from input to output.

- Enrichment and validation processing never implicitly compresses, reorders, or filters the dataset. The number of data rows in a worksheet at the end of processing equals the number of data rows at the start.

### 2.5.2 Missing and Empty Values Are Business Data State

A missing, empty, or null value encountered during enrichment or validation processing is a **business data state**, not a technical error. Specifically:

- Functions (built-in or custom) must handle empty/null input values as normal data — not raise runtime errors because a value is missing.

- Values that appear as the text "nan" (a artifact of formula evaluation or data import, typically from pandas/Excel formula cells) are treated as empty wherever empty-value handling applies — they are not treated as the literal text "nan".

- If a missing value should be flagged to the user, this is the responsibility of a dedicated validation rule (EMPTY, NULL, DATE, etc.) — not an unhandled exception from an enrichment or custom validation function.

### 2.5.3 Enrichment Runtime Failure Philosophy

If an enrichment rule fails at runtime (the registered function raises an exception during execution):

- The rule's TargetColumn is **not created** in the output worksheet.

- The enrichment rule is recorded with status `failed` in the Summary worksheet.

- The failure (including the exception details) is written to the console and log file.

- **Dependent validation rules are skipped.** A validation rule depends on a failed enrichment's TargetColumn (call it X) if, for that active validation rule:
  - `ColumnName = X`, **or**
  - `ValidationType` is REFERENCE or DUPLICATE2 **and** `ReferenceDataColumn = X`

  Such validation rules are recorded with status `skipped` in the Summary worksheet, with the reason logged (dependency on failed enrichment rule).

Enrichment rules themselves have no inter-dependencies — every EnrichmentRules row is executed independently of every other EnrichmentRules row, regardless of execution order or outcome. Similarly, ValidationRules rows never depend on other ValidationRules rows. The only dependency relationship in the framework is validation-on-enrichment, as described above.

### 2.5.4 Empty Value Exclusion Principle

For every validation type **except EMPTY**, if the target column's cell value is empty (per section 6.1), that row is excluded from the check entirely — it does not produce an issue, regardless of whether the underlying check would otherwise pass or fail. The rationale is that a single cell should not be flagged by multiple validation rules for the same underlying condition (emptiness); flagging empty cells is the dedicated responsibility of an EMPTY rule.

This applies uniformly to built-in validation types (DUPLICATE, DUPLICATE2, FORMAT, REFERENCE, DATE, NUMERIC, INTEGER). **CUSTOM validation functions are exempt from this principle** — they receive the full row context and are responsible for their own empty-value handling internally. A custom function that validates the entire row (such as `dimpolicyvalidation`) must not be silently suppressed based on the state of the anchor `ColumnName` cell, since the failure may relate to a different column on the same row. Custom functions that do want to skip empty anchor cells should implement that check themselves.

**NULL is the sole exception to this principle.** By design (section 6.2), NULL's definition deliberately includes empty values as part of what it flags — NULL is "EMPTY plus numeric zero". A worksheet that defines both an EMPTY rule and a NULL rule on the same column will therefore see the same empty cells flagged by both rules. This is expected and intentional; if this overlap is undesirable for a specific column, define only one of the two rules for that column.

For DUPLICATE and DUPLICATE2 specifically, rows with an empty `ColumnName` cell are excluded from the duplicate-combination check entirely — they do not count toward duplicate groups and are never flagged, even if multiple rows share the same empty value.

## 2.6 Extensibility Architecture

The framework maintains three separate, parallel registries so that new capabilities can be added without modifying the core engine. Each serves a distinct purpose. The first two are referenced from the metadata worksheets, each in its own way; the third, the Provider Registry, is selected entirely through configuration and is never referenced from a worksheet at all (2.6.4).

### 2.6.1 Validation Type Registry

Every built-in validation type (EMPTY, NULL, DUPLICATE, DUPLICATE2, FORMAT, REFERENCE, DATE, NUMERIC, INTEGER, AI, CUSTOM) is implemented as a self-contained module registered in the Validation Type Registry. The core validation engine is a thin dispatcher: for each active ValidationRules row, it looks up the row's `ValidationType` in the registry and delegates execution to that type's module. Adding a new built-in validation type means adding a new module to the registry — it does not require changes to the core engine.

Each validation type module may optionally declare a metadata-check function. If declared, this function is called during metadata validation (section 8.7) for every active rule using that type, and may return additional error messages specific to that type's requirements — for example, checking that a sheet or column referenced in a type-specific way actually exists. If a validation type does not declare this function, no additional checks beyond the universal ones (section 8.7) are performed for it.

### 2.6.2 Function Registry

Built-in and custom functions (used by EnrichmentRules, and by ValidationRules rows with `ValidationType = CUSTOM` or `ValidationType = AI` — the two dispatch identically via `CustomFunctionName`, section 6.8) are registered in the separate Function Registry described in section 7. As established in section 7.2.1, any function may optionally declare `REQUIRED_PARAMS` to have its parameters checked automatically during metadata validation.

In addition to `REQUIRED_PARAMS`, any function may optionally declare a metadata-check function, following the same opt-in pattern. If declared, this function is called during metadata validation for every active CUSTOM or AI ValidationRules row (or EnrichmentRules row) referencing that function by name, and may return additional error messages — for example, `dimpolicyvalidation` may check that the sheets named in its config.yaml parameters actually exist in the workbook. This mechanism allows individual custom functions to have rich, function-specific metadata validation without the function needing to be promoted to a dedicated built-in validation type.

### 2.6.3 Provider Registry

Which AI service an AI-flavored function talks to is not hardwired into the framework. A third registry, parallel to 2.6.1 and 2.6.2, holds one module per supported AI backend ("provider"), each implementing a fixed contract:

- `PROVIDER_NAME` — the string this provider is selected by, via config.yaml's `ai_provider` field (section 3.4). Matched the same way `CustomFunctionName` is matched against the Function Registry.
- `REQUIRED_SETTINGS` — optional, a list of keys this provider needs present under `ai_provider_settings.<PROVIDER_NAME>` in config.yaml (section 3.4) whenever it is the selected provider. Checked automatically at Step 2/8.2 (section 3.8) — the same automatic-checking pattern `REQUIRED_PARAMS` gives custom functions (section 7.2.1), against a different config location.
- `send(prompt, config, settings)` — sends `prompt` to this provider's API and returns the raw text response. `config` supplies the fields every provider needs in common (`ai_model`, `ai_api_key`); `settings` is this provider's own resolved `ai_provider_settings.<PROVIDER_NAME>` block, already checked against `REQUIRED_SETTINGS`. This is the only function in the framework that knows this provider's specific request/response shape — headers, body structure, how to extract text from the reply.

The shared AI client module (section 6.8) is a thin dispatcher: given `config.ai_provider`, it looks up the matching module in the Provider Registry and calls its `send()`. Everything else AI-related — `send_prompt()`/`send_batch()`'s batching, retry, and verdict-parsing logic (section 6.8), and every AI-flavored function including `ai_itemcategoryfit` (Appendix C) — is entirely provider-agnostic; none of it changes based on which provider is selected. Unlike the Function Registry, provider modules are not discovered from a user-configurable directory — they ship with the framework itself, the same distribution model as built-in validation types (2.6.1), reflecting that a provider is a core backend integration the framework maintains, not business logic anyone drops in. The framework currently ships with one provider module, `anthropic` (Appendix F).

### 2.6.4 Why Three Registries

A validation type (registry 2.6.1) is referenced directly by name in the `ValidationType` column and is meant for capabilities the framework itself defines and ships with. A function (registry 2.6.2) is referenced by name in `CustomFunctionName` and is meant for capabilities that can be added or modified by anyone without touching the validation type list at all — including one-off, experimental, or organization-specific business logic. CUSTOM remains the universal entry point for the latter category; promoting a specific custom function to a dedicated validation type is a deliberate design choice reserved for capabilities the framework considers core (such as DATE or NUMERIC), not a requirement for gaining stronger metadata validation, which is available to any function via 2.6.2.

A provider (registry 2.6.3) is a third, orthogonal kind of thing — it isn't referenced from a ValidationRules or EnrichmentRules row at all, and no rule author needs to know it exists. It's selected once per run, at the framework-configuration level (`ai_provider`, section 3.4), and every AI-flavored function shares whichever one is active without knowing or caring which it is. Where a validation type or function is "what capability runs," a provider is purely "how an AI-flavored capability reaches the outside world" — closer in spirit to `custom_functions_directory` or `log_file` (technical configuration, section 3) than to anything in ValidationRules or EnrichmentRules.

# 3. Technical Configuration

## 3.1 Configuration File and Program Arguments

Technical configuration is stored in an external YAML file. The path to the configuration file is passed as a command-line argument at execution:

python erp_migration_data_validator.py --config path/to/config.yaml

If the --config argument is omitted, the tool looks for a file named config.yaml in the same directory as the executable. If no configuration file is found at the default path, processing terminates immediately with an error message.

## 3.2 Framework

All fields in this section are mandatory. Processing terminates immediately if any of these fields are missing or invalid.

|                              |          |                                                          |
|------------------------------|----------|----------------------------------------------------------|
| **Field**                    | **Type** | **Description**                                          |
| input_file                   | string   | Path to the input workbook (xlsx)                        |
| output_file                  | string   | Path to the output workbook (xlsx)                       |
| log_file                     | string   | Path to the log file                                     |
| custom_functions\_ directory | string   | Path to the directory containing custom function modules |

## 3.3 Formatting

All fields in this section are optional.

|                      |          |                                                                                                                                                                                                                                                                                                                                                                      |
|----------------------|----------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Field**            | **Type** | **Description**                                                                                                                                                                                                                                                                                                                                                      |
| format_special_chars | string   | The set of characters treated as not permitted by the FORMAT validation type. Each character in the string is evaluated individually. Example in YAML: format_special_chars: "!@#\$%^&\*()". If omitted, the special character sub-check within FORMAT is skipped entirely — other FORMAT checks (leading/trailing whitespace, consecutive spaces) are not affected. |
| predefined_colors.\* | object   | Extends the built-in named color list with custom name-to-hex mappings. Example: predefined_colors: {PURPLE: "#7030A0", TEAL: "#008080"}. Built-in names (RED, YELLOW, GREEN, ORANGE, BLUE, GREY) are always available without configuration.                                                                                                                        |

## 3.4 AI

All fields in this section are optional except where noted below. If `ai_enabled` is omitted or set to false, all AI validation rules (`Active = Yes`, `ValidationType = AI`) are skipped, and this is recorded in the log file — this is a soft skip, not a metadata or configuration error.

If `ai_enabled = true`, `ai_model`, `ai_api_key`, and `ai_provider` must all be present, and `ai_provider_settings.<ai_provider>` must supply every key that provider's module declares as required (section 2.6.3) — missing any of these terminates processing immediately at Step 8.2 (section 3.8).

|                            |          |                                             |
|----------------------------|----------|---------------------------------------------|
| **Field**                  | **Type** | **Description**                             |
| ai_enabled                 | boolean  | Enable AI validation calls (default: false) |
| ai_model                   | string   | AI model identifier                         |
| ai_api_key                 | string   | API key for the AI service                  |
| ai_provider                | string   | Mandatory when ai_enabled = true. Selects which provider module (Provider Registry, section 2.6.3) handles AI requests for this run. Must match a `PROVIDER_NAME` registered in the framework — currently only `anthropic` (Appendix F) ships. |
| ai_provider_settings.\*    | object   | Per-provider named parameter blocks — same shape as `custom_functions.*` (section 3.5), but keyed by provider name instead of function name. Each `ai_provider_settings.<name>` block supplies the settings that specific provider's module requires (section 2.6.3); only the block matching the active `ai_provider` is checked or used. See Appendix F for the `anthropic` provider's required settings. |
| ai_connection_test_prompt  | string   | Optional. If non-empty and ai_enabled = true, this text is sent as a real request through the shared AI client module (section 6.8) at Step 2 (8.2.1), dispatched via the Provider Registry to whichever provider ai_provider selects — a framework-level check, not tied to any specific custom function. Terminates immediately if the call fails. If empty or omitted, no pre-flight connectivity check is performed — a bad key or unreachable service is only discovered the first time an AI validation rule actually runs, and surfaces as an ordinary runtime failure on that rule (section 8.9.2), not a run-terminating error. |

## 3.5 Custom Functions

All fields in this section are optional. `custom_functions` is a structured block where each key is a custom function name and its value is a set of named parameters passed to that function at runtime. The framework does not validate the contents of individual function parameter blocks — each custom function is responsible for reading and validating its own parameters from the context object.

|                     |          |                                                                                                                                                              |
|---------------------|----------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Field**           | **Type** | **Description**                                                                                                                                              |
| custom_functions.\* | object   | Per-function named parameter blocks. Each key is a function name; its value is a set of key-value parameters passed to that function via the context object. |

## 3.6 Date Validation

All fields in this section are optional. `date_validation` configures the behavior of the DATE validation type (section 6.9) globally for the entire run. See section 6.9.1 for field definitions and an example.

## 3.7 Numeric Validation

All fields in this section are optional. `numeric_validation` configures the behavior of the NUMERIC and INTEGER validation types (sections 6.10 and 6.11) globally for the entire run — both types share this single configuration block. See section 6.10.1 for field definitions and an example.

## 3.8 Configuration Validation

If the configuration file is missing, unreadable, or contains invalid values, processing terminates immediately. Required fields must all be present. The input file must exist and be a valid Excel workbook in .xlsx format (MS Excel 2010 - 365 Spreadsheet). The output file path must be writable. Any violation causes immediate termination with a brief error message to the console and full details written to the log file.

If `ai_enabled = true`: `ai_model`, `ai_api_key`, and `ai_provider` must all be present at this step, `ai_provider` must match a `PROVIDER_NAME` registered in the Provider Registry (section 2.6.3), and `ai_provider_settings.<ai_provider>` must supply every key that provider declares as `REQUIRED_SETTINGS` — a missing or unrecognized value terminates processing immediately, with the same console/log treatment as any other configuration error. This all happens at Step 2, before any workbook is opened, because provider identity and its settings are framework-level configuration exactly like `ai_model`/`ai_api_key` — nothing about them depends on which rules or functions this run's workbook happens to reference. If configuration loads successfully and `ai_connection_test_prompt` is non-empty, the framework then sends it as a real request through the shared AI client module (section 6.8), dispatched via the Provider Registry to whichever provider `ai_provider` selects (see 8.2.1) — a failed request terminates processing immediately, with the same console/log treatment as any other configuration error, still before any workbook is opened.

# 4. Workbook Structure

The input and output workbooks share the same structure. The output workbook is created by copying the input workbook to the configured output path. All subsequent processing operates on the output file only — the input file is never modified.

The Issues and Summary worksheets may already exist in the input workbook. They will be carried over by the copy operation and overwritten when the tool generates new ones.

If a file already exists at the output path, it is overwritten without prompting. This event is logged to the console and log file.

|                          |                  |                                          |
|--------------------------|------------------|------------------------------------------|
| **Worksheet**            | **Type**         | **Required**                             |
| ValidationRules          | Metadata         | Yes                                      |
| EnrichmentRules          | Metadata         | Yes — must exist even if empty           |
| ExternalFiles            | Metadata         | Yes — must exist even if empty           |
| ImportSpec               | Metadata         | No — optional; if absent, all csv external file columns import as text. Not applicable to xlsx external files, which always carry native types. |
| Business data worksheets | Data             | At least one must be referenced in rules |
| Issues                   | Generated output | Created or overwritten by the tool       |
| Summary                  | Generated output | Created or overwritten by the tool       |

# 5. Metadata Worksheets

## 5.1 ExternalFiles Worksheet

The ExternalFiles worksheet defines external reference datasets to be loaded from xlsx or csv files outside the main workbook. This is required when reference data is maintained in separate files, or when working with csv format, which can contain only one dataset per file.

|                   |              |                                                                                                                                      |
|-------------------|--------------|--------------------------------------------------------------------------------------------------------------------------------------|
| **Column**        | **Required** | **Description**                                                                                                                      |
| SequenceNum       | Yes          | Loading order — numeric, unique across Active = Yes rows                                                                             |
| FilePath          | Yes          | Absolute or relative path to the external file                                                                                       |
| SheetName         | Yes          | The name used to reference this dataset in rules and functions. Also becomes the worksheet name in the output workbook after import. |
| OriginalSheetName | Conditional  | Required when Format = xlsx. Specifies which worksheet to load from the external Excel file. Ignored for csv.                        |
| Format            | Yes          | xlsx or csv                                                                                                                          |
| CSVDelimiter      | Conditional  | Required when Format = csv. Supported values: comma (,) semicolon (;) pipe (\|) colon (:) space ( )                                             |
| Active            | Yes          | Yes or No — inactive rows are ignored completely                                                                                     |

## 5.2 ValidationRules Worksheet

Each row defines one validation rule. Inactive rows (Active = No) are ignored completely.

|                    |              |                                                                                                                                                                                  |
|--------------------|--------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column**         | **Required** | **Description**                                                                                                                                                                  |
| SequenceNum        | Yes          | Execution order — numeric, unique within ValidationRules. Validation and enrichment rules maintain independent sequences.                                                        |
| RuleCode           | Yes          | Unique identifier for the rule (case-sensitive)                                                                                                                                  |
| RuleName           | Yes          | Human-readable rule description                                                                                                                                                  |
| SheetName          | Yes          | Target worksheet name (case-insensitive)                                                                                                                                         |
| ColumnName         | Yes          | Target column name (case-insensitive)                                                                                                                                            |
| Severity           | Yes          | Error or Warning                                                                                                                                                                 |
| ValidationType     | Yes          | Built-in type or CUSTOM or AI — see section 6                                                                                                                                    |
| Color              | Yes          | Highlight color for flagged cells. Must be a built-in named color (RED, YELLOW, GREEN, ORANGE, BLUE, GREY), a name defined in config predefined_colors, or a hex code (#RRGGBB). |
| ReferenceDataSheet    | Conditional  | Required when ValidationType = REFERENCE — source worksheet name or external file SheetName. Should be differ from SheetName value. Optional when ValidationType = AI — an opt-in switch for failed-import skip protection only, not a lookup path (section 6.8). Not used for DUPLICATE2.                                                            |
| ReferenceDataColumn   | Conditional  | Required when ValidationType = REFERENCE or DUPLICATE2 — column name in the reference source. For DUPLICATE2: second column on the same sheet used for the combination check. Optional when ValidationType = AI — flags a dependency on a column pending creation by enrichment, not a lookup path; must be empty when ReferenceDataSheet is empty (section 6.8).   |
| CustomFunctionName | Conditional  | Required when ValidationType = CUSTOM or AI — registered function name                                                                                                                 |
| Active             | Yes          | Yes or No                                                                                                                                                                        |

## 5.3 EnrichmentRules Worksheet

Each row defines one enrichment rule. Enrichments generate new columns in the target worksheet. Inactive rows are ignored completely.

|                    |              |                                                                                                                           |
|--------------------|--------------|---------------------------------------------------------------------------------------------------------------------------|
| **Column**         | **Required** | **Description**                                                                                                           |
| SequenceNum        | Yes          | Execution order — numeric, unique within EnrichmentRules. Enrichment and validation rules maintain independent sequences. |
| RuleCode           | Yes          | Unique identifier for the rule (case-sensitive)                                                                           |
| RuleName           | Yes          | Human-readable description                                                                                                |
| SheetName          | Yes          | Target worksheet name (case-insensitive)                                                                                  |
| TargetColumn       | Yes          | Name of the new column to create. Must not already exist in the target sheet.                                             |
| Color              | No           | Optional highlight color applied to all cells in the generated column — both header and data cells                        |
| CustomFunctionName | Yes          | Registered function name that generates the column values                                                                 |
| FunctionArguments  | Conditional  | Required for built-in enrichment functions. Key=value pairs separated by semicolons. Lists use comma-separated values. Example: `columns=ItemName,ItemId;separator=_`. Custom functions ignore this column and use config.yaml instead. |
| Active             | Yes          | Yes or No                                                                                                                 |

## 5.4 ImportSpec Worksheet

The ImportSpec worksheet defines per-column type conversion rules applied during external file import, for csv files only (section 8.6). xlsx external files are not covered by ImportSpec — they carry native Excel types on import, the same as the primary workbook, and require no conversion step. The sheet is optional — if absent or empty, all columns from all csv external files are imported as raw text, which is the default behavior established in previous versions.

Each row defines conversion behavior for one column of one csv external file. Columns not covered by any active ImportSpec row are imported as raw text. ImportSpec rows reference external files by their `SheetName` from the ExternalFiles worksheet — the same identifier already used to uniquely identify each external file import unit elsewhere in the framework. A `SheetName` belonging to an xlsx external file is not a valid ImportSpec target (see 8.3.2).

|                    |              |                                                                                                                                                          |
|--------------------|--------------|----------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column**         | **Required** | **Description**                                                                                                                                          |
| ExternalFileSheetName | Yes        | SheetName from ExternalFiles worksheet — identifies which external file this row applies to                                                              |
| ColumnName         | Yes          | Name of the column in the external file to convert. Must match the column header exactly (case-insensitive)                                              |
| TargetType         | Yes          | Conversion target type. One of: `text` (no conversion, default), `date`, `integer`, `numeric`                                                           |
| DateFormats        | Conditional  | Required when TargetType = `date`. Ordered list of date format strings (Python strftime syntax, e.g. `%d.%m.%Y`) separated by semicolons. First matching format wins. The framework does not guess formats. |
| DecimalSeparator   | Conditional  | Used when TargetType = `numeric` or `integer`. Character used as the decimal point in the source text. Default: `.` (dot).                              |
| ThousandsSeparator | Conditional  | Used when TargetType = `numeric` or `integer`. Character used as the thousands grouping separator in the source text. Default: none. Grouping must be structurally valid (exactly 3 digits per group). |
| Active             | Yes          | Yes or No — inactive rows are ignored completely                                                                                                         |

### 5.4.1 Conversion Behavior

**date:** The source text value is parsed against each format in `DateFormats` in order. The first matching format is used to produce a native Excel date value. If no format matches, the cell is stored as text and a warning is written to the log. The converted value is a native Excel date — DATE validation (section 6.9) will accept it without further issues.

**numeric:** The source text value is parsed as a floating-point number after removing `ThousandsSeparator` (if configured) and converting `DecimalSeparator` to `.`. Thousands grouping must be structurally valid: the rightmost group(s) must each contain exactly 3 digits. If parsing fails, the cell is stored as text and a warning is written to the log.

**integer:** Same as `numeric`, but the resulting value must have no fractional part (i.e. it must equal its own integer cast). If the fractional part is non-zero, the cell is stored as text and a warning is written to the log. Values such as `"123.0"` or `"123.000"` are accepted as valid integers per the ERP migration context (section 6.11).

**text:** No conversion — the column is imported as-is. This is also the behavior for any column not covered by an active ImportSpec row.

### 5.4.2 ImportSpec and Validation Interaction

The primary purpose of ImportSpec is to produce native Excel date and number values from text-based source files — csv files specifically, since that is the only external file format that has no native types of its own. This ensures that DATE, NUMERIC, and INTEGER validation rules (sections 6.9–6.11) can operate correctly on csv-imported data — those validation types require native Excel values and do not parse text themselves. Without ImportSpec, a DATE rule applied to a csv-imported date column would always flag every row as an error. xlsx external files need no such step: they already carry native types on import (section 8.6.1), so DATE/NUMERIC/INTEGER rules work on them without any ImportSpec configuration.

### 5.4.3 Metadata Validation

ImportSpec is validated during step 8.3.2. See section 8.3.2 for the full validation table. Key points:

- `ExternalFileSheetName` must reference an existing Active = Yes row in ExternalFiles by its SheetName (referential integrity check), and that row's `Format` must be `csv`
- `TargetType` must be one of the supported values — see section 5.4.1
- `DateFormats` must not be empty when `TargetType = date`
- `ColumnName` existence in the actual file is validated at import time (section 8.6), not here, since the column list is only known after the file is opened
- Duplicate `ExternalFileSheetName` + `ColumnName` combinations across Active = Yes rows are an error
- All errors found terminate processing before any files are imported

# 6. Built-In Validation Types

Built-in validation types require no custom function. The framework executes them directly based on the ValidationType value.

## 6.1 EMPTY

Flags rows where the target column value is empty. A value is considered empty if it is:

- A missing/null cell value (no value stored in the cell), or

- A string consisting only of whitespace characters (spaces, tabs), including a zero-length string

Numeric zero (0, 0.0, etc.) is **not** considered empty by EMPTY — use NULL if zero values should also be flagged. Text values, including text representations like "nan" produced by upstream formula evaluation or data import artifacts, are treated as empty for the purposes of this validation, consistent with the framework-wide empty-value handling described in section 2.5.

## 6.2 NULL

Flags rows where the target column value is empty, null, or represents a numeric null value. This includes:

- Empty or null values (same as EMPTY)

- The numeric value 0

- Decimal variations: 0.0, 0,0 (locale-specific decimal separators), 00, etc.

Comparison treats leading zeros and different decimal separators as equivalent. For example, 0, 0.0, and 0,0 are all flagged as numeric null values. Whitespace-only values are treated as empty. This validation is useful for data imports where numeric zero values should not be present (e.g., required quantity or amount fields).

**Note:** NULL is the sole exception to the Empty Value Exclusion Principle (section 2.5.4) — unlike every other validation type, NULL deliberately flags empty cells as part of its normal operation.

## 6.3 DUPLICATE

Flags rows where the target column value appears more than once in the same worksheet. All occurrences of a duplicated value are flagged — not only the second occurrence. Comparison is case-insensitive, per section 9.

## 6.4 DUPLICATE2

Flags rows where the **combination** of two column values appears more than once in the same worksheet. The first column is defined by `ColumnName`, the second by `ReferenceDataColumn` in the rule definition. Both columns must be in the same target sheet. `ReferenceDataSheet` is not used for this type. All occurrences of a duplicated combination are flagged — not only the second occurrence. Comparison is case-insensitive for both columns, per section 9.

Example use case: flagging duplicate (ItemId, WarehouseId) combinations in a stock balance dataset where each item-warehouse pair should be unique.

## 6.5 FORMAT

Validates the format of each target column value against the following rules:

- Leading whitespace present

- Trailing whitespace present

- Consecutive internal spaces (two or more spaces between words)

- Special characters not permitted (configurable via format_special_chars in YAML)

The FORMAT function processes the entire column and returns one result per row. If format_special_chars is not defined in the configuration, the special character sub-check is skipped. All other FORMAT checks still execute normally.

## 6.6 REFERENCE

Validates that each value in the target column exists in a reference dataset. The reference source is defined by ReferenceDataSheet and ReferenceDataColumn in the rule definition. The source may be a worksheet in the workbook or an imported external file referenced by its SheetName. Comparison is case-insensitive. Empty values in the target column are not flagged by REFERENCE — use the EMPTY validation separately if needed.

## 6.7 CUSTOM

Executes a registered custom function. The function name is specified in CustomFunctionName. Custom functions receive their parameters from config.yaml under the `custom_functions` block. See section 7 for custom function requirements.

## 6.8 AI

Executes an AI-assisted validation via a registered custom function, in exactly the same manner as CUSTOM (section 6.7) — `CustomFunctionName` identifies the function, and `custom_functions.<name>` in config.yaml supplies that function's parameters (its prompt text, which sheets and columns it uses to compose the data sent to the AI, and anything else the function needs). See section 7 for custom function requirements. The only framework-level differences from CUSTOM are:

- Requires `ai_enabled = true` in configuration. If `ai_enabled` is false or omitted, the rule is skipped and this is recorded in the log file (section 3.4) — a soft skip, not a metadata or runtime error.

- `SheetName`/`ColumnName` are purely mechanical for AI rules: they determine which sheet the engine iterates and which cell gets highlighted on a flagged row — exactly as for every other validation type, with no exception. They carry no lookup meaning whatsoever for AI; how the function actually gathers its data is entirely independent of them (see below and Appendix C.3).

- `ReferenceDataSheet` and `ReferenceDataColumn` are available to AI rules on an **optional** basis, but neither performs an actual data lookup the way they do for REFERENCE — an AI-flavored function gathers 100% of its own data through its own `custom_functions.<name>` configuration (source sheet, reference sheet, key columns, text columns — everything), fully self-contained, the same way `dimpolicyvalidation` already is (section 6.7's example). `ReferenceDataSheet` and `ReferenceDataColumn` exist only as two independent, narrowly-scoped opt-in switches:

  - **`ReferenceDataSheet`**, if set, gives the rule `failed_sheets` skip protection (section 8.6.4) for that sheet — the rule is automatically skipped, not treated as a metadata or runtime error, if that sheet failed to import. It is purely a protection switch; the function's own configuration may name a completely different sheet than the one it actually reads from, though in practice they will usually match, since the whole point of setting it is to protect the sheet the function actually depends on.

  - **`ReferenceDataColumn`**, if set, flags that the AI function depends on a column that may not exist yet, because it is pending creation by an active EnrichmentRules row targeting `ReferenceDataSheet` — the same forward-reference mechanic REFERENCE and DUPLICATE2 already use (section 2.5.3). It carries no other meaning for AI.

  - **Consistency rule:** `ReferenceDataColumn` must be empty when `ReferenceDataSheet` is empty — a `ReferenceDataColumn` value with no corresponding `ReferenceDataSheet` is a metadata error (section 8.7.1). When `ReferenceDataSheet` is set, `ReferenceDataColumn` remains fully optional: leave it empty if there's no pending-column dependency, or set it if there is — either way it must be empty or reference a column that already exists (or is validly pending), never a value matching neither.

  When an AI rule has no natural fit for either switch — no meaningful sheet-failure dependency to protect against, no pending-column dependency — both may be left empty, and the function is entirely responsible for its own data access, exactly like a CUSTOM function with no reference lookup at all.

An AI validation function returns row-level pass/fail results and may additionally return a comment text per row, exactly as CUSTOM functions do. Comment text is written to the Comment column in the Issues worksheet.

AI request/response mechanics — retry, batching, and verdict-parsing logic — are provided by a shared framework module, not reimplemented per function; individual AI validation functions supply only what data to gather, what prompt to send, and how to interpret the returned result. That shared module is itself provider-agnostic: it dispatches to whichever provider module `ai_provider` selects (Provider Registry, section 2.6.3), which is the only place that knows the actual request/response shape for that specific AI service. Neither the shared module nor any AI-flavored function changes based on which provider is active. This same shared module is also what the framework itself calls directly to run the `ai_connection_test_prompt` pre-flight check at Step 2 (section 3.4, 8.2.1) — no AI-flavored function is involved in that check at all.

The shared module processes rows in batches (size configurable per function via its own `batch_size` parameter, not a global setting) rather than one AI call per row, for cost efficiency. Its response contract is standardized across every AI-flavored function: a minimal per-row verdict for every row in the batch (e.g. a compact ordered list such as `P,P,F,P,F`) with comment text present only for failing rows — this keeps every row explicitly accounted for (no row is ever silently omitted) while limiting token cost to the rows that actually need an explanation.

Because processing happens in batches, a single rule can end up with some rows successfully verdicted and others not, depending on which batches succeeded or failed. The rule's overall execution status (section 8.11, Summary worksheet) reflects this with a fourth possible value, in addition to executed/skipped/failed:

- **`executed`** — every batch succeeded; all rows have a real verdict.
- **`partial`** — at least one batch succeeded and at least one failed. Issues are recorded normally for every row in a successful batch; rows in a failed batch have no verdict and are not represented in the Issues worksheet. The Summary worksheet's Validation statistics section reports both the issue count from successful batches and the count of rows that were never checked.
- **`failed`** — every batch failed (e.g. an invalid API key, so nothing ever succeeded). No issues are recorded for the rule.

See Appendix C for a complete worked example of an AI-flavored custom function, including field roles and a concrete logic chain.

## 6.9 DATE

Flags rows where the target column value cannot be safely imported into a date field of the target system.

DATE validation requires the cell value to already be a native Excel date or datetime value — that is, a value Excel itself recognizes and stores internally as a date, regardless of its display format. **DATE does not parse text or numeric values into dates.** If the cell contains text or a plain number rather than a native date, the row fails DATE validation outright, with a message indicating the value is not a native date value.

This is a deliberate simplification: format parsing introduces ambiguity (which separator, which day/month order, two-digit vs four-digit year), and resolving that ambiguity is properly the responsibility of the data preparation step before validation runs — either by entering/importing the data as a true Excel date to begin with, or via a dedicated date-conversion enrichment step (planned for a future version; not available in V2.3) that produces a clean native-date column for DATE to validate.

DATE validation behavior is controlled globally by the `date_validation` block in config.yaml — there is no per-rule configuration.

A value fails DATE validation if any of the following is true:

- The value is empty (per the EMPTY definition in section 6.1) — empty date values are out of scope for DATE and should be covered by a separate EMPTY rule if required

- The value is not a native Excel date/datetime value (i.e. it is text or a plain number)

- The (native) date falls outside the `min_date` / `max_date` range, if either is defined in config.yaml

- The (native) date matches one of the values in `null_dates`, and `treat_null_dates_as` is set to `invalid`

If a value matches one of `null_dates` and `treat_null_dates_as` is set to `empty`, the value is treated as if it were empty (skipped by DATE, may be flagged separately by EMPTY). If set to `valid`, the value passes DATE validation without further checks.

### 6.9.1 date_validation Configuration Block

|                       |                                                                                                      |
|-----------------------|------------------------------------------------------------------------------------------------------|
| **Field**             | **Description**                                                                                       |
| min_date              | Optional. Earliest acceptable date (YYYY-MM-DD). Values before this are flagged as Error.            |
| max_date              | Optional. Latest acceptable date (YYYY-MM-DD). Values after this are flagged as Error.               |
| null_dates            | Optional. List of date values (YYYY-MM-DD) treated as "no date" placeholders by the target system.   |
| treat_null_dates_as   | Optional. One of: `empty`, `valid`, `invalid`. Default: `empty`. Controls how values matching `null_dates` are treated. |

Example:

```yaml
date_validation:
  min_date: "1900-01-02"
  max_date: "2154-12-31"
  null_dates:
    - "1900-01-01"
  treat_null_dates_as: "empty"
```

## 6.10 NUMERIC

Flags rows where the target column value cannot be safely imported into a numeric field of the target system.

NUMERIC validation requires the cell value to already be a native Excel numeric value (an int or float Excel itself recognizes and stores as a number) — not text. **NUMERIC does not parse text into numbers.** If the cell contains text rather than a native number, the row fails NUMERIC validation outright, with a message indicating the value is not a native numeric value.

As with DATE (section 6.9), this is a deliberate simplification: text-to-number parsing involves decimal/thousands-separator ambiguity that is properly resolved during data preparation — either by entering/importing the data as a true Excel number to begin with, or via a dedicated numeric-conversion enrichment step (planned for a future version; not available in V2.3) that produces a clean native-number column for NUMERIC to validate.

NUMERIC validation behavior is controlled globally by the `numeric_validation` block in config.yaml — there is no per-rule configuration. This block is shared with INTEGER (section 6.11).

A value fails NUMERIC validation if any of the following is true:

- The value is empty (per the EMPTY definition in section 6.1) — empty values are out of scope for NUMERIC and should be covered by a separate EMPTY rule if required

- The value is not a native Excel numeric value (i.e. it is text)

- The (native) value falls outside the `min_value` / `max_value` range, if either is defined in config.yaml

### 6.10.1 numeric_validation Configuration Block

|                       |                                                                                                      |
|-----------------------|------------------------------------------------------------------------------------------------------|
| **Field**             | **Description**                                                                                       |
| min_value             | Optional. Minimum acceptable numeric value. Values below this are flagged as Error.                  |
| max_value             | Optional. Maximum acceptable numeric value. Values above this are flagged as Error.                  |

Example:

```yaml
numeric_validation:
  min_value: 0
  max_value: 999999999
```

## 6.11 INTEGER

Flags rows where the target column value cannot be safely imported into a whole-number field of the target system. INTEGER validation shares the `numeric_validation` configuration block with NUMERIC (section 6.10.1) — `min_value`/`max_value` apply identically to both types, and the same native-value requirement applies (text is not parsed; see section 6.10).

A value fails INTEGER validation if any of the following is true:

- It fails NUMERIC validation (per section 6.10) for any of the same reasons (empty, not a native numeric value, or out of min/max range)

- The numeric value has a non-zero fractional part

**Important — ERP migration context:** INTEGER validation checks the *numeric value*, not its display formatting. A native Excel number such as 123.0 is **valid** for INTEGER, because its fractional part is zero and it will import correctly into a whole-number field. Only values with a genuinely non-zero fractional part (e.g. 123.45) fail INTEGER validation.

# 7. Function Library

The tool uses a single function registry for all functions — built-in and custom alike. Built-in enrichment functions receive their parameters from the `FunctionArguments` column in the EnrichmentRules sheet. Custom enrichment and validation functions receive their parameters from config.yaml under the `custom_functions` block.

## 7.1 Built-In Enrichment Functions

Built-in functions are bundled with the framework and registered automatically at startup. They are always available without any configuration. Their parameters are defined per rule in the `FunctionArguments` column of the EnrichmentRules sheet using the format `key=value;key=value`. Lists use comma-separated values: `key=val1,val2,val3`.

### 7.1.1 CONCATENATE

Concatenates values from two or more columns into a single string. Leading and trailing whitespace is trimmed from each source value before concatenation. Empty or null cell values are treated as empty string — not as "nan" or "None". Example use case: combining dimension policy columns from Item Master into a single lookup key for DimSetUp validation.

|               |                                                                             |
|---------------|-----------------------------------------------------------------------------|
| **Parameter** | **Description**                                                             |
| columns       | Ordered list of source column names to concatenate                          |
| separator     | String to place between values (default: empty string)                      |
| trim          | Trim whitespace from each source value before concatenation (default: true) |

Example FunctionArguments value: `columns=ItemName,ItemId;separator=_`

### 7.1.2 COUNTER

Generates a sequential counter starting from 1 for each row in the target worksheet. Returns a list of integers \[1, 2, 3, …, n\] where n is the number of data rows. This is useful for creating unique row identifiers or sequence numbers in the output.

|               |                                     |
|---------------|-------------------------------------|
| **Parameter** | **Description**                     |
| (none)        | This function accepts no parameters — leave FunctionArguments empty |

### 7.1.3 COUNTER4GROUP

Generates a counter within groups defined by one or more grouping columns, sorted by one or more sort columns. Within each group, the counter starts at 1 and increments for each row. When the grouping column value changes, the counter resets to 1.

|                  |                                                                                                                                                  |
|------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| **Parameter**    | **Description**                                                                                                                                  |
| grouping_columns | Ordered list of column names that define the grouping boundaries. Rows with identical values in all grouping columns belong to the same group.   |
| sort_columns     | Ordered list of column names to sort rows within each group. Sorting is applied before counter generation. Rows are sorted ascending by default. |

Example FunctionArguments value: `grouping_columns=CategoryId;sort_columns=ItemName`

Example use case: numbering invoice line items (counter resets per invoice ID) or sequencing stock movements within batches (counter resets per batch, sorted by date).

## 7.2 Custom Functions

Custom functions allow complex business logic to be implemented outside the core framework. They are loaded from the directory specified in `custom_functions_directory` and registered into the same single registry as built-in functions. The framework does not require modification to add new custom functions. Custom functions receive their parameters from config.yaml under the `custom_functions` block — the `FunctionArguments` column in EnrichmentRules is ignored for custom functions.

### 7.2.1 Function Registration

Custom functions are placed in the custom functions directory specified in the YAML configuration. Functions are automatically discovered and registered at startup. Function names are case-insensitive in the registry. If a custom function file fails to load or contains syntax errors, the function is not added to the registry. The failure is recorded and logged — metadata validation will fail if that function is referenced in any active rule.

It is strongly recommended that every custom function declares a module-level `REQUIRED_PARAMS` list. This list specifies which parameter names must be present in the function's config.yaml block. If declared, the metadata validator checks that all required parameters are present and raises a metadata error if any are missing. If `REQUIRED_PARAMS` is not declared, the argument check is skipped silently.

Example declaration at the top of a custom function file:

```python
REQUIRED_PARAMS = ["reference_sheet", "reference_key_column"]

def myfunction(context: dict):
    ...
```

### 7.2.2 Function Contract

All functions in the registry share the same contract requirements.

All functions must:

- Accept a context object containing relevant worksheet data, column values, and configuration values (see section 7.2.3)

- Not modify the workbook directly — the framework handles all workbook writes

- Not raise unhandled exceptions — all internal errors must be caught and returned as a runtime failure signal

Return values:

- Functions used in EnrichmentRules: return a list of values with exactly the same number of elements as there are data rows in the target worksheet.

- Functions used in ValidationRules: return one boolean result per data row (True = valid, False = invalid). Functions used with AI-type validation or custom functions may additionally return one comment string per row.

### 7.2.3 Function Context Object

The framework passes a context object to each function. External files are already imported as worksheets into the output workbook before any function is called — functions access them as regular worksheets, not as a separate data source.

|                     |                                                                                                       |
|---------------------|-------------------------------------------------------------------------------------------------------|
| **Context Element** | **Description**                                                                                       |
| target_column       | Values of the primary target column as a list                                                         |
| worksheet_data      | Full target worksheet as a dataframe                                                                  |
| named_columns       | Dictionary of specific column names to value lists, as specified in function parameters               |
| reference_worksheet | A specific worksheet dataframe if referenced in parameters (workbook sheet or imported external file) |
| technical_config    | Values from the YAML configuration file                                                               |
| row_ids             | Row identifier list aligned with the worksheet data rows                                              |

# 8. Program Execution Algorithm

The tool supports the following execution modes:

|                       |                                                                                  |
|-----------------------|----------------------------------------------------------------------------------|
| **Mode**              | **Command**                                                                      |
| Command line (Python) | python erp_migration_data_validator.py --config config.yaml                      |
| Compiled executable   | erp_migration_data_validator.exe --config config.yaml                            |
| Non-interactive       | No user prompts. Output worksheet conflicts are resolved by automatic overwrite. |

The program execution algorithm is described as a sequence of steps below.

## 8.1 Build AI Provider Registry

Before the configuration file is opened, the Provider Registry (section 2.6.3) is built — every provider module the framework ships with is registered by its `PROVIDER_NAME`. Unlike the Function Registry (section 8.5), this has no dependency on configuration or the workbook: provider modules are part of the framework's own source, not scanned from a user-configurable location. This is what makes it possible to validate `ai_provider` against it during the next step, immediately after — see 8.2, 3.8.

## 8.2 Reading the Technical Configuration File

The tool reads the technical configuration file on startup. Processing terminates immediately if any of the following conditions are violated:

- The configuration file is missing or unreadable

- Any required field is absent or contains an invalid value

- The input file does not exist or cannot be opened

- The output file path is not writable

On termination, a brief error message is printed to the console. Full details are also written to the log file — except when the log file path itself cannot yet be determined because configuration has not (or not yet) loaded successfully, in which case detail is available on the console only (see 8.12.2).

### 8.2.1 AI Connectivity Pre-Flight Check

If configuration loads successfully, `ai_enabled = true`, and `ai_connection_test_prompt` is non-empty, the framework sends it as a real request through the shared AI client module (section 6.8), which dispatches to whichever provider `ai_provider` selects (Provider Registry, section 2.6.3) — no function-specific code is involved; every AI-flavored custom function, including `ai_itemcategoryfit` (Appendix C), is unaware this check even happens. This confirms connectivity and credentials before any further processing begins — still at this same step, before any workbook is opened.

If the request fails, processing terminates immediately, with the same console/log treatment as any other Step 2 configuration error (3.4, 3.8). If `ai_connection_test_prompt` is empty or omitted, this check is skipped entirely — a bad key or unreachable service is only discovered the first time an AI validation rule actually runs, and surfaces as an ordinary runtime failure on that rule (section 8.9.2), not a termination error.

## 8.3 ExternalFiles and ImportSpec Metadata Validation

The tool opens the input file and reads the ExternalFiles and ImportSpec worksheets before copying or touching any files. If ExternalFiles does not exist, this step is skipped entirely (no external files will be imported). If ImportSpec does not exist, it is treated as empty (all csv external file columns import as text, no conversion applied; xlsx external file columns are unaffected either way, since they always carry native types).

If any validation error is found in either sheet, all errors are collected and reported together — processing does not stop at the first error. After all checks complete, if errors were found: all errors are printed to the console and written to the log file, and processing terminates before any file is copied or imported.

### 8.3.1 ExternalFiles Metadata Validation

For each row: if `Active` is empty — error, no further checks for that row. If `Active = No` — row skipped silently. If `Active = Yes` — all checks below apply.

|                   |                                                                                                                                                                   |
|-------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column**        | **Check**                                                                                                                                                         |
| Active            | Must not be empty. Must be Yes or No only.                                                                                                                        |
| SequenceNum       | Must not be empty. Must be numeric. Must be unique across all Active = Yes rows.                                                                                  |
| SheetName         | Must not be empty. Must be unique across all Active = Yes rows. Must not match any existing worksheet name in the input workbook.                                 |
| FilePath          | Must not be empty. Must be an existing, accessible file path (absolute or relative to the project root).                                                          |
| Format            | Must not be empty. Must be either `xlsx` or `csv` (case-insensitive).                                                                                            |
| OriginalSheetName | Required (not empty) when Format = `xlsx`. Must reference an existing worksheet name in the external xlsx file at the given FilePath. Must be empty when Format = `csv`. |
| CSVDelimiter      | Required (not empty) when Format = `csv`. Must be one of: comma `,` semicolon `;` pipe `\|` colon `:` space ` `.                                                 |

**Duplicate / consistency checks across all Active = Yes rows:**

| **Check** | **Rule** |
|---|---|
| SequenceNum uniqueness        | No two Active = Yes rows may have the same SequenceNum |
| SheetName uniqueness          | No two Active = Yes rows may have the same SheetName |
| FilePath + OriginalSheetName  | The combination of FilePath and OriginalSheetName must be unique across all Active = Yes rows. For csv rows (OriginalSheetName is empty), FilePath alone must be unique across csv rows. |

### 8.3.2 ImportSpec Metadata Validation

If the ImportSpec sheet does not exist, this subsection is skipped. For each row: if `Active = No` — row skipped silently. If `Active = Yes` — all checks below apply.

|                    |                                                                                                                                                                   |
|--------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column**         | **Check**                                                                                                                                                         |
| Active             | Must not be empty. Must be Yes or No only.                                                                                                                        |
| ExternalFileSheetName          | Must not be empty. Must reference the SheetName of an Active = Yes row in ExternalFiles — this is the referential integrity link between the two sheets. That ExternalFiles row's `Format` must be `csv`; a `SheetName` belonging to an xlsx row is a validation error here, not a runtime warning — xlsx files already carry native types (section 8.6.1) and are never eligible for ImportSpec conversion. |
| ColumnName         | Must not be empty. Existence of this column in the actual external file is validated at import time (section 8.6), not here — the column list is only known after the file is opened. |
| TargetType         | Must not be empty. Must be one of the supported conversion types: `text`, `date`, `numeric`, `integer` (see section 5.4.1 for conversion behavior of each).      |
| DateFormats        | Required (not empty) when TargetType = `date`. Must contain at least one format string in Python strftime syntax (e.g. `%d.%m.%Y`), separated by semicolons if multiple. |
| DecimalSeparator   | May be empty (defaults to `.`). If provided: must be a single character. Must differ from ThousandsSeparator if both are provided.                               |
| ThousandsSeparator | May be empty (no thousands separator applied). If provided: must be a single character. Must differ from DecimalSeparator if both are provided.                  |

**Duplicate / consistency checks across all Active = Yes rows:**

| **Check** | **Rule** |
|---|---|
| ExternalFileSheetName + ColumnName uniqueness | No two Active = Yes rows may have the same ExternalFileSheetName + ColumnName combination |

## 8.4 Copying the Input File to the Output File

The input file is copied to the configured output path. If a file already exists at the output path, it is overwritten without prompting. This event is logged to the console and log file.

## 8.5 Build Function Registry and Validation Type Registry

Two of the framework's three registries (section 2.6) are built at this step. The third, the Provider Registry (section 2.6.3), is built earlier, during Step 1 (8.1) — before configuration is even read — so `ai_provider` can be validated against it during Step 2 (8.2), immediately after, well before any workbook is opened.

**Function Registry:** built in two steps. First, all built-in functions (CONCATENATE, COUNTER, COUNTER4GROUP) are registered. Then, all Python modules (\*.py files) in the `custom_functions_directory` are scanned and loaded.

- Each module is imported and examined for callable functions

- Callable functions are registered by name (case-insensitive) in the single shared registry

- If a custom function file fails to load or contains syntax errors, the function is not added to the registry. The failure is recorded and logged. Metadata validation will later fail if that function is referenced in any active rule.

- Function names must be unique within the registry. If a custom function has the same name as a built-in function, or two custom functions share a name (case-insensitive), only the first loaded is registered and a warning is logged.

- For each registered function, its optional `REQUIRED_PARAMS` declaration (section 7.2.1) and its optional metadata-check function (section 2.6.2) are read and stored alongside it, for use during metadata validation (section 8.7).

**Validation Type Registry:** all built-in validation type modules (EMPTY, NULL, DUPLICATE, DUPLICATE2, FORMAT, REFERENCE, CUSTOM, AI, DATE, NUMERIC, INTEGER) are registered by type name. This registry is fixed at the framework level — it is not scanned from a directory like custom functions are, since validation types are a framework capability rather than a user-extensible one in V2.3. For each registered type, its optional metadata-check function (section 2.6.1) is read and stored alongside it, for use during metadata validation (section 8.7).

Both registries must be complete before metadata validation proceeds. All CustomFunctionName references in both EnrichmentRules and ValidationRules are validated against the Function Registry, and all ValidationType references in ValidationRules are validated against the Validation Type Registry, in the next step.

## 8.6 External File Import

External file import is handled by a dedicated Import Module. All active external files defined in the ExternalFiles worksheet are loaded and imported into the output workbook as new worksheets, in SequenceNum order.

xlsx files keep their native Excel types and need no conversion; csv files carry only text, so ImportSpec column conversion (8.6.2) applies to them alone.

### 8.6.1 Loading

- For xlsx files: the worksheet specified in `OriginalSheetName` is loaded from the external file, preserving each cell's native Excel type exactly as it is read for the primary workbook (section 8.1).
- For csv files: the full file content is loaded as text, using the `CSVDelimiter` specified in ExternalFiles.
- Column names are stripped of leading/trailing whitespace at load time.

### 8.6.2 ImportSpec Column Conversion (csv files only)

After loading, the Import Module applies per-column type conversions for any Active = Yes ImportSpec rows matching this file's `SheetName`, where that `SheetName` corresponds to a csv external file. ImportSpec rows are independent of each other — each converts exactly one column — so processing order between rows has no effect on the result. (An ImportSpec row targeting an xlsx `SheetName` cannot reach this step — it is rejected during metadata validation, section 8.3.2, before any file is imported.)

For each applicable ImportSpec row:

1. The specified `ColumnName` is located in the loaded data (case-insensitive match). If the column is not found, a warning is written to the log and that ImportSpec row is skipped — the column is imported as text.

2. The conversion defined by `TargetType` is applied to every cell in the column:
   - **date** — parsed against `DateFormats` in order; first match produces a native Excel date. Failed cells store original text and a per-cell warning is logged.
   - **numeric** — parsed per `DecimalSeparator`/`ThousandsSeparator` rules (section 5.4.1); produces a native Excel float. Failed cells store original text and a per-cell warning is logged.
   - **integer** — same as numeric, plus fractional-part-zero check; produces a native Excel integer. Failed cells store original text and a per-cell warning is logged.
   - **text** — no conversion applied.

3. Conversion warnings (failed cells) are counted per ImportSpec row and written to the log. They are NOT written to the Issues worksheet — conversion failure is an import-time data quality note, not a validation issue.

### 8.6.3 Writing to Output Workbook

Each successfully loaded (and optionally converted) dataset is added to the output workbook as a new worksheet named by its `SheetName` value from ExternalFiles.

After all imports, the output workbook contains:
- All original business data worksheets from the input file
- All original metadata worksheets (ValidationRules, EnrichmentRules, ExternalFiles, and ImportSpec if present)
- One imported worksheet per successfully loaded active external file
- Possibly Issues and Summary worksheets if they were present in the input file — these will be overwritten later

### 8.6.4 Import Failures

If an active external file cannot be loaded (file not found, format error, access denied, or worksheet not found for xlsx):
- The failure is recorded
- Processing continues with the remaining external files
- All validation and enrichment rules that reference this SheetName are skipped during execution
- The failure is reported to the terminal, Summary worksheet, and log file — it is NOT written to the Issues worksheet

## 8.7 Rules Metadata Validation

All detected errors are collected before reporting — processing does not stop at the first error. If any error is found after all checks complete, all errors are printed to the terminal, written to the Summary worksheet and log file, and execution terminates before any enrichment or validation runs.

For both ValidationRules and EnrichmentRules, the Active column is evaluated first for each row. If Active is empty — error, no further checks for that row. If Active = No — row skipped silently. If Active = Yes — all checks below apply.

### 8.7.1 Common Checks for ValidationRules and EnrichmentRules

|             |                                                                                                                                                                                                                    |
|-------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column**  | **Check**                                                                                                                                                                                                          |
| Active      | Must not be empty. Must be Yes or No only.                                                                                                                                                                         |
| SequenceNum | Must not be empty. Must be numeric. Must be unique across all Active = Yes rows within the same rules worksheet.                                                                                                   |
| RuleCode    | Must not be empty. Must be unique across all Active = Yes rows within the same rules worksheet.                                                                                                                    |
| RuleName    | Must not be empty. Must be unique across all Active = Yes rows within the same rules worksheet.                                                        |
| SheetName   | Must not be empty. A worksheet with this name must exist in the output workbook (including imported external file worksheets).                                                                                     |
| ColumnName  | Must not be empty. The column must exist in the worksheet named SheetName. Exception: for EnrichmentRules, ColumnName is the name of the column to be generated — it does not yet exist and this check is skipped. |
| Color       | Optional. If not empty: must be a built-in named color (RED, YELLOW, GREEN, ORANGE, BLUE, GREY), a name defined in config predefined_colors, or a valid hex code (#RRGGBB or \#RGB).                               |

### 8.7.2 Additional Checks for EnrichmentRules Only

|                            |                                                                                            |
|----------------------------|--------------------------------------------------------------------------------------------|
| **Column / Check**         | **Rule**                                                                                   |
| CustomFunctionName         | Must not be empty. Must reference an existing registered function in the function library. |
| Duplicate target detection | No two Active = Yes rows may target the same SheetName + ColumnName combination.           |

### 8.7.3 Additional Checks for ValidationRules Only

|                                |                                                                                                                                                        |
|--------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Column / Check**             | **Rule**                                                                                                                                               |
| ValidationType                 | Must not be empty. Must be one of: EMPTY, NULL, DUPLICATE, DUPLICATE2, FORMAT, REFERENCE, CUSTOM, AI, DATE, NUMERIC, INTEGER.                          |
| Severity                       | Must not be empty. Must be either Error or Warning.                                                                                                    |
| ReferenceDataSheet             | Required (not empty) only when ValidationType = REFERENCE. Should be differ from SheetName value. The worksheet or external file SheetName must exist in the output workbook. Optional when ValidationType = AI — if provided, the same existence check and must-differ-from-SheetName check apply (SheetName already receives automatic failed-import protection on its own, so a ReferenceDataSheet equal to SheetName would be redundant and is treated as an error), but this is solely for failed-import skip protection (section 8.6.4/6.8) — it performs no data lookup for AI, unlike REFERENCE. Not used for DUPLICATE2. |
| ReferenceDataColumn            | Required (not empty) when ValidationType = REFERENCE or DUPLICATE2. For REFERENCE: column must exist in ReferenceDataSheet (or be a pending EnrichmentRules TargetColumn, per section 2.5.3). For DUPLICATE2: column must exist in the same sheet as ColumnName (or likewise be a pending EnrichmentRules TargetColumn), and must not be the same column as ColumnName (use DUPLICATE instead in that case). Optional when ValidationType = AI — if non-empty, it must either already exist on ReferenceDataSheet or be a pending EnrichmentRules TargetColumn for ReferenceDataSheet (section 2.5.3); it performs no data lookup for AI, unlike REFERENCE — see the consistency check below and section 6.8 for its actual meaning. |
| ReferenceDataSheet / ReferenceDataColumn consistency (AI only) | If ValidationType = AI and ReferenceDataSheet is empty, ReferenceDataColumn must also be empty — a non-empty ReferenceDataColumn with no corresponding ReferenceDataSheet is a metadata error. If ReferenceDataSheet is not empty, ReferenceDataColumn remains fully optional in either direction: it may be left empty (no pending-column dependency), or set to a value that satisfies the ReferenceDataColumn check above. This check does not apply to REFERENCE or DUPLICATE2, where ReferenceDataColumn has its own independent requiredness rule regardless of ReferenceDataSheet. |
| CustomFunctionName             | Required (not empty) when ValidationType = CUSTOM, or when ValidationType = AI and Active = Yes and ai_enabled = true (if ai_enabled = false, the row is soft-skipped per section 6.8 and this check does not apply). Must reference an existing registered function in the function library. |
| Duplicate validation detection | No two Active = Yes rows may have an identical combination of: SheetName + ColumnName + ValidationType + CustomFunctionName.                           |
| Type-specific metadata checks  | If the rule's ValidationType module declares an optional metadata-check function (section 2.6.1), it is called here and any returned error messages are added to the metadata validation result. |
| Function-specific metadata checks | If ValidationType = CUSTOM or AI, and the referenced function declares an optional metadata-check function (section 2.6.2), it is called here and any returned error messages are added to the metadata validation result, in addition to the REQUIRED_PARAMS check (section 7.2.1). |

## 8.8 Enrichment Execution

All enrichment rules execute before any validation rules. Enrichments are processed in ascending SequenceNum order.

### 8.8.1 Generated Columns

Each enrichment rule creates one new column in the target worksheet. Generated columns are:

- Appended after the last non-empty column in the worksheet

- Added in enrichment execution order (left to right)

- Permanently retained in the output workbook and visible to the user

- Available to all subsequent validation rules that reference them

Temporary columns are not supported. All generated columns appear in the output workbook.

### 8.8.2 Enrichment Function Output

An enrichment function must return a list of values with exactly the same number of elements as there are data rows in the target worksheet. If the returned list size does not match, this is treated as a runtime failure (see section 8.8.2).

### 8.8.3 Enrichment Runtime Failures

Enrichment rules are independent of each other — no enrichment-to-enrichment dependencies exist or are permitted by the metadata structure. If an enrichment function raises an unhandled exception or returns an invalid result:

- The target generated column is not created

- The enrichment is marked as failed

- All validation rules that reference the column this enrichment was intended to generate are skipped

- The failure is written to the terminal, Summary worksheet, and log file

- Processing continues with the next enrichment rule

## 8.9 Validation Execution

Validation rules are processed in ascending SequenceNum order. For each active validation rule, the engine looks up the rule's `ValidationType` in the Validation Type Registry (section 2.6.1) and delegates execution to that type's module — the engine itself contains no type-specific logic. Execution then consists of two steps:

1.  The resolved validation type module evaluates the business data according to the rule definition (for `ValidationType = CUSTOM`, this module in turn looks up `CustomFunctionName` in the Function Registry and calls it, per section 2.6.2)
2.  Apply the Issues Worksheet Population Algorithm to the results

### 8.9.1 Issues Worksheet Population Algorithm

The Issues worksheet is populated during validation execution, not as a separate post-processing step. For each validation rule, a single combined module handles both cell highlighting and Issues row generation. This module is called once per validation rule, immediately after the rule function returns its results.

The combined module performs the following steps:

1.  Iterate over all rows in the validation result set.
2.  For each row where the result is False (validation failed): highlight the target cell in the business data worksheet using the Color defined in the validation rule.
3.  For the same failed row: append one new row to the Issues worksheet.

|                   |                                       |                                                                                                                       |
|-------------------|---------------------------------------|-----------------------------------------------------------------------------------------------------------------------|
| **Issues Column** | **Source**                            | **Value**                                                                                                             |
| IssueNum          | Auto-counter                          | Next sequential integer, incremented for each issue row added                                                         |
| SheetName         | ValidationRules                       | SheetName from the current validation rule                                                                            |
| ColumnName        | ValidationRules                       | ColumnName from the current validation rule                                                                           |
| RowNumber         | Business data                         | Row number of the failed row in the source worksheet (1-based, excluding the header row)                              |
| CellContent       | Business data                         | Current value of the target cell in the failed row                                                                    |
| RuleCode          | ValidationRules                       | RuleCode from the current validation rule                                                                             |
| RuleName          | ValidationRules                       | RuleName from the current validation rule                                                                             |
| Severity          | ValidationRules                       | Severity from the current validation rule                                                                             |
| Comment           | AI function or custom function result | Comment text returned by an AI validation or a custom validation function. Empty if the function returned no comment. |

The same business data cell may appear in multiple Issues rows if more than one validation rule flagged it. Each violation generates its own independent Issues row with its own IssueNum. If multiple rules flag the same cell with different colors, the last applied color takes effect.

Issues rows are written in the order validation rules execute (ascending SequenceNum). No sorting or grouping is applied. IssueNum reflects insertion order. The Issues worksheet header row is created once before validation execution begins.

### 8.9.2 Validation Runtime Failures

Validation rules are independent of each other — no validation-to-validation dependencies exist. If a validation function raises an unhandled exception:

- The validation is marked as failed

- No issues are generated for this rule

- The failure is written to the terminal, Summary worksheet, and log file

- Processing continues with the next validation rule

**Exception:** AI-flavored validation rules (section 6.8) process rows in batches rather than as a single unit. A failure isolated to one batch does not fail the entire rule the way a single-shot exception does here — see section 6.8 for the `partial` status, under which issues from successful batches ARE generated and written normally, alongside a failure record for the rows in the failed batch.

### 8.9.3 Dependency Skip Chain

The only cross-phase dependency supported is: a validation rule referencing a column generated by an enrichment rule. If that enrichment failed at runtime, all dependent validation rules are skipped automatically. Skip events are written to the terminal, Summary worksheet, and log file with the failed enrichment identified as the cause.

## 8.10 Issues Worksheet Formatting

After all validation rules have executed, auto-filter and auto-width are applied to all columns of the Issues worksheet.

## 8.11 Summary Worksheet Generation

The Summary worksheet is generated after Issues formatting completes. It provides high-level statistics for the entire processing run.

|                       |                                                                                                                                                                                                                   |
|-----------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **Section**           | **Contents**                                                                                                                                                                                                      |
| Run information       | Execution timestamp, input file name, output file name                                                                                                                                                            |
| Worksheet statistics  | List of processed worksheets with row counts                                                                                                                                                                      |
| Enrichment statistics | One row per enrichment rule showing RuleCode, RuleName, and execution status (executed / skipped / failed). Total row at the bottom with counts for each status.                                                  |
| Validation statistics | One row per validation rule showing RuleCode, RuleName, execution status (executed / skipped / failed / partial), and issue count for that rule. `partial` applies only to batched AI validation rules (section 6.8) where some batches succeeded and others failed — see Appendix C.3 for the exact status logic. Total row at the bottom with counts for each status and total issue count. |
| Issue statistics      | One row per validation rule showing RuleCode, RuleName, count by severity (Error / Warning), and count by worksheet. Total row at the bottom with overall counts.                                                 |
| External files        | Successfully loaded files, failed files, and (if ImportSpec is present) count of cell-level conversion warnings per file                                                                                          |

## 8.12 Logging

### 8.12.1 Console Output

The console displays processing progress at each major step, validation messages, runtime warnings, and runtime failures. All messages displayed on the console are also written to the log file — with one narrow exception during Steps 1-2, before the log file's location is known; see 8.12.2.

### 8.12.2 Log File

The log file contains full technical detail including stack traces for runtime failures, dependency skip chains, skipped rule lists, external file loading results, and AI call results. The log file path is defined in the YAML configuration.

Because the log file path is itself part of the configuration, no log file exists yet while Steps 1-2 (8.1-8.2) — building the Provider Registry and then reading and validating configuration — are in progress. During this window, all output is shown live on console and simultaneously held in memory at the same level of detail the log file would otherwise show, rather than written to a file (there is none yet to write to).

- If configuration loads successfully: everything logged since process start — not only from this point forward — is written into the now-known log file before processing continues. The log file therefore ends up complete, with no gap. Logging then continues to both console and the log file for the remainder of the run (8.12.1).
- If configuration fails to load: the termination error is shown on console only, whether it happened during Step 1 or Step 2. Nothing from either step, including that error, is persisted to a log file, because no log file destination was ever established for this run.

# 9. Case Sensitivity Rules

The framework default is **case-insensitive** for all data value comparisons. The only exception is RuleCode, which is an identifier (analogous to a variable name) and is therefore case-sensitive.

|                                            |                                       |
|---------------------------------------------|-----------------------------------------|
| **Element**                                  | **Case Sensitivity**                    |
| RuleCode values                              | Case-sensitive (identifiers)            |
| Worksheet names                              | Case-insensitive                        |
| Column names                                 | Case-insensitive                        |
| Function names                               | Case-insensitive                        |
| SheetName values in ExternalFiles            | Case-insensitive                        |
| ColumnName values in ImportSpec               | Case-insensitive                        |
| Standard (non-CUSTOM) validation execution   | Case-insensitive — framework default    |
| Built-in function execution                  | Case-insensitive — framework default    |
| Default for custom function execution        | Case-insensitive — framework default. Custom functions may implement different behavior internally if their business logic requires it, but should document any deviation. |

For all case-insensitive comparisons, leading/trailing whitespace is also trimmed before comparison.

# 10. Out of Scope for V2

- Graphical user interface

- Web interface

- Direct ERP system connectivity

- AI enrichment (AI validations only in V2)

- Parallel or multi-threaded processing

- Partial worksheet execution or row filtering

- Metadata-level conditional logic engine

- Automatic dependency declaration — all dependencies are inferred from metadata structure

# Appendix A: Validation Type Quick Reference

| **ValidationType** | **Built-in Logic** | **Requires CustomFunctionName** | **Requires ReferenceDataSheet** | **Requires ReferenceDataColumn** | **Requires AI Config**                                          |
|--------------------|--------------|---------------------------------|---------------------------------|----------------------------------|-----------------------------------------------------------------|
| EMPTY              | Yes          | No                              | No                              | No                               | No                                                              |
| NULL               | Yes          | No                              | No                              | No                               | No                                                              |
| DUPLICATE          | Yes          | No                              | No                              | No                               | No                                                              |
| DUPLICATE2         | Yes          | No                              | No                              | Yes — second column on same sheet| No                                                              |
| FORMAT             | Yes          | No                              | No                              | No                               | No                                                              |
| REFERENCE          | Yes          | No                              | Yes                             | Yes                              | No                                                              |
| CUSTOM             | No           | Yes                             | Optional                        | Optional                         | No                                                              |
| AI                 | No           | Yes (Active=Yes + ai_enabled=true) | Optional — failed-import protection switch only, no lookup (6.8) | Optional — pending-enrichment-column flag only, no lookup (6.8) | Only if ai_enabled=true — ai_model + ai_api_key required in config at load time (section 3.8); ai_enabled=false soft-skips the rule rather than erroring |
| DATE               | Yes          | No                              | No                              | No                               | No — date_validation block in config.yaml is optional; requires the cell to be a native Excel date value |
| NUMERIC            | Yes          | No                              | No                              | No                               | No — numeric_validation block in config.yaml is optional; requires the cell to be a native Excel numeric value |
| INTEGER            | Yes          | No                              | No                              | No                               | No — numeric_validation block in config.yaml is optional (shared with NUMERIC); requires the cell to be a native Excel numeric value |

# Appendix B: Stock Balance Dimension Validation

## B.1 Purpose

This is a predefined custom function included in the V2 distribution. It validates stock balance records against configurable dimension policies — the required, forbidden, and quantity-restricted attributes each stock record must carry before ERP import. It is registered as a standard custom function and is available for use, but not mandatory.

## B.2 Business Problem

When loading opening stock balances into an ERP system, each stock record must provide exactly the dimension values its governing policy requires. Missing required dimensions cause import errors. Unexpected dimension values cause loss of dimension information during import. This function catches both categories before import, and also validates quantity restrictions defined per policy where applicable.

A stock record's governing policy can be driven by different combinations of business entities depending on the ERP and the client's configuration — by the item alone, by the item together with the warehouse, or by other combinations. This function's key-composition mechanism (B.3) is configurable enough to express any of these.

## B.3 Design Overview

This function is invoked like any other custom validation function: through a ValidationRules row with `ValidationType = CUSTOM` and `CustomFunctionName = dimpolicyvalidation`. The sheet holding the stock balance records is the sheet named in that row's `SheetName` field — the same mechanical role `SheetName` plays for every other validation type. The function's own parameters, described below, are configured separately in config.yaml, under `custom_functions.dimpolicyvalidation`.

The dimension policy lookup key is composed from one or more **lookup blocks**, defined in the `policy_key_lookups` parameter. Each block specifies: a Stock Balance column whose value drives a lookup, the reference sheet and column to match it against, and one or more attribute columns to pull from the matched reference row.

All attribute values from all blocks, taken in list order (block order, then within-block column order), are concatenated with `dim_policy_key_delimiter` to form the final composed key. Example configuration shapes:

|                              |                                                                                                            |
|------------------------------|------------------------------------------------------------------------------------------------------------|
| **Policy driven by**         | **Configuration shape**                                                                                    |
| Item alone                   | One block: Item ID → Item Master → one or more key columns                                                 |
| Item and warehouse together  | Two blocks: Item ID → Item Master → key column(s); Warehouse ID → Warehouse Master → key column(s)          |
| Any other combination        | Any number of blocks, each pulling any number of key columns                                                |

Once the composed key is built, the function looks it up in the Dim Policy sheet to find the one governing policy row (B.6). That row's mapped columns are then checked, one by one, against the corresponding cells on the Stock Balance row: each mapped dimension must be present or absent depending on whether its Dim Policy value is Yes or No, and if the policy row specifies a quantity restriction, the stock quantity must match it. The full step-by-step procedure is in B.7.

## B.4 YAML Configuration Parameters

All parameters below are configured under `custom_functions.dimpolicyvalidation` in config.yaml.

**`policy_key_lookups`** — a list of one or more lookup blocks. Each block builds one segment of the composed policy key, and blocks are combined in list order (B.3). Each block contains:

- **`source_column`** — the Stock Balance column whose value drives this block's lookup
- **`reference_sheet`** — the sheet to look up that value in (nested under this block, describes where to search for the value read from `source_column`)
- **`reference_key_column`** — the column on `reference_sheet` to match against (nested under this block, together with `reference_sheet` these two fields locate the matching row)
- **`policy_key_columns`** — an ordered list of one or more column names on that same `reference_sheet`; once the matching row is found (via `reference_key_column`), these are the column(s) whose values are pulled from it into the composed key, in list order

In short: within each block, `source_column` says which stock value to look up; `reference_sheet` + `reference_key_column` say where and how to look it up; `policy_key_columns` says what to extract once found.

Top-level (scalar) parameters, alongside `policy_key_lookups`:

|                                       |                                                                                                                                    |
|----------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------|
| **Field**                              | **Role**                                                                                                                              |
| dim_policy_key_delimiter               | Separator used when concatenating all key components across all blocks, in the order described above. Optional; default is empty string. |
| dim_policy_sheet                       | Sheet holding the dimension policy table                                                                                             |
| dim_policy_key_column                  | Column on `dim_policy_sheet` holding the composed policy key, to be matched against the key built in B.3                            |
| stock_to_dim_policy_columns_mapping    | Mapping of Stock Balance column names to Dim Policy column names, for the Yes/No required/forbidden check (B.7, step 4)             |
| stock_quantity_column                  | Column on Stock Balance holding the quantity to check against the policy's Quantity Restriction (B.7, step 5)                       |

## B.5 Config Validation

### B.5.1 Config Schema Validation (structural — checked at config load time, before any file is touched)

|      |                                                                                                     |
|------|-------------------------------------------------------------------------------------------------------|
| **#**| **Check**                                                                                              |
| A.1  | `policy_key_lookups` is present and is a non-empty list                                               |
| A.2  | Every block has all four required keys: `source_column`, `reference_sheet`, `reference_key_column`, `policy_key_columns` |
| A.3  | `policy_key_columns` in every block is a non-empty list (at least one entry)                          |
| A.4  | `dim_policy_sheet` and `dim_policy_key_column` are present and non-empty strings                      |
| A.5  | `stock_to_dim_policy_columns_mapping` is present and is a non-empty mapping                           |
| A.6  | `stock_quantity_column` is present and non-empty                                                      |
| A.7  | `dim_policy_key_delimiter` — optional; if absent, defaults to empty string                            |

All schema errors found are collected and reported together, not stopped at the first one, per the framework's standard config-validation behavior (section 8.2/8.7). Processing terminates before any row is touched if any error is found.

### B.5.2 Cross-Reference Validation (checked once at load time, against the actual workbook — before row-by-row processing begins)

|      |                                                                                                     |
|------|-------------------------------------------------------------------------------------------------------|
| **#**| **Check**                                                                                              |
| B.1  | Each `source_column` exists as a column on the Stock Balance sheet                                    |
| B.2  | Each `reference_sheet` exists (worksheet or registered external file)                                 |
| B.3  | Each `reference_key_column` exists as a column on its corresponding `reference_sheet`                 |
| B.4  | Every column named in `policy_key_columns` exists as a column on that same `reference_sheet`          |
| B.5  | `dim_policy_sheet` exists                                                                              |
| B.6  | `dim_policy_key_column` exists as a column on `dim_policy_sheet`                                      |
| B.7  | Every key and value in `stock_to_dim_policy_columns_mapping` exists as a column on the Stock Balance sheet and `dim_policy_sheet` respectively |
| B.8  | `stock_quantity_column` exists as a column on the Stock Balance sheet                                 |

Same collect-and-report-together, hard-stop-before-processing behavior as schema validation.

### B.5.3 Data Integrity Validation (checked once, against the loaded reference data — not per stock row)

|      |                                            |                                                                              |
|------|--------------------------------------------|--------------------------------------------------------------------------------|
| **#**| **Check**                                  | **Why**                                                                         |
| C.1  | No duplicate values in `dim_policy_key_column` on `dim_policy_sheet` | A duplicate composed key makes policy lookup ambiguous            |
| C.2  | Every cell in the mapped dimension columns on `dim_policy_sheet` contains only `Yes` or `No` | Anything else is unparseable by the row-level Yes/No check |
| C.3  | `Quantity Restriction` column values, where non-empty, are numeric | Feeds a numeric comparison per row                          |

C.1 is a hard stop, same as schema and cross-reference validation — a duplicate policy key is a corrupted lookup table, not a row-level ambiguity.

### B.5.4 Delimiter Selection — Consultant Responsibility

The Dim Policy sheet and its composed-key column are authored manually by the consultant setting up the migration. The chosen delimiter must not collide with the actual data feeding the key — if a delimiter character can appear inside a key-component value, differently-composed keys can concatenate into the same string and silently resolve to the wrong policy. Avoiding this is the consultant's responsibility when authoring the Dim Policy sheet and choosing a delimiter; the tool does not scan for this at runtime.

## B.6 Dim Policy Sheet Structure

The Dim Policy sheet must contain at minimum the following columns:

|                          |                                                                                                                             |
|--------------------------|---------------------------------------------------------------------------------------------------------------------------|
| **Column**               | **Description**                                                                                                           |
| Dim Policy Key Column    | Policy lookup key — must match the composed key built per B.3/B.7. Column name is defined in `dim_policy_key_column`.     |
| Mapped dimension columns | One column per mapped dimension. Each cell contains Yes or No. Column names are defined in `stock_to_dim_policy_columns_mapping` values. |
| Quantity Restriction     | Optional numeric column. If a cell is not empty, the stock quantity for rows matching this policy must equal this value. |

## B.7 Logic Chain

For each row on the Stock Balance sheet, the function executes the following steps in order:

**Step 1 — Build the composed key.**

For each block in `policy_key_lookups`, in list order:

&nbsp;&nbsp;a. Read the value of `source_column` on the current row. If this value is empty (per section 2.5.2/6.1), the row is skipped entirely — it passes this validation without producing an issue. A separate EMPTY rule should be used on that column if empty values need to be flagged in their own right.

&nbsp;&nbsp;b. Look up that value in `reference_sheet`, matching `reference_key_column`. If no matching row is found, the row fails this validation, with a message naming the block's source column and reference sheet (B.8).

&nbsp;&nbsp;c. Otherwise, read the value of each column listed in `policy_key_columns` from the matched row, in list order, and append them to the growing sequence of key components.

**Step 2 — Compose the final key**, by concatenating all key components collected across all blocks — in block order, then within-block order — using `dim_policy_key_delimiter`.

**Step 3 — Look up the composed key** in `dim_policy_sheet` (`dim_policy_key_column`). If no matching row is found, the row fails this validation (B.8).

**Step 4 — Check the mapped dimension columns.** For each pair in `stock_to_dim_policy_columns_mapping`, read the Yes/No value in the corresponding Dim Policy column on the matched policy row:

- **Yes** → the mapped cell on the Stock Balance row must be **not empty**
- **No** → the mapped cell on the Stock Balance row must be **empty**

Any mismatch adds a failure message for that specific column (B.8); the check continues across all mapped columns rather than stopping at the first mismatch.

**Step 5 — Check the quantity restriction.** Read the `Quantity Restriction` value on the matched policy row — this column is numeric and optional, so an empty value means no restriction applies and this step is skipped. If not empty, the cell in `stock_quantity_column` on the Stock Balance row must equal this value; a mismatch adds a failure message (B.8).

**Step 6 — Report.** If the row failed at step 1b or step 3, that single failure message is written to the Comment column in the Issues worksheet, and steps 4–5 are not evaluated for that row. Otherwise, any failure messages accumulated across steps 4 and 5 are written together to the Comment column, semicolon-separated.

## B.8 Validation Result Messages

- '\[source_column\] value \[value\] not found in \[reference_sheet\] — dimension policy cannot be determined'

- 'Dimension policy key \[key\] not found in Dim Policy sheet'

- 'Dimension column \[Stock Column Name\] is required by policy \[policy key\] but is empty'

- 'Dimension column \[Stock Column Name\] must be absent for policy \[policy key\] but value \[value\] is present'

- 'Quantity restriction for policy \[policy key\] requires \[N\] but found \[M\] in column \[Stock Quantity Column\]'

# Appendix C: AI Item/Category Fit Validation

## C.1 Purpose

This is a predefined custom function included in the V2 distribution, registered as a standard AI-flavored custom function (`ValidationType = AI`, `CustomFunctionName = ai_itemcategoryfit`). It judges whether an item's name is semantically relevant to the category it's assigned to, using an AI model. It is available for use — but not mandatory — and serves as the reference example for writing other AI-flavored validation functions, per section 6.8.

This function is provider-agnostic — it works with whichever AI provider `ai_provider` selects (Provider Registry, section 2.6.3), and never itself touches request/response mechanics. As of this revision, `anthropic` (Appendix F) is the one provider module shipped, so `ai_model`/`ai_api_key`/`ai_provider_settings.anthropic.*` need to describe an Anthropic account for this example to work today — using a different, not-yet-shipped provider means adding a new provider module (section 2.6.3), not changing this function's `custom_functions.ai_itemcategoryfit` configuration at all.

## C.2 Business Problem

Item master data is often maintained by many different people over time, and category assignment is sometimes done quickly or inconsistently — an item can end up filed under a category that doesn't actually describe it. Rule-based validation types (REFERENCE, DUPLICATE, etc.) can only confirm a Category ID exists; they cannot judge whether the assignment actually makes sense. This function uses an AI model to make that semantic judgment and flag likely misclassifications for human review, without blocking the migration outright — it reports at Warning severity, not Error.

## C.3 Field Roles

Two completely independent layers of information are in play here, and neither one substitutes for the other:

**On the ValidationRules row — purely mechanical, no data-gathering role:**

| Field | Role | In this example |
|---|---|---|
| `SheetName` | Which sheet the engine iterates, giving the function its row data (`context["worksheet_data"]`, same as CUSTOM, section 6.7) | ItemData |
| `ColumnName` | The **highlighted cell** on a flagged row — always `ColumnName`, exactly as for every other validation type, with no exception for AI | CategoryId — **not** Item Name, even though Item Name is the text actually being judged |
| `ReferenceDataSheet` | Optional. An opt-in switch for `failed_sheets` protection only (section 6.8) — no lookup role | CategoryMaster |
| `ReferenceDataColumn` | Optional. Flags a pending-enrichment-column dependency only (section 6.8) — no lookup role | (empty — Category Description already exists, nothing pending) |

**In `custom_functions.ai_itemcategoryfit` — everything the function actually needs to gather and compare data, fully self-contained:**

| Parameter | Role | In this example |
|---|---|---|
| `source_sheet` | Sheet the function reads its source text from | ItemData |
| `source_key_column` | Join key on the source side | CategoryId |
| `source_text_column` | Column holding the text to evaluate | Item Name |
| `reference_sheet` | Sheet the function reads its reference text from | CategoryMaster |
| `reference_key_column` | Join key on the reference side | CategoryId |
| `reference_text_column` | Column holding the text to compare against | Category Description |

**One hard constraint, not just a convention:** `source_sheet` **must** match `SheetName` exactly. `SheetName` is what the framework already loads into `context["worksheet_data"]`, and CUSTOM's existing dispatch mechanism (section 6.7) maps a function's returned per-row results back to that data *positionally* — row 1 of the response to row 1 of `context["worksheet_data"]`, and so on. If `source_sheet` named a different sheet, the function would need to discard the data it was already handed and reload something else independently, and the positional mapping back to `SheetName`'s rows would no longer correspond to anything meaningful. `source_key_column`, by contrast, has no such constraint — it doesn't need to relate to `ColumnName` at all, since `ColumnName` is purely a highlight target and never participates in the function's own lookup logic.

`reference_sheet` has a softer version of the same concern: it isn't positionally mapped the way `source_sheet` is (reference rows are found by key lookup, not by row order), so it *can* technically differ from `ReferenceDataSheet` without breaking anything mechanically — but if it does, `ReferenceDataSheet`'s `failed_sheets` protection would be protecting the wrong sheet, which defeats its purpose. In practice, keep them matching.

## C.4 Logic Chain

For each active row on `SheetName` (ItemData):

1.  If `source_key_column`'s value (CategoryId) is empty, the row is skipped — it passes without producing an issue. A separate EMPTY rule should be used if empty CategoryId values need to be flagged.
2.  The CategoryId is looked up in `reference_sheet` (CategoryMaster), matching `reference_key_column`. **If not found, the row is skipped** — it passes without producing an issue. A separate REFERENCE rule should be used if invalid CategoryId values need to be flagged; this function's only concern is relevance, not existence.
3.  The function reads the text in `source_text_column` on the primary row, and the text in `reference_text_column` on the matched reference row.
4.  Matched (source text, reference text) pairs are collected across rows and sent to the configured AI service in batches of `batch_size` rows (default 20), via the shared AI client module — see section 6.8 for the batching mechanism, verdict wire format, and the resulting rule-status logic (`executed`/`partial`/`failed`), all of which are standard across every AI-flavored function, not specific to this one.
5.  Alongside each batch's row pairs, the function also sends the **full list of `reference_text_column` values across every row of `reference_sheet`** (Category Description only — `reference_key_column`/Category ID values are never sent) — sent once per batch request, not once per row. This gives the AI the complete set of real categories to recommend from when it judges a row's assigned category doesn't fit, rather than only ever seeing the one category already assigned to that row.
6.  Rows verdicted `F` (per section 6.8's format) fail with the severity configured on the rule (Warning, in the reference example); the AI's comment text — including any recommended category, in the AI's own words — is written to the Comment column in the Issues worksheet exactly as returned, with no verification against the real category list (the same plain pass-through every other validation type already uses for its own comment/message text). `ColumnName`'s cell (CategoryId on ItemData, per C.3 — never `source_text_column`'s cell) is highlighted.
7.  If the AI service call fails for a batch, section 6.8 and 8.9.2's standard runtime-failure and `partial`-status handling applies — nothing specific to this function.

**Prompt requirement:** the `prompt` parameter must explicitly instruct the AI to (a) use a category's exact name as given in the supplied list when recommending it, so the comment stays grounded in real categories even without code-level verification, and (b) say so plainly if it judges that *no* category in the supplied list fits the item at all, rather than forcing a recommendation from a list that doesn't actually contain a good match. See C.5 for the reference example's exact wording.

## C.5 YAML Configuration Parameters

|                        |                                                                                          |
|------------------------|------------------------------------------------------------------------------------------|
| **Parameter**          | **Description**                                                                          |
| prompt                 | Natural-language instruction describing what the AI should judge for each row.           |
| source_sheet           | Must match SheetName on the ValidationRules row exactly (see C.3 — this is a hard requirement, not a convention). |
| source_key_column      | Join key column on source_sheet. Independent of ColumnName — no relationship required.   |
| source_text_column     | Column on source_sheet holding the text to evaluate.                                     |
| reference_sheet        | Sheet holding the reference data. Should match ReferenceDataSheet if that's set, so its failed-import protection actually covers the sheet in use (see C.3). |
| reference_key_column   | Join key column on reference_sheet.                                                       |
| reference_text_column  | Column on reference_sheet holding the text to compare against. Also the source of the full category list sent once per batch (C.4, step 5), across every row of reference_sheet, not just matched rows. |
| batch_size             | Optional. Number of rows sent per AI API call (default: 20 — proposed default, adjust to taste). |

## C.6 Validation Result Messages

- The AI's own returned comment text, written verbatim to the Comment column, for `F`-verdicted rows only — no verification or matching against the real category list, per C.4. Typical examples, depending on what the AI decides:
  - A recommendation: *"Item name doesn't match Widgets — closer to Fasteners."*
  - No fit found in the supplied list at all: *"Item name doesn't match Widgets, and no category in the supplied list looks like a good fit either."*

- On an AI service failure for a batch: standard runtime failure handling and message per section 8.9.2, not a custom message specific to this function.

# Appendix D: Example YAML Configuration

The following example shows a complete configuration for the Stock Balance Dimension Validation scenario, using an item+warehouse policy (B.3) — the shape needed by ERPs such as SAP or Infor LN, where dimension control isn't a pure item attribute:

```yaml
# --- Framework (mandatory) ---
input_file: data/stock_migration.xlsx
output_file: data/stock_migration_checked.xlsx
log_file: logs/run_20260501.log
custom_functions_directory: functions/

# --- Formatting (optional) ---
format_special_chars: "!@#$%^&*()"
predefined_colors:
  PURPLE: "#7030A0"
  TEAL: "#008080"

# --- AI (optional) ---
ai_enabled: false
# ai_provider: anthropic
# ai_model: claude-haiku-4-5-20251001
# ai_api_key: your-api-key-here
# ai_connection_test_prompt: "Respond with the single word OK."

# ai_provider_settings:
#   anthropic:
#     ai_endpoint: https://api.anthropic.com/v1/messages
#     ai_api_version: "2023-06-01"

# --- Custom Functions (optional) ---
custom_functions:
  dimpolicyvalidation:
    policy_key_lookups:
      - source_column: Item ID
        reference_sheet: ItemsMaster
        reference_key_column: Item ID
        policy_key_columns:
          - Tracking Numbers
      - source_column: Warehouse ID
        reference_sheet: WarehousesMaster
        reference_key_column: Warehouse ID
        policy_key_columns:
          - Location Control
    dim_policy_key_delimiter: "|"
    dim_policy_sheet: DimPolicy
    dim_policy_key_column: DimPolicySetUp
    stock_to_dim_policy_columns_mapping:
      Lot Number: Lot Number
      Serial Number: Serial Number
      Location ID: Location ID
    stock_quantity_column: On-Hand
  ai_itemcategoryfit:
    prompt: "Evaluate whether the item name is relevant to the assigned
      category description. If not, recommend a better-fitting category
      from the supplied category list, using its name exactly as given
      in that list. If no category in the supplied list fits the item
      at all, say so directly rather than forcing a recommendation from
      the list."
    source_sheet: ItemData
    source_key_column: CategoryId
    source_text_column: "Item Name"
    reference_sheet: CategoryMaster
    reference_key_column: CategoryId
    reference_text_column: "Category Description"
    batch_size: 20
```

See Appendix C for the full worked example of this function — logic chain, business rationale, and configuration reference.

The `ai_itemcategoryfit` function above pairs with a ValidationRules row like:

```
SheetName=ItemData, ColumnName=CategoryId,
ReferenceDataSheet=CategoryMaster, ReferenceDataColumn=(empty),
ValidationType=AI, CustomFunctionName=ai_itemcategoryfit,
Severity=Warning
```

`ReferenceDataColumn` is empty here because Category Description already exists — there's no pending enrichment dependency to flag. `ReferenceDataSheet` is set purely to give this rule failed-import skip protection for CategoryMaster; the function's own `reference_sheet`/`reference_key_column`/`reference_text_column` above are what it actually uses to gather data.

A pure item-level policy (the AX/D365 shape, B.3) needs only a single lookup block, pulling multiple columns from Item Master instead of combining two sheets:

```yaml
custom_functions:
  dimpolicyvalidation:
    policy_key_lookups:
      - source_column: Item ID
        reference_sheet: ItemsMaster
        reference_key_column: Item ID
        policy_key_columns:
          - Tracking Group
          - Storage Group
    dim_policy_key_delimiter: "|"
    dim_policy_sheet: DimPolicy
    dim_policy_key_column: DimPolicySetUp
    stock_to_dim_policy_columns_mapping:
      Lot Number: Lot Number
      Serial Number: Serial Number
      Location ID: Location ID
    stock_quantity_column: On-Hand
```

Both examples share identical Stock Balance and Dim Policy sheet/column naming — only the shape of `policy_key_lookups` differs, per the mechanism described in B.3.

# Appendix E: Glossary

|                      |                                                                                                                                                  |
|----------------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| **Term**             | **Definition**                                                                                                                                   |
| Active rule          | A rule row where Active = Yes                                                                                                                    |
| DimSetUp             | Worksheet defining mandatory dimension profiles per policy key                                                                                   |
| Enrichment           | The process of generating new derived columns in the workbook                                                                                    |
| Generated column     | A column created by an enrichment function; not present in the original input                                                                    |
| Item Master          | Worksheet or external file containing item definitions including dimension policy                                                                |
| Metadata             | Rules and configuration stored within the workbook that control processing                                                                       |
| Runtime failure      | An unexpected error during function execution — distinct from a data validation issue                                                            |
| Severity             | Classification of a validation issue: Error (blocking) or Warning (advisory)                                                                     |
| SheetName (external) | The name assigned to an external file dataset. Used to reference it in rules and becomes the worksheet name in the output workbook after import. |
| Stock Balance        | Worksheet containing opening inventory quantities to be loaded into the ERP system                                                               |
| ValidationType       | The category of built-in or custom validation logic to apply to a column                                                                         |

# Appendix F: Anthropic Provider Reference

## F.1 Purpose

The reference implementation of the Provider Registry contract (section 2.6.3), and the one provider module the framework currently ships with. Selected via `ai_provider: anthropic` (section 3.4). Every AI-flavored function, including `ai_itemcategoryfit` (Appendix C), works against this module — or any other correctly-implementing provider module — without any function-level changes; a function never knows which provider is active.

## F.2 Required Settings

`PROVIDER_NAME = "anthropic"`. `REQUIRED_SETTINGS` — both mandatory in `ai_provider_settings.anthropic` (section 3.4) whenever `ai_provider: anthropic` is selected and `ai_enabled = true`:

|                |          |                                                                                          |
|----------------|----------|------------------------------------------------------------------------------------------|
| **Setting**    | **Type** | **Description**                                                                          |
| ai_endpoint    | string   | The Anthropic Messages API URL requests are sent to. No built-in default — always visible in config.yaml rather than implied by code, since it's where this run's business data (via AI-flavored validation calls) is actually sent. |
| ai_api_version | string   | The Anthropic API version this provider module was written against, sent as the `anthropic-version` request header. Not a business decision — it's a wire-protocol detail specific to this one provider, so it lives here rather than anywhere function- or framework-level. |

## F.3 Request/Response Mechanics

`send(prompt, config, settings)` builds a POST request to `settings["ai_endpoint"]` with headers `x-api-key: config.ai_api_key`, `anthropic-version: settings["ai_api_version"]`, and a JSON body of `{"model": config.ai_model, "max_tokens": <fixed>, "messages": [{"role": "user", "content": prompt}]}`. The response is expected as Anthropic's Messages API shape — a `content` list of typed blocks — and this module extracts and concatenates the `text` blocks to produce the raw string `send_prompt()`/`send_batch()` (section 6.8) work with. No other part of the framework parses or constructs Anthropic-specific request or response structure.
