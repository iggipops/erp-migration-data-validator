# ERP Migration Data Validator

A metadata-driven validation and enrichment tool for Excel master and transactional data, run at the final checkpoint before ERP import.

## What it does

The tool reads an Excel workbook containing your business data alongside metadata rules that define what "valid" means for that data. It runs enrichment functions to generate derived columns, then runs validation rules to check data quality — flagging issues with cell-level highlighting and a Summary worksheet, so problems are visible and traceable before the data ever reaches the ERP system.

Built-in validation types include EMPTY, DUPLICATE, FORMAT, REFERENCE, DATE, NUMERIC, and INTEGER checks, plus an AI-assisted validation type for judgment calls — like flagging an item that looks miscategorized — that plain rule matching can't catch. A ready-to-use AI validation function (`ai_itemcategoryfit`) ships with the tool — usable as-is, and also the reference implementation to follow when writing your own AI-flavored validation functions. Custom, project-specific validation and enrichment functions can be added through an extensibility framework, including a predefined dimension-policy validator covering common ERP dimensional control patterns (SAP, Infor LN, Dynamics AX/D365 style key structures).

## What it does not do

- Replicate ERP business logic
- Connect directly to ERP systems
- Perform data cleansing, transformation, or general ETL
- Act as a workflow or BPM engine

This tool sits at one specific point in a migration project: the last validation pass on Excel data before import, not a data preparation or transformation platform.

## Why on-site execution matters

The tool runs entirely on your own machine or infrastructure. No data is uploaded anywhere as part of core validation — input files, configuration, and results all stay local. This is intentional: it's built for engagements (particularly in European industrial environments) where client data cannot leave client infrastructure. The one exception is the optional AI validation feature, which, when explicitly enabled, sends the specific data needed for that check to your configured AI provider's API under your own API key — everything else is fully local and this feature can be left disabled entirely.

## AI provider support

The AI backend is provider-agnostic by design (a pluggable provider registry), but only one provider adapter ships today: Anthropic. Using a different AI service means adding a new provider module — the rest of the framework doesn't need to change.

## Installation

Requires Python 3.14 — the only version this has been tested against. Older versions may or may not work; none have been verified.

```bash
git clone https://github.com/iggipops/erp-migration-data-validator.git
cd erp-migration-data-validator
pip install -r requirements.txt
```

## Configuration

The tool is driven by a YAML configuration file. Two example configs are provided as starting points:

- `src/v2/config_universal.example.yaml` — general-purpose starting point
- `src/v2/config_ax.example.yaml` — geared toward Dynamics AX/D365-style dimension policy setups

Copy whichever fits your situation to `config.yaml` and adjust from there:

```bash
cp src/v2/config_universal.example.yaml src/v2/config.yaml
```

Edit `config.yaml` to point at your input workbook, output path, and any optional features you want enabled (AI validation, custom functions, etc.). See the Functional Specification in `functional_spec/` for the full configuration reference.

The tool does not create its output directory automatically — make sure the folder set as `output_file`'s directory already exists before running, e.g.:

```bash
mkdir -p output/v2
```

## Usage

Run from the repository root, so that the paths you set in `config.yaml` (input workbook, output file, custom functions directory) resolve correctly:

```bash
python src/v2/erp_migration_data_validator.py
```

With no `--config` flag, the tool looks for `config.yaml` next to the script itself (`src/v2/config.yaml`) — which is exactly where the Configuration step above puts it, regardless of your current directory.

`--config` accepts a path to use a different config file instead, resolved relative to wherever you run the command from (not relative to the script) — for example `--config /path/to/other-config.yaml`.

The tool copies your input workbook to the configured output path, applies enrichments and validations in sequence, and writes the results back into that output file — the original input is never modified.

## Documentation

The full behavior of the tool — every validation type, configuration field, and processing rule — is documented in the Functional Specification:

- [`functional_spec/`](functional_spec/) — the current version is the single file in that folder

## License

Licensed under the [Apache License 2.0](LICENSE).

## About

Built and maintained by [Igor Gindin](https://github.com/iggipops) ([email](mailto:igbgindin@gmail.com)), an independent ERP migration consultant. This tool grew out of real client migration work and is shared publicly under Apache 2.0.
