# Changelog

All notable changes to this project are documented here, in the style of
[Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

Versions follow the `vN_L_rK_M` scheme defined in
[etc/versioning_scheme.md](etc/versioning_scheme.md) — `vN_L_rK` is the
Functional Specification version the code implements, `M` counts
code-only change rounds against that same FS version and resets to 0
whenever `N`, `L`, or `K` changes.

Tracking in this file starts at `v2_9_r0_2` (2026-09-18) — no one had
visited the public repo before that point, so earlier history isn't
reconstructed here; see `git log` and the repo's tags for anything before
it.

## [Unreleased]

## [v2_10_r0_1] - 2026-09-25

Code-only round against FS v2_10_r0; no behavior change.

### Changed

- CI workflow (`.github/workflows/tests.yml`) now runs with read-only
  repository permissions (`permissions: contents: read`).
- `.gitignore`: two `!config_*.example.yaml` lines reordered; `desktop.ini`
  stays ignored. No change in what is ignored.
- Reworded a comment in `src/v2/tests/test_read_layer.py` that referred to
  an internal document.

## [v2_10_r0_0] - 2026-09-25

Implements FS v2_10_r0.

### Changed

- `ValidationRules.Color` is now required (FS 8.7.3). An active row with an
  empty Color is a metadata error. Before this, an empty Color passed
  validation and then crashed the IssuesWriter the first time the rule
  flagged a cell. **Existing workbooks with blank ValidationRules Color
  cells must fill them in.**
- The accepted Color format is tighter for both rules sheets (FS 8.7.2,
  8.7.3). Color must be a built-in color name, a name from config
  `predefined_colors`, or a plain 6-digit hex code with no `#` (e.g.
  `FF0000`). `#RRGGBB`, `RGB` and `#RGB` were accepted before and are now
  metadata errors. openpyxl rejects all three when it writes the fill, so
  they previously got through validation and crashed at run time.
  **Existing workbooks using these forms must switch to `RRGGBB`.** Config
  `predefined_colors` values are unaffected: a leading `#` there is still
  accepted.

### Fixed

- `EnrichmentRules.Color` was not validated at all, although FS 8.7.1
  already listed Color as a common check. It is now checked with the same
  format rule when non-empty and stays optional (FS 8.7.2). Before this, an
  invalid value failed the enrichment at run time, after the generated
  column had already been written.

## [v2_9_r2_1] - 2026-09-24

Code-only round against FS v2_9_r2; no behavior change.

### Changed

- Added `pyproject.toml` as the source of truth for dependency ranges, and
  brought `requirements.txt` and `src/v2/requirements-dev.txt` (used by CI)
  in line with it: `pandas>=3.0.3,<4`, `openpyxl>=3.1.5`, `pyyaml>=6.0.3`,
  `pytest>=9.1.1`. Floors are the versions verified against the full suite;
  pandas is capped below 4 so a future major is a deliberate, tested change.
  `requires-python` is `>=3.14`, the only interpreter tested.

## [v2_9_r2_0] - 2026-09-24

Implements FS v2_9_r2.

### Added

- New required `CSVEncoding` column on `ExternalFiles` (FS 5.1). Supported
  values: `utf-8`, `utf-8-sig`, `cp1251`, `cp1252` (case-insensitive).
  `PreflightValidator` (FS 8.3.1) requires it, non-empty, when `Format = csv`,
  rejects any other value, and requires it empty when `Format = xlsx`; its
  header must exist in the sheet, like `CSVDelimiter`. **Existing workbooks
  must add the column** — without it, preflight terminates with a
  "column not found" error.
- `Importer` decodes csv files using the declared `CSVEncoding`. A file that
  cannot be decoded is a named import failure (FS 8.6.4): the message names
  the file and the declared encoding, the sheet is added to `failed_sheets`,
  and dependent rules are skipped like any other failed import.

### Fixed

- csv import no longer blanks literal text that looks like a missing value
  (`NULL`, `NA`, `N/A`, `None`, `#N/A`, ...). `pd.read_csv` now runs with
  `keep_default_na=False, na_values=[""]`, so that text survives exactly as
  written and only a truly empty cell is read as missing (FS 8.6.1).
  xlsx import is unaffected.

## [v2_9_r1_0] - 2026-09-19

Implements FS v2_9_r1.

### Changed

- ImportSpec applies to csv external files only. An ImportSpec row whose
  `ExternalFileSheetName` points to an xlsx row in `ExternalFiles` is now a
  metadata validation error (`PreflightValidator`, FS 8.3.2). External xlsx
  files are no longer read as text and converted: `Importer` copies the
  worksheet cell by cell with native types and formulas (FS 8.6.1). csv
  loading is unchanged.
- xlsx workbooks (primary and external alike) are read one way (FS 8.4.1):
  - Formulas vs. calculated values: `open_workbook()` loads each file twice;
    validators, enrichment and metadata readers read the last-calculated
    value (`cell_value()`, `sheet_to_dataframe()`), while the output keeps
    the formulas. A formula cell with no cached value logs a warning.
  - Last row: `last_data_row()` replaces `max_row` for reading business and
    metadata sheets — rows that only carry formatting are no longer scanned.
  - `sheet_to_dataframe()` builds the frame with `dtype=object`, so ints and
    floats keep the type they were read with.
- Six existing tests now include a filled second column: with the new last-row
  rule, a trailing row that is blank in every column is not a data row.

## [v2_9_r0_3] - 2026-09-18

### Added

- `parse_bool()` helper (`src/v2/utils/parsing.py`) for strict boolean
  coercion (`true/false`, `yes/no`, `1/0`), used for `ai_enabled` in
  `config_loader.py` and the `trim` parameter in `concatenate.py`.
- CI workflow (`.github/workflows/tests.yml`) running pytest on every
  push and pull request.

### Fixed

- `config_loader.py`: `date_validation.min_date`/`max_date`/`null_dates`
  now accept both native date/datetime objects (how PyYAML parses an
  unquoted date) and strings, instead of crashing when a date is left
  unquoted in YAML.
- `config_loader.py`: a non-numeric `numeric_validation.min_value`/
  `max_value` is now reported as a normal config validation error
  instead of raising an unhandled `TypeError`.
- `excel_utils.append_row_safe()` guards against formula injection —
  any string value starting with `=`, `+`, `-`, or `@` is stored as
  literal text rather than a live formula. Applied in `Importer` and
  `IssuesWriter`, wherever rows are built from external/input data.
- Failed-enrichment dependency skip in `validation_engine.py` is now
  scoped per `(SheetName, ColumnName)` instead of a bare column name, so
  a failed enrichment on one sheet no longer wrongly suppresses an
  unrelated validation on a different sheet with a same-named column.
  `REFERENCE` rules are checked against `ReferenceDataSheet`;
  `DUPLICATE2` rules against the rule's own `SheetName` (FS 6.4).
- `counter4group()` now returns `int64` group counters instead of
  occasionally float-typed ones.

### Changed

- Dropped redundant `decimal_separator`/`thousands_separator` from
  `numeric_validation` in both example configs.

## [v2_9_r0_2] - 2026-09-18

### Added

- `src/v2/__version__.py` as the single source of truth for the codebase
  version; the entry script imports it, exposes it via `--version`, and
  logs it on startup.
