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
