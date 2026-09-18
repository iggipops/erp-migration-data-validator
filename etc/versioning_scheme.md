# ERP Migration Data Validator — Versioning Scheme

## 1. Purpose

Defines how the Business Specification (BS), Functional Specification (FS), and codebase are versioned, and how their version numbers stay in sync with each other.

## 2. Components

| Symbol | Meaning |
|---|---|
| `v` | Literal prefix for "version" |
| `N` | Business-requirement version number. Increments **only** when the Business Specification document itself changes. If BS is untouched, N never changes. |
| `L` | Functional-difference version number. Increments when FS changes in a way that has no corresponding BS change. |
| `r` | Literal prefix marking a revision |
| `K` | Revision number, for non-functional/editorial changes (wording, formatting, clarification) that don't change behavior. Starts at 0. |
| `M` | Codebase change counter, for code changes within an otherwise-unchanged FS version. Starts at 0. |

## 3. Reset Rule

Incrementing any component resets every component to its right to 0.

## 4. Business Specification (BS)

**Mask:** `vN_rK`

Example progression: `v2_r0`, `v2_r1`, `v2_r2`

## 5. Functional Specification (FS)

**Mask:** `vN_L_rK`

- `N` always matches the BS version it's based on.
- `L` increments for a functional change that doesn't require a BS change.
- `K` increments for a revision to the same functional content (wording, restructuring, clarification — no behavior change).

Example progression: `v2_4_r0`, `v2_4_r1`, `v2_5_r0`
(two revisions of the same functional content, then a functional change with no BS change)

## 6. Codebase

**Mask:** `vN_L_rK_M`

- `vN_L_rK` matches the FS version the code implements.
- `M` increments for each round of code changes made against that same FS version.
- Per the reset rule (section 3): if any part of `vN_L_rK` changes, `M` resets to 0.

## 7. Naming Convention

Use underscores consistently — in document titles, version-history tables, and filenames alike. No dots, no mixing.
