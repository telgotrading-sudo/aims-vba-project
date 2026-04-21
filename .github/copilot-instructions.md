# Copilot Instructions for aims-vba-project

## Project Overview

This repository contains VBA modules extracted from Excel workbooks that form the AIMS (Advanced Investment Management System) suite. The project uses a **source-control-first approach**: Excel workbooks (.xlsm files) in the `excel/` directory are treated as binaries and excluded from Git, while all VBA code is exported and tracked as plain-text `.bas` files in the `vba/` directory tree.

The workflow involves:
1. Exporting VBA code from workbooks to version-controllable `.bas` files
2. Modifying and testing VBA logic in Git-tracked source files
3. Importing updated modules back into the workbooks for execution

## Directory Structure

- **`excel/`** — Binary Excel workbooks (.xlsm files) for running VBA code. Excluded from Git tracking (see `.gitignore`).
- **`vba/`** — Exported VBA modules organized by workbook name:
  - `{WorkbookName}/modules/` — Standard code modules (.bas files)
  - `{WorkbookName}/classes/` — Class modules (.cls files) if present
  - `{WorkbookName}/forms/` — Form modules (.frm files) if present
- **`docs/`** — Additional documentation (if any)

### Key Workbooks

- **aimsAll2025 / aimsAll2026** — Main data consolidation workbooks with multi-step workflows:
  - Data cleanup, preparation, formula pasting, final reconciliation
  - Handles fund name standardization and per-policy calculations
  
- **aimswrap** — Wrapper workbook with integrity validation workflows:
  - Bidirectional reconciliation between "aims" and "aimswrap" sheets
  - Four-step validation with staging data and comparison logic
  
- **Monthly investor workbooks** — Company-specific monthly consolidation:
  - `psg monthly`, `sanlam monthly`, `bci monthly`, `investec monthly`
  - Similar patterns: data cleanup, filtering, formula pasting, reconciliation
  
- **Utility workbooks**:
  - `ExportAllModules` — Main tool for exporting all VBA to text files
  - `ImportAllModules` — Tool for importing text modules back into workbooks
  - `companies` — Lookup and account setup utilities

## Build, Test, and Import/Export

### Export VBA Modules to Text

Run from any workbook containing the `ModuleExport` code (typically `ExportAllModules.xlsm`):

```vba
Sub ExportAllWorkbooks()
```

- Prompts user to export a single workbook or all workbooks at once
- Exports modules, classes, and forms to the `vba/{WorkbookName}/` directory tree
- **Root path** is hardcoded in the module — update if repository location changes

### Import VBA Modules into Workbooks

Run from any workbook containing the `ModuleImport` code (typically `ImportAllModules.xlsm`):

```vba
Sub ImportAllModules()
```

- Reads all `.bas`, `.cls`, and `.frm` files from the `vba/` directory tree
- Imports them back into the corresponding workbooks in the `excel/` directory
- Creates workbooks if they don't exist

### Running Individual Workflows

Each workbook contains numbered workflow steps (e.g., `Step01a`, `Step02a`, `Step03`, `Step04`). Open the workbook in Excel and:

1. Open the VBA editor (Alt+F11)
2. Select the module containing your step (e.g., `DataPreparation`, `ReconcileWorkbooks`)
3. Position cursor on the desired sub and press F5, or call it from the Immediate window

Example (in Immediate window):
```vba
Call Step01aNewCleanFundNames()
Call Step02aCalculatePerPolicyTotals()
```

## Key Conventions

### Module Organization

- **Modules named `DataPreparation`** contain steps for cleaning, standardizing, and calculating fund data
- **Modules named `DataCleanup`** handle row/column deletions and normalization
- **Modules named `ReconcileWorkbooks`** implement bidirectional existence checks and reconciliation logic
- **Modules named `PasteFormulas`** extend formulas from a seed range down to the last data row
- **Modules named `FinalPaste`** handle the final consolidation and summary calculations
- **Modules named `NameFilter`, `BalanceFilter`** apply domain-specific filters to workbook data
- **Utility modules** (e.g., `CompanyLookup`, `AccountSetup`) provide lookup tables and reference data

### Naming and Comments

- All subroutines follow a numbered step naming convention: `Sub Step01a...`, `Sub Step02b...`, etc.
- Inline comments are sparse but descriptive — look for `'` at module scope for high-level purpose
- Functions that transform data (fund name parsing, filtering) include the transformation logic as a comment block

### Column Conventions

Fund and investment data typically uses lettered columns:
- Column R: Product/fund type
- Column T: Cleaned fund name (derived in Step01a)
- Column U: Per-policy totals (calculated in Step02a)
- Column I, K: Reference columns for fund names (used in cleaning logic)

Lookup workbooks (`companies`, `aimswrap`) use different column structures — examine the module headers for specifics.

### Temporary Staging Areas

- Modules with names ending in `ToDel` write temporary staging data (e.g., formula blocks in rows 1502–2817 of "aims" sheet)
- These should be deleted after validation is complete
- See `IntegrityWorkflow.bas` for examples of how staging data flows through a multi-step validation

### Hard-Coded Paths

- **Export/Import modules** reference the repository root path directly:
  ```vba
  projectRoot = "C:\Users\andriesvt\OneDrive\ExcelGitProjects\aims-vba-project"
  ```
- When moving the repository, update this path in both `ModuleExport` and `ModuleImport`

## Workflow Patterns

### Data Preparation Workflow

Typical for `aimsAll` and monthly investor workbooks:

1. **Step01** — Remove header rows and leftmost column from raw export; standardize fund names
2. **Step02a** — Calculate per-policy totals
3. **Step02b** — Copy rows with totals to summary sheet
4. **Step03** — Paste formulas for further calculations
5. **Step04** — Final consolidation and cleanup

### Integrity Validation Workflow

Used in `aimswrap`:

1. **ReconcileAimsWrap** — Bidirectional existence checks (aims → aimswrap, aimswrap → aims)
2. **Step02A** — Extend formulas in "aims" sheet staging area
3. **Step02B** — Copy "aimswrap" data to staging area as values
4. **Step03** — Aggregate active totals by account
5. **Step04** — Add percentage difference and flag formulas for final review

### Filtering Patterns

- **NameFilter** — Looks up and filters by fund name (see `psg monthly`)
- **BalanceFilter** — Filters rows based on balance thresholds (see `bci monthly`, `psg monthly`)

## Common Tasks

### Adding a New Monthly Investor Workbook

1. Create a new `.xlsm` file in `excel/`
2. Copy an existing monthly workbook template and rename appropriately
3. In `ImportAllModules`, import the baseline modules:
   - `DataPreparation.bas` (or adapt from a similar month)
   - `ReconcileWorkbooks.bas`
   - `PasteFormulas.bas` (if needed)
   - `FinalPaste.bas`
4. Run `ExportAllModules` to commit the modules to Git

### Modifying Fund Name Cleaning Logic

Edit `DataPreparation.bas` in the corresponding workbook folder:
- The `Step01aNewCleanFundNames()` sub contains a `Select Case` statement matching fund types
- Add new fund names or modify cleaning logic there
- Test by running the step in Excel, then export to Git

### Debugging Reconciliation Failures

Check `ReconcileAimsWrap.bas` or `ReconcileWorkbooks.bas`:
- These modules flag rows in the first sheet that don't exist in the second sheet
- Look for columns with formulas containing `MATCH` or `COUNTIF` used to find mismatches
- Examine the "ToDel" staging sheets to see which rows failed the bidirectional check

## Notes

- This is a VBA-based system; no build artifacts or compiled outputs exist
- The export/import tooling maintains the source-control workflow and prevents binary file diffs in Git
- All workbooks should be closed before running import operations to avoid file-locking issues
- Formula-heavy calculations mean recalculation time can be significant; set `Application.ScreenUpdating = False` during long operations
