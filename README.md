# Rent Roll Standardizer v0.3

A practical, explainable prototype for turning a `.csv`, `.xlsx`, or `.xls`
rent roll into a standardized JSON representation.

The architecture deliberately separates two jobs:

- Deterministic software profiles the workbook, validates a parse plan, groups
  units, calculates rents, calculates rollups, and runs validations.
- OpenRouter may propose document regions, uncertain column meanings, and
  charge-code classifications. It never performs arithmetic or directly
  creates canonical units.

The output is designed to be inspected, not taken on faith. Every run creates
a log and numbered artifacts that show what the parser did at each stage.

The intended QA workflow is:

```text
Original Rent Roll → standardized_rent_roll.xlsx
```

The workbook is the primary artifact for business review. `final_output.json`
remains the canonical source of truth for debugging and downstream processing.

## Quick start

### 1. Open Terminal in this directory

```bash
cd "/Users/mohitmotwani/Documents/Rent Roll Standardizer"
```

### 2. Create the project environment

Python 3.10 or newer is recommended.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

This creates an isolated Python environment inside the project. The pinned
dependencies in `requirements.txt` include:

- `openpyxl` for `.xlsx` files
- `xlrd` for legacy `.xls` files

When the environment is active, the Terminal prompt normally starts with
`(.venv)`.

### 3. Configure OpenRouter

Edit the local `.env` file:

```dotenv
OPENROUTER_API_KEY=your_openrouter_key_here
OPENROUTER_MODEL=anthropic/claude-haiku-4.5
```

The script loads this file automatically. `.env` is ignored by Git and should
never be shared or committed.

### 4. Run a rent roll

```bash
python rent_roll_standardizer.py \
  --input "test_files/retreat_canal_rent_roll.xlsx" \
  --output-dir "output"
```

Every run gets its own timestamped folder, for example:

```text
output/retreat_canal_rent_roll_20260702_144914/
```

The terminal prints the exact folder when the run finishes.

### Every time you open a new Terminal

Return to the project and activate the existing environment:

```bash
cd "/Users/mohitmotwani/Documents/Rent Roll Standardizer"
source .venv/bin/activate
```

Then run the tool with `python rent_roll_standardizer.py ...`.

You do **not** recreate the venv or reinstall dependencies each time.

If you prefer not to activate it, call its Python executable directly:

```bash
.venv/bin/python rent_roll_standardizer.py \
  --input "test_files/retreat_canal_rent_roll.xlsx" \
  --output-dir "output"
```

## Useful commands

Run Mirada:

```bash
python rent_roll_standardizer.py \
  --input "test_files/mirada_rent_roll.xlsx" \
  --output-dir "output"
```

Run a legacy Excel file:

```bash
python rent_roll_standardizer.py \
  --input "test_files/nova_central_rent_roll.xls" \
  --output-dir "output"
```

Show detailed console logs and preserve raw rows in the grouping artifact:

```bash
python rent_roll_standardizer.py \
  --input "test_files/retreat_canal_rent_roll.xlsx" \
  --output-dir "output" \
  --verbose \
  --debug
```

Force a particular worksheet:

```bash
python rent_roll_standardizer.py \
  --input "/path/to/rent_roll.xlsx" \
  --output-dir "output" \
  --sheet "Rent Roll"
```

Override the model for one run:

```bash
python rent_roll_standardizer.py \
  --input "/path/to/rent_roll.xlsx" \
  --output-dir "output" \
  --model "anthropic/claude-haiku-4.5"
```

Skip numbered artifacts when only final output is needed:

```bash
python rent_roll_standardizer.py \
  --input "/path/to/rent_roll.xlsx" \
  --output-dir "output" \
  --no-save-intermediate
```

## What a run produces

| File | What to inspect |
|---|---|
| `debug.log` | The complete chronological story of the run |
| `01_input_summary.json` | File type, workbook dimensions, sheets, model configuration |
| `02_sheet_detection.json` | Which sheet was selected and why |
| `03_header_detection.json` | Raw header rows, combined headers, confidence, alternatives |
| `04_structure_detection.json` | Redacted document profile, validated parse plan, regions, boundaries, and row model |
| `05_unit_grouping.json` | Units found, source rows, skipped rows, and accounting for every non-empty primary-region row |
| `06_column_mapping.json` | Initial mappings and the column roles actually used after parse-plan validation |
| `07_charge_classification.json` | Charge categories plus the file-level financial profile and validated treatment plan |
| `08_validation_report.json` | Validation results and deterministic per-unit calculations |
| `final_output.json` | Final standardized rent roll and property rollups |
| `standardized_rent_roll.xlsx` | Primary QA workbook with one row per unit and a Summary sheet |
| `standardized_rent_roll.csv` | Optional flat CSV containing the same 16 unit columns |

The Excel file contains:

- `Standardized Rent Roll`: exactly one row per canonical JSON unit
- `Summary`: property identity, rent roll date, unit/occupancy totals, flagged
  units, and validation warnings

Excel generation does not calculate or reinterpret values. It only formats and
flattens the canonical JSON.

## How to evaluate a run visually

Open the original workbook and the run folder side by side. Review the artifacts
in numerical order; this tells the same story as the pipeline.

### 1. Confirm the source and selected sheet

Open `01_input_summary.json` and `02_sheet_detection.json`.

Check:

- The expected filename was loaded.
- The expected sheet was selected.
- Row and column counts look plausible.
- `openrouter_enabled` is `true`.

If the wrong sheet was selected, rerun with `--sheet "Sheet Name"`.

### 2. Confirm the header block

Open `03_header_detection.json` and compare `raw_header_rows` with the workbook.

Check:

- The detected header row or two-row header block is correct.
- Combined names make sense, such as `Unit | Sq Ft`.
- Metadata rows like property name and report date were not treated as headers.
- Confidence is reasonably high.

### 3. Confirm the parse plan and unit grouping

Open `04_structure_detection.json` first.

Check:

- `parse_plan.primary_region` starts on the first current-unit row and ends
  before property totals, summaries, or future-resident tables.
- `parse_plan.regions` describes the other visible parts of the workbook.
- `unit_column_index` points to the actual unit/building-unit column.
- `method` shows whether the deterministic fallback or a validated OpenRouter
  proposal was used.
- Any `warnings` are understandable and acceptable.

Then open `05_unit_grouping.json`.

Check:

- `unit_count` matches the source summary.
- `row_accounting.unaccounted_rows` is empty.
- The row-role counts make sense for the source format.
- The first, middle, and last few unit IDs exist.
- Repeated charge rows belong to the correct unit.
- Per-unit `source_row_numbers` point to the expected workbook rows.
- Section headings and totals appear under `skipped_rows`, not as units.
- Future applicants appear as `future_source_rows` and do not overwrite current
  residents or current charges.

This is the most important structural check. If grouping is wrong, do not trust
the financial output even if later validations pass.

### 4. Review column mappings

Open `06_column_mapping.json`.

For every decision, inspect:

- `source_header`
- `canonical_field`
- `confidence`
- `method`
- `reason`

`method: "deterministic"` means an alias/rule mapped it.
`method: "openrouter"` means the model handled uncertain semantics.
`canonical_field: "unknown"` means the tool deliberately did not guess.

It is normal for IDs or unused duplicate columns to remain unknown or be mapped
to `ignore`. It is not normal for the unit number, tenant name, charge amount,
or lease dates to be incorrectly mapped.

### 5. Review charge classifications

Open `07_charge_classification.json`.

Check each distinct charge code against its sample amounts:

- Base rent should be `BASE_RENT`.
- Concessions and rent credits should be `CONCESSION`.
- Garage charges should normally be `PARKING`.
- Pet rent should normally be `PET_FEE`.
- Recurring services may be `UTILITY`, `OTHER_RENT`, or another fee category.
- Anything genuinely unclear should remain `UNKNOWN`.

Pay extra attention to AI decisions with lower confidence. The model suggests
semantics; deterministic software still performs all arithmetic afterward.

For the Retreat test file, the current model classified codes such as `cbl`,
`con`, `gar`, `ptr`, `mtm`, and `stl`. Decisions such as `mtm` and `stl` are
exactly the kind a human should review before production use.

### 6. Recalculate a few units manually

Open `08_validation_report.json` and find `calculated_totals`.

Pick at least:

- One ordinary occupied unit
- One unit with a concession
- One unit with multiple fees
- One vacant or model unit

Use the source row numbers to find the same unit in Excel.

The deterministic formulas are:

```text
effective rent = BASE_RENT + CONCESSION
total rent     = recurring charges + CONCESSION
```

Concessions are normalized as negative values. One-time fees and deposits are
not treated as recurring rent. Unknown charges are preserved but excluded from
financial totals, which should cause reconciliation warnings rather than a
silent guess.

### 7. Review validation results

Open `08_validation_report.json`.

The strongest demo outcome is:

```json
{
  "status": "passed",
  "failed": 0
}
```

Review every failed check even when the overall unit count looks right.
Especially important checks are:

- `parse_plan_integrity`
- `primary_row_accounting`
- `source_summary_reconciliation`
- `duplicate_unit_numbers`
- `unit_charge_total_reconciliation`
- `source_total_reconciliation`
- `effective_not_above_total`
- `vacant_units_without_tenants`
- `admin_model_units_flagged`

A warning is not necessarily a parser failure. It means a human decision or
source-data investigation is required.

`source_summary_reconciliation` compares the computed unit count, occupied
count, vacant count, non-revenue count, and occupancy rate with any summary
values detected in the source workbook. This is intended to catch plausible-
looking output that accidentally includes a footer or misclassifies vacancies.

### 8. Review the standardized workbook

Open `standardized_rent_roll.xlsx` beside the original rent roll.

Check:

- The `Summary` sheet agrees with the source property summary.
- The detail sheet contains the exact expected unit count plus one header row.
- Every physical unit appears exactly once.
- Tenant, status, type, size, rents, dates, balance, and deposit match the source.
- Currency, percentages, and dates display consistently.
- Flagged rows are easy to spot and review.
- Blank canonical values remain blank in Excel.

Use `final_output.json` only when you need to trace a spreadsheet value back to
source rows, charge decisions, or confidence details.

## Green, yellow, and red signals

### Green: suitable for the demo

- Correct sheet and header block
- Unit count agrees with the source
- First/middle/last units are grouped correctly
- OpenRouter calls are limited to column mapping and charge classification
- Per-unit and property totals reconcile
- Zero failed validations
- Any remaining unknown columns are genuinely irrelevant

### Yellow: investigate before presenting

- A few low-confidence mappings
- An unfamiliar charge remains `UNKNOWN`
- Missing lease dates on otherwise valid units
- Future applicants or unusual admin/model units
- A small source-total difference with a clear explanation

### Red: do not trust the financial output yet

- Wrong sheet or header row
- Section headings appear as unit IDs
- Charge rows attach to the wrong unit
- Unit count differs materially from the source
- Unit or property charge totals do not reconcile
- Unit number, tenant, rent, or charge-amount columns are mapped incorrectly

## Verified examples in `test_files`

### Retreat at Canal

The AI-assisted test run produced:

- 140 total units
- 123 occupied units
- 16 vacant units
- 1 admin/model unit
- 87.86% occupancy
- 9 validation checks, including source-summary reconciliation
- 2 OpenRouter calls: column mapping and charge classification

This is the best file for demonstrating AI classification of opaque charge
codes followed by deterministic rent calculations.

### Mirada

The deterministic test run produced:

- 256 total units
- 234 occupied units
- 21 vacant units
- 1 admin/model unit
- 91.41% occupancy
- 9 validation checks, including source-summary reconciliation
- 10 future-resident records preserved without overwriting current records

This is the best file for demonstrating two-row headers, repeated unit blocks,
footer detection, and future-resident handling.

## Suggested live meeting flow

1. Run Retreat with `--verbose`.
2. Point out the header and grouping messages in the terminal.
3. Open `06_column_mapping.json`.
4. Open `07_charge_classification.json` and show deterministic vs. AI methods.
5. Open one concession unit in `08_validation_report.json`.
6. Show that the source summary and all other validations pass.
7. Open `standardized_rent_roll.xlsx` beside the original workbook.
8. Start on `Summary`, then spot-check units on `Standardized Rent Roll`.
9. Use `final_output.json` only for deeper debugging.

The product message should be visible throughout:

> AI helps understand the document, but deterministic software produces the
> final financial output.

## Deterministic row and occupancy rules

The parser applies the following rules before a row can become a unit:

- Property-name rows immediately below the headers are treated as labels, not
  unit records.
- Footer rows such as `203 Units`, `Total 203 Units`, `Unit Count: 200`,
  `Total Number of Units ...`, and `Total Market Rent ...` stop unit grouping.
- A source summary is parsed independently and compared with the computed
  rollup during validation.
- Tenant placeholders such as `Vacant Unit`, `Vacant`, and `No Tenant` produce
  a blank tenant name and `vacant` status.
- An explicit source status such as `Current`, `Occupied`, or `Vacant` takes
  precedence over descriptive words inside a tenant name.
- Exact placeholders such as `MODEL`, `ADMIN`, or `OFFICE` are treated as
  admin/model units when no explicit occupied/vacant status is available.

Property occupancy is calculated deterministically as:

```text
occupied units / total physical units
```

Admin/model and unknown-status units remain in the denominator but are not
counted as occupied. For example, 220 occupied units in a 232-unit property
produce 94.83% occupancy even when two additional units are models.

## Parse-plan architecture (v0.3)

v0.3 treats structure discovery as a separate, inspectable compilation step.
The parser does not immediately interpret every row beneath the header as unit
data.

The pipeline is:

```text
Workbook
  → deterministic document profile
  → AI-proposed parse plan (when OpenRouter is available)
  → deterministic plan validation
  → deterministic grouping and calculations
  → canonical JSON
  → QA spreadsheet
```

The document profile scans the full selected sheet for header boundaries,
unit-ID patterns, totals, summaries, secondary tables, repeated rows, charge
evidence, and date shapes. Data-like text is redacted before the structural
profile is sent to the model.

The parse plan declares:

- The inclusive row range containing current physical units
- The unit-number column
- Whether rows are flat, repeated, or continuation-based unit blocks
- Column roles used during execution
- Other regions such as metadata, summaries, charge summaries, and future
  residents

After grouping, v0.3 builds a second, file-level `financial_plan`. This keeps
charge category and financial treatment separate. Identifying a row as a
`CONCESSION` does not by itself establish whether it is a recurring monthly
adjustment, a one-time move-in credit, or a current-period full-month
concession.

The financial profile compares each charge label across the property, including
its frequency, sample amounts, and size relative to base rent. OpenRouter may
propose the financial behavior, but deterministic guards validate the
proposal. A credit that offsets nearly a full month of base rent cannot be
treated as recurring without affirmative recurring evidence.

The recurring-rent contract is:

```text
effective rent = recurring base rent + concessions established as recurring
total rent     = recurring base and ancillary charges
                 + concessions established as recurring
```

One-time concessions remain visible in the canonical charge ledger and source
reconciliations but do not erase contractual recurring rent. If periodicity
cannot be established, the charge is excluded from recurring rent and the unit
receives an `UNRESOLVED_CONCESSION_PERIODICITY` warning rather than a silent
guess.

The model's proposal is advisory. Deterministic guards reject invalid row
ranges and column indexes, prevent a primary region from crossing a known
summary/footer boundary, and reject plans that lose a material portion of the
deterministically observed unit candidates. If the model is unavailable or its
plan fails validation, the deterministic plan is used.

During grouping, each non-empty row inside the accepted primary region is
assigned a role such as `unit_start`, `unit_continuation`, `control_total`,
`excluded_label`, or `excluded_unresolved`. The
`primary_row_accounting` validation fails when a row has no role. This makes
silent row loss visible during QA.

Only the accepted `primary_units` region may create canonical units. Summary
tables and future-resident sections remain available as evidence and control
data but cannot silently inflate the physical unit count.

## Adaptive field handling (v0.2, retained in v0.3)

The parser no longer assumes that every nonblank cell under the unit-number
header is a unit. It first separates the workbook into logical regions and
admits rows into the canonical unit list only when they contain a plausible
unit identifier.

### Primary rent-roll boundaries

The primary unit region ends when the parser encounters structural evidence
such as:

- A property-level total with several numeric totals
- A status, charge-code, unit-type, or other summary section
- A repeated summary header such as Description / Unit Count / Scheduled
- A future-resident or applicant section

Text labels such as `STATUS SUMMARY`, `DESCRIPTION`, `OCCUPIED NO NOTICE`, and
`NOTICE RENTED` therefore cannot become unit numbers merely because they appear
in the same column as real unit IDs.

Plausible unit IDs are compact identifiers containing digits, short
single/two-token identifiers, or exact admin markers such as `MODEL`. Long
headings and descriptive prose are rejected and preserved in debug artifacts.

### Future residents

Future-resident sections are scanned independently from the current rent roll.
Their rows are attached to matching current units through
`source.future_record_rows`; they never create extra physical units or replace
the current tenant and current charges.

### Combined lease dates

Columns named `Lease Dates`, `Lease Term`, or `Lease Period` are represented by
the canonical `lease_date_range` role. When a cell contains two dates, for
example:

```text
05/01/2025
04/30/2026
```

the first date becomes `lease_start_date` and the second becomes
`lease_end_date`. The parser does not ask the model to split or calculate
dates.

### Debit, credit, and reported totals

Charge semantics now distinguish:

- `charge_amount`: a debit or scheduled charge
- `credit_amount`: a credit, stored as a negative signed charge
- `reported_unit_total`: a source-provided per-unit billing or scheduled total

This prevents a credit amount from winning a duplicate column mapping and
becoming total rent.

Loss-to-lease rows such as `LTOR`, `LTOL`, and `Loss To Old Lease` are classified
as `LOSS_TO_LEASE`. They remain visible in the canonical charge ledger but are
not treated as recurring rent.

If a source provides `reported_unit_total`, it is preserved as the unit's total
rent source and reconciled against the signed debit/credit ledger. Otherwise,
the existing deterministic recurring-charge calculation is used. A mismatch
is surfaced in validation instead of being silently resolved by AI.

## Troubleshooting

### OpenRouter is skipped

If the log says:

```text
OpenRouter skipped: OPENROUTER_API_KEY is not set
```

Confirm that `.env` contains a non-empty `OPENROUTER_API_KEY` and that the file
is in the same directory as `rent_roll_standardizer.py`.

### OpenRouter fails but the run continues

The script intentionally falls back to deterministic behavior. Check:

- Internet connectivity
- OpenRouter credit and key validity
- The configured model ID

The failure and unresolved semantics will be visible in the log and artifacts.

### `.xlsx` will not open

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

### `.xls` will not open

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Confirm that the tool is using the project environment:

```bash
which python
python -c "import openpyxl, xlrd; print('Excel readers installed')"
```

`which python` should end with:

```text
Rent Roll Standardizer/.venv/bin/python
```

### No units are found

Review `03_header_detection.json` and `06_column_mapping.json`. The most likely
cause is an unfamiliar unit-number column or an incorrect sheet. Try a manual
`--sheet` override and inspect the unmapped columns.

## Data handling

Rent rolls contain tenant names and potentially sensitive financial data.

- Keep `.env` private.
- Treat the entire output folder as sensitive.
- Do not upload source or output files to public repositories.
- Rotate an API key immediately if it is pasted into chat, email, or a document.
