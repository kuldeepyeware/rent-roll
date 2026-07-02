# Rent Roll Standardizer v0

A practical, explainable prototype for turning a `.csv`, `.xlsx`, or `.xls`
rent roll into a standardized JSON representation.

The architecture deliberately separates two jobs:

- Deterministic software detects structure, groups units, calculates rents,
  calculates rollups, and runs validations.
- OpenRouter is used only for uncertain column meanings and charge-code
  classification.

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
| `04_structure_detection.json` | Flat rows vs. repeating unit blocks, start row, footer detection |
| `05_unit_grouping.json` | Units found, source row numbers, skipped rows, future-record rows |
| `06_column_mapping.json` | Every source column, canonical mapping, confidence, method |
| `07_charge_classification.json` | Every charge code, category, confidence, samples, method |
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

### 3. Confirm unit grouping

Open `05_unit_grouping.json`.

Check:

- `unit_count` matches the source summary.
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

- `duplicate_unit_numbers`
- `unit_charge_total_reconciliation`
- `source_total_reconciliation`
- `effective_not_above_total`
- `vacant_units_without_tenants`
- `admin_model_units_flagged`

A warning is not necessarily a parser failure. It means a human decision or
source-data investigation is required.

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
- 88.49% occupancy
- 8 of 8 validations passed
- 2 OpenRouter calls: column mapping and charge classification

This is the best file for demonstrating AI classification of opaque charge
codes followed by deterministic rent calculations.

### Mirada

The deterministic test run produced:

- 256 total units
- 234 occupied units
- 21 vacant units
- 1 admin/model unit
- 91.76% occupancy
- 8 of 8 validations passed
- 10 future-resident records preserved without overwriting current records

This is the best file for demonstrating two-row headers, repeated unit blocks,
footer detection, and future-resident handling.

## Suggested live meeting flow

1. Run Retreat with `--verbose`.
2. Point out the header and grouping messages in the terminal.
3. Open `06_column_mapping.json`.
4. Open `07_charge_classification.json` and show deterministic vs. AI methods.
5. Open one concession unit in `08_validation_report.json`.
6. Show that all eight validations pass.
7. Open `standardized_rent_roll.xlsx` beside the original workbook.
8. Start on `Summary`, then spot-check units on `Standardized Rent Roll`.
9. Use `final_output.json` only for deeper debugging.

The product message should be visible throughout:

> AI helps understand the document, but deterministic software produces the
> final financial output.

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
