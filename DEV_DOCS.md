# Dev Docs

Random notes and things to know useful for maintainers and contributors.

## Google Sheets Compatibility Testing

HyperFormula ships a compatibility test suite that measures how closely its formula output matches Google Sheets when running with `compatibilityMode: 'googleSheets'`. All test data lives in `test/gsheets-compat/`.

### How it works

```
script/generate-formula-compat-tests.ts
  └─ reads  script/formula-functions.json
  └─ writes test/gsheets-compat/__fixtures__/formula-compat-tests.json
         (512 test cases, no expected values yet)

script/generate-formula-test-csv.ts
  └─ reads  formula-compat-tests.json
  └─ writes test/gsheets-compat/__fixtures__/formula-compat-input.csv
         (771 rows; column C holds =FORMULA() strings for GSheets to evaluate)

──── MANUAL STEP ────────────────────────────────────────────────────────────
  Import formula-compat-input.csv into Google Sheets.
  Google Sheets evaluates column C automatically.
  Export the sheet as CSV → formula-compat-gsheets.csv
  Place it at: test/gsheets-compat/__fixtures__/formula-compat-gsheets.csv
─────────────────────────────────────────────────────────────────────────────

script/patch-expected-values.ts
  └─ reads  formula-compat-gsheets.csv  (required — exits if missing)
  └─ writes expectedValue fields back into formula-compat-tests.json

test/gsheets-compat/gsheets-compat.spec.ts
  └─ evaluates each formula through HyperFormula
  └─ compares against expectedValue from formula-compat-tests.json
  └─ fails if any function in MUST_MATCH_FUNCTIONS mismatches
  └─ prints a per-category compatibility table
```

Key files:

| File | Purpose |
|------|---------|
| `script/formula-functions.json` | Master list of all functions and example formulas |
| `test/gsheets-compat/__fixtures__/formula-compat-tests.json` | Generated test cases (with `expectedValue` after patching) |
| `test/gsheets-compat/__fixtures__/formula-compat-input.csv` | CSV to upload to Google Sheets |
| `test/gsheets-compat/__fixtures__/formula-compat-gsheets.csv` | CSV exported from Google Sheets (required for testing) |
| `test/gsheets-compat/gsheets-compat-config.ts` | `MUST_MATCH_FUNCTIONS`, `VOLATILE_FUNCTIONS`, `GSHEETS_ONLY_FUNCTIONS` |
| `test/gsheets-compat/helpers.ts` | Sheet building, value comparison, report formatting |
| `script/parse-gsheets-value.ts` | Parses GSheets CSV values (currency, %, booleans, errors) |

### Running the compat test

`formula-compat-gsheets.csv` must exist before the test can compare values.

```bash
# 1. (One-time / after adding functions) Regenerate test cases
npm run gsheets:generate-tests

# 2. (One-time / after regenerating) Rebuild the upload CSV
npm run gsheets:generate-csv
# → Import formula-compat-input.csv into Google Sheets
# → Export evaluated sheet as CSV → formula-compat-gsheets.csv

# 3. Patch expected values from the GSheets export
npm run gsheets:patch-values

# 4. Run the compatibility test
npm run test:gsheets-compat
```

Steps 1–3 only need to be repeated when `formula-functions.json` changes or when you want to refresh the GSheets reference values.

### Checking compatibility without running tests

For a quick report that includes overall %, per-category breakdown, and a list of functions with no GSheets reference value:

```bash
npm run gsheets:compat-report
```

This evaluates all formulas directly without Jest — useful for a fast sanity check during development.

### Updating the GSheets reference CSV

The `formula-compat-gsheets.csv` file is committed to the repository so the test can run without manual GSheets access. Update it when:
- New functions are added to `formula-functions.json`
- Existing test formulas change
- You want to capture updated GSheets behaviour

Workflow to update:
1. `npm run gsheets:generate-csv` — regenerates the input CSV
2. Import `formula-compat-input.csv` into a fresh Google Sheets document
3. Export the sheet: **File → Download → Comma-separated values (.csv)**
4. Save as `test/gsheets-compat/__fixtures__/formula-compat-gsheets.csv`
5. `npm run gsheets:patch-values` — writes `expectedValue` into `formula-compat-tests.json`
6. Commit both changed fixture files

### Adding a function to MUST_MATCH_FUNCTIONS

When a function's compatibility is fixed and verified, add it to `MUST_MATCH_FUNCTIONS` in `test/gsheets-compat/gsheets-compat-config.ts`. The test will then fail if it regresses.

---

## Sources of the function translations

HF supports internationalization and provides the localized function names for all built-in languages. When looking for the valid translations for the new functions, try these sources:
- https://support.microsoft.com/en-us/office/excel-functions-translator-f262d0c0-991c-485b-89b6-32cc8d326889
- http://dolf.trieschnigg.nl/excel/index.php
