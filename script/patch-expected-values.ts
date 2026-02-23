#!/usr/bin/env ts-node
/**
 * Patches expected values from a Google Sheets CSV export back into
 * formula-compat-tests.json.
 *
 * Usage: npx ts-node --project tsconfig.test.json script/patch-expected-values.ts
 *
 * Reads formula-compat-gsheets.csv (exported from GSheets after evaluation),
 * matches each row by name to the JSON, and writes expectedValue fields.
 */

import { readFileSync, writeFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import Papa from "papaparse";

const JSON_PATH = resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-compat-tests.json');
const CSV_PATH = resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-compat-gsheets.csv');

if (!existsSync(CSV_PATH)) {
  console.error(
    `Missing: ${CSV_PATH}\n` +
      "To generate it:\n" +
      "  1. Run: npx ts-node script/generate-formula-test-csv.ts\n" +
      "  2. Import formula-compat-input.csv into Google Sheets\n" +
      "  3. Export as CSV → formula-compat-gsheets.csv",
  );
  process.exit(1);
}

interface FormulaTestCase {
  formula: string;
  cellData?: Record<string, string | number | boolean | null>;
  expectedValue?: string | number | boolean | null;
  note?: string;
}

interface FormulaTestEntry {
  name: string;
  category: string;
  tests: FormulaTestCase[];
  volatile?: boolean;
  gSheetsOnly?: boolean;
  needsCellRefs?: boolean;
}

interface FormulaCompatTests {
  version: number;
  functions: Record<string, FormulaTestEntry>;
}

const jsonRaw = readFileSync(JSON_PATH, "utf-8");
const data: FormulaCompatTests = JSON.parse(jsonRaw);

const csvText = readFileSync(CSV_PATH, "utf-8");
const parsed = Papa.parse<string[]>(csvText, { header: false });
const rows = parsed.data;

// Build a map: row name → evaluated value (column C)
const resultMap = new Map<string, string>();

// Skip header row
for (let i = 1; i < rows.length; i++) {
  const row = rows[i];
  if (!row || row.length < 3) continue;
  const name = row[0]?.trim();
  const value = row[2] ?? "";
  if (name) {
    resultMap.set(name, value);
  }
}

let patched = 0;
let notFound = 0;
const missing: string[] = [];

for (const entry of Object.values(data.functions)) {
  // Skip volatile/gSheetsOnly/needsCellRefs — their CSV values aren't evaluated
  if (entry.volatile || entry.gSheetsOnly || entry.needsCellRefs) continue;

  for (let i = 0; i < entry.tests.length; i++) {
    const test = entry.tests[i];
    if (!test) continue;

    const suffix = i === 0 ? "" : ` #${i + 1}`;
    const rowName = `${entry.name}${suffix}`;
    const csvValue = resultMap.get(rowName);

    if (csvValue === undefined) {
      notFound++;
      missing.push(rowName);
      continue;
    }

    // Parse the CSV value into an appropriate type
    test.expectedValue = parseGSheetsValue(csvValue);
    patched++;
  }
}

// Write back
writeFileSync(JSON_PATH, JSON.stringify(data, null, 2) + "\n", "utf-8");

console.log(`Patched ${patched} expected values into ${JSON_PATH}`);
if (notFound > 0) {
  console.log(`  ${notFound} rows not found in CSV:`);
  for (const m of missing.slice(0, 10)) {
    console.log(`    - ${m}`);
  }
  if (missing.length > 10) {
    console.log(`    ... and ${missing.length - 10} more`);
  }
}

/**
 * Parses a raw string value from GSheets CSV into a typed value.
 * GSheets exports: numbers as numbers, TRUE/FALSE as booleans,
 * errors as #ERROR! strings, and everything else as strings.
 */
function parseGSheetsValue(raw: string): string | number | boolean | null {
  const trimmed = raw.trim();

  // Empty → null
  if (trimmed === "") return null;

  // Booleans
  if (trimmed === "TRUE") return true;
  if (trimmed === "FALSE") return false;

  // Error strings (keep as-is)
  if (trimmed.startsWith("#")) return trimmed;

  // Try numeric
  const num = Number(trimmed);
  if (!isNaN(num) && trimmed !== "") return num;

  // Default: string
  return trimmed;
}
