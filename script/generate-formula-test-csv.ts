#!/usr/bin/env ts-node
/**
 * Generates a CSV for formula compatibility testing with Google Sheets.
 *
 * Reads formula-compat-tests.json and produces a CSV where column C
 * contains =FORMULA expressions that GSheets will evaluate.
 *
 * Rows 1-513:  formula rows (header + test entries)
 * Rows 514-699: empty padding
 * Rows 700+:   cell data blocks for formulas that need real cell ranges
 *
 * Output: test/gsheets-compat/__fixtures__/formula-compat-input.csv
 *
 * Usage: npx ts-node --transpile-only -O '{"module":"commonjs"}' script/generate-formula-test-csv.ts
 */

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import Papa from "papaparse";

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

const JSON_PATH = resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-compat-tests.json');
const OUTPUT_PATH = resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-compat-input.csv');

const raw = readFileSync(JSON_PATH, "utf-8");
const data: FormulaCompatTests = JSON.parse(raw);

const rows: string[][] = [];

// Header
rows.push(["name", "formula_text", "formula_eval"]);

let evalCount = 0;
let skipCount = 0;

for (const entry of Object.values(data.functions)) {
  const skip = entry.volatile || entry.gSheetsOnly || entry.needsCellRefs;

  for (let i = 0; i < entry.tests.length; i++) {
    const test = entry.tests[i];
    if (!test) continue;

    const suffix = i === 0 ? "" : ` #${i + 1}`;
    const rowName = `${entry.name}${suffix}`;
    const formulaText = test.formula;

    if (skip) {
      // Include as plain text (no = prefix) so GSheets doesn't evaluate
      rows.push([rowName, formulaText, formulaText]);
      skipCount++;
    } else {
      // Include as evaluable formula
      rows.push([rowName, formulaText, `=${formulaText}`]);
      evalCount++;
    }
  }
}

// ---------------------------------------------------------------------------
// Pad to row 700 and emit cell data blocks
// ---------------------------------------------------------------------------

// rows array is 0-indexed; rows[0] is the CSV header (row 1 in GSheets).
// To place data starting at row 700 (1-indexed), we need rows[699].
const DATA_START_ROW = 700; // 1-indexed row number

// Pad with empty rows until we reach row 699 (0-indexed: 698)
while (rows.length < DATA_START_ROW - 1) {
  rows.push([]);
}

/**
 * Sets a value in the rows array at the given 1-indexed row and 0-indexed column.
 * Extends the row with empty strings if needed.
 */
function setCell(row1: number, col: number, value: string): void {
  const idx = row1 - 1; // 0-indexed
  while (rows.length <= idx) {
    rows.push([]);
  }
  const row = rows[idx]!;
  while (row.length <= col) {
    row.push("");
  }
  row[col] = value;
}

// Column indices: A=0, B=1, C=2, D=3, E=4, F=5

// ── Numeric block (rows 700-710) ──────────────────────────────────────
setCell(700, 0, "[NUMERIC DATA]");

// Col A: numeric (10-55)
const numericA = [10, 20, 30, 40, 50, 15, 25, 35, 45, 55];
// Col B: numeric2
const numericB = [12, 22, 28, 42, 48, 18, 24, 38, 44, 52];
// Col C: secondary range (100-1000)
const numericC = [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000];
// Col D: criteria values (5-50)
const numericD = [5, 10, 15, 20, 25, 30, 35, 40, 45, 50];
// Col E: booleans
const numericE = ["TRUE", "FALSE"];

for (let i = 0; i < 10; i++) {
  const row = 701 + i;
  setCell(row, 0, String(numericA[i]));
  setCell(row, 1, String(numericB[i]));
  setCell(row, 2, String(numericC[i]));
  setCell(row, 3, String(numericD[i]));
  if (i < numericE.length) {
    setCell(row, 4, numericE[i]!);
  }
}

// ── Lookup block (rows 720-725) ───────────────────────────────────────
setCell(720, 0, "[LOOKUP DATA]");

const lookupIds = [10001, 10002, 10003, 10004, 10005];
const lookupNames = ["Alpha", "Bravo", "Charlie", "Delta", "Echo"];

for (let i = 0; i < 5; i++) {
  const row = 721 + i;
  setCell(row, 0, String(lookupIds[i]));
  setCell(row, 1, lookupNames[i]!);
}

// ── 3x3 Matrix (rows 730-733) ────────────────────────────────────────
setCell(730, 0, "[3x3 MATRIX]");

const matrix3x3 = [
  [1, 2, 3],
  [4, 5, 6],
  [7, 8, 9],
];
for (let r = 0; r < 3; r++) {
  for (let c = 0; c < 3; c++) {
    setCell(731 + r, c, String(matrix3x3[r]![c]));
  }
}

// ── 4x4 Matrix (rows 735-739) ────────────────────────────────────────
setCell(735, 0, "[4x4 MATRIX]");

const matrix4x4 = [
  [2, 1, 0, 3],
  [1, 0, 2, 1],
  [0, 2, 1, 0],
  [1, 1, 1, 2],
];
for (let r = 0; r < 4; r++) {
  for (let c = 0; c < 4; c++) {
    setCell(736 + r, c, String(matrix4x4[r]![c]));
  }
}

// ── MMULT matrices (rows 741-748) ────────────────────────────────────
setCell(741, 0, "[MMULT A 3x2]");

const mmultA = [
  [1, 2],
  [3, 4],
  [5, 6],
];
for (let r = 0; r < 3; r++) {
  for (let c = 0; c < 2; c++) {
    setCell(742 + r, c, String(mmultA[r]![c]));
  }
}

setCell(746, 0, "[MMULT B 2x4]");

const mmultB = [
  [7, 8, 9, 10],
  [11, 12, 13, 14],
];
for (let r = 0; r < 2; r++) {
  for (let c = 0; c < 4; c++) {
    setCell(747 + r, c, String(mmultB[r]![c]));
  }
}

// ── Database block (rows 750-756) ────────────────────────────────────
setCell(750, 0, "[DATABASE]");

// Header row
const dbHeaders = ["Name", "Age", "Dept", "Salary", "Rating", "Active"];
for (let c = 0; c < dbHeaders.length; c++) {
  setCell(751, c, dbHeaders[c]!);
}

// Data rows
const dbData: (string | number | boolean)[][] = [
  ["Alice", 30, "Eng", 80000, 4, "TRUE"],
  ["Bob", 25, "Sales", 60000, 3, "TRUE"],
  ["Charlie", 35, "Eng", 90000, 5, "FALSE"],
  ["Diana", 28, "Sales", 65000, 4, "TRUE"],
  ["Eve", 32, "Eng", 85000, 3, "TRUE"],
];
for (let r = 0; r < dbData.length; r++) {
  for (let c = 0; c < dbData[r]!.length; c++) {
    setCell(752 + r, c, String(dbData[r]![c]));
  }
}

// ── DB Criteria blocks ───────────────────────────────────────────────
// Multi-match criteria (Dept=Eng) at rows 758-760
setCell(758, 0, "[DB CRITERIA - multi-match]");
setCell(759, 0, "Dept");
setCell(760, 0, "Eng");

// Unique match criteria (Name=Alice) for DGET at rows 762-764
setCell(762, 0, "[DB CRITERIA - unique match for DGET]");
setCell(763, 0, "Name");
setCell(764, 0, "Alice");

// ── HLOOKUP horizontal block (rows 766-768) ─────────────────────────
setCell(766, 0, "[HLOOKUP - horizontal]");

for (let c = 0; c < 5; c++) {
  setCell(767, c, String(lookupIds[c]));
  setCell(768, c, lookupNames[c]!);
}

// ── Special cells (rows 770-772) ─────────────────────────────────────
setCell(770, 0, "[SPECIAL CELLS]");
setCell(771, 0, "42");
setCell(771, 1, "=1+1");  // formula cell
setCell(772, 0, "");       // empty cell
setCell(772, 1, "hello");

// ---------------------------------------------------------------------------
// Write CSV
// ---------------------------------------------------------------------------

const csv = Papa.unparse(rows);
mkdirSync(dirname(OUTPUT_PATH), { recursive: true });
writeFileSync(OUTPUT_PATH, csv, "utf-8");

// eslint-disable-next-line no-console
console.log(
  `Wrote ${rows.length} rows to ${OUTPUT_PATH} (${evalCount} evaluable, ${skipCount} skipped, data blocks at row ${DATA_START_ROW}+)`,
);
