#!/usr/bin/env ts-node

/**
 * Generates formula-compat-tests.json from formula-functions.json.
 *
 * This is a deterministic rewriter that converts cell-reference formulas
 * into self-contained formulas using inline arrays, so they can be evaluated
 * in HyperFormula without needing real sheet data.
 *
 * Usage:
 *   npx ts-node script/generate-formula-compat-tests.ts
 */

import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";

const formulaFunctions = JSON.parse(
  readFileSync(resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-functions.json'), 'utf-8')
);

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

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

interface FormulaRewrite {
  formula: string;
  note?: string;
  cellData?: Record<string, string | number | boolean | null>;
}

// ---------------------------------------------------------------------------
// Inline data constants
// ---------------------------------------------------------------------------

const INLINE_DATA = {
  // 10-element numeric column
  numeric: "{10;20;30;40;50;15;25;35;45;55}",
  // Second numeric column (for correlation, paired tests)
  numeric2: "{12;22;28;42;48;18;24;38;44;52}",
  // Small 5-element set
  numericSmall: "{10;20;30;40;50}",
  // For COUNTIF/SUMIF style: mix of values including ones that match ">20"
  numericMixed: "{5;10;15;20;25;30;35;40;45;50}",
  // For AVERAGEIF second range
  numericMixed2: "{100;200;300;400;500;600;700;800;900;1000}",
  // Lookup table: 2 columns (key, value)
  lookupKeys: "{10001;10002;10003;10004;10005}",
  lookupValues: '{"Alpha";"Bravo";"Charlie";"Delta";"Echo"}',
  // For HLOOKUP: row-oriented lookup
  hlookupRow1: "{10001,10002,10003,10004,10005}",
  hlookupRow2: '{"Alpha","Bravo","Charlie","Delta","Echo"}',
  // Matrix 3x3
  matrix3x3: "{1,2,3;4,5,6;7,8,9}",
  // Matrix 4x4 (for MDETERM, MINVERSE)
  matrix4x4: "{2,1,0,3;1,0,2,1;0,2,1,0;1,1,1,2}",
  // Matrix for MMULT: 3x2 and 2x4
  matrixA: "{1,2;3,4;5,6}",
  matrixB: "{7,8,9,10;11,12,13,14}",
  // Dates for XIRR/XNPV (as date strings)
  cashflows: "{-10000;2750;4250;3250;2750}",
  dates: '{"1/1/2008";"3/1/2008";"10/30/2008";"2/15/2009";"4/1/2009"}',
  // For IRR/MIRR: investment cashflows
  irrCashflows: "{-70000;12000;15000;18000;21000;26000}",
  // Frequency bins
  freqData: "{1;3;5;7;9;2;4;6;8;10}",
  freqBins: "{3;6;9}",
  // For GROWTH/TREND
  yValues: "{33100;47300;69000;102000;150000;220000}",
  xValues: "{1;2;3;4;5;6}",
  newXValues: "{7;8;9}",
  // For paired arrays (8 elements each, used by SUMX2MY2 etc.)
  paired1: "{2;3;9;1;8;7;5;6}",
  paired2: "{6;5;11;7;5;4;4;3}",
  // Series coefficients for SERIESSUM
  seriesCoeffs:
    "{1;-0.5;0.1667;-0.0417;0.0083;-0.0014;0.0002;-0.00002;0.000003}",
  // For FVSCHEDULE
  interestRates: "{0.09;0.11;0.10}",
  // GCD/LCM small set
  gcdSet: "{24;36;48;12}",
  // For T.TEST / paired tests (4 elements)
  testA: "{3;4;1;2}",
  testB: "{5;3;4;2}",
  // Chi-square test data
  chiObserved: "{10;15;20;25;30}",
  chiExpected: "{12;14;18;26;30}",
  // For PROB
  probValues: "{0;1;2;3;4}",
  probProbs: "{0.1;0.2;0.3;0.25;0.15}",
  // Boolean-ish mixed data for COUNTA etc
  mixedData: '{1;"a";TRUE;0;"";2;"b";FALSE;3;"c"}',
  // Holiday dates for NETWORKDAYS/WORKDAY
  holidays: '{"7/20/1969";"7/21/1969"}',
} as const;

// ---------------------------------------------------------------------------
// Classification sets
// ---------------------------------------------------------------------------

const VOLATILE_FUNCTIONS = new Set([
  "NOW",
  "TODAY",
  "RAND",
  "RANDBETWEEN",
  "RANDARRAY",
]);

const GSHEETS_ONLY = new Set([
  "GOOGLEFINANCE",
  "GOOGLETRANSLATE",
  "IMPORTDATA",
  "IMPORTFEED",
  "IMPORTHTML",
  "IMPORTRANGE",
  "IMPORTXML",
  "IMAGE",
  "SPARKLINE",
  "QUERY",
  "GETPIVOTDATA",
  "DETECTLANGUAGE",
]);

// These were formerly in NEEDS_CELL_REFS — now handled by CELL_REF_FORMULAS
const NEEDS_CELL_REFS = new Set<string>([]);

// ---------------------------------------------------------------------------
// Formula rewrites map
// ---------------------------------------------------------------------------

const FORMULA_REWRITES: Record<string, FormulaRewrite> = {
  // ── Single range → numeric ──────────────────────────────────────────
  SUM: { formula: `SUM(${INLINE_DATA.numeric}, 101)` },
  PRODUCT: { formula: `PRODUCT(${INLINE_DATA.numericSmall}, 2)` },
  SUMSQ: { formula: `SUMSQ(${INLINE_DATA.numeric}, 10)` },
  COUNTBLANK: {
    formula: `COUNTBLANK(${INLINE_DATA.mixedData})`,
    note: "Mixed data includes empty string",
  },
  TRIMMEAN: { formula: `TRIMMEAN(${INLINE_DATA.numeric}, 0.1)` },
  PERCENTILE: { formula: `PERCENTILE(${INLINE_DATA.numeric}, 0.95)` },
  "PERCENTILE.EXC": {
    formula: `PERCENTILE.EXC(${INLINE_DATA.numeric}, 0.25)`,
  },
  "PERCENTILE.INC": {
    formula: `PERCENTILE.INC(${INLINE_DATA.numeric}, 0.95)`,
  },
  QUARTILE: { formula: `QUARTILE(${INLINE_DATA.numeric}, 3)` },
  "QUARTILE.EXC": { formula: `QUARTILE.EXC(${INLINE_DATA.numeric}, 3)` },
  "QUARTILE.INC": { formula: `QUARTILE.INC(${INLINE_DATA.numeric}, 3)` },
  MEDIAN: {
    formula: `MEDIAN(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  MAX: { formula: `MAX(${INLINE_DATA.numeric}, 42)` },
  MAXA: { formula: `MAXA(${INLINE_DATA.numeric}, 42)` },
  MIN: { formula: `MIN(${INLINE_DATA.numeric}, 5)` },
  MINA: { formula: `MINA(${INLINE_DATA.numeric}, 5)` },
  MODE: { formula: `MODE(${INLINE_DATA.numeric}, 5)` },
  "MODE.MULT": { formula: `MODE.MULT(${INLINE_DATA.numeric}, 5)` },
  "MODE.SNGL": { formula: `MODE.SNGL(${INLINE_DATA.numeric}, 5)` },
  LARGE: { formula: `LARGE(${INLINE_DATA.numeric}, 4)` },
  SMALL: { formula: `SMALL(${INLINE_DATA.numeric}, 4)` },
  "Z.TEST": { formula: `Z.TEST(${INLINE_DATA.numeric}, 95, 1.2)` },
  ZTEST: { formula: `ZTEST(${INLINE_DATA.numeric}, 95, 1.2)` },

  // ── Two ranges → numeric + numeric2 ────────────────────────────────
  AVERAGE: {
    formula: `AVERAGE(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  AVERAGEA: {
    formula: `AVERAGEA(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  COUNT: {
    formula: `COUNT(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  COUNTA: {
    formula: `COUNTA(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  CORREL: {
    formula: `CORREL(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  COVAR: {
    formula: `COVAR(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  "COVARIANCE.P": {
    formula: `COVARIANCE.P(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  "COVARIANCE.S": {
    formula: `COVARIANCE.S(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  PEARSON: {
    formula: `PEARSON(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  RSQ: { formula: `RSQ(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})` },
  INTERCEPT: {
    formula: `INTERCEPT(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  SLOPE: {
    formula: `SLOPE(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  STEYX: {
    formula: `STEYX(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  FORECAST: {
    formula: `FORECAST(5, ${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },
  "FORECAST.LINEAR": {
    formula: `FORECAST.LINEAR(5, ${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2})`,
  },

  // ── Conditional single range ────────────────────────────────────────
  COUNTIF: { formula: `COUNTIF(${INLINE_DATA.numericMixed}, ">20")` },
  // SUMIF, AVERAGEIF → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── Conditional multi-range ─────────────────────────────────────────
  COUNTIFS: {
    formula: `COUNTIFS(${INLINE_DATA.numericMixed}, ">20", ${INLINE_DATA.numericMixed2}, "<800")`,
  },
  // AVERAGEIFS, SUMIFS, MAXIFS, MINIFS → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── Statistical tests (two equal-length ranges) ─────────────────────
  "CHISQ.TEST": {
    formula: `CHISQ.TEST(${INLINE_DATA.chiObserved}, ${INLINE_DATA.chiExpected})`,
  },
  CHITEST: {
    formula: `CHITEST(${INLINE_DATA.chiObserved}, ${INLINE_DATA.chiExpected})`,
  },
  "F.TEST": {
    formula: `F.TEST(${INLINE_DATA.chiObserved}, ${INLINE_DATA.chiExpected})`,
  },
  FTEST: {
    formula: `FTEST(${INLINE_DATA.chiObserved}, ${INLINE_DATA.chiExpected})`,
  },
  "T.TEST": {
    formula: `T.TEST(${INLINE_DATA.testA}, ${INLINE_DATA.testB}, 2, 1)`,
  },
  TTEST: {
    formula: `TTEST(${INLINE_DATA.testA}, ${INLINE_DATA.testB}, 2, 1)`,
  },

  // ── Paired arrays ───────────────────────────────────────────────────
  SUMX2MY2: {
    formula: `SUMX2MY2(${INLINE_DATA.paired1}, ${INLINE_DATA.paired2})`,
  },
  SUMX2PY2: {
    formula: `SUMX2PY2(${INLINE_DATA.paired1}, ${INLINE_DATA.paired2})`,
  },
  SUMXMY2: {
    formula: `SUMXMY2(${INLINE_DATA.paired1}, ${INLINE_DATA.paired2})`,
  },
  SUMPRODUCT: {
    formula: `SUMPRODUCT(${INLINE_DATA.paired1}, ${INLINE_DATA.paired2})`,
    note: "Using 8-element paired arrays",
  },

  // ── Lookup functions ────────────────────────────────────────────────
  // VLOOKUP, HLOOKUP → moved to CELL_REF_FORMULAS (HF returns #VALUE! with inline arrays)
  XLOOKUP: {
    formula: `XLOOKUP(10003, ${INLINE_DATA.lookupKeys}, ${INLINE_DATA.lookupValues}, "missing", 0, 1)`,
  },
  LOOKUP: {
    formula: `LOOKUP(10003, ${INLINE_DATA.lookupKeys}, ${INLINE_DATA.lookupValues})`,
  },
  MATCH: {
    formula: `MATCH("Sunday", {"Monday";"Tuesday";"Wednesday";"Thursday";"Friday";"Saturday";"Sunday"}, 0)`,
  },
  INDEX: {
    formula: `INDEX({10,20,30;40,50,60;70,80,90}, 2, 3)`,
    note: "Returns 60",
  },

  // ── Matrix ──────────────────────────────────────────────────────────
  MDETERM: { formula: `MDETERM(${INLINE_DATA.matrix4x4})` },
  MINVERSE: { formula: `MINVERSE(${INLINE_DATA.matrix4x4})` },
  // MMULT → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── GROWTH/TREND/LINEST/LOGEST ──────────────────────────────────────
  GROWTH: {
    formula: `GROWTH(${INLINE_DATA.yValues}, ${INLINE_DATA.xValues}, ${INLINE_DATA.newXValues}, TRUE)`,
  },
  TREND: {
    formula: `TREND(${INLINE_DATA.yValues}, ${INLINE_DATA.xValues}, ${INLINE_DATA.newXValues}, TRUE)`,
  },
  LINEST: {
    formula: `LINEST(${INLINE_DATA.yValues}, ${INLINE_DATA.xValues}, FALSE, TRUE)`,
  },
  LOGEST: {
    formula: `LOGEST(${INLINE_DATA.yValues}, ${INLINE_DATA.xValues}, TRUE, TRUE)`,
  },

  // ── FREQUENCY ───────────────────────────────────────────────────────
  FREQUENCY: {
    formula: `FREQUENCY(${INLINE_DATA.freqData}, ${INLINE_DATA.freqBins})`,
  },

  // ── Financial: IRR/MIRR/XIRR/XNPV ──────────────────────────────────
  IRR: { formula: `IRR(${INLINE_DATA.irrCashflows}, 1%)` },
  MIRR: { formula: `MIRR(${INLINE_DATA.irrCashflows}, 8%, 11%)` },
  XIRR: {
    formula: `XIRR(${INLINE_DATA.cashflows}, ${INLINE_DATA.dates}, 1%)`,
  },
  XNPV: {
    formula: `XNPV(8%, ${INLINE_DATA.cashflows}, ${INLINE_DATA.dates})`,
  },
  FVSCHEDULE: {
    formula: `FVSCHEDULE(10000, ${INLINE_DATA.interestRates})`,
  },

  // ── PROB ─────────────────────────────────────────────────────────────
  PROB: {
    formula: `PROB(${INLINE_DATA.probValues}, ${INLINE_DATA.probProbs}, 2, 4)`,
  },

  // ── RANK ─────────────────────────────────────────────────────────────
  RANK: { formula: `RANK(42, ${INLINE_DATA.numeric}, 1)` },
  "RANK.AVG": { formula: `RANK.AVG(42, ${INLINE_DATA.numeric}, TRUE)` },
  "RANK.EQ": { formula: `RANK.EQ(42, ${INLINE_DATA.numeric}, TRUE)` },

  // ── PERCENTRANK ─────────────────────────────────────────────────────
  PERCENTRANK: { formula: `PERCENTRANK(${INLINE_DATA.numeric}, 42, 4)` },
  "PERCENTRANK.EXC": {
    formula: `PERCENTRANK.EXC(${INLINE_DATA.numeric}, 42, 4)`,
  },
  "PERCENTRANK.INC": {
    formula: `PERCENTRANK.INC(${INLINE_DATA.numeric}, 42, 4)`,
  },

  // ── ROWS/COLUMNS ────────────────────────────────────────────────────
  ROWS: {
    formula: `ROWS({1;2;3;4;5;6;7;8;9;10;11;12;13;14;15;16;17;18;19;20;21;22;23;24;25;26;27;28;29;30;31;32;33;34;35;36;37;38;39;40;41;42;43;44;45;46;47;48;49;50;51;52;53;54})`,
    note: "Returns 54",
  },
  COLUMNS: {
    formula: `COLUMNS({1,2,3,4,5,6,7,8,9})`,
    note: "Returns 9",
  },
  COLUMN: {
    formula: `COLUMN(C9)`,
    note: "Returns 3 — HyperFormula can resolve the ref without cell data",
  },
  ROW: {
    formula: `ROW(A9)`,
    note: "Returns 9 — HyperFormula can resolve the ref without cell data",
  },

  // SUBTOTAL → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // NETWORKDAYS, NETWORKDAYS.INTL, WORKDAY, WORKDAY.INTL → moved to CELL_REF_FORMULAS
  // (GSheets returns #VALUE! when holiday arrays are passed as inline arrays)

  // ── GCD/LCM ─────────────────────────────────────────────────────────
  GCD: { formula: `GCD(${INLINE_DATA.gcdSet}, 6)` },
  LCM: { formula: `LCM(${INLINE_DATA.gcdSet}, 6)` },

  // ── COUNTUNIQUE ─────────────────────────────────────────────────────
  COUNTUNIQUE: {
    formula: `COUNTUNIQUE({1;2;3;2;1;4;5;3;6;7}, 100)`,
    note: "8 unique values in the array + 100 = 9 unique",
  },

  // ── JOIN ─────────────────────────────────────────────────────────────
  JOIN: {
    formula: `JOIN("-", {"a";"b";"c"}, {"d";"e";"f"})`,
  },

  // ── LET ──────────────────────────────────────────────────────────────
  LET: {
    formula: `LET(x, SUM({1;2;3;4;5}), y, AVERAGE({1;2;3;4;5}), x*x + y*y)`,
  },

  // FILTER → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── SORT / SORTN ────────────────────────────────────────────────────
  SORT: {
    formula: `SORT({5,"e";3,"c";1,"a";4,"d";2,"b"}, 1, TRUE)`,
  },
  SORTN: {
    formula: `SORTN({5,"e";3,"c";1,"a";4,"d";2,"b"}, 2, 0, 1, TRUE)`,
  },

  // ── UNIQUE ──────────────────────────────────────────────────────────
  UNIQUE: {
    formula: `UNIQUE({1;2;3;2;1;4;5;3;6;7})`,
  },

  // ── FLATTEN ─────────────────────────────────────────────────────────
  FLATTEN: {
    formula: `FLATTEN({1,2;3,4}, {5;6})`,
  },

  // TRANSPOSE → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── TOCOL / TOROW ───────────────────────────────────────────────────
  TOCOL: {
    formula: `TOCOL({1,2,3;4,5,6}, 0, FALSE)`,
  },
  TOROW: {
    formula: `TOROW({1,2,3;4,5,6}, 0, FALSE)`,
  },

  // ── HSTACK / VSTACK ─────────────────────────────────────────────────
  HSTACK: {
    formula: `HSTACK({1;2;3}, {4;5;6})`,
  },
  VSTACK: {
    formula: `VSTACK({1,2,3}, {4,5,6})`,
  },

  // ── WRAPCOLS / WRAPROWS ─────────────────────────────────────────────
  WRAPCOLS: {
    formula: `WRAPCOLS({1;2;3;4;5;6;7;8}, 3, "Pad")`,
  },
  WRAPROWS: {
    formula: `WRAPROWS({1;2;3;4;5;6;7;8}, 3, "Pad")`,
  },

  // ── CHOOSECOLS / CHOOSEROWS ─────────────────────────────────────────
  CHOOSECOLS: {
    formula: `CHOOSECOLS(${INLINE_DATA.matrix3x3}, 2)`,
  },
  CHOOSEROWS: {
    formula: `CHOOSEROWS(${INLINE_DATA.matrix3x3}, 2)`,
  },

  // ARRAY_CONSTRAIN → moved to CELL_REF_FORMULAS (GSheets needs real ranges)

  // ── ARRAYFORMULA ────────────────────────────────────────────────────
  ARRAYFORMULA: {
    formula: `ARRAYFORMULA({1,2,3}+{4,5,6})`,
  },

  // ── BYCOL / BYROW ──────────────────────────────────────────────────
  BYCOL: {
    formula: `BYCOL(${INLINE_DATA.matrix3x3}, LAMBDA(column, SUM(column)))`,
  },
  BYROW: {
    formula: `BYROW(${INLINE_DATA.matrix3x3}, LAMBDA(row, MAX(row)))`,
  },

  // ── MAP / REDUCE / SCAN ─────────────────────────────────────────────
  MAP: {
    formula: `MAP({1;2;3}, {4;5;6}, LAMBDA(cell1, cell2, MAX(cell1, cell2)))`,
  },
  REDUCE: {
    formula: `REDUCE(0, {1;2;3;4;5;6;7;8;9}, LAMBDA(total, value, total + value))`,
  },
  SCAN: {
    formula: `SCAN(0, {1;2;3;4;5;6;7;8;9}, LAMBDA(cumulative, value, cumulative + value))`,
  },

  // ── MARGINOFERROR ───────────────────────────────────────────────────
  MARGINOFERROR: {
    formula: `MARGINOFERROR({10;20;30;40;50;60}, 0.95)`,
  },

  // ── IMPRODUCT / IMSUM ───────────────────────────────────────────────
  IMPRODUCT: {
    formula: `IMPRODUCT({"3+4i";"1+2i";"2-i"}, "4+3i")`,
  },
  IMSUM: {
    formula: `IMSUM({"3+4i";"1+2i";"2-i"}, "4+3i")`,
  },

  // ── AVERAGE.WEIGHTED ────────────────────────────────────────────────
  "AVERAGE.WEIGHTED": {
    formula: `AVERAGE.WEIGHTED(${INLINE_DATA.numeric}, ${INLINE_DATA.numeric2}, 5, 0.5)`,
    note: "Weighted average with inline arrays and extra scalar args",
  },

  // ── SERIESSUM ───────────────────────────────────────────────────────
  SERIESSUM: {
    formula: `SERIESSUM(3, 0, 2, ${INLINE_DATA.seriesCoeffs})`,
  },

  // ── Logical with cell refs → literal values ─────────────────────────
  AND: { formula: `AND(TRUE, FALSE)` },
  OR: { formula: `OR(TRUE, FALSE)` },
  XOR: { formula: `XOR(TRUE, FALSE)` },
  IF: {
    formula: `IF("foo"="foo", "is foo", "is not foo")`,
    note: "Rewritten from cell ref example",
  },
  IFS: {
    formula: `IFS(95>90, "A", 95>80, "B")`,
    note: "Rewritten from cell ref example",
  },
  SWITCH: {
    formula: `SWITCH(1, 0, "Zero", 1, "One")`,
    note: "Rewritten from cell ref example",
  },

  // ── IS* functions that can use literal values ───────────────────────
  ISBLANK: { formula: `ISBLANK("")` },
  ISNUMBER: { formula: `ISNUMBER(42)` },
  ISTEXT: { formula: `ISTEXT("hello")` },
  ISLOGICAL: { formula: `ISLOGICAL(TRUE)` },
  ISNA: {
    formula: `ISNA(MATCH(0, {1}, 0))`,
    note: "MATCH returns #N/A when not found",
  },
  ISERROR: { formula: `ISERROR(1/0)` },
  ISERR: { formula: `ISERR(1/0)` },
  ISNONTEXT: { formula: `ISNONTEXT(42)` },
  "ERROR.TYPE": { formula: `ERROR.TYPE(1/0)` },
  TYPE: { formula: `TYPE(42)`, note: "Returns 1 for numbers" },

  // Special cell functions + Database functions → moved to CELL_REF_FORMULAS

  // ── NPV ─────────────────────────────────────────────────────────────
  // NPV example doesn't have cell refs ("NPV(8%, 200, 250)"), but keeping
  // it here for visibility — it will be self-contained by default.

  // ── TEXTJOIN with range → inline ────────────────────────────────────
  // Original example has no cell refs: TEXTJOIN(" ", TRUE, "hello", "world")
  // Self-contained, no rewrite needed.

  // ── SPLIT — no cell refs ────────────────────────────────────────────
  // Self-contained, no rewrite needed.
};

// ---------------------------------------------------------------------------
// Cell-ref formulas: use real cell ranges at row 700+ in the CSV
// ---------------------------------------------------------------------------

// Shared cell data for the numeric block (rows 701-710, cols A-E)
const NUMERIC_CELL_DATA: Record<string, number | boolean> = {
  // Col A: numeric (10-55)
  A701: 10, A702: 20, A703: 30, A704: 40, A705: 50,
  A706: 15, A707: 25, A708: 35, A709: 45, A710: 55,
  // Col B: numeric2
  B701: 12, B702: 22, B703: 28, B704: 42, B705: 48,
  B706: 18, B707: 24, B708: 38, B709: 44, B710: 52,
  // Col C: secondary range (100-1000)
  C701: 100, C702: 200, C703: 300, C704: 400, C705: 500,
  C706: 600, C707: 700, C708: 800, C709: 900, C710: 1000,
  // Col D: criteria values (5-50)
  D701: 5, D702: 10, D703: 15, D704: 20, D705: 25,
  D706: 30, D707: 35, D708: 40, D709: 45, D710: 50,
  // Col E: booleans
  E701: true, E702: false,
};

// Lookup block (rows 721-725)
const LOOKUP_CELL_DATA: Record<string, number | string> = {
  A721: 10001, B721: "Alpha",
  A722: 10002, B722: "Bravo",
  A723: 10003, B723: "Charlie",
  A724: 10004, B724: "Delta",
  A725: 10005, B725: "Echo",
};

// 3x3 matrix (rows 731-733)
const MATRIX_3X3_CELL_DATA: Record<string, number> = {
  A731: 1, B731: 2, C731: 3,
  A732: 4, B732: 5, C732: 6,
  A733: 7, B733: 8, C733: 9,
};

// MMULT matrices: A (3x2, rows 742-744) and B (2x4, rows 747-748)
const MMULT_CELL_DATA: Record<string, number> = {
  A742: 1, B742: 2,
  A743: 3, B743: 4,
  A744: 5, B744: 6,
  A747: 7, B747: 8, C747: 9, D747: 10,
  A748: 11, B748: 12, C748: 13, D748: 14,
};

// Database block (rows 751-756) + criteria (759-760, 763-764)
const DATABASE_CELL_DATA: Record<string, string | number | boolean> = {
  // Header
  A751: "Name", B751: "Age", C751: "Dept", D751: "Salary", E751: "Rating", F751: "Active",
  // Data rows
  A752: "Alice", B752: 30, C752: "Eng", D752: 80000, E752: 4, F752: true,
  A753: "Bob", B753: 25, C753: "Sales", D753: 60000, E753: 3, F753: true,
  A754: "Charlie", B754: 35, C754: "Eng", D754: 90000, E754: 5, F754: false,
  A755: "Diana", B755: 28, C755: "Sales", D755: 65000, E755: 4, F755: true,
  A756: "Eve", B756: 32, C756: "Eng", D756: 85000, E756: 3, F756: true,
  // Criteria: Dept = Eng
  A759: "Dept", A760: "Eng",
  // Criteria: Name = Alice (unique match for DGET)
  A763: "Name", A764: "Alice",
};

// HLOOKUP horizontal block (rows 767-768)
const HLOOKUP_CELL_DATA: Record<string, number | string> = {
  A767: 10001, B767: 10002, C767: 10003, D767: 10004, E767: 10005,
  A768: "Alpha", B768: "Bravo", C768: "Charlie", D768: "Delta", E768: "Echo",
};

// Special cells (rows 771-772)
const SPECIAL_CELL_DATA: Record<string, number | string> = {
  A771: 42,
  B771: "=1+1",  // formula cell
  A772: "",       // empty cell
  B772: "hello",
};

// Holiday dates for NETWORKDAYS/WORKDAY (rows 781-782)
// GSheets requires real cell ranges for the holidays parameter; inline arrays return #VALUE!
const HOLIDAY_CELL_DATA: Record<string, string> = {
  A781: "7/20/1969",
  A782: "7/21/1969",
};

/**
 * Formulas that need real cell ranges in the CSV (row 700+).
 * Each entry provides the formula referencing fixed cell positions
 * and cellData so the local HyperFormula test can evaluate them too.
 */
const CELL_REF_FORMULAS: Record<string, FormulaRewrite> = {
  // ── GSheets inline array failures → cell ranges (10) ─────────────
  SUMIF: {
    formula: `SUMIF(D701:D710, ">20", C701:C710)`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  SUMIFS: {
    formula: `SUMIFS(C701:C710, D701:D710, ">20", C701:C710, "<800")`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  AVERAGEIF: {
    formula: `AVERAGEIF(D701:D710, ">20", C701:C710)`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  MAXIFS: {
    formula: `MAXIFS(A701:A710, A701:A710, "<100", A701:A710, ">15")`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  MINIFS: {
    formula: `MINIFS(A701:A710, A701:A710, "<100", A701:A710, ">15")`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  AVERAGEIFS: {
    formula: `AVERAGEIFS(C701:C710, D701:D710, ">20", C701:C710, "<800")`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  SUBTOTAL: {
    formula: `SUBTOTAL(1, A701:A705)`,
    cellData: { ...NUMERIC_CELL_DATA },
    note: "Function 1 = AVERAGE",
  },
  FILTER: {
    formula: `FILTER(A701:A710, A701:A710>30)`,
    cellData: { ...NUMERIC_CELL_DATA },
  },
  TRANSPOSE: {
    formula: `TRANSPOSE(A731:C733)`,
    cellData: { ...MATRIX_3X3_CELL_DATA },
  },
  MMULT: {
    formula: `MMULT(A742:B744, A747:D748)`,
    cellData: { ...MMULT_CELL_DATA },
  },
  ARRAY_CONSTRAIN: {
    formula: `ARRAY_CONSTRAIN(A731:C733, 2, 2)`,
    cellData: { ...MATRIX_3X3_CELL_DATA },
  },

  // ── Lookup → cell ranges (2) ─────────────────────────────────────
  VLOOKUP: {
    formula: `VLOOKUP(10003, A721:B725, 2, FALSE)`,
    cellData: { ...LOOKUP_CELL_DATA },
  },
  HLOOKUP: {
    formula: `HLOOKUP(10003, A767:E768, 2, FALSE)`,
    cellData: { ...HLOOKUP_CELL_DATA },
  },

  // ── Database functions (12) ───────────────────────────────────────
  DAVERAGE: {
    formula: `DAVERAGE(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DCOUNT: {
    formula: `DCOUNT(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DCOUNTA: {
    formula: `DCOUNTA(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DGET: {
    formula: `DGET(A751:F756, D751, A763:A764)`,
    cellData: { ...DATABASE_CELL_DATA },
    note: "Uses unique criteria (Name=Alice) for single-row match",
  },
  DMAX: {
    formula: `DMAX(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DMIN: {
    formula: `DMIN(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DPRODUCT: {
    formula: `DPRODUCT(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DSTDEV: {
    formula: `DSTDEV(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DSTDEVP: {
    formula: `DSTDEVP(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DSUM: {
    formula: `DSUM(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DVAR: {
    formula: `DVAR(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },
  DVARP: {
    formula: `DVARP(A751:F756, D751, A759:A760)`,
    cellData: { ...DATABASE_CELL_DATA },
  },

  // ── Special cell functions (6) ────────────────────────────────────
  INDIRECT: {
    formula: `INDIRECT("A771")`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "INDIRECT resolves string ref to cell value",
  },
  OFFSET: {
    formula: `OFFSET(A771, 1, 0)`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "OFFSET(A771, 1, 0) returns A772 (empty)",
  },
  CELL: {
    formula: `CELL("contents", A771)`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "Returns info about a cell",
  },
  ISFORMULA: {
    formula: `ISFORMULA(B771)`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "B771 contains =1+1",
  },
  ISREF: {
    formula: `ISREF(A771)`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "ISREF checks if the argument is a valid reference",
  },
  FORMULATEXT: {
    formula: `FORMULATEXT(B771)`,
    cellData: { ...SPECIAL_CELL_DATA },
    note: "B771 contains =1+1",
  },

  // ── NETWORKDAYS / WORKDAY with cell-ref holidays (rows 781-782) ───
  // GSheets returns #VALUE! when the holidays parameter is an inline array.
  NETWORKDAYS: {
    formula: `NETWORKDAYS("7/16/1969", "7/24/1969", A781:A782)`,
    cellData: { ...HOLIDAY_CELL_DATA },
  },
  "NETWORKDAYS.INTL": {
    formula: `NETWORKDAYS.INTL("7/16/1969", "7/24/1969", 1, A781:A782)`,
    cellData: { ...HOLIDAY_CELL_DATA },
  },
  WORKDAY: {
    formula: `WORKDAY("7/16/1969", 4, A781:A782)`,
    cellData: { ...HOLIDAY_CELL_DATA },
  },
  "WORKDAY.INTL": {
    formula: `WORKDAY.INTL("7/16/1969", 4, 1, A781:A782)`,
    cellData: { ...HOLIDAY_CELL_DATA },
  },
};

// ---------------------------------------------------------------------------
// Main logic
// ---------------------------------------------------------------------------

function generate(): FormulaCompatTests {
  const result: FormulaCompatTests = {
    version: 1,
    functions: {},
  };

  const stats = {
    total: 0,
    selfContained: 0,
    rewritten: 0,
    volatile: 0,
    gSheetsOnly: 0,
    needsCellRefs: 0,
    warnings: [] as string[],
  };

  const functionNames = Object.keys(formulaFunctions) as string[];

  for (const name of functionNames) {
    const fn = (formulaFunctions as Record<string, any>)[name];
    stats.total++;

    const entry: FormulaTestEntry = {
      name,
      category: fn.category ?? "Unknown",
      tests: [],
    };

    // Mark volatile
    if (VOLATILE_FUNCTIONS.has(name)) {
      entry.volatile = true;
      stats.volatile++;
      // For volatile functions, keep the original example but note it
      const examples: string[] = fn.examples ?? [];
      for (const ex of examples) {
        entry.tests.push({
          formula: ex,
          note: "Volatile — value changes on each evaluation",
        });
      }
      result.functions[name] = entry;
      continue;
    }

    // Mark gSheetsOnly
    if (GSHEETS_ONLY.has(name)) {
      entry.gSheetsOnly = true;
      stats.gSheetsOnly++;
      const examples: string[] = fn.examples ?? [];
      for (const ex of examples) {
        entry.tests.push({
          formula: ex,
          note: "Google Sheets only — not available in HyperFormula",
        });
      }
      result.functions[name] = entry;
      continue;
    }

    // Check for cell-ref formula (real cell ranges at row 700+)
    const cellRefRewrite = CELL_REF_FORMULAS[name];
    if (cellRefRewrite) {
      stats.rewritten++;
      const testCase: FormulaTestCase = {
        formula: cellRefRewrite.formula,
      };
      if (cellRefRewrite.cellData) {
        testCase.cellData = cellRefRewrite.cellData;
      }
      if (cellRefRewrite.note) {
        testCase.note = cellRefRewrite.note;
      }
      entry.tests.push(testCase);
      result.functions[name] = entry;
      continue;
    }

    // Check for explicit rewrite (inline arrays)
    const rewrite = FORMULA_REWRITES[name];
    if (rewrite) {
      const isNeedsCellRefs = NEEDS_CELL_REFS.has(name);
      if (isNeedsCellRefs) {
        entry.needsCellRefs = true;
        stats.needsCellRefs++;
      } else {
        stats.rewritten++;
      }

      const testCase: FormulaTestCase = {
        formula: rewrite.formula,
      };
      if (rewrite.cellData) {
        testCase.cellData = rewrite.cellData;
      }
      if (rewrite.note) {
        testCase.note = rewrite.note;
      }
      entry.tests.push(testCase);
      result.functions[name] = entry;
      continue;
    }

    // No rewrite — check if original examples have cell refs
    const examples: string[] = fn.examples ?? [];
    if (examples.length === 0) {
      entry.tests.push({
        formula: `${name}()`,
        note: "No example available",
      });
      stats.selfContained++;
      result.functions[name] = entry;
      continue;
    }

    // Detect cell references in formula (but not inside function names or strings)
    const hasCellRef = exampleHasCellRef(examples[0], name);

    if (hasCellRef) {
      // This function has cell refs but no explicit rewrite — warn
      stats.warnings.push(name);
      entry.tests.push({
        formula: examples[0],
        note: "WARNING: Contains cell references but has no rewrite rule",
      });
    } else {
      stats.selfContained++;
      for (const ex of examples) {
        entry.tests.push({ formula: ex });
      }
    }

    result.functions[name] = entry;
  }

  // Print stats
  /* eslint-disable no-console */
  console.log("\n=== Formula Compat Tests Generation ===");
  console.log(`Total functions:   ${stats.total}`);
  console.log(`Self-contained:    ${stats.selfContained}`);
  console.log(`Rewritten:         ${stats.rewritten}`);
  console.log(`Volatile:          ${stats.volatile}`);
  console.log(`GSheets-only:      ${stats.gSheetsOnly}`);
  console.log(`Needs cell refs:   ${stats.needsCellRefs}`);
  console.log(`Warnings:          ${stats.warnings.length}`);
  if (stats.warnings.length > 0) {
    console.log(
      `\nFunctions with cell refs but no rewrite rule:\n  ${stats.warnings.join(", ")}`
    );
  }
  console.log("");
  /* eslint-enable no-console */

  return result;
}

/**
 * Detects whether a formula example contains cell references (like A1, B2:C10)
 * that are NOT part of a function name (like ATAN2, BIN2DEC, LOG10, etc.)
 * and NOT inside string literals.
 *
 * Cell ref pattern: one or more uppercase letters followed by one or more digits,
 * optionally followed by `:` and another cell ref (for ranges).
 */
function exampleHasCellRef(formula: string, _fnName: string): boolean {
  // Known function name substrings that look like cell refs
  const FALSE_POSITIVE_FUNCTIONS = new Set([
    "ATAN2",
    "BIN2DEC",
    "BIN2HEX",
    "BIN2OCT",
    "DEC2BIN",
    "DEC2HEX",
    "DEC2OCT",
    "HEX2BIN",
    "HEX2DEC",
    "HEX2OCT",
    "IMLOG10",
    "IMLOG2",
    "LOG10",
    "OCT2BIN",
    "OCT2DEC",
    "OCT2HEX",
    "DAYS360",
    "T2",
    "F2",
    "X2",
    "S2",
    "P2",
  ]);

  // Strip string literals first to avoid false positives inside quotes
  const stripped = formula.replace(/"[^"]*"/g, '""');

  // Match potential cell references: 1-3 uppercase letters + 1-5 digits
  const cellRefPattern = /\b([A-Z]{1,3})(\d{1,5})\b/g;
  let match;

  while ((match = cellRefPattern.exec(stripped)) !== null) {
    const fullMatch = match[0];
    const beforeIndex = match.index;

    // Check if this match is part of a function name by looking at surrounding context
    // A function name would have more letters before or be part of a known function
    const before = stripped.slice(Math.max(0, beforeIndex - 20), beforeIndex);

    // If preceded by more uppercase letters, it's part of a function name
    if (/[A-Z]$/.test(before)) {
      continue;
    }

    // Check against known false positive function name fragments
    let isFalsePositive = false;
    for (const fp of FALSE_POSITIVE_FUNCTIONS) {
      if (fp.includes(fullMatch)) {
        // Check if the function name appears in the formula
        if (stripped.includes(fp)) {
          isFalsePositive = true;
          break;
        }
      }
    }
    if (isFalsePositive) {
      continue;
    }

    // This looks like a genuine cell reference
    return true;
  }

  return false;
}

// ---------------------------------------------------------------------------
// Run
// ---------------------------------------------------------------------------

const output = generate();

const outputPath = resolve(
  __dirname,
  '../test/gsheets-compat/__fixtures__/formula-compat-tests.json'
);
mkdirSync(dirname(outputPath), { recursive: true });
writeFileSync(outputPath, JSON.stringify(output, null, 2) + "\n");

/* eslint-disable no-console */
console.log(`Written to: ${outputPath}`);
/* eslint-enable no-console */
