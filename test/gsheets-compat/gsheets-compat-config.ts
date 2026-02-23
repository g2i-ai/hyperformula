/**
 * Configuration for GSheets compatibility testing.
 *
 * MUST_MATCH_FUNCTIONS: Functions that are expected to match GSheets output.
 * If any of these produce a mismatch, the test fails.
 * Add functions here as compatibility fixes are verified.
 */
export const MUST_MATCH_FUNCTIONS: string[] = [
  // Start empty — add functions as fixes are verified
  // e.g., 'ABS', 'SUM', 'IF', ...
]

/** Functions whose results change each evaluation. Cannot be compared. */
export const VOLATILE_FUNCTIONS = new Set([
  'NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY',
])

/** Functions only available in Google Sheets, not in HyperFormula. */
export const GSHEETS_ONLY_FUNCTIONS = new Set([
  'GOOGLEFINANCE', 'GOOGLETRANSLATE', 'IMPORTDATA', 'IMPORTFEED',
  'IMPORTHTML', 'IMPORTRANGE', 'IMPORTXML', 'IMAGE', 'SPARKLINE',
  'QUERY', 'GETPIVOTDATA', 'DETECTLANGUAGE',
])
