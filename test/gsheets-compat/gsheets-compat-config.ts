/**
 * Configuration for GSheets compatibility testing.
 *
 * MUST_MATCH_FUNCTIONS: Functions that are expected to match GSheets output.
 * If any of these produce a mismatch, the test fails.
 * Add functions here as compatibility fixes are verified.
 */
export const MUST_MATCH_FUNCTIONS: string[] = [
  // Math (93.8% match rate)
  'ABS', 'ACOS', 'ACOSH', 'ASIN', 'ASINH', 'ATAN', 'ATAN2', 'ATANH',
  'CEILING', 'COS', 'COSH', 'COT', 'COTH', 'DEGREES', 'EVEN', 'EXP',
  'FACT', 'FACTDOUBLE', 'FLOOR', 'INT', 'LN', 'LOG', 'LOG10',
  'MOD', 'MROUND', 'ODD', 'PI', 'POWER', 'QUOTIENT', 'RADIANS',
  'ROUND', 'ROUNDDOWN', 'ROUNDUP', 'SIGN', 'SIN', 'SINH',
  'SQRT', 'SQRTPI', 'TAN', 'TANH', 'TRUNC',
  'SUM', 'PRODUCT', 'SUMSQ', 'SUMPRODUCT',
  'MAX', 'MIN', 'AVERAGE', 'MEDIAN',
  'GCD', 'LCM', 'COMBIN', 'COMBINA',
  'MULTINOMIAL', 'SERIESSUM',
  // Logical
  'IF', 'NOT', 'TRUE', 'FALSE', 'IFERROR', 'IFNA',
  // Text
  'LEN', 'UPPER', 'LOWER', 'TRIM', 'LEFT', 'RIGHT', 'MID',
  'REPT', 'SUBSTITUTE', 'REPLACE', 'EXACT', 'FIND', 'SEARCH',
  'CONCATENATE', 'T', 'VALUE',
  // Engineering
  'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'DELTA',
  'OCT2BIN', 'OCT2DEC', 'OCT2HEX',
  // Statistical
  'COUNT', 'COUNTA',
  'LARGE', 'SMALL',
  'STDEV', 'VAR', 'GEOMEAN', 'HARMEAN',
  'CORREL', 'COVARIANCE.P', 'COVARIANCE.S',
  'NORM.DIST', 'NORM.INV', 'NORM.S.INV',
  // Lookup
  'MATCH', 'INDEX',
  // Date
  'YEAR', 'MONTH', 'HOUR', 'MINUTE', 'SECOND',
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
