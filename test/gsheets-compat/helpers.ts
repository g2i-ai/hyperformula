import {DetailedCellError} from '../../src'

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface FormulaTestCase {
  formula: string
  cellData?: Record<string, string | number | boolean | null>
  expectedValue?: string | number | boolean | null
  note?: string
}

export interface FormulaTestEntry {
  name: string
  category: string
  tests: FormulaTestCase[]
  volatile?: boolean
  gSheetsOnly?: boolean
  needsCellRefs?: boolean
}

export interface FormulaCompatTests {
  version: number
  functions: Record<string, FormulaTestEntry>
}

export interface TestResult {
  name: string
  category: string
  formula: string
  ourValue: string
  refValue: string
  status: 'match' | 'mismatch' | 'error' | 'skip'
  errorMessage?: string
}

// ---------------------------------------------------------------------------
// Cell Data Helpers
// ---------------------------------------------------------------------------

/**
 * Parses an A1-notation reference into [col, row] (0-indexed).
 * E.g., "A1" -> [0, 0], "C2" -> [2, 1]
 */
export function parseA1(ref: string): [number, number] {
  const match = /^([A-Z]+)(\d+)$/i.exec(ref)
  if (!match || !match[1] || !match[2]) {
    throw new Error(`Invalid A1 reference: ${ref}`)
  }
  const letters = match[1].toUpperCase()
  let col = 0
  for (let i = 0; i < letters.length; i++) {
    col = col * 26 + (letters.charCodeAt(i) - 64)
  }
  col -= 1 // 0-indexed
  const row = parseInt(match[2], 10) - 1
  return [col, row]
}

/**
 * Builds a 2D sheet array from a cellData map.
 * When no cellData, the formula goes in A1.
 * When cellData is present, the formula goes two rows after the last data row.
 * Cell data is placed at the referenced positions.
 */
export function buildSheetData(
  formula: string,
  cellData?: Record<string, string | number | boolean | null>,
): (string | number | boolean | null)[][] {
  const cells: Array<{col: number, row: number, value: string | number | boolean | null}> = []

  if (cellData) {
    for (const [ref, value] of Object.entries(cellData)) {
      const [col, row] = parseA1(ref)
      cells.push({col, row, value})
    }
  }

  // No cell data: simple case, formula in A1
  if (cells.length === 0) {
    return [['=' + formula]]
  }

  // Determine dimensions needed
  let maxRow = 0
  let maxCol = 0
  for (const {col, row} of cells) {
    if (row > maxRow) maxRow = row
    if (col > maxCol) maxCol = col
  }

  // Formula goes at a row after all data (safe position)
  const formulaRow = maxRow + 2
  maxRow = formulaRow

  // Build the 2D array
  const sheet: (string | number | boolean | null)[][] = []
  for (let r = 0; r <= maxRow; r++) {
    const row: (string | number | boolean | null)[] = []
    for (let c = 0; c <= maxCol; c++) {
      row.push(null)
    }
    sheet.push(row)
  }

  // Place cell data
  for (const {col, row, value} of cells) {
    while (sheet[row]!.length <= col) {
      sheet[row]!.push(null)
    }
    sheet[row]![col] = value
  }

  // Place formula
  while (sheet[formulaRow]!.length <= 0) {
    sheet[formulaRow]!.push(null)
  }
  sheet[formulaRow]![0] = '=' + formula

  return sheet
}

/**
 * Returns the address string for the formula cell given the sheet data.
 * The formula is always placed at maxRow + 2 (0-indexed), which is maxRow + 3 in A1 notation.
 */
export function getFormulaAddress(
  cellData?: Record<string, string | number | boolean | null>,
): string {
  if (!cellData || Object.keys(cellData).length === 0) {
    return 'A1'
  }
  let maxRow = 0
  for (const ref of Object.keys(cellData)) {
    const [, row] = parseA1(ref)
    if (row > maxRow) maxRow = row
  }
  return `A${maxRow + 3}` // formulaRow = maxRow + 2, so A1-notation = maxRow + 3
}

// ---------------------------------------------------------------------------
// Value Normalization
// ---------------------------------------------------------------------------

/**
 * Convert a HyperFormula cell value to a comparable string.
 * Uses .value for DetailedCellError which gives the standard error string
 * (e.g., "#DIV/0!", "#N/A", "#VALUE!") matching Google Sheets format.
 */
export function cellValueToString(value: unknown): string {
  if (value instanceof DetailedCellError) {
    return value.value
  }
  if (value === null || value === undefined) return ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number') return String(value)
  return String(value)
}

/** Normalize a value for comparison */
function normalize(value: string): string {
  return value.trim()
}

/** Compare two values with appropriate tolerance */
export function valuesMatch(ours: string, ref: string): boolean {
  const a = normalize(ours)
  const b = normalize(ref)

  // Exact match
  if (a === b) return true

  // Case-insensitive match for booleans only
  const BOOL_VALS = new Set(['TRUE', 'FALSE'])
  if (BOOL_VALS.has(a.toUpperCase()) && a.toUpperCase() === b.toUpperCase()) return true

  // Numeric comparison with relative tolerance.
  // Use Number() (not parseFloat) to avoid false positives: parseFloat("4-9i")
  // and parseFloat("7/20/1969") extract leading digits, making different values match.
  // Guard against empty strings: Number("") === 0, which would falsely match "0".
  const numA = a === '' ? NaN : Number(a)
  const numB = b === '' ? NaN : Number(b)
  if (!isNaN(numA) && !isNaN(numB)) {
    if (numA === 0 && numB === 0) return true
    const denom = Math.max(Math.abs(numA), Math.abs(numB))
    if (denom === 0) return true
    return Math.abs(numA - numB) / denom < 1e-6
  }

  return false
}

// ---------------------------------------------------------------------------
// Report Formatting
// ---------------------------------------------------------------------------

export function printReport(results: TestResult[]): void {
  const categories = new Map<string, {
    total: number
    match: number
    mismatch: number
    error: number
    skip: number
    mismatches: TestResult[]
  }>()

  for (const r of results) {
    let cat = categories.get(r.category)
    if (!cat) {
      cat = {total: 0, match: 0, mismatch: 0, error: 0, skip: 0, mismatches: []}
      categories.set(r.category, cat)
    }
    cat.total++
    if (r.status === 'match') cat.match++
    else if (r.status === 'skip') cat.skip++
    else {
      if (r.status === 'mismatch') cat.mismatch++
      else cat.error++
      cat.mismatches.push(r)
    }
  }

  const sorted = [...categories.entries()].sort((a, b) => a[0].localeCompare(b[0]))

  let totalAll = 0
  let matchAll = 0
  let mismatchAll = 0
  let errorAll = 0
  let skipAll = 0

  const lines: string[] = []
  lines.push('')
  lines.push('Formula Compatibility Report (HyperFormula vs Google Sheets)')
  lines.push('='.repeat(75))
  lines.push(
    `${'Category'.padEnd(15)}| ${'Total'.padStart(5)} | ${'Match'.padStart(5)} | ${'Mismatch'.padStart(8)} | ${'Error'.padStart(5)} | ${'Skip'.padStart(5)} | ${'Match%'.padStart(7)}`
  )
  lines.push(
    '-'.repeat(15) + '+' + '-'.repeat(7) + '+' + '-'.repeat(7) + '+' +
    '-'.repeat(10) + '+' + '-'.repeat(7) + '+' + '-'.repeat(7) + '+' + '-'.repeat(9)
  )

  for (const [name, cat] of sorted) {
    totalAll += cat.total
    matchAll += cat.match
    mismatchAll += cat.mismatch
    errorAll += cat.error
    skipAll += cat.skip
    const comparable = cat.total - cat.skip
    const pct = comparable > 0 ? ((cat.match / comparable) * 100).toFixed(1) : 'N/A'
    lines.push(
      `${name.padEnd(15)}| ${String(cat.total).padStart(5)} | ${String(cat.match).padStart(5)} | ${String(cat.mismatch).padStart(8)} | ${String(cat.error).padStart(5)} | ${String(cat.skip).padStart(5)} | ${(pct === 'N/A' ? pct : pct + '%').padStart(7)}`
    )
  }

  lines.push('='.repeat(75))
  const comparableAll = totalAll - skipAll
  const totalPct = comparableAll > 0 ? ((matchAll / comparableAll) * 100).toFixed(1) : 'N/A'
  lines.push(
    `${'TOTAL'.padEnd(15)}| ${String(totalAll).padStart(5)} | ${String(matchAll).padStart(5)} | ${String(mismatchAll).padStart(8)} | ${String(errorAll).padStart(5)} | ${String(skipAll).padStart(5)} | ${(totalPct === 'N/A' ? totalPct : totalPct + '%').padStart(7)}`
  )

  // Mismatch details
  lines.push('')
  lines.push('Mismatches (first 5 per category):')
  for (const [name, cat] of sorted) {
    const details = cat.mismatches.slice(0, 5)
    if (details.length === 0) continue
    lines.push(`  ${name}:`)
    for (const d of details) {
      if (d.status === 'error') {
        lines.push(`    ${d.formula} -- ENGINE ERROR: ${d.errorMessage ?? 'unknown'}`)
      } else {
        lines.push(`    ${d.formula} -- ours: ${d.ourValue}, ref: ${d.refValue}`)
      }
    }
  }

  // eslint-disable-next-line no-console
  console.log(lines.join('\n'))
}
