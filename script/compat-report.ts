#!/usr/bin/env ts-node
/**
 * Compatibility Report — HyperFormula vs Google Sheets
 *
 * Reads formula-compat-tests.json and evaluates all non-volatile, non-GSheets-only
 * formulas through HyperFormula (compatibilityMode: 'googleSheets'), then prints:
 *   - Overall compatibility percentage
 *   - Per-category breakdown (match / mismatch / error / skip)
 *   - List of untested functions (no GSheets reference value)
 *
 * Usage:
 *   npx ts-node --project tsconfig.test.json script/compat-report.ts
 *
 * Requires formula-compat-tests.json to have expectedValue fields populated.
 * If not, run: npx ts-node --project tsconfig.test.json script/patch-expected-values.ts
 */

import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { HyperFormula } from '../src'
import { VOLATILE_FUNCTIONS, GSHEETS_ONLY_FUNCTIONS } from '../test/gsheets-compat/gsheets-compat-config'
import {
  buildSheetData,
  cellValueToString,
  getFormulaAddress,
  valuesMatch,
} from '../test/gsheets-compat/helpers'
import type { FormulaCompatTests, TestResult } from '../test/gsheets-compat/helpers'

const JSON_PATH = resolve(__dirname, '../test/gsheets-compat/__fixtures__/formula-compat-tests.json')

if (!existsSync(JSON_PATH)) {
  console.error(
    `Missing: ${JSON_PATH}\n` +
    'Run: npx ts-node --project tsconfig.test.json script/generate-formula-compat-tests.ts'
  )
  process.exit(1)
}

const data: FormulaCompatTests = JSON.parse(readFileSync(JSON_PATH, 'utf-8'))

const results: TestResult[] = []
const untestedFunctions: string[] = []

for (const entry of Object.values(data.functions)) {
  const isSkipped =
    entry.volatile ||
    entry.gSheetsOnly ||
    VOLATILE_FUNCTIONS.has(entry.name) ||
    GSHEETS_ONLY_FUNCTIONS.has(entry.name)

  const allTestsHaveNoExpected = entry.tests.every(
    (t) => t.expectedValue === undefined || t.expectedValue === null
  )

  if (!isSkipped && allTestsHaveNoExpected) {
    untestedFunctions.push(entry.name)
  }

  for (let i = 0; i < entry.tests.length; i++) {
    const test = entry.tests[i]
    if (!test) continue

    const suffix = i === 0 ? '' : ` #${i + 1}`
    const displayName = `${entry.name}${suffix}`

    if (isSkipped) {
      results.push({
        name: displayName,
        category: entry.category,
        formula: test.formula,
        ourValue: '',
        refValue: '',
        status: 'skip',
      })
      continue
    }

    const refValue =
      test.expectedValue !== undefined && test.expectedValue !== null
        ? String(test.expectedValue)
        : ''

    let ourValue: string
    let status: TestResult['status']
    let errorMessage: string | undefined

    try {
      const sheetData = buildSheetData(test.formula, test.cellData)
      const formulaAddr = getFormulaAddress(test.cellData)

      // Parse A1-notation address into {sheet, row, col}
      const addrMatch = /^([A-Z]+)(\d+)$/i.exec(formulaAddr)
      if (!addrMatch || !addrMatch[1] || !addrMatch[2]) {
        throw new Error(`Invalid formula address: ${formulaAddr}`)
      }
      const addrLetters = addrMatch[1].toUpperCase()
      let addrCol = 0
      for (let j = 0; j < addrLetters.length; j++) {
        addrCol = addrCol * 26 + (addrLetters.charCodeAt(j) - 64)
      }
      addrCol -= 1
      const addrRow = parseInt(addrMatch[2], 10) - 1

      const hf = HyperFormula.buildFromArray(sheetData, {
        licenseKey: 'gpl-v3',
        compatibilityMode: 'googleSheets',
      })

      try {
        const cellValue = hf.getCellValue({ sheet: 0, row: addrRow, col: addrCol })
        ourValue = cellValueToString(cellValue)
      } finally {
        hf.destroy()
      }

      if (refValue === '') {
        status = 'skip'
      } else if (valuesMatch(ourValue, refValue)) {
        status = 'match'
      } else {
        status = 'mismatch'
      }
    } catch (err) {
      ourValue = ''
      status = 'error'
      errorMessage = err instanceof Error ? err.message : String(err)
    }

    results.push({
      name: displayName,
      category: entry.category,
      formula: test.formula,
      ourValue,
      refValue,
      status,
      errorMessage,
    })
  }
}

// ─── Print report ─────────────────────────────────────────────────────────────

type CategoryStats = {
  total: number
  match: number
  mismatch: number
  error: number
  skip: number
  mismatches: TestResult[]
}

const categories = new Map<string, CategoryStats>()

for (const r of results) {
  let cat = categories.get(r.category)
  if (!cat) {
    cat = { total: 0, match: 0, mismatch: 0, error: 0, skip: 0, mismatches: [] }
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

let totalAll = 0, matchAll = 0, mismatchAll = 0, errorAll = 0, skipAll = 0

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

// Untested functions
lines.push('')
if (untestedFunctions.length === 0) {
  lines.push('Untested functions: none — all functions have GSheets reference values.')
} else {
  lines.push(`Untested functions (${untestedFunctions.length}) — no GSheets reference value:`)
  for (const fn of untestedFunctions) {
    lines.push(`  - ${fn}`)
  }
  lines.push('')
  lines.push('To populate reference values:')
  lines.push('  1. npx ts-node --project tsconfig.test.json script/generate-formula-test-csv.ts')
  lines.push('  2. Import formula-compat-input.csv into Google Sheets and evaluate')
  lines.push('  3. Export as CSV → test/gsheets-compat/__fixtures__/formula-compat-gsheets.csv')
  lines.push('  4. npx ts-node --project tsconfig.test.json script/patch-expected-values.ts')
}

console.log(lines.join('\n'))
