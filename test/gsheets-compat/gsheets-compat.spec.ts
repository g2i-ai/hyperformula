/**
 * GSheets Compatibility Baseline Test
 *
 * Evaluates all formula functions through HyperFormula with
 * compatibilityMode: 'googleSheets', compares results against
 * a Google Sheets reference export, and prints a compatibility table.
 *
 * Functions in MUST_MATCH_FUNCTIONS must match or the test fails.
 * All other functions are tracked for reporting only.
 */

import {existsSync, readFileSync} from 'fs'
import {resolve} from 'path'
import {HyperFormula} from '../../src'
import {adr} from '../testUtils'
import {MUST_MATCH_FUNCTIONS, VOLATILE_FUNCTIONS, GSHEETS_ONLY_FUNCTIONS} from './gsheets-compat-config'
import {
  buildSheetData,
  cellValueToString,
  getFormulaAddress,
  printReport,
  valuesMatch,
} from './helpers'
import type {FormulaCompatTests, TestResult} from './helpers'

const JSON_PATH = resolve(__dirname, '__fixtures__/formula-compat-tests.json')

describe('GSheets Compatibility Baseline', () => {
  jest.setTimeout(60000)

  it('evaluates all formulas and compares against GSheets reference', () => {
    if (!existsSync(JSON_PATH)) {
      // eslint-disable-next-line no-console
      console.log(
        '\n  Skipping: formula-compat-tests.json not found.\n' +
        '  Run: npx ts-node script/generate-formula-compat-tests.ts\n'
      )
      return
    }

    const raw = readFileSync(JSON_PATH, 'utf-8')
    const data: FormulaCompatTests = JSON.parse(raw)

    const hasAnyExpected = Object.values(data.functions).some((entry) =>
      entry.tests.some((t) => t.expectedValue !== undefined && t.expectedValue !== null)
    )

    if (!hasAnyExpected) {
      // eslint-disable-next-line no-console
      console.log(
        '\n  No expected values found in formula-compat-tests.json.\n' +
        '  To populate:\n' +
        '  1. Run: npx ts-node script/generate-formula-test-csv.ts\n' +
        '  2. Import formula-compat-input.csv into Google Sheets\n' +
        '  3. Export as CSV -> formula-compat-gsheets.csv\n' +
        '  4. Run: npx ts-node script/patch-expected-values.ts\n'
      )
    }

    const results: TestResult[] = []
    const mustMatchFailures: TestResult[] = []

    for (const entry of Object.values(data.functions)) {
      for (let i = 0; i < entry.tests.length; i++) {
        const test = entry.tests[i]
        if (!test) continue

        const suffix = i === 0 ? '' : ` #${i + 1}`
        const displayName = `${entry.name}${suffix}`

        // Skip volatile and GSheets-only functions (check both entry flags and config sets)
        if (entry.volatile || entry.gSheetsOnly || VOLATILE_FUNCTIONS.has(entry.name) || GSHEETS_ONLY_FUNCTIONS.has(entry.name)) {
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

        const refValue = test.expectedValue !== undefined && test.expectedValue !== null
          ? String(test.expectedValue)
          : ''

        let ourValue: string
        let status: TestResult['status']
        let errorMessage: string | undefined

        try {
          // Build sheet data with cell data (if any) and formula
          const sheetData = buildSheetData(test.formula, test.cellData)
          const formulaAddr = getFormulaAddress(test.cellData)

          const hf = HyperFormula.buildFromArray(sheetData, {
            licenseKey: 'gpl-v3',
            compatibilityMode: 'googleSheets',
          })

          const cellValue = hf.getCellValue(adr(formulaAddr))
          ourValue = cellValueToString(cellValue)
          hf.destroy()

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

        const result: TestResult = {
          name: displayName,
          category: entry.category,
          formula: test.formula,
          ourValue,
          refValue,
          status,
          errorMessage,
        }
        results.push(result)

        // Track allowlist failures
        if (MUST_MATCH_FUNCTIONS.includes(entry.name) && status !== 'match' && status !== 'skip') {
          mustMatchFailures.push(result)
        }
      }
    }

    printReport(results)

    // Gate: fail if any allowlisted function mismatches
    if (mustMatchFailures.length > 0) {
      const details = mustMatchFailures
        .map((r) => `  ${r.name}: ${r.formula} -- ours: ${r.ourValue}, ref: ${r.refValue}`)
        .join('\n')
      fail(
        `${mustMatchFailures.length} MUST_MATCH function(s) failed:\n${details}`
      )
    }
  })
})
