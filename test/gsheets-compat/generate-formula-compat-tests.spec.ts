/**
 * Tests for generate-formula-compat-tests script output.
 *
 * Verifies that functions known to fail in GSheets with inline arrays
 * are configured to use real cell ranges instead.
 */

import {existsSync, readFileSync} from 'fs'
import {resolve} from 'path'

const JSON_PATH = resolve(__dirname, '__fixtures__/formula-compat-tests.json')

describe('generate-formula-compat-tests output', () => {
  let data: {
    version: number
    functions: Record<string, {
      name: string
      category: string
      tests: Array<{
        formula: string
        cellData?: Record<string, string | number | boolean | null>
        note?: string
      }>
    }>
  }

  beforeAll(() => {
    if (!existsSync(JSON_PATH)) {
      throw new Error(
        'formula-compat-tests.json not found. Run: npx ts-node script/generate-formula-compat-tests.ts'
      )
    }
    data = JSON.parse(readFileSync(JSON_PATH, 'utf-8'))
  })

  describe('AVERAGEIFS', () => {
    it('uses cell references instead of inline arrays (GSheets returns #N/A with inline arrays)', () => {
      const entry = data.functions['AVERAGEIFS']
      expect(entry).toBeDefined()
      const test = entry!.tests[0]!
      // Must have cellData (cell-ref based), not raw inline arrays
      expect(test.cellData).toBeDefined()
      expect(Object.keys(test.cellData!).length).toBeGreaterThan(0)
    })

    it('formula references cell ranges, not inline arrays', () => {
      const entry = data.functions['AVERAGEIFS']
      const test = entry!.tests[0]!
      // Formula must NOT contain inline array syntax like {5;10;...}
      expect(test.formula).not.toMatch(/\{[^}]+;[^}]+\}/)
      // Formula must contain a cell range like A701:A710
      expect(test.formula).toMatch(/[A-Z]\d+:[A-Z]\d+/)
    })
  })

  describe('NETWORKDAYS', () => {
    it('uses cell references for holidays parameter (GSheets returns #VALUE! with inline arrays)', () => {
      const entry = data.functions['NETWORKDAYS']
      expect(entry).toBeDefined()
      const test = entry!.tests[0]!
      expect(test.cellData).toBeDefined()
      expect(Object.keys(test.cellData!).length).toBeGreaterThan(0)
    })

    it('holidays parameter uses a cell range, not an inline array', () => {
      const entry = data.functions['NETWORKDAYS']
      const test = entry!.tests[0]!
      // Formula must NOT contain inline array for holidays
      expect(test.formula).not.toMatch(/\{[^}]*"[^"]+\/[^"]+\/\d{4}"[^}]*\}/)
      // Formula must reference a cell range for holidays
      expect(test.formula).toMatch(/[A-Z]\d+:[A-Z]\d+/)
    })
  })

  describe('NETWORKDAYS.INTL', () => {
    it('uses cell references for holidays parameter', () => {
      const entry = data.functions['NETWORKDAYS.INTL']
      expect(entry).toBeDefined()
      const test = entry!.tests[0]!
      expect(test.cellData).toBeDefined()
      expect(Object.keys(test.cellData!).length).toBeGreaterThan(0)
    })

    it('holidays parameter uses a cell range, not an inline array', () => {
      const entry = data.functions['NETWORKDAYS.INTL']
      const test = entry!.tests[0]!
      expect(test.formula).not.toMatch(/\{[^}]*"[^"]+\/[^"]+\/\d{4}"[^}]*\}/)
      expect(test.formula).toMatch(/[A-Z]\d+:[A-Z]\d+/)
    })
  })

  describe('WORKDAY', () => {
    it('uses cell references for holidays parameter (GSheets returns #VALUE! with inline arrays)', () => {
      const entry = data.functions['WORKDAY']
      expect(entry).toBeDefined()
      const test = entry!.tests[0]!
      expect(test.cellData).toBeDefined()
      expect(Object.keys(test.cellData!).length).toBeGreaterThan(0)
    })

    it('holidays parameter uses a cell range, not an inline array', () => {
      const entry = data.functions['WORKDAY']
      const test = entry!.tests[0]!
      expect(test.formula).not.toMatch(/\{[^}]*"[^"]+\/[^"]+\/\d{4}"[^}]*\}/)
      expect(test.formula).toMatch(/[A-Z]\d+:[A-Z]\d+/)
    })
  })

  describe('WORKDAY.INTL', () => {
    it('uses cell references for holidays parameter', () => {
      const entry = data.functions['WORKDAY.INTL']
      expect(entry).toBeDefined()
      const test = entry!.tests[0]!
      expect(test.cellData).toBeDefined()
      expect(Object.keys(test.cellData!).length).toBeGreaterThan(0)
    })

    it('holidays parameter uses a cell range, not an inline array', () => {
      const entry = data.functions['WORKDAY.INTL']
      const test = entry!.tests[0]!
      expect(test.formula).not.toMatch(/\{[^}]*"[^"]+\/[^"]+\/\d{4}"[^}]*\}/)
      expect(test.formula).toMatch(/[A-Z]\d+:[A-Z]\d+/)
    })
  })
})
