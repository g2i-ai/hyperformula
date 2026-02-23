import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Google Sheets mixed range syntax', () => {
  const gsConfig = {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' as const}

  describe('CellRef:Column (e.g., A1:B)', () => {
    it('=SUM(A1:B) sums from A1 to B{maxRow}', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, '=SUM(A1:B)'],
      ], gsConfig)
      // A1:B means A1 to B{maxRow}; only A1=1, B1=2 have values
      expect(hf.getCellValue(adr('C1'))).toBe(3)
      hf.destroy()
    })

    it('=SUM(A1:C) includes column C', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, 3, '=SUM(A1:C)'],
      ], gsConfig)
      expect(hf.getCellValue(adr('D1'))).toBe(6)
      hf.destroy()
    })

    it('handles absolute start reference $A$1:B', () => {
      const hf = HyperFormula.buildFromArray([
        [10, 20, '=SUM($A$1:B)'],
      ], gsConfig)
      expect(hf.getCellValue(adr('C1'))).toBe(30)
      hf.destroy()
    })
  })

  describe('CellRef:Row (e.g., A1:2)', () => {
    it('=SUM(A1:2) sums from A1 to {maxCol}2', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, 3],
        [4, 5, 6],
        [null, null, null, '=SUM(A1:2)'],
      ], gsConfig)
      // A1:2 = A1 to {maxCol}2. Contains A1..C1 and A2..C2 = 1+2+3+4+5+6 = 21
      expect(hf.getCellValue(adr('D3'))).toBe(21)
      hf.destroy()
    })
  })

  describe('Row:CellRef (e.g., 1:B2)', () => {
    it('=SUM(1:B2) sums from A1 to B2', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, '=SUM(1:B2)'],
        [3, 4],
      ], gsConfig)
      // 1:B2 = A1:B2 → 1+2+3+4 = 10
      expect(hf.getCellValue(adr('C1'))).toBe(10)
      hf.destroy()
    })
  })

  describe('Column:CellRef via ColumnRange+Number (e.g., A:B2)', () => {
    it('=SUM(A:B2) sums from A1 to B2', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, '=SUM(A:B2)'],
        [3, 4],
      ], gsConfig)
      // A:B2: lexed as ColumnRange(A:B) NumberLiteral(2) → A1:B2 → 1+2+3+4 = 10
      expect(hf.getCellValue(adr('C1'))).toBe(10)
      hf.destroy()
    })
  })

  describe('standard ranges still work', () => {
    it('=SUM(A1:B2) in both modes', () => {
      for (const mode of ['default', 'googleSheets'] as const) {
        const hf = HyperFormula.buildFromArray([
          [1, 2],
          [3, 4, `=SUM(A1:B2)`],
        ], {licenseKey: 'gpl-v3', compatibilityMode: mode})
        expect(hf.getCellValue(adr('C2'))).toBe(10)
        hf.destroy()
      }
    })

    it('=SUM(A:B) column range in both modes', () => {
      for (const mode of ['default', 'googleSheets'] as const) {
        const hf = HyperFormula.buildFromArray([
          [1, 2],
          [3, 4, `=SUM(A:B)`],
        ], {licenseKey: 'gpl-v3', compatibilityMode: mode})
        expect(hf.getCellValue(adr('C2'))).toBe(10)
        hf.destroy()
      }
    })

    it('=SUM(1:2) row range in both modes', () => {
      for (const mode of ['default', 'googleSheets'] as const) {
        const hf = HyperFormula.buildFromArray([
          [1, 2],
          [3, 4],
          [null, null, `=SUM(1:2)`],
        ], {licenseKey: 'gpl-v3', compatibilityMode: mode})
        expect(hf.getCellValue(adr('C3'))).toBe(10)
        hf.destroy()
      }
    })
  })

  describe('named expressions still work', () => {
    it('=myVar resolves as named expression in GSheets mode', () => {
      const hf = HyperFormula.buildFromArray([
        ['=myVar'],
      ], gsConfig, [{name: 'myVar', expression: '=42'}])
      expect(hf.getCellValue(adr('A1'))).toBe(42)
      hf.destroy()
    })
  })

  describe('mixed ranges in arithmetic', () => {
    it('=SUM(A1:B) + 1 works', () => {
      const hf = HyperFormula.buildFromArray([
        [1, 2, '=SUM(A1:B) + 1'],
      ], gsConfig)
      expect(hf.getCellValue(adr('C1'))).toBe(4)
      hf.destroy()
    })
  })
})
