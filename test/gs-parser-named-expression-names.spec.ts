import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Google Sheets named expression naming rules', () => {
  describe('isNameValid relaxation', () => {
    it('accepts long column-like names in GSheets mode (e.g., ProductPrice1)', () => {
      // ProductPrice1 looks like a cell reference, but "ProductPrice" is way beyond valid column range
      const hf = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      // This should not throw
      hf.addNamedExpression('ProductPrice1', '=42')
      expect(hf.getNamedExpressionValue('ProductPrice1')).toBe(42)
      hf.destroy()
    })

    it('rejects long column-like names in default mode (e.g., ProductPrice1)', () => {
      const hf = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'default',
      })
      expect(() => hf.addNamedExpression('ProductPrice1', '=42')).toThrow()
      hf.destroy()
    })

    it('rejects short valid cell references in GSheets mode (e.g., AB1)', () => {
      const hf = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      // AB is column 27 — well within valid range
      expect(() => hf.addNamedExpression('AB1', '=42')).toThrow()
      hf.destroy()
    })

    it('accepts 5-letter column names in GSheets mode (e.g., ABCDE1)', () => {
      const hf = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      // ABCDE is column 494,264 — way beyond default maxColumns (18278)
      hf.addNamedExpression('ABCDE1', '=100')
      expect(hf.getNamedExpressionValue('ABCDE1')).toBe(100)
      hf.destroy()
    })

    it('rejects R1C1-pattern names in both modes', () => {
      const hfGS = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(() => hfGS.addNamedExpression('R1C1', '=42')).toThrow()
      hfGS.destroy()

      const hfDefault = HyperFormula.buildEmpty({
        licenseKey: 'gpl-v3', compatibilityMode: 'default',
      })
      expect(() => hfDefault.addNamedExpression('R1C1', '=42')).toThrow()
      hfDefault.destroy()
    })
  })

  describe('formula evaluation with long-name expressions', () => {
    it('resolves a long-column-like named expression in formulas', () => {
      const hf = HyperFormula.buildFromArray([
        ['=TotalRevenue2024'],
      ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' },
        [{ name: 'TotalRevenue2024', expression: '=42' }])

      expect(hf.getCellValue(adr('A1'))).toBe(42)
      hf.destroy()
    })

    it('does not affect standard cell references', () => {
      const hf = HyperFormula.buildFromArray([
        [10, '=A1'],
      ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

      expect(hf.getCellValue(adr('B1'))).toBe(10)
      hf.destroy()
    })
  })
})
