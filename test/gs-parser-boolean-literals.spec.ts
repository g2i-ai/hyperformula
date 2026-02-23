import {HyperFormula} from '../src'
import {DetailedCellError} from '../src/CellValue'
import {deDE} from '../src/i18n/languages'
import {adr} from './testUtils'

describe('Google Sheets boolean literals (parser-level)', () => {
  afterAll(() => {
    try { HyperFormula.unregisterLanguage('deDE') } catch { /* ignore */ }
  })

  describe('basic TRUE/FALSE evaluation', () => {
    it('=TRUE evaluates to true in GSheets mode', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('=FALSE evaluates to false in GSheets mode', () => {
      const hf = HyperFormula.buildFromArray([['=FALSE']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('=true evaluates to true (case insensitive)', () => {
      const hf = HyperFormula.buildFromArray([['=true']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('=True evaluates to true (mixed case)', () => {
      const hf = HyperFormula.buildFromArray([['=True']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })

  describe('function form still works', () => {
    it('=TRUE() evaluates to true in GSheets mode', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE()']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('=FALSE() evaluates to false in GSheets mode', () => {
      const hf = HyperFormula.buildFromArray([['=FALSE()']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('boolean in expressions', () => {
    it('=IF(TRUE, 1, 2) evaluates to 1', () => {
      const hf = HyperFormula.buildFromArray([['=IF(TRUE, 1, 2)']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('=TRUE + 1 evaluates to 2 (numeric coercion)', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE + 1']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('=TRUE * FALSE evaluates to 0', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE * FALSE']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('=NOT(TRUE) evaluates to false', () => {
      const hf = HyperFormula.buildFromArray([['=NOT(TRUE)']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('default mode behavior', () => {
    it('=TRUE produces NAME error in default mode', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'default',
      })
      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('=TRUE() still works in default mode (function call)', () => {
      const hf = HyperFormula.buildFromArray([['=TRUE()']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'default',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })

  describe('i18n support', () => {
    it('=TRUE resolves through functionMapping for non-English locale', () => {
      HyperFormula.registerLanguage('deDE', deDE)
      const hf = HyperFormula.buildFromArray([['=TRUE', '=FALSE']], {
        licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', language: 'deDE',
      })
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      expect(hf.getCellValue(adr('B1'))).toBe(false)
      hf.destroy()
    })
  })
})
