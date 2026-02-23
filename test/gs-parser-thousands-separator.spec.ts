import { HyperFormula } from '../src'
import { Config } from '../src/Config'
import { adr } from './testUtils'

const gsConfig = { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' as const }
const defaultConfig = { licenseKey: 'gpl-v3' }

describe('Thousands separator - GSheets mode', () => {
  describe('basic thousands grouping', () => {
    it('parses 1,000 as 1000', () => {
      const hf = HyperFormula.buildFromArray([['=1,000']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000)
      hf.destroy()
    })

    it('parses chained groups: 1,000,000', () => {
      const hf = HyperFormula.buildFromArray([['=1,000,000']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000000)
      hf.destroy()
    })

    it('parses triple chain: 1,000,000,000', () => {
      const hf = HyperFormula.buildFromArray([['=1,000,000,000']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000000000)
      hf.destroy()
    })

    it('parses thousands with decimal: 1,000.5', () => {
      const hf = HyperFormula.buildFromArray([['=1,000.5']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000.5)
      hf.destroy()
    })

    it('parses chained thousands with decimal: 1,000,000.123', () => {
      const hf = HyperFormula.buildFromArray([['=1,000,000.123']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000000.123)
      hf.destroy()
    })

    it('parses 10,000', () => {
      const hf = HyperFormula.buildFromArray([['=10,000']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(10000)
      hf.destroy()
    })

    it('parses 100,000', () => {
      const hf = HyperFormula.buildFromArray([['=100,000']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(100000)
      hf.destroy()
    })
  })

  describe('non-thousands comma (argument separator)', () => {
    it('SUM(1,2) treats comma as arg separator', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,2)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(3)
      hf.destroy()
    })

    it('SUM(1,2,3) treats commas as arg separators', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,2,3)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(6)
      hf.destroy()
    })

    it('SUM(1,00) - 2 digits after comma → arg separator', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,00)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('SUM(1,0000) - 4 digits after comma → arg separator', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,0000)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })
  })

  describe('ambiguous / tricky cases', () => {
    it('SUM(1,000) - 3 digits → thousands grouping', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,000)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000)
      hf.destroy()
    })

    it('SUM(1,000,000) - greedy chain → single number', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,000,000)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1000000)
      hf.destroy()
    })

    it('SUM(1,000,00) - partial chain: 1,000 is thousands, ,00 is arg + 0', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1,000,00)']], gsConfig)
      // 1,000 = 1000, then comma separates, 00 = 0 → SUM(1000, 0) = 1000
      expect(hf.getCellValue(adr('A1'))).toBe(1000)
      hf.destroy()
    })

    it('IF(TRUE,1,000,0) - 1,000 is thousands, ,0 is arg', () => {
      const hf = HyperFormula.buildFromArray([['=IF(TRUE,1,000,0)']], gsConfig)
      // IF(TRUE, 1000, 0) = 1000
      expect(hf.getCellValue(adr('A1'))).toBe(1000)
      hf.destroy()
    })

    it('1,000+2 - thousands in arithmetic', () => {
      const hf = HyperFormula.buildFromArray([['=1,000+2']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(1002)
      hf.destroy()
    })

    it('1,000e2 - thousands with scientific notation', () => {
      const hf = HyperFormula.buildFromArray([['=1,000e2']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(100000)
      hf.destroy()
    })

    it('1E5 - uppercase E in scientific notation (default mode, no thousandSeparator)', () => {
      const hf = HyperFormula.buildFromArray([['=1E5']], defaultConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(100000)
      hf.destroy()
    })

    it('1E5 - uppercase E in scientific notation (GSheets mode with thousandSeparator)', () => {
      const hf = HyperFormula.buildFromArray([['=1E5']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(100000)
      hf.destroy()
    })

    it('1,000E2 - thousands with uppercase E scientific notation', () => {
      const hf = HyperFormula.buildFromArray([['=1,000E2']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(100000)
      hf.destroy()
    })
  })

  describe('leading group >3 digits (no thousands grouping)', () => {
    it('SUM(1000,000) - 4-digit leading group → two args', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(1000,000)']], gsConfig)
      // 1000 is a number, 000 is a number → SUM(1000, 0) = 1000
      expect(hf.getCellValue(adr('A1'))).toBe(1000)
      hf.destroy()
    })

    it('SUM(12345,678) - 5-digit leading group → two args', () => {
      const hf = HyperFormula.buildFromArray([['=SUM(12345,678)']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(13023)
      hf.destroy()
    })
  })

  describe('array literals', () => {
    it('SUM({1,000;2,000}) - thousands in array literal', () => {
      const hf = HyperFormula.buildFromArray([['=SUM({1,000;2,000})']], gsConfig)
      expect(hf.getCellValue(adr('A1'))).toBe(3000)
      hf.destroy()
    })
  })
})

describe('Thousands separator - default mode', () => {
  it('SUM(1,000) in default mode treats comma as arg sep', () => {
    const hf = HyperFormula.buildFromArray([['=SUM(1,000)']], defaultConfig)
    // SUM(1, 0) = 1
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    hf.destroy()
  })
})

describe('Thousands separator - config validation', () => {
  it('allows thousandSeparator to equal functionArgSeparator', () => {
    expect(() => new Config({
      licenseKey: 'gpl-v3',
      thousandSeparator: ',',
      functionArgSeparator: ',',
    })).not.toThrow()
  })

  it('rejects thousandSeparator equal to decimalSeparator (comma)', () => {
    expect(() => new Config({
      licenseKey: 'gpl-v3',
      thousandSeparator: ',',
      decimalSeparator: ',',
    })).toThrow()
  })

  it('rejects thousandSeparator equal to decimalSeparator (dot)', () => {
    expect(() => new Config({
      licenseKey: 'gpl-v3',
      thousandSeparator: '.',
      decimalSeparator: '.',
    })).toThrow()
  })

  it('allows empty thousandSeparator (no conflict check)', () => {
    expect(() => new Config({
      licenseKey: 'gpl-v3',
      thousandSeparator: '',
    })).not.toThrow()
  })
})

describe('Cross-instance isolation - singleton state corruption', () => {
  it('GSheets instance correctly parses thousands after a default instance reconfigures the shared matcher', () => {
    // Create a GSheets instance first, then a default-config instance.
    // The default instance's buildLexerConfig call reconfigures the shared singleton
    // (thousandSeparator → '').  Now mutate hfGs — setCellContents re-lexes the
    // formula using the shared singleton, so if the bug is present the comma in
    // "1,000" is treated as an arg separator and the result is wrong.
    const hfGs = HyperFormula.buildFromArray([['']], gsConfig)
    HyperFormula.buildFromArray([['']], defaultConfig).destroy()

    hfGs.setCellContents(adr('A1'), '=1,000')
    expect(hfGs.getCellValue(adr('A1'))).toBe(1000)

    hfGs.destroy()
  })

  it('default instance correctly treats comma as arg separator after a GSheets instance reconfigures the shared matcher', () => {
    // Create a default instance first, then a GSheets instance.
    // The GSheets buildLexerConfig sets thousandSeparator → ',', which would cause
    // the default instance's lexer (still bound to the shared singleton) to
    // misinterpret "1,000" as a thousands-grouped number instead of two args.
    const hfDefault = HyperFormula.buildFromArray([['']], defaultConfig)
    HyperFormula.buildFromArray([['']], gsConfig).destroy()

    hfDefault.setCellContents(adr('A1'), '=SUM(1,000)')
    expect(hfDefault.getCellValue(adr('A1'))).toBe(1) // SUM(1, 0) = 1

    hfDefault.destroy()
  })
})

describe('Thousands separator - GSheets preset auto-config', () => {
  it('sets thousandSeparator to comma in GSheets mode', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    expect(config.thousandSeparator).toBe(',')
  })

  it('allows explicit override of thousandSeparator in GSheets mode', () => {
    const config = new Config({
      licenseKey: 'gpl-v3',
      compatibilityMode: 'googleSheets',
      thousandSeparator: '',
    })
    expect(config.thousandSeparator).toBe('')
  })

  it('keeps empty thousandSeparator in default mode', () => {
    const config = new Config({ licenseKey: 'gpl-v3' })
    expect(config.thousandSeparator).toBe('')
  })
})
