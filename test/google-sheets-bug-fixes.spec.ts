import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

const GS_CONFIG = {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'} as const

describe('GoogleSheetsEngineeringFixesPlugin — HEX2BIN', () => {
  it('converts lowercase hex to binary', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2BIN("f3", 8)']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('11110011')
    hf.destroy()
  })

  it('converts uppercase hex to binary (regression: must still work)', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2BIN("F3", 8)']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('11110011')
    hf.destroy()
  })

  it('converts mixed-case hex to binary', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2BIN("aB")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('10101011')
    hf.destroy()
  })

  it('returns NUM error for invalid hex characters', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2BIN("xyz")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('returns NUM error for places shorter than result', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2BIN("FF", 2)']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('GoogleSheetsEngineeringFixesPlugin — HEX2DEC', () => {
  it('converts lowercase hex to decimal', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2DEC("f3")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe(243)
    hf.destroy()
  })

  it('converts uppercase hex to decimal (regression: must still work)', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2DEC("F3")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe(243)
    hf.destroy()
  })

  it('handles single lowercase digit', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2DEC("a")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe(10)
    hf.destroy()
  })

  it('handles twos complement negative hex', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2DEC("ffffffffff")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe(-1)
    hf.destroy()
  })

  it('returns NUM error for invalid hex string', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2DEC("zz")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('GoogleSheetsEngineeringFixesPlugin — HEX2OCT', () => {
  it('converts lowercase hex to octal', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2OCT("f3", 8)']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('00000363')
    hf.destroy()
  })

  it('converts uppercase hex to octal (regression: must still work)', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2OCT("F3", 8)']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('00000363')
    hf.destroy()
  })

  it('converts hex without places', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2OCT("1f")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBe('37')
    hf.destroy()
  })

  it('returns NUM error for invalid hex string', () => {
    const hf = HyperFormula.buildFromArray([['=HEX2OCT("xyz")']], GS_CONFIG)
    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('GoogleSheetsStatisticalFixesPlugin — NORM.S.DIST', () => {
  it('defaults cumulative to true when omitted', () => {
    const hf = HyperFormula.buildFromArray([['=NORM.S.DIST(2.4)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.9918024641, 6)
    hf.destroy()
  })

  it('returns cumulative CDF when cumulative is TRUE', () => {
    const hf = HyperFormula.buildFromArray([['=NORM.S.DIST(2.4, TRUE())']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.9918024641, 6)
    hf.destroy()
  })

  it('returns PDF when cumulative is FALSE', () => {
    const hf = HyperFormula.buildFromArray([['=NORM.S.DIST(0, FALSE())']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    // PDF at 0 for standard normal is 1/sqrt(2*pi) ≈ 0.3989
    expect(val).toBeCloseTo(0.3989422804, 6)
    hf.destroy()
  })

  it('NORMSDIST alias also defaults cumulative to true', () => {
    const hf = HyperFormula.buildFromArray([['=NORMSDIST(2.4)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.9918024641, 6)
    hf.destroy()
  })
})

describe('GoogleSheetsStatisticalFixesPlugin — HYPGEOM.DIST', () => {
  it('defaults cumulative to true when omitted (returns CDF, not PDF)', () => {
    const hf = HyperFormula.buildFromArray([['=HYPGEOM.DIST(4, 12, 20, 40)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    // CDF value — must NOT equal PDF value (0.1092430027)
    expect(val).toBeCloseTo(0.15042239125, 6)
    hf.destroy()
  })

  it('returns cumulative CDF when cumulative is TRUE (matches default)', () => {
    const hfDefault = HyperFormula.buildFromArray([['=HYPGEOM.DIST(4, 12, 20, 40)']], GS_CONFIG)
    const hfExplicit = HyperFormula.buildFromArray([['=HYPGEOM.DIST(4, 12, 20, 40, TRUE())']], GS_CONFIG)
    const defaultVal = hfDefault.getCellValue(adr('A1')) as number
    const explicitVal = hfExplicit.getCellValue(adr('A1')) as number
    expect(defaultVal).toBeCloseTo(explicitVal, 10)
    hfDefault.destroy()
    hfExplicit.destroy()
  })

  it('PDF differs from CDF (validates cumulative flag works)', () => {
    const hfCDF = HyperFormula.buildFromArray([['=HYPGEOM.DIST(4, 12, 20, 40, TRUE())']], GS_CONFIG)
    const hfPDF = HyperFormula.buildFromArray([['=HYPGEOM.DIST(4, 12, 20, 40, FALSE())']], GS_CONFIG)
    const cdfVal = hfCDF.getCellValue(adr('A1')) as number
    const pdfVal = hfPDF.getCellValue(adr('A1')) as number
    expect(cdfVal).not.toBeCloseTo(pdfVal, 3)
    hfCDF.destroy()
    hfPDF.destroy()
  })

  it('HYPGEOMDIST alias also defaults cumulative to true', () => {
    const hf = HyperFormula.buildFromArray([['=HYPGEOMDIST(4, 12, 20, 40)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.15042239125, 6)
    hf.destroy()
  })
})

describe('GoogleSheetsStatisticalFixesPlugin — LOGNORM.DIST', () => {
  it('defaults cumulative to true when omitted', () => {
    const hf = HyperFormula.buildFromArray([['=LOGNORM.DIST(4, 4, 6)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.3315570972, 6)
    hf.destroy()
  })

  it('LOGNORMDIST alias also defaults cumulative to true', () => {
    const hf = HyperFormula.buildFromArray([['=LOGNORMDIST(4, 4, 6)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.3315570972, 6)
    hf.destroy()
  })
})

describe('GoogleSheetsStatisticalFixesPlugin — NEGBINOM.DIST', () => {
  it('defaults cumulative to true when omitted (returns CDF, not PDF)', () => {
    const hf = HyperFormula.buildFromArray([['=NEGBINOM.DIST(4, 2, 0.1)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    // CDF value — 0.032805 is the PDF, CDF is larger
    expect(val).toBeCloseTo(0.114265, 5)
    hf.destroy()
  })

  it('returns cumulative CDF when cumulative is TRUE (matches default)', () => {
    const hfDefault = HyperFormula.buildFromArray([['=NEGBINOM.DIST(4, 2, 0.1)']], GS_CONFIG)
    const hfExplicit = HyperFormula.buildFromArray([['=NEGBINOM.DIST(4, 2, 0.1, TRUE())']], GS_CONFIG)
    const defaultVal = hfDefault.getCellValue(adr('A1')) as number
    const explicitVal = hfExplicit.getCellValue(adr('A1')) as number
    expect(defaultVal).toBeCloseTo(explicitVal, 10)
    hfDefault.destroy()
    hfExplicit.destroy()
  })

  it('PDF differs from CDF (validates cumulative flag works)', () => {
    const hfCDF = HyperFormula.buildFromArray([['=NEGBINOM.DIST(4, 2, 0.1, TRUE())']], GS_CONFIG)
    const hfPDF = HyperFormula.buildFromArray([['=NEGBINOM.DIST(4, 2, 0.1, FALSE())']], GS_CONFIG)
    const cdfVal = hfCDF.getCellValue(adr('A1')) as number
    const pdfVal = hfPDF.getCellValue(adr('A1')) as number
    expect(cdfVal).not.toBeCloseTo(pdfVal, 3)
    hfCDF.destroy()
    hfPDF.destroy()
  })

  it('NEGBINOMDIST alias also defaults cumulative to true', () => {
    const hf = HyperFormula.buildFromArray([['=NEGBINOMDIST(4, 2, 0.1)']], GS_CONFIG)
    const val = hf.getCellValue(adr('A1')) as number
    expect(val).toBeCloseTo(0.114265, 5)
    hf.destroy()
  })
})
