import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

describe('GoogleSheetsTextPlugin - SPLIT', () => {
  it('should split by delimiter and return array', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("hello world", " ")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe('hello')
    expect(hf.getCellValue(adr('B1'))).toBe('world')
    hf.destroy()
  })

  it('should split by each character in delimiter when split_by_each is true', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a-b.c", "-.")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // Default split_by_each is TRUE in GSheets
    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('b')
    expect(hf.getCellValue(adr('C1'))).toBe('c')
    hf.destroy()
  })

  it('should split by whole delimiter when split_by_each is false', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a-b-.c-d", "-.", FALSE())'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // With split_by_each=FALSE, "-." is treated as a single delimiter
    expect(hf.getCellValue(adr('A1'))).toBe('a-b')
    expect(hf.getCellValue(adr('B1'))).toBe('c-d')
    hf.destroy()
  })

  it('should remove empty text by default', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a,,b", ",")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // Default remove_empty_text is TRUE — empty strings are removed
    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('b')
    hf.destroy()
  })

  it('should keep empty text when remove_empty_text is false', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a,,b", ",", TRUE(), FALSE())'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('')
    expect(hf.getCellValue(adr('C1'))).toBe('b')
    hf.destroy()
  })

  it('should return VALUE error for empty delimiter', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("abc", "")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('should not load GoogleSheets SPLIT in default mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a b", 1)'],
    ], {licenseKey: 'gpl-v3'})

    // HF built-in SPLIT: SPLIT(string, index) — index 1 returns second word
    expect(hf.getCellValue(adr('A1'))).toBe('b')
    hf.destroy()
  })

  it('should split with cell reference as text argument', () => {
    const hf = HyperFormula.buildFromArray([
      ['one,two,three', '=SPLIT(A1, ",")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('B1'))).toBe('one')
    expect(hf.getCellValue(adr('C1'))).toBe('two')
    expect(hf.getCellValue(adr('D1'))).toBe('three')
    hf.destroy()
  })

  it('should handle regex-special characters in delimiter', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("a(b)c", "()")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    // split_by_each=TRUE by default, so "(" and ")" are separate delimiters
    expect(hf.getCellValue(adr('A1'))).toBe('a')
    expect(hf.getCellValue(adr('B1'))).toBe('b')
    expect(hf.getCellValue(adr('C1'))).toBe('c')
    hf.destroy()
  })

  it('should return single-element array when delimiter not found', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT("hello", ",")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe('hello')
    hf.destroy()
  })

  it('should return VALUE error when all parts are empty after removeEmptyText', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SPLIT(",,,", ",")'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})
