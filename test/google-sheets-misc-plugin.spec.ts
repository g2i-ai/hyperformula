import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

describe('GoogleSheetsMiscPlugin', () => {
  describe('GESTEP', () => {
    it('should return 1 if number >= step', () => {
      const hf = HyperFormula.buildFromArray([
        ['=GESTEP(5, 2)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('should return 0 if number < step', () => {
      const hf = HyperFormula.buildFromArray([
        ['=GESTEP(1, 2)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('should return 1 if number equals step', () => {
      const hf = HyperFormula.buildFromArray([
        ['=GESTEP(2, 2)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('should use default step value of 0', () => {
      const hf = HyperFormula.buildFromArray([
        ['=GESTEP(5)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('should return 1 for any positive number with default step=0', () => {
      const hf = HyperFormula.buildFromArray([
        ['=GESTEP(-1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })
  })

  describe('LOOKUP with 3 arguments', () => {
    it('should find value in sorted search range and return from result range', () => {
      const hf = HyperFormula.buildFromArray([
        ['=LOOKUP(10003, {10001; 10002; 10003}, {"Alice"; "Bob"; "Charlie"})'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('Charlie')
      hf.destroy()
    })

    it('should find largest value <= search key', () => {
      const hf = HyperFormula.buildFromArray([
        ['=LOOKUP(10002.5, {10001; 10002; 10003}, {"Alice"; "Bob"; "Charlie"})'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('Bob')
      hf.destroy()
    })

    it('should return NA if search key is smaller than all values', () => {
      const hf = HyperFormula.buildFromArray([
        ['=LOOKUP(9999, {10001; 10002; 10003}, {"Alice"; "Bob"; "Charlie"})'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('LOOKUP with 2 arguments', () => {
    it('should search in 2-column array and return from second column', () => {
      const hf = HyperFormula.buildFromArray([
        ['=LOOKUP(10002, {10001, "Alice"; 10002, "Bob"; 10003, "Charlie"})'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('Bob')
      hf.destroy()
    })

    it('should find largest value <= key in 2-column array', () => {
      const hf = HyperFormula.buildFromArray([
        ['=LOOKUP(10002.5, {10001, "Alice"; 10002, "Bob"; 10003, "Charlie"})'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('Bob')
      hf.destroy()
    })
  })

  describe('CELL', () => {
    it('should return row number for "row" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['', '=CELL("row", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('should return column number for "col" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['', '=CELL("col", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('should return correct address for "address" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['', '=CELL("address", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe('$A$1')
      hf.destroy()
    })

    it('should return contents of cell for "contents" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['Hello', '=CELL("contents", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe('Hello')
      hf.destroy()
    })

    it('should return "v" for numeric cells with "type" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['42', '=CELL("type", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe('v')
      hf.destroy()
    })

    it('should return "t" for text cells with "type" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        ['hello', '=CELL("type", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe('t')
      hf.destroy()
    })

    it('should return "b" for empty cells with "type" info_type', () => {
      const hf = HyperFormula.buildFromArray([
        [null, '=CELL("type", A1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('B1'))).toBe('b')
      hf.destroy()
    })
  })

  describe('ENCODEURL', () => {
    it('should URL-encode a string', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ENCODEURL("http://www.google.com")'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('http%3A%2F%2Fwww.google.com')
      hf.destroy()
    })

    it('should encode spaces as %20', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ENCODEURL("hello world")'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('hello%20world')
      hf.destroy()
    })

    it('should encode special characters', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ENCODEURL("a&b=c")'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('a%26b%3Dc')
      hf.destroy()
    })
  })

  describe('EPOCHTODATE', () => {
    it('should convert Unix timestamp (seconds) to date serial number', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(0)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(25569)
      hf.destroy()
    })

    it('should handle positive Unix timestamps', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(86400)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(25570)
      hf.destroy()
    })

    it('should convert milliseconds to date serial when unit=2', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(86400000, 2)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(25570)
      hf.destroy()
    })

    it('should convert microseconds to date serial when unit=3', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(86400000000, 3)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(25570)
      hf.destroy()
    })

    it('should use default unit=1 (seconds)', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(86400)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(25570)
      hf.destroy()
    })
  })

  describe('ISDATE', () => {
    it('should return TRUE for date values', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE(DATE(2023, 1, 1))'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return FALSE for numeric values', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE(42)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return FALSE for text values', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE("2023-01-01")'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return FALSE for boolean values', () => {
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE(TRUE())'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return TRUE for DateTimeNumber values (NOW returns date+time)', () => {
      // DateTimeNumber is a sibling of DateNumber — both represent dates
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE(NOW())'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should not parse string arguments as dates', () => {
      // Google Sheets ISDATE only checks if a value IS a date, not whether
      // a string can be parsed into one
      const hf = HyperFormula.buildFromArray([
        ['=ISDATE("January 1, 2023")'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('CELL address beyond Z', () => {
    it('should return correct absolute address for column AA (index 26)', () => {
      const hf = HyperFormula.buildFromArray([
        ['=CELL("address", AA1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('$AA$1')
      hf.destroy()
    })

    it('should return correct absolute address for column AB (index 27)', () => {
      const hf = HyperFormula.buildFromArray([
        ['=CELL("address", AB1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('$AB$1')
      hf.destroy()
    })

    it('should return correct absolute address for column AZ (index 51)', () => {
      const hf = HyperFormula.buildFromArray([
        ['=CELL("address", AZ1)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A1'))).toBe('$AZ$1')
      hf.destroy()
    })
  })

  describe('LOOKUP with string keys', () => {
    it('should find exact string match in 3-argument form', () => {
      const hf = HyperFormula.buildFromArray([
        [null, null],
        ['apple', 10],
        ['banana', 20],
        ['cherry', 30],
        ['=LOOKUP("banana", A2:A4, B2:B4)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A5'))).toBe(20)
      hf.destroy()
    })

    it('should return largest value <= string key in 3-argument form', () => {
      const hf = HyperFormula.buildFromArray([
        [null, null],
        ['apple', 10],
        ['banana', 20],
        ['cherry', 30],
        // "b" comes after "apple" but before "banana" — should return apple's value
        ['=LOOKUP("b", A2:A4, B2:B4)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      expect(hf.getCellValue(adr('A5'))).toBe(10)
      hf.destroy()
    })

    it('should return NA for string key before all entries in 3-argument form', () => {
      const hf = HyperFormula.buildFromArray([
        [null, null],
        ['banana', 20],
        ['cherry', 30],
        ['=LOOKUP("aaa", A2:A3, B2:B3)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      const val = hf.getCellValue(adr('A4'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('EPOCHTODATE invalid unit', () => {
    it('should return VALUE error for unit=0', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(0, 0)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('should return VALUE error for unit=4', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(0, 4)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('should return VALUE error for unit=99', () => {
      const hf = HyperFormula.buildFromArray([
        ['=EPOCHTODATE(0, 99)'],
      ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })
})
