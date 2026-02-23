import {HyperFormula, DetailedCellError} from '../src'
import {adr} from './testUtils'

describe('GoogleSheetsInfoPlugin', () => {
  const buildWithPlugin = (data: any[][]) =>
    HyperFormula.buildFromArray(data, {
      compatibilityMode: 'googleSheets',
      licenseKey: 'gpl-v3',
    })

  describe('dependency tracking', () => {
    it('ERROR.TYPE recalculates when referenced cell changes from error to non-error', () => {
      const hf = buildWithPlugin([['=1/0', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(2) // DIV/0 = 2
      hf.setCellContents(adr('A1'), [[42]])
      const result = hf.getCellValue(adr('B1')) as DetailedCellError
      expect(result.type).toBe('NA') // non-error → #N/A
      hf.destroy()
    })

    it('ERROR.TYPE recalculates when referenced cell changes error type', () => {
      const hf = buildWithPlugin([['=1/0', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(2) // DIV/0 = 2
      hf.setCellContents(adr('A1'), [['=NA()']])
      expect(hf.getCellValue(adr('B1'))).toBe(7) // NA = 7
      hf.destroy()
    })

    it('TYPE recalculates when referenced cell changes from number to text', () => {
      const hf = buildWithPlugin([[42, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1) // number
      hf.setCellContents(adr('A1'), [['hello']])
      expect(hf.getCellValue(adr('B1'))).toBe(2) // text
      hf.destroy()
    })

    it('TYPE recalculates when referenced cell changes from number to boolean', () => {
      const hf = buildWithPlugin([[42, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1) // number
      hf.setCellContents(adr('A1'), [['=TRUE()']])
      expect(hf.getCellValue(adr('B1'))).toBe(4) // boolean
      hf.destroy()
    })

    it('TYPE recalculates when referenced cell changes from number to error', () => {
      const hf = buildWithPlugin([[42, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1) // number
      hf.setCellContents(adr('A1'), [['=1/0']])
      expect(hf.getCellValue(adr('B1'))).toBe(16) // error
      hf.destroy()
    })
  })

  describe('ERROR.TYPE', () => {
    it('returns 2 for DIV/0 error', () => {
      const hf = buildWithPlugin([['=1/0', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(2)
      hf.destroy()
    })

    it('returns 3 for VALUE error', () => {
      const hf = buildWithPlugin([['=1+"text"', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(3)
      hf.destroy()
    })

    it('returns 5 for NAME error', () => {
      const hf = buildWithPlugin([['=UNDEFINED_FUNCTION()', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(5)
      hf.destroy()
    })

    it('returns 6 for NUM error', () => {
      const hf = buildWithPlugin([['=SQRT(-1)', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(6)
      hf.destroy()
    })

    it('returns 7 for N/A error', () => {
      const hf = buildWithPlugin([['=NA()', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(7)
      hf.destroy()
    })

    it('returns 8 for generic #ERROR! (parse error)', () => {
      // A parse error like "=+" produces ErrorType.ERROR which maps to #ERROR! in Google Sheets
      const hf = buildWithPlugin([['=+', '=ERROR.TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(8)
      hf.destroy()
    })

    it('returns #N/A for non-error value', () => {
      const hf = buildWithPlugin([['=42', '=ERROR.TYPE(A1)']])
      const result = hf.getCellValue(adr('B1')) as DetailedCellError
      expect(result.type).toBe('NA')
      hf.destroy()
    })

    it('returns #N/A for text value', () => {
      const hf = buildWithPlugin([['="hello"', '=ERROR.TYPE(A1)']])
      const result = hf.getCellValue(adr('B1')) as DetailedCellError
      expect(result.type).toBe('NA')
      hf.destroy()
    })

    it('handles missing argument', () => {
      const hf = buildWithPlugin([['=ERROR.TYPE()']])
      const result = hf.getCellValue(adr('A1')) as DetailedCellError
      expect(result.type).toBe('NA')
      hf.destroy()
    })
  })

  describe('TYPE', () => {
    it('returns 1 for number', () => {
      const hf = buildWithPlugin([[42, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('returns 1 for empty cell', () => {
      const hf = buildWithPlugin([[null, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('returns 2 for empty string literal (not empty cell)', () => {
      const hf = buildWithPlugin([['=""', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(2) // empty string is text, not empty
      hf.destroy()
    })


    it('returns 2 for text', () => {
      const hf = buildWithPlugin([['="hello"', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(2)
      hf.destroy()
    })

    it('returns 4 for boolean true', () => {
      const hf = buildWithPlugin([['=TRUE()', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(4)
      hf.destroy()
    })

    it('returns 4 for boolean false', () => {
      const hf = buildWithPlugin([['=FALSE()', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(4)
      hf.destroy()
    })

    it('returns 16 for error', () => {
      const hf = buildWithPlugin([['=1/0', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(16)
      hf.destroy()
    })

    it('returns 64 for array', () => {
      const hf = buildWithPlugin([['={1,2,3}', '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(64)
      hf.destroy()
    })

    it('returns value type for non-root spill cells (not 64)', () => {
      // ={1,2,3} spills into A1, B1, C1 — only A1 is the root (array type 64)
      // B1 and C1 are spill cells and should return the type of their individual value
      const hf = buildWithPlugin([['={1,2,3}', null, null, '=TYPE(A1)', '=TYPE(B1)', '=TYPE(C1)']])
      expect(hf.getCellValue(adr('D1'))).toBe(64) // A1 is root of array
      expect(hf.getCellValue(adr('E1'))).toBe(1)  // B1 is spill cell with number value
      expect(hf.getCellValue(adr('F1'))).toBe(1)  // C1 is spill cell with number value
      hf.destroy()
    })

    it('returns text type for spill cell containing text', () => {
      // ={"hello","world"} spills text values
      const hf = buildWithPlugin([['={"hello","world"}', null, '=TYPE(A1)', '=TYPE(B1)']])
      expect(hf.getCellValue(adr('C1'))).toBe(64) // A1 is root of array
      expect(hf.getCellValue(adr('D1'))).toBe(2)  // B1 is spill cell with text value
      hf.destroy()
    })


    it('returns 1 for decimal number', () => {
      const hf = buildWithPlugin([[3.14, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('returns 1 for negative number', () => {
      const hf = buildWithPlugin([[-42, '=TYPE(A1)']])
      expect(hf.getCellValue(adr('B1'))).toBe(1)
      hf.destroy()
    })

    it('handles missing argument', () => {
      const hf = buildWithPlugin([['=TYPE()']])
      const result = hf.getCellValue(adr('A1')) as DetailedCellError
      expect(result.type).toBe('NA')
      hf.destroy()
    })
  })

  describe('ISEMAIL', () => {
    it('returns true for valid email', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user@example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for email with multiple domain levels', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user@mail.example.co.uk")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for email with numbers', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user123@example456.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for email with special characters', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user.name+tag@example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns false for email without @ symbol', () => {
      const hf = buildWithPlugin([['=ISEMAIL("userexample.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for email without domain', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user@")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for email without local part', () => {
      const hf = buildWithPlugin([['=ISEMAIL("@example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for email without dot in domain', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user@example")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for email with spaces', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user name@example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for empty string', () => {
      const hf = buildWithPlugin([['=ISEMAIL("")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for multiple @ symbols', () => {
      const hf = buildWithPlugin([['=ISEMAIL("user@example@com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('ISURL', () => {
    it('returns true for valid http URL', () => {
      const hf = buildWithPlugin([['=ISURL("http://www.google.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for https URL', () => {
      const hf = buildWithPlugin([['=ISURL("https://www.google.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for URL with path', () => {
      const hf = buildWithPlugin([['=ISURL("https://example.com/path/to/resource")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for URL with query parameters', () => {
      const hf = buildWithPlugin([['=ISURL("https://example.com?param=value")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for URL with fragments', () => {
      const hf = buildWithPlugin([['=ISURL("https://example.com#section")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for HTTP (uppercase)', () => {
      const hf = buildWithPlugin([['=ISURL("HTTP://example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns true for HTTPS (uppercase)', () => {
      const hf = buildWithPlugin([['=ISURL("HTTPS://example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns false for URL without protocol', () => {
      const hf = buildWithPlugin([['=ISURL("www.google.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for FTP URL', () => {
      const hf = buildWithPlugin([['=ISURL("ftp://example.com")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for empty string', () => {
      const hf = buildWithPlugin([['=ISURL("")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for plain text', () => {
      const hf = buildWithPlugin([['=ISURL("hello world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for URL with only protocol', () => {
      const hf = buildWithPlugin([['=ISURL("http://")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('returns false for https without domain', () => {
      const hf = buildWithPlugin([['=ISURL("https://")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })
})
