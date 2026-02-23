/**
 * Tests for parseGSheetsValue in script/parse-gsheets-value.ts.
 *
 * Verifies that GSheets-formatted values (currency, percent) are correctly
 * parsed to their numeric equivalents rather than left as strings.
 */

import {parseGSheetsValue} from '../../script/parse-gsheets-value'

describe('parseGSheetsValue', () => {
  describe('empty and null', () => {
    it('returns null for empty string', () => {
      expect(parseGSheetsValue('')).toBeNull()
    })

    it('returns null for whitespace-only string', () => {
      expect(parseGSheetsValue('   ')).toBeNull()
    })
  })

  describe('booleans', () => {
    it('returns true for TRUE', () => {
      expect(parseGSheetsValue('TRUE')).toBe(true)
    })

    it('returns false for FALSE', () => {
      expect(parseGSheetsValue('FALSE')).toBe(false)
    })
  })

  describe('error strings', () => {
    it('returns error string as-is for #VALUE!', () => {
      expect(parseGSheetsValue('#VALUE!')).toBe('#VALUE!')
    })

    it('returns error string as-is for #N/A', () => {
      expect(parseGSheetsValue('#N/A')).toBe('#N/A')
    })

    it('returns error string as-is for #DIV/0!', () => {
      expect(parseGSheetsValue('#DIV/0!')).toBe('#DIV/0!')
    })
  })

  describe('plain numbers', () => {
    it('parses integer', () => {
      expect(parseGSheetsValue('42')).toBe(42)
    })

    it('parses decimal', () => {
      expect(parseGSheetsValue('3.14')).toBe(3.14)
    })

    it('parses negative', () => {
      expect(parseGSheetsValue('-5.052998806')).toBe(-5.052998806)
    })
  })

  describe('currency values (GSheets financial function output)', () => {
    it('parses $6.25 as 6.25 (DB function)', () => {
      expect(parseGSheetsValue('$6.25')).toBe(6.25)
    })

    it('parses $17.44 as 17.44 (DDB function)', () => {
      expect(parseGSheetsValue('$17.44')).toBe(17.44)
    })

    it('parses -$1,848.51 as -1848.51 (FV function)', () => {
      expect(parseGSheetsValue('-$1,848.51')).toBe(-1848.51)
    })

    it('parses -$1,000.00 as -1000 (IPMT function)', () => {
      expect(parseGSheetsValue('-$1,000.00')).toBe(-1000)
    })

    it('parses $399.52 as 399.52 (NPV function)', () => {
      expect(parseGSheetsValue('$399.52')).toBe(399.52)
    })

    it('parses -$1,028.61 as -1028.61 (PMT function)', () => {
      expect(parseGSheetsValue('-$1,028.61')).toBe(-1028.61)
    })

    it('parses -$28.61 as -28.61 (PPMT function)', () => {
      expect(parseGSheetsValue('-$28.61')).toBe(-28.61)
    })

    it('parses -$1,057.53 as -1057.53 (PV function)', () => {
      expect(parseGSheetsValue('-$1,057.53')).toBe(-1057.53)
    })

    it('parses $5.00 as 5 (SLN function)', () => {
      expect(parseGSheetsValue('$5.00')).toBe(5)
    })

    it('parses $8.18 as 8.18 (SYD function)', () => {
      expect(parseGSheetsValue('$8.18')).toBe(8.18)
    })

    it('parses $3.49 as 3.49 (VDB function)', () => {
      expect(parseGSheetsValue('$3.49')).toBe(3.49)
    })

    it('parses $10,000.00 as 10000 (TO_DOLLARS function)', () => {
      expect(parseGSheetsValue('$10,000.00')).toBe(10000)
    })
  })

  describe('percentage values (GSheets financial function output)', () => {
    it('parses 9% as 0.09 (IRR function)', () => {
      expect(parseGSheetsValue('9%')).toBeCloseTo(0.09)
    })

    it('parses 10% as 0.10 (MIRR function)', () => {
      expect(parseGSheetsValue('10%')).toBeCloseTo(0.10)
    })

    it('parses 23% as 0.23 (RATE function)', () => {
      expect(parseGSheetsValue('23%')).toBeCloseTo(0.23)
    })

    it('parses 41% as 0.41 (TO_PERCENT function)', () => {
      expect(parseGSheetsValue('41%')).toBeCloseTo(0.41)
    })

    it('parses 0.5% as 0.005', () => {
      expect(parseGSheetsValue('0.5%')).toBeCloseTo(0.005)
    })
  })

  describe('plain string fallback', () => {
    it('returns plain text as string', () => {
      expect(parseGSheetsValue('hello')).toBe('hello')
    })

    it('returns url-encoded text as string', () => {
      expect(parseGSheetsValue('http%3A%2F%2Fwww.google.com')).toBe('http%3A%2F%2Fwww.google.com')
    })
  })
})
