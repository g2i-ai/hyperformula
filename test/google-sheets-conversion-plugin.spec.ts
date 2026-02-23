import {DetailedCellError, ErrorType, HyperFormula} from '../src'
import {adr} from './testUtils'
import {GoogleSheetsConversionPlugin} from '../src/interpreter/plugin/googleSheets/GoogleSheetsConversionPlugin'
import {NumberType} from '../src/interpreter/InterpreterValue'

const GS_OPTIONS = {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' as const}

describe('GoogleSheetsConversionPlugin', () => {
  it('is exported from the googleSheets plugin index', () => {
    expect(GoogleSheetsConversionPlugin).toBeDefined()
    expect(typeof GoogleSheetsConversionPlugin).toBe('function')
  })

  describe('TO_DATE', () => {
    it('returns a numeric date value for a serial number', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DATE(25405)']], GS_OPTIONS)
      const val = hf.getCellValue(adr('A1'))
      expect(typeof val === 'number' || (val != null && typeof val === 'object')).toBe(true)
      expect(Number(val)).toBeGreaterThan(0)
      hf.destroy()
    })

    it('preserves the numeric value of the serial', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DATE(1)']], GS_OPTIONS)
      const raw = hf.getCellValue(adr('A1'))
      // The raw numeric value should still be 1
      expect(Number(raw)).toBe(1)
      hf.destroy()
    })

    it('returns an error for a non-numeric argument', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DATE("not a number")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns a NUMBER_DATE typed value (not just a plain number)', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DATE(25405)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_DATE)
      hf.destroy()
    })
  })

  describe('TO_DOLLARS', () => {
    it('returns a NUMBER_CURRENCY typed value (not a string)', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DOLLARS(10000)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_CURRENCY)
      hf.destroy()
    })

    it('preserves the raw numeric value', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DOLLARS(42.5)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(42.5)
      hf.destroy()
    })

    it('handles zero', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DOLLARS(0)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_CURRENCY)
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('handles negative values', () => {
      const hf = HyperFormula.buildFromArray([['=TO_DOLLARS(-1)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_CURRENCY)
      expect(hf.getCellValue(adr('A1'))).toBe(-1)
      hf.destroy()
    })
  })

  describe('TO_PERCENT', () => {
    it('returns a NUMBER_PERCENT typed value (not a string)', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PERCENT(0.40826)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_PERCENT)
      hf.destroy()
    })

    it('preserves the raw numeric value', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PERCENT(0.5)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(0.5)
      hf.destroy()
    })

    it('handles zero', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PERCENT(0)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_PERCENT)
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('handles values greater than 1', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PERCENT(2)']], GS_OPTIONS)
      expect(hf.getCellValueDetailedType(adr('A1'))).toBe(NumberType.NUMBER_PERCENT)
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })
  })

  describe('TO_PURE_NUMBER', () => {
    it('returns a raw number for a plain integer', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PURE_NUMBER(50)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(50)
      hf.destroy()
    })

    it('returns the raw number for a decimal', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PURE_NUMBER(3.14)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(3.14)
      hf.destroy()
    })

    it('returns 0 for zero', () => {
      const hf = HyperFormula.buildFromArray([['=TO_PURE_NUMBER(0)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })
  })

  describe('TO_TEXT', () => {
    it('converts a number to its string representation', () => {
      const hf = HyperFormula.buildFromArray([['=TO_TEXT(24)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe('24')
      hf.destroy()
    })

    it('converts a decimal to its string representation', () => {
      const hf = HyperFormula.buildFromArray([['=TO_TEXT(3.14)']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe('3.14')
      hf.destroy()
    })

    it('converts TRUE() to "true"', () => {
      const hf = HyperFormula.buildFromArray([['=TO_TEXT(TRUE())']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe('true')
      hf.destroy()
    })

    it('converts a string to itself', () => {
      const hf = HyperFormula.buildFromArray([['=TO_TEXT("hello")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe('hello')
      hf.destroy()
    })
  })

  describe('CONVERT', () => {
    it('converts grams to kilograms', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(5.1,"g","kg")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.0051, 7)
      hf.destroy()
    })

    it('converts miles to kilometers', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"mi","km")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1.609344, 5)
      hf.destroy()
    })

    it('converts meters to feet', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"m","ft")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1 / 0.3048, 5)
      hf.destroy()
    })

    it('converts Celsius to Fahrenheit: 100C = 212F', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(100,"C","F")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(212, 5)
      hf.destroy()
    })

    it('converts Fahrenheit to Celsius: 32F = 0C', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(32,"F","C")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0, 5)
      hf.destroy()
    })

    it('converts Celsius to Kelvin: 0C = 273.15K', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(0,"C","K")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(273.15, 5)
      hf.destroy()
    })

    it('converts Kelvin to Fahrenheit: 273.15K = 32F', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(273.15,"K","F")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(32, 4)
      hf.destroy()
    })

    it('returns same value when from_unit equals to_unit', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(42,"m","m")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(42)
      hf.destroy()
    })

    it('returns #N/A when both units are the same but unknown', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"xyz","xyz")']], GS_OPTIONS)
      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      expect((val as DetailedCellError).type).toBe(ErrorType.NA)
      hf.destroy()
    })

    it('converts liters to gallons', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"l","gal")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1 / 3.785411784, 5)
      hf.destroy()
    })

    it('returns #N/A for incompatible units (length vs mass)', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"m","kg")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns #N/A for an unknown from_unit', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"xyz","m")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns #N/A for an unknown to_unit', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"m","xyz")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns #N/A (not #NUM) when unit is a prototype property like "toString"', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"toString","m")']], GS_OPTIONS)
      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      expect((val as DetailedCellError).type).toBe(ErrorType.NA)
      hf.destroy()
    })

    it('returns #N/A (not #NUM) when unit is a prototype property like "constructor"', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"constructor","m")']], GS_OPTIONS)
      const val = hf.getCellValue(adr('A1'))
      expect(val).toBeInstanceOf(DetailedCellError)
      expect((val as DetailedCellError).type).toBe(ErrorType.NA)
      hf.destroy()
    })

    it('converts hours to seconds', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"hr","sec")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBe(3600)
      hf.destroy()
    })

    it('converts 1 erg (e) to joules: 1 erg = 1e-7 J', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"e","J")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1e-7, 20)
      hf.destroy()
    })

    it('converts 1e7 ergs to 1 joule', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1e7,"e","J")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1, 10)
      hf.destroy()
    })

    it('converts 1 joule to 1e7 ergs', () => {
      const hf = HyperFormula.buildFromArray([['=CONVERT(1,"J","e")']], GS_OPTIONS)
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1e7, 0)
      hf.destroy()
    })
  })
})
