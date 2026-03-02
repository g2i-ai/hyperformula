import {HyperFormula} from '../src'
import {adr} from './testUtils'
import {DetailedCellError} from '../src'

describe('GoogleSheetsOperatorPlugin', () => {
  const buildWithPlugin = (data: any[][]) =>
    HyperFormula.buildFromArray(data, {
      compatibilityMode: 'googleSheets',
      licenseKey: 'gpl-v3',
    })

  describe('ADD', () => {
    it('should add two positive numbers', () => {
      const hf = buildWithPlugin([['=ADD(2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(5)
      hf.destroy()
    })

    it('should add negative numbers', () => {
      const hf = buildWithPlugin([['=ADD(-5,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-2)
      hf.destroy()
    })

    it('should add decimals', () => {
      const hf = buildWithPlugin([['=ADD(1.5,2.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })
  })

  describe('MINUS', () => {
    it('should subtract two positive numbers', () => {
      const hf = buildWithPlugin([['=MINUS(8,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(5)
      hf.destroy()
    })

    it('should subtract resulting in negative', () => {
      const hf = buildWithPlugin([['=MINUS(3,8)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-5)
      hf.destroy()
    })

    it('should subtract with decimals', () => {
      const hf = buildWithPlugin([['=MINUS(5.5,2.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(3)
      hf.destroy()
    })
  })

  describe('MULTIPLY', () => {
    it('should multiply two positive numbers', () => {
      const hf = buildWithPlugin([['=MULTIPLY(6,7)']])
      expect(hf.getCellValue(adr('A1'))).toBe(42)
      hf.destroy()
    })

    it('should multiply positive and negative', () => {
      const hf = buildWithPlugin([['=MULTIPLY(6,-7)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-42)
      hf.destroy()
    })

    it('should multiply two negative numbers', () => {
      const hf = buildWithPlugin([['=MULTIPLY(-6,-7)']])
      expect(hf.getCellValue(adr('A1'))).toBe(42)
      hf.destroy()
    })

    it('should multiply by zero', () => {
      const hf = buildWithPlugin([['=MULTIPLY(10,0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('should multiply decimals', () => {
      const hf = buildWithPlugin([['=MULTIPLY(2.5,4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(10)
      hf.destroy()
    })
  })

  describe('DIVIDE', () => {
    it('should divide two positive numbers', () => {
      const hf = buildWithPlugin([['=DIVIDE(4,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('should divide resulting in decimal', () => {
      const hf = buildWithPlugin([['=DIVIDE(5,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2.5)
      hf.destroy()
    })

    it('should return DIV_BY_ZERO error when dividing by zero', () => {
      const hf = buildWithPlugin([['=DIVIDE(10,0)']])
      const result = hf.getCellValue(adr('A1')) as DetailedCellError
      expect(result).toBeInstanceOf(DetailedCellError)
      expect(result.type).toBe('DIV_BY_ZERO')
      hf.destroy()
    })

    it('should divide with negative numbers', () => {
      const hf = buildWithPlugin([['=DIVIDE(-10,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-5)
      hf.destroy()
    })
  })

  describe('POW', () => {
    it('should compute power', () => {
      const hf = buildWithPlugin([['=POW(2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(8)
      hf.destroy()
    })

    it('should compute square root as fractional exponent', () => {
      const hf = buildWithPlugin([['=POW(4,0.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('should handle zero exponent', () => {
      const hf = buildWithPlugin([['=POW(5,0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('should handle negative exponent', () => {
      const hf = buildWithPlugin([['=POW(2,-2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0.25)
      hf.destroy()
    })
  })

  describe('GT', () => {
    it('should return true when a > b', () => {
      const hf = buildWithPlugin([['=GT(5,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when a <= b', () => {
      const hf = buildWithPlugin([['=GT(2,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return false when a = b', () => {
      const hf = buildWithPlugin([['=GT(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should compare strings', () => {
      const hf = buildWithPlugin([['=GT("b","a")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })

  describe('GTE', () => {
    it('should return true when a >= b', () => {
      const hf = buildWithPlugin([['=GTE(5,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true when a = b', () => {
      const hf = buildWithPlugin([['=GTE(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when a < b', () => {
      const hf = buildWithPlugin([['=GTE(2,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('LT', () => {
    it('should return true when a < b', () => {
      const hf = buildWithPlugin([['=LT(2,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when a >= b', () => {
      const hf = buildWithPlugin([['=LT(5,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return false when a = b', () => {
      const hf = buildWithPlugin([['=LT(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('LTE', () => {
    it('should return true when a <= b', () => {
      const hf = buildWithPlugin([['=LTE(2,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true when a = b', () => {
      const hf = buildWithPlugin([['=LTE(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when a > b', () => {
      const hf = buildWithPlugin([['=LTE(5,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('EQ', () => {
    it('should return true when values are equal', () => {
      const hf = buildWithPlugin([['=EQ(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when values are different', () => {
      const hf = buildWithPlugin([['=EQ(5,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should compare strings', () => {
      const hf = buildWithPlugin([['=EQ("hello","hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false for different strings', () => {
      const hf = buildWithPlugin([['=EQ("hello","world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should handle boolean comparison', () => {
      const hf = buildWithPlugin([['=EQ(TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })

  describe('NE', () => {
    it('should return true when values are different', () => {
      const hf = buildWithPlugin([['=NE(5,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when values are equal', () => {
      const hf = buildWithPlugin([['=NE(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should compare strings', () => {
      const hf = buildWithPlugin([['=NE("hello","world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false for equal strings', () => {
      const hf = buildWithPlugin([['=NE("hello","hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })
  })

  describe('CONCAT', () => {
    it('should concatenate two strings', () => {
      const hf = buildWithPlugin([['=CONCAT("hello","goodbye")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hellogoodbye')
      hf.destroy()
    })

    it('should concatenate with space in between', () => {
      const hf = buildWithPlugin([['=CONCAT("hello"," world")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello world')
      hf.destroy()
    })

    it('should concatenate numbers as strings', () => {
      // In GSheets mode the comma is also the thousands separator, so numeric
      // literals must be provided via cell references to avoid ambiguity.
      const hf = buildWithPlugin([[123, 456, '=CONCAT(A1,B1)']])
      expect(hf.getCellValue(adr('C1'))).toBe('123456')
      hf.destroy()
    })

    it('should handle empty strings', () => {
      const hf = buildWithPlugin([['=CONCAT("hello","")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello')
      hf.destroy()
    })
  })

  describe('UMINUS', () => {
    it('should negate positive number', () => {
      const hf = buildWithPlugin([['=UMINUS(4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-4)
      hf.destroy()
    })

    it('should negate negative number', () => {
      const hf = buildWithPlugin([['=UMINUS(-4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })

    it('should negate zero', () => {
      const hf = buildWithPlugin([['=UMINUS(0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('should negate decimal', () => {
      const hf = buildWithPlugin([['=UMINUS(3.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-3.5)
      hf.destroy()
    })
  })

  describe('UPLUS', () => {
    it('should return positive number unchanged', () => {
      const hf = buildWithPlugin([['=UPLUS(4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })

    it('should return negative number unchanged', () => {
      const hf = buildWithPlugin([['=UPLUS(-4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-4)
      hf.destroy()
    })

    it('should return zero', () => {
      const hf = buildWithPlugin([['=UPLUS(0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('should return decimal unchanged', () => {
      const hf = buildWithPlugin([['=UPLUS(3.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(3.5)
      hf.destroy()
    })
  })

  describe('UNARY_PERCENT', () => {
    it('should convert to percentage decimal', () => {
      const hf = buildWithPlugin([['=UNARY_PERCENT(50)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0.5)
      hf.destroy()
    })

    it('should convert decimal percentage', () => {
      const hf = buildWithPlugin([['=UNARY_PERCENT(0.5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0.005)
      hf.destroy()
    })

    it('should handle zero', () => {
      const hf = buildWithPlugin([['=UNARY_PERCENT(0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('should handle negative percentage', () => {
      const hf = buildWithPlugin([['=UNARY_PERCENT(-50)']])
      expect(hf.getCellValue(adr('A1'))).toBe(-0.5)
      hf.destroy()
    })
  })

  describe('ISBETWEEN', () => {
    it('should return true when value is within inclusive range', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,4.5,7.9,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true when value equals lower bound (inclusive)', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,5,10,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true when value equals upper bound (inclusive)', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(10,5,10,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when value is below range', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(3,5,10,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return false when value is above range', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(15,5,10,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should handle exclusive lower bound', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,5,10,FALSE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should handle exclusive upper bound', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(10,5,10,TRUE,FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should handle both bounds exclusive', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(7.5,5,10,FALSE,FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should use default inclusive bounds when omitted', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,5,10)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should handle decimal values', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5.5,5,6,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true for value just above exclusive lower bound', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5.01,5,10,FALSE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true for value just below exclusive upper bound', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(9.99,5,10,TRUE,FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })
})
