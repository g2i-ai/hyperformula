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

    // Epsilon rounding: when the result is near-zero relative to the operands,
    // it should be rounded to 0 (e.g. floating point cancellation artifacts)
    it('should apply epsilon rounding to near-zero sums', () => {
      // In raw JS, 0.1 + 0.2 - 0.3 = 5.551115123125783e-17, not 0.
      // With epsilon rounding, ADD(0.1+0.2, -0.3) rounds this to 0.
      const hf = buildWithPlugin([['=ADD(ADD(0.1,0.2),-0.3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
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

    // Epsilon rounding: subtracting nearly equal numbers should give 0
    it('should apply epsilon rounding when result is near zero', () => {
      // 1 - (1 + 1e-16) should be 0 after epsilon rounding
      const hf = buildWithPlugin([['=MINUS(1,1)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
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

    it('should compare strings using locale-aware collation', () => {
      const hf = buildWithPlugin([['=GT("b","a")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    // Cross-type ordering: numbers < strings (by cell type ordinal)
    it('should order number less than string in cross-type comparison', () => {
      const hf = buildWithPlugin([['=GT("text",1)']])
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

    // Cross-type ordering: string > number
    it('should return true when string is compared to number', () => {
      const hf = buildWithPlugin([['=GTE("text",1)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
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

    // Cross-type: number < string
    it('should return true when number is less than string', () => {
      const hf = buildWithPlugin([['=LT(1,"text")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
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

    // Number vs string should never be equal
    it('should return false for number vs string', () => {
      const hf = buildWithPlugin([['=EQ(1,"1")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    // Case-insensitive by default (Google Sheets EQ is case-insensitive)
    it('should compare strings case-insensitively by default', () => {
      const hf = buildWithPlugin([['=EQ("HELLO","hello")']])
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

    // Number vs string: always not equal due to type mismatch
    it('should return true for number vs string', () => {
      const hf = buildWithPlugin([['=NE(1,"1")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    // Case-insensitive: "HELLO" == "hello" so NE returns false
    it('should treat equal-collation strings as equal', () => {
      const hf = buildWithPlugin([['=NE("HELLO","hello")']])
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
      const hf = buildWithPlugin([['=CONCAT(123,456)']])
      expect(hf.getCellValue(adr('A1'))).toBe('123456')
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

    // Google Sheets returns #NUM! error when lower_value > upper_value
    it('should return NUM error when lower bound exceeds upper bound', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,10,1)']])
      const result = hf.getCellValue(adr('A1')) as DetailedCellError
      expect(result).toBeInstanceOf(DetailedCellError)
      expect(result.type).toBe('NUM')
      hf.destroy()
    })

    it('should return NUM error when lower equals upper and exclusive on one side', () => {
      const hf = buildWithPlugin([['=ISBETWEEN(5,10,5)']])
      const result = hf.getCellValue(adr('A1')) as DetailedCellError
      expect(result).toBeInstanceOf(DetailedCellError)
      expect(result.type).toBe('NUM')
      hf.destroy()
    })

    // Epsilon-awareness: ISBETWEEN must use floatCmp (not raw >/>=/</<= operators)
    // so that floating-point boundary values are treated consistently with LTE/GTE.
    it('should return true when value is epsilon-equal to upper bound (inclusive)', () => {
      // 0.1 + 0.2 = 0.30000000000000004 in IEEE 754, which is epsilon-equal to 0.3.
      // Raw `val <= hi` would be false, but floatCmp treats them as equal.
      const hf = buildWithPlugin([['=ISBETWEEN(ADD(0.1,0.2),0,0.3,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return true when value is epsilon-equal to lower bound (inclusive)', () => {
      // Symmetric: lo = 0.1+0.2 = 0.30000000000000004, val = 0.3.
      // Raw `val >= lo` would be false, but floatCmp treats them as equal.
      const hf = buildWithPlugin([['=ISBETWEEN(0.3,ADD(0.1,0.2),1,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('should return false when value is epsilon-equal to exclusive upper bound', () => {
      // When hiInc=FALSE, a value epsilon-equal to hi should NOT be inside the range.
      const hf = buildWithPlugin([['=ISBETWEEN(ADD(0.1,0.2),0,0.3,TRUE,FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should return false when value is epsilon-equal to exclusive lower bound', () => {
      // When loInc=FALSE, a value epsilon-equal to lo should NOT be inside the range.
      const hf = buildWithPlugin([['=ISBETWEEN(0.3,ADD(0.1,0.2),1,FALSE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('should not return NUM error when lo and hi are epsilon-equal (lo <= hi)', () => {
      // floatCmp(lo, hi) === 0 when they are epsilon-equal, so lo > hi check should not fire.
      const hf = buildWithPlugin([['=ISBETWEEN(0.3,ADD(0.1,0.2),0.3,TRUE,TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })
  })
})
