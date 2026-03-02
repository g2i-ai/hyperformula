import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

const GS_CONFIG = {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' as const}

function buildGS(data: (string | number | boolean | null)[][]): HyperFormula {
  return HyperFormula.buildFromArray(data, GS_CONFIG)
}

describe('GoogleSheetsStatisticalPlugin', () => {

  // ---------------------------------------------------------------------------
  // PERMUT
  // ---------------------------------------------------------------------------
  describe('PERMUT', () => {
    it('computes P(4,2) = 12', () => {
      const hf = buildGS([['=PERMUT(4,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(12)
      hf.destroy()
    })

    it('computes P(5,0) = 1', () => {
      const hf = buildGS([['=PERMUT(5,0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('computes P(5,5) = 120', () => {
      const hf = buildGS([['=PERMUT(5,5)']])
      expect(hf.getCellValue(adr('A1'))).toBe(120)
      hf.destroy()
    })

    it('returns NUM error when k > n', () => {
      const hf = buildGS([['=PERMUT(2,4)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // PERMUTATIONA
  // ---------------------------------------------------------------------------
  describe('PERMUTATIONA', () => {
    it('computes PA(4,2) = 16', () => {
      const hf = buildGS([['=PERMUTATIONA(4,2)']])
      expect(hf.getCellValue(adr('A1'))).toBe(16)
      hf.destroy()
    })

    it('computes PA(3,3) = 27', () => {
      const hf = buildGS([['=PERMUTATIONA(3,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(27)
      hf.destroy()
    })

    it('computes PA(5,0) = 1', () => {
      const hf = buildGS([['=PERMUTATIONA(5,0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // KURT
  // ---------------------------------------------------------------------------
  describe('KURT', () => {
    it('returns a finite number for 5 values', () => {
      const hf = buildGS([['=KURT(1,2,3,4,5)']])
      const val = hf.getCellValue(adr('A1'))
      expect(typeof val).toBe('number')
      expect(Number.isFinite(val as number)).toBe(true)
      hf.destroy()
    })

    it('returns DIV0 error for fewer than 4 values', () => {
      const hf = buildGS([['=KURT(1,2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns DIV0 error for identical values (zero stddev)', () => {
      const hf = buildGS([['=KURT(2,2,2,2)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // TRIMMEAN
  // ---------------------------------------------------------------------------
  describe('TRIMMEAN', () => {
    it('trims 10% from each end of [1,2,3,4,5,6,7,8,9,10]', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5, 6, 7, 8, 9, 10], ['=TRIMMEAN(A1:J1, 0.1)']],
        GS_CONFIG
      )
      // Trim 1 from each end: mean of [2,3,4,5,6,7,8,9] = 44/8 = 5.5
      expect(hf.getCellValue(adr('A2'))).toBe(5.5)
      hf.destroy()
    })

    it('returns full mean when percent=0', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=TRIMMEAN(A1:E1, 0)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(3)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // QUARTILE / QUARTILE.INC / QUARTILE.EXC
  // ---------------------------------------------------------------------------
  describe('QUARTILE', () => {
    it('returns min for quart=0', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4], ['=QUARTILE(A1:D1, 0)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(1)
      hf.destroy()
    })

    it('returns median for quart=2', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4], ['=QUARTILE(A1:D1, 2)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(2.5)
      hf.destroy()
    })

    it('returns max for quart=4', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4], ['=QUARTILE(A1:D1, 4)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(4)
      hf.destroy()
    })

    it('QUARTILE.INC matches QUARTILE for Q1', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=QUARTILE(A1:E1, 1)', '=QUARTILE.INC(A1:E1, 1)']],
        GS_CONFIG
      )
      const q1 = hf.getCellValue(adr('A2'))
      const q1inc = hf.getCellValue(adr('B2'))
      expect(q1).toBe(q1inc)
      hf.destroy()
    })

    it('QUARTILE.EXC Q1 of [1,2,3,4,5] is 1.5', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=QUARTILE.EXC(A1:E1, 1)']],
        GS_CONFIG
      )
      // EXC: position = (1/4) * (5+1) - 1 = 0.5, interpolate between sorted[0]=1 and sorted[1]=2 → 1.5
      expect(hf.getCellValue(adr('A2'))).toBe(1.5)
      hf.destroy()
    })

    it('QUARTILE.EXC returns NUM error for quart=0', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=QUARTILE.EXC(A1:E1, 0)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // MODE / MODE.SNGL / MODE.MULT
  // ---------------------------------------------------------------------------
  describe('MODE', () => {
    it('returns the most frequent value', () => {
      const hf = buildGS([['=MODE(1,2,2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('returns NA error when all values are unique', () => {
      const hf = buildGS([['=MODE(1,2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('MODE.SNGL behaves identically to MODE', () => {
      const hf = buildGS([['=MODE.SNGL(1,2,2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('MODE.MULT returns all modes as a vertical array', () => {
      const hf = HyperFormula.buildFromArray(
        [['=MODE.MULT(1,1,2,2,3)'], [null]],
        GS_CONFIG
      )
      // Both 1 and 2 appear twice — modes are [1, 2]
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      expect(hf.getCellValue(adr('A2'))).toBe(2)
      hf.destroy()
    })

    it('MODE.MULT returns NA when no duplicates', () => {
      const hf = buildGS([['=MODE.MULT(1,2,3)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // FORECAST / FORECAST.LINEAR
  // ---------------------------------------------------------------------------
  describe('FORECAST / FORECAST.LINEAR', () => {
    it('predicts y for x=5 given linear data', () => {
      // y = 2x + 1: known_y=[3,5,7], known_x=[1,2,3], predict x=5 → y=11
      const hf = HyperFormula.buildFromArray(
        [
          [3, 5, 7],
          [1, 2, 3],
          ['=FORECAST(5, A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(11, 8)
      hf.destroy()
    })

    it('FORECAST.LINEAR gives same result as FORECAST', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [3, 5, 7],
          [1, 2, 3],
          ['=FORECAST(5, A1:C1, A2:C2)', '=FORECAST.LINEAR(5, A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBe(hf.getCellValue(adr('B3')))
      hf.destroy()
    })

    it('returns NA error when array lengths differ', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [3, 5],
          [1, 2, 3],
          ['=FORECAST(5, A1:B1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('preserves pair alignment when known_y/known_x contain non-numeric cells', () => {
      // known_y=[3, "text", 7], known_x=[1, 2, "text"] — only indices 0 match:
      // Row 0: y=3, x=1 — numeric pair
      // Row 1: y="text", x=2 — y is non-numeric, skip
      // Row 2: y=7, x="text" — x is non-numeric, skip
      // With only 1 numeric pair, regression requires ≥2, so NA expected.
      const hf = HyperFormula.buildFromArray(
        [
          [3, 'text', 7],
          [1, 2, 'text'],
          ['=FORECAST(5, A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('uses correct numeric pairs when ranges have non-numeric cells at different positions', () => {
      // known_y=[3, "skip", 5, 7], known_x=[1, 3, "skip", 4] — indices 0 and 3 have valid pairs:
      // (y=3, x=1), (y=7, x=4)
      // Regression on (1,3) and (4,7): slope=(7-3)/(4-1)=4/3, intercept=3-4/3=5/3
      // At x=7: y = 4/3*7 + 5/3 = 28/3 + 5/3 = 11
      const hf = HyperFormula.buildFromArray(
        [
          [3, 'skip', 5, 7],
          [1, 3, 'skip', 4],
          ['=FORECAST(7, A1:D1, A2:D2)'],
        ],
        GS_CONFIG
      )
      // With correct pairing: (y=3,x=1) and (y=7,x=4) → at x=7, y=11
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(11, 5)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // INTERCEPT
  // ---------------------------------------------------------------------------
  describe('INTERCEPT', () => {
    it('computes y-intercept of y=2x+1', () => {
      // y=[3,5,7], x=[1,2,3] → intercept=1
      const hf = HyperFormula.buildFromArray(
        [
          [3, 5, 7],
          [1, 2, 3],
          ['=INTERCEPT(A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(1, 8)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // PERCENTRANK / PERCENTRANK.INC / PERCENTRANK.EXC
  // ---------------------------------------------------------------------------
  describe('PERCENTRANK', () => {
    it('returns 0 for the minimum value', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK(A1:E1, 1)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(0)
      hf.destroy()
    })

    it('returns 1 for the maximum value', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK(A1:E1, 5)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(1)
      hf.destroy()
    })

    it('returns 0.5 for the median of [1,2,3,4,5]', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK(A1:E1, 3)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(0.5)
      hf.destroy()
    })

    it('returns NA when x is outside data range', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3], ['=PERCENTRANK(A1:C1, 10)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('PERCENTRANK.INC matches PERCENTRANK', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK(A1:E1, 2)', '=PERCENTRANK.INC(A1:E1, 2)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(hf.getCellValue(adr('B2')))
      hf.destroy()
    })

    it('PERCENTRANK.EXC excludes 0 and 1', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK.EXC(A1:E1, 3)']],
        GS_CONFIG
      )
      // rank=2, (2+1)/(5+1) = 0.5
      expect(hf.getCellValue(adr('A2'))).toBeCloseTo(0.5, 3)
      hf.destroy()
    })

    it('PERCENTRANK.EXC returns NA for boundary values', () => {
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=PERCENTRANK.EXC(A1:E1, 1)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('PERCENTRANK.INC returns 0 for single-element dataset', () => {
      // Google Sheets: PERCENTRANK.INC({5}, 5) = 0, not NaN
      const hf = HyperFormula.buildFromArray(
        [[5], ['=PERCENTRANK.INC(A1:A1, 5)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(0)
      hf.destroy()
    })

    it('PERCENTRANK returns 0 for single-element dataset', () => {
      const hf = HyperFormula.buildFromArray(
        [[5], ['=PERCENTRANK(A1:A1, 5)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBe(0)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // PROB
  // ---------------------------------------------------------------------------
  describe('PROB', () => {
    it('sums probabilities between lower and upper limits', () => {
      // x=[1,2,3,4], prob=[0.1,0.4,0.4,0.1], lower=2, upper=3 → 0.8
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3, 4],
          [0.1, 0.4, 0.4, 0.1],
          ['=PROB(A1:D1, A2:D2, 2, 3)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(0.8, 10)
      hf.destroy()
    })

    it('returns single probability when upper_limit omitted', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3, 4],
          [0.1, 0.4, 0.4, 0.1],
          ['=PROB(A1:D1, A2:D2, 2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(0.4, 10)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // AVERAGE.WEIGHTED
  // ---------------------------------------------------------------------------
  describe('AVERAGE.WEIGHTED', () => {
    it('computes weighted average', () => {
      // (1*2 + 2*3 + 3*4) / (2+3+4) = (2+6+12)/9 = 20/9
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3],
          [2, 3, 4],
          ['=AVERAGE.WEIGHTED(A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeCloseTo(20 / 9, 8)
      hf.destroy()
    })

    it('returns DIV0 when all weights are zero', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3],
          [0, 0, 0],
          ['=AVERAGE.WEIGHTED(A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns VALUE error when any weight is negative', () => {
      // Google Sheets: "Weights cannot be negative" → #VALUE!
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3],
          [2, -1, 4],
          ['=AVERAGE.WEIGHTED(A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns VALUE error when all weights are negative', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [1, 2, 3],
          [-2, -3, -4],
          ['=AVERAGE.WEIGHTED(A1:C1, A2:C2)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // MARGINOFERROR
  // ---------------------------------------------------------------------------
  describe('MARGINOFERROR', () => {
    it('computes margin of error', () => {
      // data=[1,2,3,4,5], mean=3, stddev=√(10/4)=√2.5, moe = 0.95 * √2.5 / √5
      const hf = HyperFormula.buildFromArray(
        [[1, 2, 3, 4, 5], ['=MARGINOFERROR(A1:E1, 0.95)']],
        GS_CONFIG
      )
      const val = hf.getCellValue(adr('A2'))
      expect(typeof val).toBe('number')
      expect(val as number).toBeGreaterThan(0)
      hf.destroy()
    })

    it('returns DIV0 for single-element range', () => {
      const hf = HyperFormula.buildFromArray(
        [[5], ['=MARGINOFERROR(A1:A1, 0.95)']],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // ERF.PRECISE / ERFC.PRECISE
  // ---------------------------------------------------------------------------
  describe('ERF.PRECISE', () => {
    it('ERF.PRECISE(0) ≈ 0', () => {
      const hf = buildGS([['=ERF.PRECISE(0)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0, 6)
      hf.destroy()
    })

    it('ERF.PRECISE(1) ≈ 0.8427', () => {
      const hf = buildGS([['=ERF.PRECISE(1)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.8427, 3)
      hf.destroy()
    })

    it('ERF.PRECISE(-1) ≈ -0.8427 (odd function)', () => {
      const hf = buildGS([['=ERF.PRECISE(-1)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(-0.8427, 3)
      hf.destroy()
    })
  })

  describe('ERFC.PRECISE', () => {
    it('ERFC.PRECISE(0) ≈ 1', () => {
      const hf = buildGS([['=ERFC.PRECISE(0)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1, 6)
      hf.destroy()
    })

    it('ERFC.PRECISE(1) ≈ 0.1573', () => {
      const hf = buildGS([['=ERFC.PRECISE(1)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.1573, 3)
      hf.destroy()
    })

    it('ERF.PRECISE(x) + ERFC.PRECISE(x) = 1', () => {
      const hf = buildGS([['=ERF.PRECISE(2)+ERFC.PRECISE(2)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1, 8)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // ERF.PRECISE precision — must match the high-precision jstat implementation
  // (max error < 1e-14), not the Abramowitz & Stegun approximation (max ~1.5e-7)
  // ---------------------------------------------------------------------------
  describe('ERF.PRECISE high-precision requirements', () => {
    // Known high-precision values from Wolfram Alpha / mathematical tables
    it('ERF.PRECISE(0.5) matches high-precision value to 10 decimal places', () => {
      // erf(0.5) = 0.5204998778130465...
      const hf = buildGS([['=ERF.PRECISE(0.5)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.5204998778130465, 10)
      hf.destroy()
    })

    it('ERF.PRECISE(1.5) matches high-precision value to 10 decimal places', () => {
      // erf(1.5) = 0.9661051464753108...
      const hf = buildGS([['=ERF.PRECISE(1.5)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.9661051464753108, 10)
      hf.destroy()
    })

    it('ERFC.PRECISE(0.5) matches high-precision value to 10 decimal places', () => {
      // erfc(0.5) = 0.4795001221869535...
      const hf = buildGS([['=ERFC.PRECISE(0.5)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.4795001221869535, 10)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // AVERAGEIFS
  // ---------------------------------------------------------------------------
  describe('AVERAGEIFS', () => {
    it('averages cells where criterion is met', () => {
      // avg_range=[10,20,30], crit_range=[1,2,1], criterion=1 → average(10,30) = 20
      const hf = HyperFormula.buildFromArray(
        [
          [10, 20, 30],
          [1, 2, 1],
          ['=AVERAGEIFS(A1:C1, A2:C2, 1)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBe(20)
      hf.destroy()
    })

    it('returns DIV0 when no cells match', () => {
      const hf = HyperFormula.buildFromArray(
        [
          [10, 20, 30],
          [1, 2, 1],
          ['=AVERAGEIFS(A1:C1, A2:C2, 99)'],
        ],
        GS_CONFIG
      )
      expect(hf.getCellValue(adr('A3'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // ---------------------------------------------------------------------------
  // Ensure functions are NOT available in default mode
  // ---------------------------------------------------------------------------
  describe('functions not registered in default mode', () => {
    it('PERMUTATIONA is not available in default mode', () => {
      const hf = HyperFormula.buildFromArray(
        [['=PERMUTATIONA(4,2)']],
        {licenseKey: 'gpl-v3'}
      )
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })
})
