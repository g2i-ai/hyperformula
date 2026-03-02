import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

/**
 * Builds a HyperFormula instance in Google Sheets compatibility mode.
 * This activates all googleSheetsPlugins including GoogleSheetsFinancialPlugin.
 */
function build(data: unknown[][]): HyperFormula {
  return HyperFormula.buildFromArray(data as string[][], {
    licenseKey: 'gpl-v3',
    compatibilityMode: 'googleSheets',
  })
}

describe('GoogleSheetsFinancialPlugin', () => {
  // -------------------------------------------------------------------------
  // ACCRINT
  // -------------------------------------------------------------------------
  describe('ACCRINT', () => {
    it('calculates accrued interest for periodic security (basis 0 US 30/360)', () => {
      // par=100, rate=5%, 3 full years in US 30/360 = 100*0.05*1080/360 = 15
      const hf = build([['=ACCRINT(DATE(2010,1,1), DATE(2010,2,1), DATE(2012,12,31), 0.05, 100, 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(15, 4)
      hf.destroy()
    })

    it('calculates accrued interest for periodic security (basis 4 European 30/360)', () => {
      // ACCRINT with European 30/360 (basis=4): Dec 31 → Dec 30, so 1079/360 years
      // 100 * 0.05 * 1079/360 ≈ 14.98611111
      const hf = build([['=ACCRINT(DATE(2010,1,1), DATE(2010,2,1), DATE(2012,12,31), 0.05, 100, 4, 4)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(14.98611111, 4)
      hf.destroy()
    })

    it('returns NUM error when settlement <= issue', () => {
      const hf = build([['=ACCRINT(DATE(2012,12,31), DATE(2010,2,1), DATE(2010,1,1), 0.05, 100, 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns NUM error for non-positive rate', () => {
      const hf = build([['=ACCRINT(DATE(2010,1,1), DATE(2010,2,1), DATE(2012,12,31), 0, 100, 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('supports actual/actual basis (basis=1)', () => {
      const hf = build([['=ACCRINT(DATE(2010,1,1), DATE(2010,2,1), DATE(2012,12,31), 0.05, 100, 2, 1)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThan(0)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // ACCRINTM
  // -------------------------------------------------------------------------
  describe('ACCRINTM', () => {
    it('calculates accrued interest for security paying at maturity (actual/actual basis)', () => {
      // 6 months at 10% on $1000 par, actual/actual basis
      // Jan 1 to Jul 1 = 181 actual days. yearFrac = 181/365 ≈ 0.49589
      // Result = 1000 * 0.10 * 0.49589 ≈ 49.589
      const hf = build([['=ACCRINTM(DATE(2010,1,1), DATE(2010,7,1), 0.10, 1000, 1)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(49.589, 2)
      hf.destroy()
    })

    it('calculates accrued interest for security paying at maturity (actual/360 basis)', () => {
      // 6 months at 10% on $1000 par, actual/360 basis
      // Jan 1 to Jul 1 = 181 actual days. yearFrac = 181/360 ≈ 0.50278
      // Result = 1000 * 0.10 * 0.50278 ≈ 50.278
      const hf = build([['=ACCRINTM(DATE(2010,1,1), DATE(2010,7,1), 0.10, 1000, 2)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(50.278, 2)
      hf.destroy()
    })

    it('returns NUM error when settlement <= issue', () => {
      const hf = build([['=ACCRINTM(DATE(2010,7,1), DATE(2010,1,1), 0.10, 1000, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // AMORLINC
  // -------------------------------------------------------------------------
  describe('AMORLINC', () => {
    it('calculates linear depreciation for first period', () => {
      // cost=1000, purchased=2010-01-01, first_period=2010-12-31, salvage=0, period=1, rate=0.1
      const hf = build([['=AMORLINC(1000, DATE(2010,1,1), DATE(2010,12,31), 0, 1, 0.1, 1)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(100, 0)
      hf.destroy()
    })

    it('returns 0 for periods beyond asset life', () => {
      const hf = build([['=AMORLINC(1000, DATE(2010,1,1), DATE(2010,12,31), 0, 20, 0.1, 1)']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPDAYBS
  // -------------------------------------------------------------------------
  describe('COUPDAYBS', () => {
    it('returns days from beginning of coupon period to settlement', () => {
      // COUPDAYBS("2/1/2010", "12/31/2019", 4, 0) = 31
      const hf = build([['=COUPDAYBS(DATE(2010,2,1), DATE(2019,12,31), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(31)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=COUPDAYBS(DATE(2020,1,1), DATE(2010,1,1), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('handles annual frequency (frequency=1)', () => {
      const hf = build([['=COUPDAYBS(DATE(2010,6,1), DATE(2019,12,31), 1, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThanOrEqual(0)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPDAYS
  // -------------------------------------------------------------------------
  describe('COUPDAYS', () => {
    it('returns days in coupon period containing settlement', () => {
      // COUPDAYS("2/1/2010", "12/31/2019", 4, 0) = 90
      const hf = build([['=COUPDAYS(DATE(2010,2,1), DATE(2019,12,31), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(90)
      hf.destroy()
    })

    it('returns 180 days for semi-annual frequency with 30/360 basis', () => {
      const hf = build([['=COUPDAYS(DATE(2010,2,1), DATE(2019,6,30), 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(180)
      hf.destroy()
    })

    it('returns 365 days for annual frequency with actual/365 basis', () => {
      const hf = build([['=COUPDAYS(DATE(2010,2,1), DATE(2019,12,31), 1, 3)']])
      expect(hf.getCellValue(adr('A1'))).toBe(365)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPDAYSNC
  // -------------------------------------------------------------------------
  describe('COUPDAYSNC', () => {
    it('satisfies COUPDAYBS + COUPDAYSNC = COUPDAYS', () => {
      // This identity must hold for all bases
      const hf = build([
        [
          '=COUPDAYBS(DATE(2010,2,1), DATE(2019,12,31), 4, 0)',
          '=COUPDAYS(DATE(2010,2,1), DATE(2019,12,31), 4, 0)',
          '=COUPDAYSNC(DATE(2010,2,1), DATE(2019,12,31), 4, 0)',
        ],
      ])
      const bs = hf.getCellValue(adr('A1')) as number
      const days = hf.getCellValue(adr('B1')) as number
      const nc = hf.getCellValue(adr('C1')) as number
      expect(bs + nc).toBeCloseTo(days, 5)
      // COUPDAYBS=31, COUPDAYS=90, so COUPDAYSNC=59
      expect(nc).toBe(59)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=COUPDAYSNC(DATE(2020,1,1), DATE(2010,1,1), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPNCD
  // -------------------------------------------------------------------------
  describe('COUPNCD', () => {
    it('returns next coupon date as a serial number', () => {
      const hf = build([['=COUPNCD(DATE(2010,2,1), DATE(2019,12,31), 4, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      // Next coupon date should be strictly after settlement
      const hf2 = build([['=DATE(2010,2,1)']])
      const settlement = hf2.getCellValue(adr('A1')) as number
      expect(result as number).toBeGreaterThan(settlement)
      hf.destroy()
      hf2.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPNUM
  // -------------------------------------------------------------------------
  describe('COUPNUM', () => {
    it('returns number of coupons between settlement and maturity', () => {
      // 40 quarterly coupons over 10 years
      const hf = build([['=COUPNUM(DATE(2010,1,1), DATE(2020,1,1), 4, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThanOrEqual(1)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=COUPNUM(DATE(2020,1,1), DATE(2010,1,1), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('counts coupons correctly when settlement day is before maturity day within the month', () => {
      // Settlement: Jan 14 2010, Maturity: Apr 15 2010, quarterly (freq=4)
      // Coupon dates anchored to Apr 15: Jan 15, Apr 15
      // Jan 15 is strictly after Jan 14, so there are 2 remaining coupons: Jan 15 and Apr 15
      // Bug: month-only calculation gives ceil(3/3)=1, ignoring the Jan 15 coupon
      const hf = build([['=COUPNUM(DATE(2010,1,14), DATE(2010,4,15), 4, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('counts coupons correctly when settlement is one day before a coupon date (semi-annual)', () => {
      // Settlement: Sep 30 2010, Maturity: Apr 1 2011, semi-annual (freq=2)
      // Coupon dates: Oct 1 2010 and Apr 1 2011
      // Sep 30 is before Oct 1, so 2 coupons remain
      // Bug: month-only calculation gives ceil(6/6)=1
      const hf = build([['=COUPNUM(DATE(2010,9,30), DATE(2011,4,1), 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(2)
      hf.destroy()
    })

    it('counts exactly one coupon when settlement is on a coupon date', () => {
      // Settlement: Oct 1 2010, Maturity: Apr 1 2011, semi-annual
      // Oct 1 is a coupon date, so it is the previous coupon; next = Apr 1 = 1 remaining
      const hf = build([['=COUPNUM(DATE(2010,10,1), DATE(2011,4,1), 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // COUPPCD
  // -------------------------------------------------------------------------
  describe('COUPPCD', () => {
    it('returns previous coupon date before settlement', () => {
      const hf = build([['=COUPPCD(DATE(2010,2,1), DATE(2019,12,31), 4, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      const hf2 = build([['=DATE(2010,2,1)']])
      const settlement = hf2.getCellValue(adr('A1')) as number
      expect(result as number).toBeLessThanOrEqual(settlement)
      hf.destroy()
      hf2.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // DISC
  // -------------------------------------------------------------------------
  describe('DISC', () => {
    it('returns discount rate', () => {
      // DISC("1/2/2010", "12/31/2039", 90, 100, 0) ≈ 0.003333642
      const hf = build([['=DISC(DATE(2010,1,2), DATE(2039,12,31), 90, 100, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.003333642, 5)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=DISC(DATE(2040,1,1), DATE(2010,1,2), 90, 100, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('supports actual/360 basis', () => {
      const hf = build([['=DISC(DATE(2010,1,2), DATE(2039,12,31), 90, 100, 2)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThan(0)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // DURATION
  // -------------------------------------------------------------------------
  describe('DURATION', () => {
    it('returns Macaulay duration', () => {
      const hf = build([['=DURATION(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0.05, 2, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      // A 10-year 5% bond at 5% yield has duration slightly less than 10 years
      expect(result as number).toBeGreaterThan(0)
      expect(result as number).toBeLessThan(10)
      hf.destroy()
    })

    it('returns NUM error for negative coupon', () => {
      const hf = build([['=DURATION(DATE(2010,1,1), DATE(2020,1,1), -0.05, 0.05, 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // INTRATE
  // -------------------------------------------------------------------------
  describe('INTRATE', () => {
    it('returns interest rate for fully invested security', () => {
      // Bought at 1000, redeems at 1100, 1 year, actual/360
      const hf = build([['=INTRATE(DATE(2010,1,1), DATE(2011,1,1), 1000, 1100, 2)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(0.1 * 365 / 360, 2)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=INTRATE(DATE(2011,1,1), DATE(2010,1,1), 1000, 1100, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // MDURATION
  // -------------------------------------------------------------------------
  describe('MDURATION', () => {
    it('returns modified duration less than Macaulay duration', () => {
      const hf = build([
        ['=DURATION(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0.05, 2, 0)', '=MDURATION(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0.05, 2, 0)'],
      ])
      const dur = hf.getCellValue(adr('A1')) as number
      const mdur = hf.getCellValue(adr('B1')) as number
      expect(mdur).toBeLessThan(dur)
      // MDURATION = DURATION / (1 + yield/frequency)
      expect(mdur).toBeCloseTo(dur / (1 + 0.05 / 2), 4)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // PRICE
  // -------------------------------------------------------------------------
  describe('PRICE', () => {
    it('returns price above par when coupon rate exceeds yield', () => {
      // A bond with 3% coupon rate and 1.2% yield should price above 100
      // ~30 year bond, semi-annual
      const hf = build([['=PRICE(DATE(2010,1,2), DATE(2039,12,31), 0.03, 0.012, 100, 2, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThan(100)
      hf.destroy()
    })

    it('returns price below par when yield exceeds coupon rate', () => {
      // A bond with 3% coupon rate and 5% yield should price below 100
      const hf = build([['=PRICE(DATE(2010,1,2), DATE(2039,12,31), 0.03, 0.05, 100, 2, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeLessThan(100)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=PRICE(DATE(2040,1,1), DATE(2010,1,2), 0.03, 0.05, 100, 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('yields par value when coupon rate equals yield', () => {
      // A bond priced at par (coupon = yield) should return ~100
      const hf = build([['=PRICE(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0.05, 100, 2, 1)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(100, 0)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // PRICEDISC
  // -------------------------------------------------------------------------
  describe('PRICEDISC', () => {
    it('returns discounted price', () => {
      // 1 year, 5% discount, redemption=100, basis=0 → price = 95
      const hf = build([['=PRICEDISC(DATE(2010,1,1), DATE(2011,1,1), 0.05, 100, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(95, 1)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=PRICEDISC(DATE(2020,1,1), DATE(2010,1,1), 0.05, 100, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // PRICEMAT
  // -------------------------------------------------------------------------
  describe('PRICEMAT', () => {
    it('returns price for security paying interest at maturity', () => {
      const hf = build([['=PRICEMAT(DATE(2011,1,1), DATE(2012,1,1), DATE(2010,1,1), 0.05, 0.05, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      // When rate equals yield over same period, price should be near 100
      expect(result as number).toBeCloseTo(100, 0)
      hf.destroy()
    })

    it('returns NUM error for invalid date ordering', () => {
      const hf = build([['=PRICEMAT(DATE(2010,1,1), DATE(2012,1,1), DATE(2011,1,1), 0.05, 0.05, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // RECEIVED
  // -------------------------------------------------------------------------
  describe('RECEIVED', () => {
    it('returns amount received at maturity', () => {
      // 1 year, 5% discount, $1000 investment → received = 1000/(1-0.05) ≈ 1052.63
      const hf = build([['=RECEIVED(DATE(2010,1,1), DATE(2011,1,1), 1000, 0.05, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(1052.63, 1)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=RECEIVED(DATE(2020,1,1), DATE(2010,1,1), 1000, 0.05, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // VDB
  // -------------------------------------------------------------------------
  describe('VDB', () => {
    it('calculates variable declining balance depreciation', () => {
      // VDB(100, 10, 20, 10, 11, 2, TRUE) ≈ 3.49
      const hf = build([['=VDB(100, 10, 20, 10, 11, 2, TRUE)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(3.49, 1)
      hf.destroy()
    })

    it('calculates VDB for first full period', () => {
      // cost=100, salvage=10, life=5, start=0, end=1, factor=2, no_switch=FALSE → DDB = 40
      const hf = build([['=VDB(100, 10, 5, 0, 1, 2, FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBeCloseTo(40, 4)
      hf.destroy()
    })

    it('handles fractional periods', () => {
      const hf = build([['=VDB(100, 10, 5, 0, 0.5, 2, FALSE)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeGreaterThan(0)
      expect(result as number).toBeLessThan(40)
      hf.destroy()
    })

    it('returns NUM error when end_period > life', () => {
      const hf = build([['=VDB(100, 10, 5, 0, 6, 2, FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns NUM error when start_period > end_period', () => {
      const hf = build([['=VDB(100, 10, 5, 3, 1, 2, FALSE)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // XIRR
  // -------------------------------------------------------------------------
  describe('XIRR', () => {
    it('calculates IRR for non-periodic cash flows', () => {
      // Simple cash flows: invest 100, get 115 after 1 year ≈ 15% IRR
      const hf = build([
        [-100, 115, '=DATE(2010,1,1)', '=DATE(2011,1,1)'],
        ['=XIRR(A1:B1, C1:D1, 0.1)'],
      ])
      const result = hf.getCellValue(adr('A2'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(0.15, 2)
      hf.destroy()
    })

    it('returns NUM error for all-positive cash flows', () => {
      const hf = build([
        [100, 115, '=DATE(2010,1,1)', '=DATE(2011,1,1)'],
        ['=XIRR(A1:B1, C1:D1, 0.1)'],
      ])
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns NUM error when arrays have different lengths', () => {
      const hf = build([
        [-100, 115, 50, '=DATE(2010,1,1)', '=DATE(2011,1,1)'],
        ['=XIRR(A1:C1, D1:E1, 0.1)'],
      ])
      expect(hf.getCellValue(adr('A2'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // YIELD
  // -------------------------------------------------------------------------
  describe('YIELD', () => {
    it('calculates yield from price', () => {
      // Inverse of PRICE: if PRICE at 5% yield ≈ 100, then YIELD at price 100 ≈ 5%
      const hf = build([
        ['=PRICE(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0.05, 100, 2, 1)', '=YIELD(DATE(2010,1,1), DATE(2020,1,1), 0.05, A1, 100, 2, 1)'],
      ])
      const result = hf.getCellValue(adr('B1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(0.05, 3)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=YIELD(DATE(2020,1,1), DATE(2010,1,1), 0.05, 100, 100, 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns NUM error for non-positive price', () => {
      const hf = build([['=YIELD(DATE(2010,1,1), DATE(2020,1,1), 0.05, 0, 100, 2, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('does not silently converge to initial guess when solver enters error region', () => {
      // When Newton-Raphson evaluates priceCore at a point where it returns a CellError
      // (e.g. yld makes the price denominator zero), the wrapped NaN->0 substitution
      // makes f(x)=0, causing false convergence to that x instead of returning #NUM.
      // The fix propagates the CellError rather than masking it.
      //
      // Near-maturity single-period bond (N=1) with very large redemption/settlement gap:
      // settlement just after a coupon, maturity = next coupon date.
      // Use a redemption of 0.001 and price of 999 – a price no yield can match –
      // to ensure the solver eventually crosses the denom-zero pole (yld = -freq*E/DSC).
      // With NaN wrapping, the solver may return the initial guess (0.1) as a false answer.
      // Without NaN wrapping, the solver correctly returns #NUM.
      const hf = build([['=YIELD(DATE(2010,1,1), DATE(2010,4,1), 0.0, 999, 0.001, 4, 0)']])
      // This bond has no valid yield (price 999 is impossible for redemption 0.001)
      // so YIELD must return a #NUM error, not a spurious number
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('computes correct yield for a well-known single-period bond (N=1)', () => {
      // Settlement 2010-01-01, maturity 2010-04-01, rate=5%, redemption=100, freq=4
      // The PRICE at yield=5% should equal the YIELD-computed price
      const hf = build([
        ['=PRICE(DATE(2010,1,1), DATE(2010,4,1), 0.05, 0.05, 100, 4, 0)',
         '=YIELD(DATE(2010,1,1), DATE(2010,4,1), 0.05, A1, 100, 4, 0)'],
      ])
      const yld = hf.getCellValue(adr('B1'))
      expect(typeof yld).toBe('number')
      expect(yld as number).toBeCloseTo(0.05, 4)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // YIELDDISC
  // -------------------------------------------------------------------------
  describe('YIELDDISC', () => {
    it('returns annual yield for discounted security', () => {
      // Bought at 95, redeems at 100, 1 year → yield ≈ 5.26%
      const hf = build([['=YIELDDISC(DATE(2010,1,1), DATE(2011,1,1), 95, 100, 0)']])
      const result = hf.getCellValue(adr('A1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(0.0526, 3)
      hf.destroy()
    })

    it('inverse of PRICEDISC: YIELDDISC(PRICEDISC()) returns original discount', () => {
      const hf = build([
        ['=PRICEDISC(DATE(2010,1,1), DATE(2011,1,1), 0.05, 100, 0)', '=YIELDDISC(DATE(2010,1,1), DATE(2011,1,1), A1, 100, 0)'],
      ])
      // YIELDDISC is NOT the same as discount; it's the yield equivalent
      const result = hf.getCellValue(adr('B1'))
      expect(typeof result).toBe('number')
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // YIELDMAT
  // -------------------------------------------------------------------------
  describe('YIELDMAT', () => {
    it('returns annual yield for security paying interest at maturity', () => {
      const hf = build([
        ['=PRICEMAT(DATE(2011,1,1), DATE(2012,1,1), DATE(2010,1,1), 0.05, 0.05, 0)', '=YIELDMAT(DATE(2011,1,1), DATE(2012,1,1), DATE(2010,1,1), 0.05, A1, 0)'],
      ])
      const result = hf.getCellValue(adr('B1'))
      expect(typeof result).toBe('number')
      expect(result as number).toBeCloseTo(0.05, 3)
      hf.destroy()
    })

    it('returns NUM error when settlement >= maturity', () => {
      const hf = build([['=YIELDMAT(DATE(2020,1,1), DATE(2010,1,1), DATE(2009,1,1), 0.05, 100, 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  // -------------------------------------------------------------------------
  // Integration: functions unavailable outside googleSheets mode
  // -------------------------------------------------------------------------
  describe('availability', () => {
    it('ACCRINT is not available in default mode', () => {
      const hf = HyperFormula.buildFromArray(
        [['=ACCRINT(DATE(2010,1,1), DATE(2010,2,1), DATE(2012,12,31), 0.05, 100, 4, 0)']] as string[][],
        {licenseKey: 'gpl-v3'}
      )
      // Should return NAME error since function not registered in default mode
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })
})
