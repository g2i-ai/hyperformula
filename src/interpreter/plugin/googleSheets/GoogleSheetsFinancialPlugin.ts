/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {getRawValue, InterpreterValue, isExtendedNumber, NumberType} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'
import {SimpleDate, offsetMonth, toBasisEU} from '../../../DateTimeHelper'

/**
 * Day-count basis identifiers as specified by Excel/Google Sheets.
 */
const enum Basis {
  US_30_360 = 0,
  ACTUAL_ACTUAL = 1,
  ACTUAL_360 = 2,
  ACTUAL_365 = 3,
  EUROPEAN_30_360 = 4,
}

/**
 * Valid coupon frequencies (payments per year).
 */
type CouponFrequency = 1 | 2 | 4

/**
 * Google Sheets-compatible financial function implementations.
 *
 * Implements bond, coupon, yield, discount, and depreciation functions
 * that are not present in the base FinancialPlugin.
 */
export class GoogleSheetsFinancialPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsFinancialPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'ACCRINT': {
      method: 'accrint',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'ACCRINTM': {
      method: 'accrintm',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'AMORLINC': {
      method: 'amorlinc',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'COUPDAYBS': {
      method: 'coupdaybs',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'COUPDAYS': {
      method: 'coupdays',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'COUPDAYSNC': {
      method: 'coupdaysnc',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'COUPNCD': {
      method: 'coupncd',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_DATE,
    },
    'COUPNUM': {
      method: 'coupnum',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'COUPPCD': {
      method: 'couppcd',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_DATE,
    },
    'DISC': {
      method: 'disc',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'DURATION': {
      method: 'duration',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'INTRATE': {
      method: 'intrate',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'MDURATION': {
      method: 'mduration',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
    },
    'PRICE': {
      method: 'price',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'PRICEDISC': {
      method: 'pricedisc',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'PRICEMAT': {
      method: 'pricemat',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'RECEIVED': {
      method: 'received',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'VDB': {
      method: 'vdb',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0, defaultValue: 2},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false},
      ],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'XIRR': {
      method: 'xirr',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 0.1},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'YIELD': {
      method: 'yield',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 4},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'YIELDDISC': {
      method: 'yielddisc',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'YIELDMAT': {
      method: 'yieldmat',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4, defaultValue: 0},
      ],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
  }

  // ---------------------------------------------------------------------------
  // Public function methods (one per implemented function)
  // ---------------------------------------------------------------------------

  /**
   * ACCRINT(issue, first_interest, settlement, rate, par, frequency, [basis])
   *
   * Returns the accrued interest for a security that pays periodic interest.
   */
  public accrint(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ACCRINT'),
      (issue: number, firstInterest: number, settlement: number, rate: number, par: number, frequency: number, basis: number) => {
        if (settlement <= issue || rate <= 0 || par <= 0 || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(issue, settlement, basis)
        if (yearFrac instanceof CellError) return yearFrac
        return par * rate * yearFrac
      }
    )
  }

  /**
   * ACCRINTM(issue, settlement, rate, par, [basis])
   *
   * Returns the accrued interest for a security that pays interest at maturity.
   */
  public accrintm(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ACCRINTM'),
      (issue: number, settlement: number, rate: number, par: number, basis: number) => {
        if (settlement <= issue || rate <= 0 || par <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(issue, settlement, basis)
        if (yearFrac instanceof CellError) return yearFrac
        return par * rate * yearFrac
      }
    )
  }

  /**
   * AMORLINC(cost, date_purchased, first_period, salvage, period, rate, [basis])
   *
   * Returns the depreciation for each accounting period using a prorated linear
   * depreciation coefficient (French accounting system).
   */
  public amorlinc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('AMORLINC'),
      (cost: number, datePurchased: number, firstPeriod: number, salvage: number, period: number, rate: number, basis: number) => {
        if (cost < 0 || salvage < 0 || rate <= 0 || period < 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const lifeSpan = 1.0 / rate
        const yearFrac = this.yearFraction(datePurchased, firstPeriod, basis)
        if (yearFrac instanceof CellError) return yearFrac
        // Depreciation in first period (prorated)
        const firstDep = cost * rate * yearFrac
        const fullDep = cost * rate
        const depreciatedCost = cost - salvage

        if (period === 0) {
          return Math.min(firstDep, depreciatedCost)
        }

        let acc = firstDep
        for (let p = 1; p < period; p++) {
          acc += fullDep
          if (acc >= depreciatedCost) return 0
        }

        // Last period may be partial
        if (period >= Math.ceil(lifeSpan)) {
          return Math.max(0, depreciatedCost - acc)
        }
        return Math.min(fullDep, depreciatedCost - acc)
      }
    )
  }

  /**
   * COUPDAYBS(settlement, maturity, frequency, [basis])
   *
   * Returns the number of days from the beginning of the coupon period to the
   * settlement date.
   */
  public coupdaybs(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPDAYBS'),
      (settlement: number, maturity: number, frequency: number, basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const pcd = this.prevCouponDate(settlement, maturity, frequency as CouponFrequency)
        return this.daysBetween(this.dateTimeHelper.dateToNumber(pcd), settlement, basis)
      }
    )
  }

  /**
   * COUPDAYS(settlement, maturity, frequency, [basis])
   *
   * Returns the number of days in the coupon period that contains the settlement
   * date.
   */
  public coupdays(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPDAYS'),
      (settlement: number, maturity: number, frequency: number, basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const pcd = this.prevCouponDate(settlement, maturity, frequency as CouponFrequency)
        const ncd = this.nextCouponDate(settlement, maturity, frequency as CouponFrequency)
        if (basis === Basis.ACTUAL_ACTUAL) {
          return this.dateTimeHelper.dateToNumber(ncd) - this.dateTimeHelper.dateToNumber(pcd)
        }
        return daysInCouponPeriod(frequency as CouponFrequency, basis)
      }
    )
  }

  /**
   * COUPDAYSNC(settlement, maturity, frequency, [basis])
   *
   * Returns the number of days from the settlement date to the next coupon date.
   *
   * Computed as COUPDAYS - COUPDAYBS to ensure the three coupon-day functions
   * satisfy the identity COUPDAYBS + COUPDAYSNC = COUPDAYS for all bases.
   */
  public coupdaysnc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPDAYSNC'),
      (settlement: number, maturity: number, frequency: number, basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const pcd = this.prevCouponDate(settlement, maturity, frequency as CouponFrequency)
        const ncd = this.nextCouponDate(settlement, maturity, frequency as CouponFrequency)

        let coupDays: number
        if (basis === Basis.ACTUAL_ACTUAL) {
          coupDays = this.dateTimeHelper.dateToNumber(ncd) - this.dateTimeHelper.dateToNumber(pcd)
        } else {
          coupDays = daysInCouponPeriod(frequency as CouponFrequency, basis)
        }

        const pcdSerial = this.dateTimeHelper.dateToNumber(pcd)
        const coupdaybs = this.daysBetween(pcdSerial, settlement, basis)
        return coupDays - coupdaybs
      }
    )
  }

  /**
   * COUPNCD(settlement, maturity, frequency, [basis])
   *
   * Returns the next coupon date after the settlement date.
   */
  public coupncd(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPNCD'),
      (settlement: number, maturity: number, frequency: number, _basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const ncd = this.nextCouponDate(settlement, maturity, frequency as CouponFrequency)
        return this.dateTimeHelper.dateToNumber(ncd)
      }
    )
  }

  /**
   * COUPNUM(settlement, maturity, frequency, [basis])
   *
   * Returns the number of coupon payments between the settlement date and the
   * maturity date.
   */
  public coupnum(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPNUM'),
      (settlement: number, maturity: number, frequency: number, _basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const mat = this.dateTimeHelper.numberToSimpleDate(maturity)
        const set = this.dateTimeHelper.numberToSimpleDate(settlement)
        const monthsDiff = (mat.year - set.year) * 12 + (mat.month - set.month)
        const monthsPerCoupon = 12 / frequency
        return Math.max(1, Math.ceil(monthsDiff / monthsPerCoupon))
      }
    )
  }

  /**
   * COUPPCD(settlement, maturity, frequency, [basis])
   *
   * Returns the previous coupon date before the settlement date.
   */
  public couppcd(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('COUPPCD'),
      (settlement: number, maturity: number, frequency: number, _basis: number) => {
        if (settlement >= maturity || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const pcd = this.prevCouponDate(settlement, maturity, frequency as CouponFrequency)
        return this.dateTimeHelper.dateToNumber(pcd)
      }
    )
  }

  /**
   * DISC(settlement, maturity, pr, redemption, [basis])
   *
   * Returns the discount rate for a security.
   */
  public disc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DISC'),
      (settlement: number, maturity: number, pr: number, redemption: number, basis: number) => {
        if (settlement >= maturity || pr <= 0 || redemption <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(settlement, maturity, basis)
        if (yearFrac instanceof CellError) return yearFrac
        if (yearFrac === 0) return new CellError(ErrorType.DIV_BY_ZERO)
        return (1 - pr / redemption) / yearFrac
      }
    )
  }

  /**
   * DURATION(settlement, maturity, coupon, yield, frequency, [basis])
   *
   * Returns the Macaulay duration of a security with an assumed par value of $100.
   */
  public duration(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DURATION'),
      (settlement: number, maturity: number, coupon: number, yld: number, frequency: number, basis: number) => {
        if (settlement >= maturity || coupon < 0 || yld < 0 || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        return durationCore(settlement, maturity, coupon, yld, frequency as CouponFrequency, basis, this.yearFraction.bind(this), this.prevCouponDate.bind(this), this.dateTimeHelper.dateToNumber.bind(this.dateTimeHelper))
      }
    )
  }

  /**
   * INTRATE(settlement, maturity, investment, redemption, [basis])
   *
   * Returns the interest rate for a fully invested security.
   */
  public intrate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('INTRATE'),
      (settlement: number, maturity: number, investment: number, redemption: number, basis: number) => {
        if (settlement >= maturity || investment <= 0 || redemption <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(settlement, maturity, basis)
        if (yearFrac instanceof CellError) return yearFrac
        if (yearFrac === 0) return new CellError(ErrorType.DIV_BY_ZERO)
        return (redemption / investment - 1) / yearFrac
      }
    )
  }

  /**
   * MDURATION(settlement, maturity, coupon, yield, frequency, [basis])
   *
   * Returns the Macaulay modified duration for a security with an assumed par
   * value of $100.
   */
  public mduration(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MDURATION'),
      (settlement: number, maturity: number, coupon: number, yld: number, frequency: number, basis: number) => {
        if (settlement >= maturity || coupon < 0 || yld < 0 || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const dur = durationCore(settlement, maturity, coupon, yld, frequency as CouponFrequency, basis, this.yearFraction.bind(this), this.prevCouponDate.bind(this), this.dateTimeHelper.dateToNumber.bind(this.dateTimeHelper))
        if (dur instanceof CellError) return dur
        return dur / (1 + yld / frequency)
      }
    )
  }

  /**
   * PRICE(settlement, maturity, rate, yield, redemption, frequency, [basis])
   *
   * Returns the price per $100 face value of a security that pays periodic
   * interest.
   */
  public price(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PRICE'),
      (settlement: number, maturity: number, rate: number, yld: number, redemption: number, frequency: number, basis: number) => {
        if (settlement >= maturity || rate < 0 || yld < 0 || redemption <= 0 || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        return priceCore(settlement, maturity, rate, yld, redemption, frequency as CouponFrequency, basis, this.yearFraction.bind(this), this.prevCouponDate.bind(this), this.nextCouponDate.bind(this), this.dateTimeHelper.dateToNumber.bind(this.dateTimeHelper), this.daysBetween.bind(this))
      }
    )
  }

  /**
   * PRICEDISC(settlement, maturity, discount, redemption, [basis])
   *
   * Returns the price per $100 face value of a discounted security.
   */
  public pricedisc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PRICEDISC'),
      (settlement: number, maturity: number, discount: number, redemption: number, basis: number) => {
        if (settlement >= maturity || redemption <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(settlement, maturity, basis)
        if (yearFrac instanceof CellError) return yearFrac
        return redemption * (1 - discount * yearFrac)
      }
    )
  }

  /**
   * PRICEMAT(settlement, maturity, issue, rate, yield, [basis])
   *
   * Returns the price per $100 face value of a security that pays interest at
   * maturity.
   */
  public pricemat(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PRICEMAT'),
      (settlement: number, maturity: number, issue: number, rate: number, yld: number, basis: number) => {
        if (settlement >= maturity || issue > settlement) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const dimYF = this.yearFraction(issue, maturity, basis)
        if (dimYF instanceof CellError) return dimYF
        const disYF = this.yearFraction(issue, settlement, basis)
        if (disYF instanceof CellError) return disYF
        const smYF = this.yearFraction(settlement, maturity, basis)
        if (smYF instanceof CellError) return smYF
        const denom = 1 + yld * smYF
        if (denom === 0) return new CellError(ErrorType.DIV_BY_ZERO)
        return (100 * (1 + rate * dimYF)) / denom - 100 * rate * disYF
      }
    )
  }

  /**
   * RECEIVED(settlement, maturity, investment, discount, [basis])
   *
   * Returns the amount received at maturity for a fully invested security.
   */
  public received(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('RECEIVED'),
      (settlement: number, maturity: number, investment: number, discount: number, basis: number) => {
        if (settlement >= maturity || investment <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(settlement, maturity, basis)
        if (yearFrac instanceof CellError) return yearFrac
        const denom = 1 - discount * yearFrac
        if (denom <= 0) return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        return investment / denom
      }
    )
  }

  /**
   * VDB(cost, salvage, life, start_period, end_period, [factor], [no_switch])
   *
   * Returns the depreciation of an asset for any period you specify, including
   * partial periods, using the double-declining balance method or another method
   * you specify. VDB stands for Variable Declining Balance.
   */
  public vdb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('VDB'),
      (cost: number, salvage: number, life: number, startPeriod: number, endPeriod: number, factor: number, noSwitch: boolean) => {
        if (startPeriod > endPeriod || endPeriod > life || salvage > cost) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        return vdbCore(cost, salvage, life, startPeriod, endPeriod, factor, noSwitch)
      }
    )
  }

  /**
   * XIRR(values, dates, [guess])
   *
   * Returns the internal rate of return for a schedule of cash flows that is
   * not necessarily periodic.
   */
  public xirr(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('XIRR'),
      (valuesRange: SimpleRangeValue, datesRange: SimpleRangeValue, guess: number) => {
        const values = extractNumbers(valuesRange)
        if (values instanceof CellError) return values
        const dates = extractNumbers(datesRange)
        if (dates instanceof CellError) return dates

        if (values.length !== dates.length) {
          return new CellError(ErrorType.NUM, ErrorMessage.EqualLength)
        }
        if (values.length < 2) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const hasPositive = values.some(v => v > 0)
        const hasNegative = values.some(v => v < 0)
        if (!hasPositive || !hasNegative) {
          return new CellError(ErrorType.NUM)
        }

        const d0 = dates[0]
        const xirrF = (r: number) =>
          values.reduce((sum, v, i) => sum + v / Math.pow(1 + r, (dates[i] - d0) / 365), 0)
        const xirrDF = (r: number) =>
          values.reduce((sum, v, i) => sum - (dates[i] - d0) / 365 * v / Math.pow(1 + r, (dates[i] - d0) / 365 + 1), 0)

        return newtonRaphson(xirrF, xirrDF, guess)
      }
    )
  }

  /**
   * YIELD(settlement, maturity, rate, price, redemption, frequency, [basis])
   *
   * Returns the yield on a security that pays periodic interest.
   */
  public yield(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('YIELD'),
      (settlement: number, maturity: number, rate: number, pr: number, redemption: number, frequency: number, basis: number) => {
        if (settlement >= maturity || rate < 0 || pr <= 0 || redemption <= 0 || !isValidFrequency(frequency)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const priceAtYield = (yld: number) =>
          priceCore(settlement, maturity, rate, yld, redemption, frequency as CouponFrequency, basis, this.yearFraction.bind(this), this.prevCouponDate.bind(this), this.nextCouponDate.bind(this), this.dateTimeHelper.dateToNumber.bind(this.dateTimeHelper), this.daysBetween.bind(this))

        const f = (yld: number) => {
          const p = priceAtYield(yld)
          if (p instanceof CellError) return NaN
          return p - pr
        }
        const df = (yld: number) => {
          const delta = 1e-7
          const p1 = priceAtYield(yld + delta)
          const p0 = priceAtYield(yld - delta)
          if (p1 instanceof CellError || p0 instanceof CellError) return NaN
          return (p1 - p0) / (2 * delta)
        }

        const fWrapped = (x: number) => { const v = f(x); return isNaN(v) ? 0 : v }
        const dfWrapped = (x: number) => { const v = df(x); return isNaN(v) ? 1 : v }

        return newtonRaphson(fWrapped, dfWrapped, 0.1)
      }
    )
  }

  /**
   * YIELDDISC(settlement, maturity, pr, redemption, [basis])
   *
   * Returns the annual yield for a discounted security.
   */
  public yielddisc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('YIELDDISC'),
      (settlement: number, maturity: number, pr: number, redemption: number, basis: number) => {
        if (settlement >= maturity || pr <= 0 || redemption <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const yearFrac = this.yearFraction(settlement, maturity, basis)
        if (yearFrac instanceof CellError) return yearFrac
        if (yearFrac === 0) return new CellError(ErrorType.DIV_BY_ZERO)
        return (redemption / pr - 1) / yearFrac
      }
    )
  }

  /**
   * YIELDMAT(settlement, maturity, issue, rate, price, [basis])
   *
   * Returns the annual yield of a security that pays interest at maturity.
   */
  public yieldmat(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('YIELDMAT'),
      (settlement: number, maturity: number, issue: number, rate: number, pr: number, basis: number) => {
        if (settlement >= maturity || issue > settlement || pr <= 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }
        const dimYF = this.yearFraction(issue, maturity, basis)
        if (dimYF instanceof CellError) return dimYF
        const disYF = this.yearFraction(issue, settlement, basis)
        if (disYF instanceof CellError) return disYF
        const smYF = this.yearFraction(settlement, maturity, basis)
        if (smYF instanceof CellError) return smYF
        if (smYF === 0) return new CellError(ErrorType.DIV_BY_ZERO)
        return ((1 + rate * dimYF) / (pr / 100 + rate * disYF) - 1) / smYF
      }
    )
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  /**
   * Computes the year fraction between two date serial numbers using the
   * specified day-count basis.
   *
   * Basis 0: US 30/360 (NASD)
   * Basis 1: Actual/Actual
   * Basis 2: Actual/360
   * Basis 3: Actual/365
   * Basis 4: European 30/360
   */
  private yearFraction(startSerial: number, endSerial: number, basis: number): number | CellError {
    const start = this.dateTimeHelper.numberToSimpleDate(Math.floor(startSerial))
    const end = this.dateTimeHelper.numberToSimpleDate(Math.floor(endSerial))
    const actualDays = Math.floor(endSerial) - Math.floor(startSerial)

    switch (basis) {
      case Basis.US_30_360: {
        const [s, e] = this.dateTimeHelper.toBasisUS({...start}, {...end})
        return days30_360(s, e) / 360
      }
      case Basis.ACTUAL_ACTUAL: {
        const yearLen = this.dateTimeHelper.yearLengthForBasis(start, end)
        return actualDays / yearLen
      }
      case Basis.ACTUAL_360:
        return actualDays / 360
      case Basis.ACTUAL_365:
        return actualDays / 365
      case Basis.EUROPEAN_30_360: {
        const s = toBasisEU({...start})
        const e = toBasisEU({...end})
        return days30_360(s, e) / 360
      }
      default:
        return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
    }
  }

  /**
   * Returns the number of days between two date serials for the given basis.
   * For basis 0 and 4 uses 30/360 day count; otherwise uses actual days.
   */
  private daysBetween(startSerial: number, endSerial: number, basis: number): number {
    const actualDays = Math.floor(endSerial) - Math.floor(startSerial)
    if (basis === Basis.US_30_360) {
      const start = this.dateTimeHelper.numberToSimpleDate(Math.floor(startSerial))
      const end = this.dateTimeHelper.numberToSimpleDate(Math.floor(endSerial))
      const [s, e] = this.dateTimeHelper.toBasisUS({...start}, {...end})
      return days30_360(s, e)
    }
    if (basis === Basis.EUROPEAN_30_360) {
      const start = this.dateTimeHelper.numberToSimpleDate(Math.floor(startSerial))
      const end = this.dateTimeHelper.numberToSimpleDate(Math.floor(endSerial))
      return days30_360(toBasisEU({...start}), toBasisEU({...end}))
    }
    return actualDays
  }

  /**
   * Returns the previous coupon date (as SimpleDate) before or on `settlement`.
   *
   * Coupon dates are anchored to the maturity date, going back in increments of
   * (12 / frequency) months.
   */
  private prevCouponDate(settlementSerial: number, maturitySerial: number, frequency: CouponFrequency): SimpleDate {
    const maturity = this.dateTimeHelper.numberToSimpleDate(maturitySerial)
    const settlement = this.dateTimeHelper.numberToSimpleDate(settlementSerial)
    const monthsPerCoupon = 12 / frequency

    // Find how many full coupon periods fit between settlement and maturity
    const totalMonths = (maturity.year - settlement.year) * 12 + (maturity.month - settlement.month)
    const couponsBefore = Math.ceil(totalMonths / monthsPerCoupon)

    let candidate = offsetMonth(maturity, -couponsBefore * monthsPerCoupon)
    candidate = {year: candidate.year, month: candidate.month, day: Math.min(candidate.day, maturity.day)}

    // Step forward until we find the latest coupon date <= settlement
    while (this.dateTimeHelper.dateToNumber(candidate) > settlementSerial) {
      candidate = offsetMonth(candidate, -monthsPerCoupon)
      candidate = {year: candidate.year, month: candidate.month, day: Math.min(candidate.day, maturity.day)}
    }
    while (true) {
      const next = offsetMonth(candidate, monthsPerCoupon)
      const nextFixed = {year: next.year, month: next.month, day: Math.min(maturity.day, daysInMonth(next.year, next.month))}
      if (this.dateTimeHelper.dateToNumber(nextFixed) > settlementSerial) break
      candidate = nextFixed
    }

    return candidate
  }

  /**
   * Returns the next coupon date (as SimpleDate) strictly after `settlement`.
   */
  private nextCouponDate(settlementSerial: number, maturitySerial: number, frequency: CouponFrequency): SimpleDate {
    const pcd = this.prevCouponDate(settlementSerial, maturitySerial, frequency)
    const monthsPerCoupon = 12 / frequency
    const next = offsetMonth(pcd, monthsPerCoupon)
    const maturity = this.dateTimeHelper.numberToSimpleDate(maturitySerial)
    return {year: next.year, month: next.month, day: Math.min(maturity.day, daysInMonth(next.year, next.month))}
  }
}

// ---------------------------------------------------------------------------
// Module-level pure helper functions
// ---------------------------------------------------------------------------

/**
 * Validates that a frequency value is 1, 2, or 4.
 */
function isValidFrequency(frequency: number): frequency is CouponFrequency {
  return frequency === 1 || frequency === 2 || frequency === 4
}

/**
 * Returns whether a year is a Gregorian leap year.
 */
function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

/**
 * Returns the number of days in a given month/year.
 */
function daysInMonth(year: number, month: number): number {
  const dims = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
  if (month === 2 && isLeapYear(year)) return 29
  return dims[month - 1]
}

/**
 * Computes the 30/360 day count between two already-adjusted dates.
 * Formula: (Y2-Y1)*360 + (M2-M1)*30 + (D2-D1)
 */
function days30_360(start: SimpleDate, end: SimpleDate): number {
  return (end.year - start.year) * 360 + (end.month - start.month) * 30 + (end.day - start.day)
}

/**
 * Returns the number of days in one coupon period for non-actual/actual bases.
 * For basis 0, 2, 3, 4: period = 360 / frequency or 365 / frequency.
 * In practice Excel/GSheets uses 360/freq for basis 0,2,4 and 365/freq for basis 3.
 */
function daysInCouponPeriod(frequency: CouponFrequency, basis: number): number {
  if (basis === Basis.ACTUAL_365) {
    return 365 / frequency
  }
  return 360 / frequency
}

type YearFractionFn = (start: number, end: number, basis: number) => number | CellError
type PrevCouponFn = (settlement: number, maturity: number, frequency: CouponFrequency) => SimpleDate
type NextCouponFn = (settlement: number, maturity: number, frequency: CouponFrequency) => SimpleDate
type DateToNumberFn = (date: SimpleDate) => number
type DaysBetweenFn = (start: number, end: number, basis: number) => number

/**
 * Core PRICE computation shared by PRICE and YIELD (via Newton-Raphson).
 *
 * Based on the standard Excel/Google Sheets price formula:
 * PRICE = (redemption / (1+yld/freq)^(N-1+DSC/E)) +
 *         sum_{k=1}^{N}[ (100*rate/freq) / (1+yld/freq)^(k-1+DSC/E) ]
 *         - 100 * rate/freq * A/E
 * where:
 *   N   = number of coupon periods remaining
 *   DSC = days from settlement to next coupon date
 *   E   = days in the coupon period
 *   A   = days from beginning of coupon period to settlement (accrued)
 */
function priceCore(
  settlement: number,
  maturity: number,
  rate: number,
  yld: number,
  redemption: number,
  frequency: CouponFrequency,
  basis: number,
  yearFractionFn: YearFractionFn,
  prevCouponFn: PrevCouponFn,
  nextCouponFn: NextCouponFn,
  dateToNumberFn: DateToNumberFn,
  daysBetweenFn: DaysBetweenFn,
): number | CellError {
  const pcd = prevCouponFn(settlement, maturity, frequency)
  const ncd = nextCouponFn(settlement, maturity, frequency)
  const pcdSerial = dateToNumberFn(pcd)
  const ncdSerial = dateToNumberFn(ncd)

  let E: number
  if (basis === Basis.ACTUAL_ACTUAL) {
    E = ncdSerial - pcdSerial
  } else {
    E = daysInCouponPeriod(frequency, basis)
  }

  const A = daysBetweenFn(pcdSerial, settlement, basis)
  const DSC = daysBetweenFn(settlement, ncdSerial, basis)

  // Number of remaining full coupon periods after next coupon date
  const N = Math.round((maturity - ncdSerial) / (365 / frequency)) + 1

  const couponPayment = 100 * rate / frequency
  const yldPerPeriod = yld / frequency
  const firstExponent = DSC / E

  let price: number
  if (N === 1) {
    // Single remaining period
    const denom = 1 + firstExponent * yldPerPeriod
    if (denom === 0) return new CellError(ErrorType.DIV_BY_ZERO)
    price = (redemption + couponPayment) / denom - couponPayment * A / E
  } else {
    price = redemption / Math.pow(1 + yldPerPeriod, N - 1 + firstExponent)
    for (let k = 1; k <= N; k++) {
      price += couponPayment / Math.pow(1 + yldPerPeriod, k - 1 + firstExponent)
    }
    price -= couponPayment * A / E
  }

  return price
}

/**
 * Macaulay duration core computation.
 */
function durationCore(
  settlement: number,
  maturity: number,
  coupon: number,
  yld: number,
  frequency: CouponFrequency,
  basis: number,
  yearFractionFn: YearFractionFn,
  prevCouponFn: PrevCouponFn,
  dateToNumberFn: DateToNumberFn,
): number | CellError {
  const pcd = prevCouponFn(settlement, maturity, frequency)
  const pcdSerial = dateToNumberFn(pcd)

  const yldPerPeriod = yld / frequency
  const couponPayment = 100 * coupon / frequency

  // Get year fraction from settlement to maturity
  const totalYF = yearFractionFn(settlement, maturity, basis)
  if (totalYF instanceof CellError) return totalYF

  // Get year fraction from previous coupon to settlement (accrual fraction)
  const accYF = yearFractionFn(pcdSerial, settlement, basis)
  if (accYF instanceof CellError) return accYF

  const N = Math.round(totalYF * frequency)
  if (N < 1) return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)

  // Fractional first coupon period remaining
  const w = accYF * frequency

  let pvSum = 0
  let weightedSum = 0
  for (let k = 1; k <= N; k++) {
    const t = (k - w) / frequency
    const cf = k < N ? couponPayment : couponPayment + 100
    const pv = cf / Math.pow(1 + yldPerPeriod, k - w)
    pvSum += pv
    weightedSum += t * pv
  }

  if (pvSum === 0) return new CellError(ErrorType.DIV_BY_ZERO)
  return weightedSum / pvSum
}

/**
 * Newton-Raphson root finder.
 *
 * Finds x such that f(x) = 0 using derivative df.
 */
function newtonRaphson(
  f: (x: number) => number,
  df: (x: number) => number,
  guess: number,
  maxIter = 100,
  tol = 1e-10,
): number | CellError {
  let x = guess
  for (let i = 0; i < maxIter; i++) {
    const fx = f(x)
    const dfx = df(x)
    if (Math.abs(dfx) < 1e-15) return new CellError(ErrorType.NUM)
    const nextX = x - fx / dfx
    if (Math.abs(nextX - x) < tol) return nextX
    x = nextX
    if (!isFinite(x)) return new CellError(ErrorType.NUM)
  }
  return new CellError(ErrorType.NUM)
}

/**
 * Extracts an array of numbers from a SimpleRangeValue, returning a CellError
 * if any non-numeric value is encountered.
 *
 * Handles ExtendedNumber types (e.g. DateNumber, CurrencyNumber) by unwrapping
 * them via getRawValue.
 */
function extractNumbers(range: SimpleRangeValue): number[] | CellError {
  const vals = range.valuesFromTopLeftCorner()
  const result: number[] = []
  for (const v of vals) {
    if (v instanceof CellError) return v
    if (isExtendedNumber(v)) {
      result.push(getRawValue(v))
    } else if (typeof v === 'number') {
      result.push(v)
    } else {
      return new CellError(ErrorType.VALUE, ErrorMessage.NumberExpected)
    }
  }
  return result
}

/**
 * Variable Declining Balance depreciation core.
 *
 * Integrates DDB depreciation from startPeriod to endPeriod, switching to
 * straight-line when SLN > DDB (unless noSwitch is true).
 */
function vdbCore(
  cost: number,
  salvage: number,
  life: number,
  startPeriod: number,
  endPeriod: number,
  factor: number,
  noSwitch: boolean,
): number {
  // Integrate by computing cumulative depreciation up to each endpoint
  return cumulativeVDB(cost, salvage, life, endPeriod, factor, noSwitch) -
         cumulativeVDB(cost, salvage, life, startPeriod, factor, noSwitch)
}

/**
 * Computes cumulative VDB depreciation from period 0 to `period`.
 */
function cumulativeVDB(
  cost: number,
  salvage: number,
  life: number,
  period: number,
  factor: number,
  noSwitch: boolean,
): number {
  if (period === 0) return 0

  let bookValue = cost
  let cumDep = 0
  const rate = factor / life

  // Full integer periods
  const fullPeriods = Math.floor(period)
  let switchedToSLN = false

  for (let p = 1; p <= fullPeriods; p++) {
    const remainingLife = life - (p - 1)
    const slnDep = remainingLife > 0 ? (bookValue - salvage) / remainingLife : 0
    const ddbDep = bookValue * rate
    const actualDep = (!noSwitch && slnDep > ddbDep && !switchedToSLN)
      ? (() => { switchedToSLN = true; return slnDep })()
      : (switchedToSLN ? slnDep : ddbDep)
    const dep = Math.min(actualDep, Math.max(bookValue - salvage, 0))
    cumDep += dep
    bookValue -= dep
  }

  // Fractional remainder
  const fraction = period - fullPeriods
  if (fraction > 0) {
    const remainingLife = life - fullPeriods
    const slnDep = remainingLife > 0 ? (bookValue - salvage) / remainingLife : 0
    const ddbDep = bookValue * rate
    const periodDep = (!noSwitch && slnDep > ddbDep && !switchedToSLN) ? slnDep : ddbDep
    const dep = Math.min(periodDep * fraction, Math.max(bookValue - salvage, 0))
    cumDep += dep
  }

  return cumDep
}
