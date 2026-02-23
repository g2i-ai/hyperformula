/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InternalScalarValue, InterpreterValue, getRawValue, isExtendedNumber, NumberType} from '../../InterpreterValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/** Temperature unit identifiers for special-case conversion handling. */
type TemperatureUnit = 'C' | 'F' | 'K'

/**
 * Google Sheets-compatible conversion functions.
 *
 * Implements: TO_DATE, TO_DOLLARS, TO_PERCENT, TO_PURE_NUMBER, TO_TEXT, CONVERT
 */
export class GoogleSheetsConversionPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsConversionPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'TO_DATE': {
      method: 'toDate',
      parameters: [{argumentType: FunctionArgumentType.NUMBER}],
      returnNumberType: NumberType.NUMBER_DATE,
    },
    'TO_DOLLARS': {
      method: 'toDollars',
      parameters: [{argumentType: FunctionArgumentType.NUMBER}],
      returnNumberType: NumberType.NUMBER_CURRENCY,
    },
    'TO_PERCENT': {
      method: 'toPercent',
      parameters: [{argumentType: FunctionArgumentType.NUMBER}],
      returnNumberType: NumberType.NUMBER_PERCENT,
    },
    'TO_PURE_NUMBER': {
      method: 'toPureNumber',
      parameters: [{argumentType: FunctionArgumentType.NUMBER}],
    },
    'TO_TEXT': {
      method: 'toText',
      parameters: [{argumentType: FunctionArgumentType.SCALAR}],
    },
    'CONVERT': {
      method: 'convert',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
  }

  /**
   * Conversion factors relative to a base unit per category.
   * Temperature uses string markers because it requires formula-based conversion.
   */
  private static readonly CONVERSION_TABLE: Record<string, Record<string, number | string>> = {
    length: {
      m: 1, km: 1000, cm: 0.01, mm: 0.001, mi: 1609.344, yd: 0.9144,
      ft: 0.3048, in: 0.0254, Nmi: 1852, um: 1e-6, ang: 1e-10,
    },
    mass: {
      g: 1, kg: 1000, mg: 0.001, lbm: 453.59237, ozm: 28.349523125,
      stone: 6350.29318, ton: 907184.74, uk_ton: 1016046.9088,
      sg: 14593.90294, u: 1.66053906660e-24,
    },
    time: {
      sec: 1, s: 1, min: 60, mn: 60, hr: 3600, day: 86400, yr: 31557600,
    },
    temperature: {
      C: 'C', F: 'F', K: 'K',
    },
    volume: {
      l: 1, lt: 1, ml: 0.001, gal: 3.785411784, qt: 0.946352946,
      pt: 0.473176473, cup: 0.236588236, oz: 0.0295735296, tbs: 0.0147867648,
      tsp: 0.00492892159, m3: 1000, ft3: 28.316846592, in3: 0.016387064,
      yd3: 764.554857984,
    },
    area: {
      m2: 1, km2: 1e6, cm2: 1e-4, mm2: 1e-6, ft2: 0.09290304,
      in2: 0.00064516, yd2: 0.83612736, mi2: 2589988.110336,
      ha: 10000, acre: 4046.8564224, Morgen: 2500, ar: 100,
    },
    speed: {
      'm/s': 1, 'm/h': 1 / 3600, 'mph': 0.44704, 'kn': 0.514444444,
      'admkn': 0.514773333,
    },
    pressure: {
      Pa: 1, atm: 101325, mmHg: 133.322, psi: 6894.757, Torr: 133.322,
    },
    energy: {
      J: 1, e: 1e-7, cal: 4.1868, eV: 1.602176634e-19, HPh: 2684519.5368,
      Wh: 3600, BTU: 1055.05585262,
    },
    force: {
      N: 1, dyn: 1e-5, lbf: 4.4482216152605, pond: 0.00980665,
    },
  }

  /**
   * TO_DATE(serial) — wraps a serial number as a date-typed value.
   */
  public toDate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TO_DATE'),
      (serial: number) => this.arithmeticHelper.ExtendedNumberFactory(serial, {type: NumberType.NUMBER_DATE})
    )
  }

  /**
   * TO_DOLLARS(number) — wraps a number as a currency-typed value.
   */
  public toDollars(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TO_DOLLARS'),
      (value: number) => value
    )
  }

  /**
   * TO_PERCENT(number) — wraps a number as a percent-typed value.
   */
  public toPercent(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TO_PERCENT'),
      (value: number) => value
    )
  }

  /**
   * TO_PURE_NUMBER(value) — strips formatting and returns the raw numeric value.
   */
  public toPureNumber(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TO_PURE_NUMBER'),
      (value: number) => getRawValue(value)
    )
  }

  /**
   * TO_TEXT(value) — converts any scalar value to its string representation.
   */
  public toText(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TO_TEXT'),
      (value: InternalScalarValue) => String(isExtendedNumber(value) ? getRawValue(value) : value)
    )
  }

  /**
   * CONVERT(value, from_unit, to_unit) — converts a numeric value between compatible units.
   *
   * Supports: length, mass, time, temperature, volume, area, speed, pressure, energy, force.
   * Returns #N/A if the units are unknown or belong to different categories.
   */
  public convert(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('CONVERT'),
      (value: number, fromUnit: string, toUnit: string) => {
        if (fromUnit === toUnit) {
          return value
        }

        const fromCategory = this.findCategory(fromUnit)
        const toCategory = this.findCategory(toUnit)

        if (fromCategory === null || toCategory === null || fromCategory !== toCategory) {
          return new CellError(ErrorType.NA, ErrorMessage.ValueNotFound)
        }

        const category = GoogleSheetsConversionPlugin.CONVERSION_TABLE[fromCategory]

        if (fromCategory === 'temperature') {
          return this.convertTemperature(value, fromUnit as TemperatureUnit, toUnit as TemperatureUnit)
        }

        const fromFactor = category[fromUnit] as number
        const toFactor = category[toUnit] as number

        return value * fromFactor / toFactor
      }
    )
  }

  /**
   * Finds the category name that contains the given unit.
   *
   * @returns category key or null if not found
   */
  private findCategory(unit: string): string | null {
    return Object.keys(GoogleSheetsConversionPlugin.CONVERSION_TABLE).find(
      category => unit in GoogleSheetsConversionPlugin.CONVERSION_TABLE[category]
    ) ?? null
  }

  /**
   * Converts a temperature value between Celsius, Fahrenheit, and Kelvin.
   */
  private convertTemperature(value: number, from: TemperatureUnit, to: TemperatureUnit): number | CellError {
    const conversions: Record<string, (v: number) => number> = {
      'C->F': v => v * 9 / 5 + 32,
      'F->C': v => (v - 32) * 5 / 9,
      'C->K': v => v + 273.15,
      'K->C': v => v - 273.15,
      'F->K': v => (v - 32) * 5 / 9 + 273.15,
      'K->F': v => (v - 273.15) * 9 / 5 + 32,
    }

    const key = `${from}->${to}`
    const convert = conversions[key]

    if (convert === undefined) {
      return new CellError(ErrorType.NA, ErrorMessage.ValueNotFound)
    }

    return convert(value)
  }
}
