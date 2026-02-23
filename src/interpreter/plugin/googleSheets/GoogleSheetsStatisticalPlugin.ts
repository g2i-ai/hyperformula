/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {ArraySize} from '../../../ArraySize'
import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {getRawValue, InternalScalarValue, InterpreterValue, isExtendedNumber, RawScalarValue} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'
import {Condition, CriterionFunctionCompute} from '../../CriterionFunctionCompute'
import {Maybe} from '../../../Maybe'
import {erf as jstatErf} from '../3rdparty/jstat/jstat'

/** Result accumulator for AVERAGEIFS. */
class AverageResult {
  public static empty = new AverageResult(0, 0)

  constructor(
    public readonly sum: number,
    public readonly count: number,
  ) {}

  public static single(arg: number): AverageResult {
    return new AverageResult(arg, 1)
  }

  public compose(other: AverageResult): AverageResult {
    return new AverageResult(this.sum + other.sum, this.count + other.count)
  }

  public averageValue(): Maybe<number> {
    return this.count > 0 ? this.sum / this.count : undefined
  }
}

function averageifsCacheKey(conditions: Condition[]): string {
  // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
  const parts = conditions.map(c => `${c.conditionRange.range!.sheet},${c.conditionRange.range!.start.col},${c.conditionRange.range!.start.row}`)
  return ['AVERAGEIFS', ...parts].join(',')
}

/**
 * Extracts numeric values from a SimpleRangeValue's data grid.
 */
function extractNumbersFromRange(range: SimpleRangeValue): number[] {
  const nums: number[] = []
  for (const row of range.data) {
    for (const cell of row) {
      if (typeof cell === 'number') {
        nums.push(cell)
      } else if (isExtendedNumber(cell)) {
        nums.push(cell.val)
      }
    }
  }
  return nums
}

/**
 * Extracts matched numeric pairs from two ranges of equal shape.
 *
 * Iterates both ranges in lockstep and includes a pair only when both
 * corresponding cells are numeric. This preserves alignment when either
 * range contains non-numeric (text, blank, error) cells.
 *
 * @returns An object with two arrays of equal length containing the matched numeric values.
 */
function extractNumericPairs(
  rangeA: SimpleRangeValue,
  rangeB: SimpleRangeValue,
): { a: number[], b: number[] } {
  const a: number[] = []
  const b: number[] = []

  const cellsA = rangeA.data.flatMap(row => row)
  const cellsB = rangeB.data.flatMap(row => row)

  for (let i = 0; i < Math.min(cellsA.length, cellsB.length); i++) {
    const va = cellsA[i]
    const vb = cellsB[i]

    const numA = typeof va === 'number' ? va : isExtendedNumber(va) ? va.val : null
    const numB = typeof vb === 'number' ? vb : isExtendedNumber(vb) ? vb.val : null

    if (numA !== null && numB !== null) {
      a.push(numA)
      b.push(numB)
    }
  }

  return {a, b}
}

/**
 * Computes mean and sample standard deviation for an array of numbers.
 * Returns undefined stddev if fewer than 2 values.
 */
function computeMeanAndStdDev(nums: number[]): { mean: number, stddev: number | undefined } {
  const n = nums.length
  const mean = nums.reduce((a, b) => a + b, 0) / n
  if (n < 2) {
    return {mean, stddev: undefined}
  }
  const variance = nums.reduce((sum, x) => sum + (x - mean) ** 2, 0) / (n - 1)
  return {mean, stddev: Math.sqrt(variance)}
}

/**
 * Performs linear interpolation for percentile calculations.
 */
function interpolate(sortedData: number[], position: number): number {
  const lower = Math.floor(position)
  const upper = Math.ceil(position)
  if (lower === upper) {
    return sortedData[lower]
  }
  return sortedData[lower] + (position - lower) * (sortedData[upper] - sortedData[lower])
}

/**
 * Google Sheets-specific statistical functions.
 *
 * Implements functions either absent from or behaving differently in standard HyperFormula:
 * AVERAGEIFS, FORECAST, FORECAST.LINEAR, INTERCEPT, KURT, TRIMMEAN,
 * PERCENTRANK, PERCENTRANK.INC, PERCENTRANK.EXC, PERMUT, PERMUTATIONA,
 * PROB, QUARTILE, QUARTILE.INC, QUARTILE.EXC, MODE, MODE.SNGL, MODE.MULT,
 * AVERAGE.WEIGHTED, MARGINOFERROR, ERF.PRECISE, ERFC.PRECISE.
 */
export class GoogleSheetsStatisticalPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsStatisticalPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'AVERAGEIFS': {
      method: 'averageifs',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
      repeatLastArgs: 2,
    },
    'FORECAST': {
      method: 'forecast',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
      ],
    },
    'FORECAST.LINEAR': {
      method: 'forecast',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
      ],
    },
    'INTERCEPT': {
      method: 'intercept',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
      ],
    },
    'KURT': {
      method: 'kurt',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'TRIMMEAN': {
      method: 'trimmean',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0, lessThan: 1},
      ],
    },
    'PERCENTRANK': {
      method: 'percentrankInc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, defaultValue: 3},
      ],
    },
    'PERCENTRANK.INC': {
      method: 'percentrankInc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, defaultValue: 3},
      ],
    },
    'PERCENTRANK.EXC': {
      method: 'percentrankExc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, defaultValue: 3},
      ],
    },
    'PERMUT': {
      method: 'permut',
      parameters: [
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0},
      ],
    },
    'PERMUTATIONA': {
      method: 'permutationa',
      parameters: [
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0},
      ],
    },
    'PROB': {
      method: 'prob',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER, optionalArg: true},
      ],
    },
    'QUARTILE': {
      method: 'quartileInc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4},
      ],
    },
    'QUARTILE.INC': {
      method: 'quartileInc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 0, maxValue: 4},
      ],
    },
    'QUARTILE.EXC': {
      method: 'quartileExc',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.INTEGER, minValue: 1, maxValue: 3},
      ],
    },
    'MODE': {
      method: 'mode',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'MODE.SNGL': {
      method: 'mode',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'MODE.MULT': {
      method: 'modeMult',
      sizeOfResultArrayMethod: 'modeMultArraySize',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
      vectorizationForbidden: true,
    },
    'AVERAGE.WEIGHTED': {
      method: 'averageWeighted',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.RANGE},
      ],
      repeatLastArgs: 2,
    },
    'MARGINOFERROR': {
      method: 'marginOfError',
      parameters: [
        {argumentType: FunctionArgumentType.RANGE},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0, lessThan: 1},
      ],
    },
    'ERF.PRECISE': {
      method: 'erfPrecise',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'ERFC.PRECISE': {
      method: 'erfcPrecise',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
  }

  /**
   * AVERAGEIFS(avg_range, criteria_range1, criterion1, [criteria_range2, criterion2, ...])
   *
   * Like SUMIFS but returns the average of cells in avg_range that meet all criteria.
   */
  public averageifs(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('AVERAGEIFS'),
      (avgRange: SimpleRangeValue, ...args: unknown[]) => {
        const conditions: Condition[] = []
        for (let i = 0; i < args.length; i += 2) {
          const conditionRange = args[i] as SimpleRangeValue
          const criterion = args[i + 1] as RawScalarValue
          const criterionPackage = this.interpreter.criterionBuilder.fromCellValue(criterion, this.arithmeticHelper)
          if (criterionPackage === undefined) {
            return new CellError(ErrorType.VALUE, ErrorMessage.BadCriterion)
          }
          conditions.push(new Condition(conditionRange, criterionPackage))
        }

        const result = new CriterionFunctionCompute<AverageResult>(
          this.interpreter,
          averageifsCacheKey,
          AverageResult.empty,
          (left, right) => left.compose(right),
          (arg) => isExtendedNumber(arg) ? AverageResult.single(getRawValue(arg)) : AverageResult.empty,
        ).compute(avgRange, conditions)

        if (result instanceof CellError) {
          return result
        }
        return result.averageValue() ?? new CellError(ErrorType.DIV_BY_ZERO)
      }
    )
  }

  /**
   * FORECAST(x, known_y, known_x) / FORECAST.LINEAR(x, known_y, known_x)
   *
   * Predicts a y-value for the given x using linear regression on the known data.
   * Ranges must have the same total cell count. Within equal-sized ranges,
   * non-numeric cells are skipped in lockstep to preserve pairing.
   */
  public forecast(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata(ast.procedureName),
      (x: number, knownYRange: SimpleRangeValue, knownXRange: SimpleRangeValue) => {
        const totalY = knownYRange.data.reduce((sum, row) => sum + row.length, 0)
        const totalX = knownXRange.data.reduce((sum, row) => sum + row.length, 0)
        if (totalY !== totalX) {
          return new CellError(ErrorType.NA, ErrorMessage.EqualLength)
        }

        const {a: knownY, b: knownX} = extractNumericPairs(knownYRange, knownXRange)

        if (knownY.length < 2) {
          return new CellError(ErrorType.NA, ErrorMessage.EqualLength)
        }

        const {slope, intercept} = this.linearRegression(knownY, knownX)
        if (!isFinite(slope) || !isFinite(intercept)) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }
        return slope * x + intercept
      }
    )
  }

  /**
   * INTERCEPT(known_y, known_x)
   *
   * Returns the y-intercept of the linear regression line through the known data points.
   * Ranges must have the same total cell count. Within equal-sized ranges,
   * non-numeric cells are skipped in lockstep to preserve pairing.
   */
  public intercept(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('INTERCEPT'),
      (knownYRange: SimpleRangeValue, knownXRange: SimpleRangeValue) => {
        const totalY = knownYRange.data.reduce((sum, row) => sum + row.length, 0)
        const totalX = knownXRange.data.reduce((sum, row) => sum + row.length, 0)
        if (totalY !== totalX) {
          return new CellError(ErrorType.NA, ErrorMessage.EqualLength)
        }

        const {a: knownY, b: knownX} = extractNumericPairs(knownYRange, knownXRange)

        if (knownY.length < 2) {
          return new CellError(ErrorType.NA, ErrorMessage.EqualLength)
        }

        const {intercept} = this.linearRegression(knownY, knownX)
        if (!isFinite(intercept)) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }
        return intercept
      }
    )
  }

  /**
   * KURT(value1, value2, ...)
   *
   * Returns the kurtosis of a data set. Requires at least 4 values.
   */
  public kurt(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('KURT'),
      (...args: number[]) => {
        const n = args.length
        if (n < 4) {
          return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.ThreeValues)
        }

        const {mean, stddev} = computeMeanAndStdDev(args)
        if (stddev === undefined || stddev === 0) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }

        const sum4 = args.reduce((sum, x) => sum + ((x - mean) / stddev) ** 4, 0)
        const kurtosis =
          (n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3)) * sum4 -
          (3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
        return kurtosis
      }
    )
  }

  /**
   * TRIMMEAN(range, percent)
   *
   * Returns the mean after trimming the given fraction of data points from both ends.
   */
  public trimmean(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TRIMMEAN'),
      (range: SimpleRangeValue, percent: number) => {
        const nums = extractNumbersFromRange(range)
        if (nums.length === 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.OneValue)
        }

        const sorted = [...nums].sort((a, b) => a - b)
        const trimCount = Math.floor(sorted.length * (percent / 2))
        const trimmed = sorted.slice(trimCount, sorted.length - trimCount)

        if (trimmed.length === 0) {
          return new CellError(ErrorType.NUM)
        }
        return trimmed.reduce((a, b) => a + b, 0) / trimmed.length
      }
    )
  }

  /**
   * PERCENTRANK.INC(data, x, [significance]) / PERCENTRANK(data, x, [significance])
   *
   * Returns the rank of x in data as a percentage (0 to 1 inclusive).
   * Uses rank / (n - 1) formula. For a single-element dataset where x equals
   * that element, returns 0 (matching Google Sheets behavior).
   */
  public percentrankInc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const name = ast.procedureName === 'PERCENTRANK.INC' ? 'PERCENTRANK.INC' : 'PERCENTRANK'
    return this.runFunction(ast.args, state, this.metadata(name),
      (range: SimpleRangeValue, x: number, significance: number) => {
        const nums = extractNumbersFromRange(range).sort((a, b) => a - b)
        if (nums.length === 0) {
          return new CellError(ErrorType.NUM)
        }
        if (x < nums[0] || x > nums[nums.length - 1]) {
          return new CellError(ErrorType.NA)
        }
        if (nums.length === 1) {
          return 0
        }

        const rank = this.computePercentRank(nums, x)
        const pct = rank / (nums.length - 1)
        const factor = Math.pow(10, significance)
        return Math.floor(pct * factor) / factor
      }
    )
  }

  /**
   * PERCENTRANK.EXC(data, x, [significance])
   *
   * Returns the rank of x in data as a percentage (exclusive of 0 and 1).
   * Uses (rank + 1) / (n + 1) formula.
   */
  public percentrankExc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PERCENTRANK.EXC'),
      (range: SimpleRangeValue, x: number, significance: number) => {
        const nums = extractNumbersFromRange(range).sort((a, b) => a - b)
        if (nums.length === 0) {
          return new CellError(ErrorType.NUM)
        }
        if (x <= nums[0] || x >= nums[nums.length - 1]) {
          return new CellError(ErrorType.NA)
        }

        const rank = this.computePercentRank(nums, x)
        const pct = (rank + 1) / (nums.length + 1)
        const factor = Math.pow(10, significance)
        return Math.floor(pct * factor) / factor
      }
    )
  }

  /**
   * PERMUT(n, k)
   *
   * Returns the number of permutations for n objects taken k at a time: n! / (n-k)!
   */
  public permut(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PERMUT'),
      (n: number, k: number) => {
        if (k > n) {
          return new CellError(ErrorType.NUM, ErrorMessage.WrongOrder)
        }
        return this.fallingFactorial(n, k)
      }
    )
  }

  /**
   * PERMUTATIONA(n, k)
   *
   * Returns the number of permutations with repetition: n^k
   */
  public permutationa(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PERMUTATIONA'),
      (n: number, k: number) => Math.pow(n, k)
    )
  }

  /**
   * PROB(x_range, prob_range, lower_limit, [upper_limit])
   *
   * Returns the probability that values in x_range are between lower and upper limits.
   * If upper_limit is omitted, returns probability that x equals lower_limit exactly.
   */
  public prob(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('PROB'),
      (xRange: SimpleRangeValue, probRange: SimpleRangeValue, lowerLimit: number, upperLimit: number | undefined) => {
        const xValues = extractNumbersFromRange(xRange)
        const probValues = extractNumbersFromRange(probRange)

        if (xValues.length !== probValues.length) {
          return new CellError(ErrorType.NA, ErrorMessage.EqualLength)
        }

        const totalProb = probValues.reduce((sum, p) => sum + p, 0)
        if (Math.abs(totalProb - 1) > 1e-10) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }

        if (probValues.some(p => p < 0 || p > 1)) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueSmall)
        }

        const upper = upperLimit ?? lowerLimit
        return xValues.reduce((sum, x, i) => {
          if (x >= lowerLimit && x <= upper) {
            return sum + probValues[i]
          }
          return sum
        }, 0)
      }
    )
  }

  /**
   * QUARTILE.INC(data, quart) / QUARTILE(data, quart)
   *
   * Returns the quartile of a data set using inclusive interpolation (0–4).
   */
  public quartileInc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const name = ast.procedureName === 'QUARTILE' ? 'QUARTILE' : 'QUARTILE.INC'
    return this.runFunction(ast.args, state, this.metadata(name),
      (range: SimpleRangeValue, quart: number) => {
        const nums = extractNumbersFromRange(range).sort((a, b) => a - b)
        if (nums.length === 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.OneValue)
        }
        // quart 0 = min, 4 = max
        const position = (quart / 4) * (nums.length - 1)
        return interpolate(nums, position)
      }
    )
  }

  /**
   * QUARTILE.EXC(data, quart)
   *
   * Returns the quartile using exclusive interpolation (1–3 only).
   */
  public quartileExc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('QUARTILE.EXC'),
      (range: SimpleRangeValue, quart: number) => {
        const nums = extractNumbersFromRange(range).sort((a, b) => a - b)
        if (nums.length === 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.OneValue)
        }
        const position = (quart / 4) * (nums.length + 1) - 1
        if (position < 0 || position > nums.length - 1) {
          return new CellError(ErrorType.NUM)
        }
        return interpolate(nums, position)
      }
    )
  }

  /**
   * MODE(value1, value2, ...) / MODE.SNGL(value1, ...)
   *
   * Returns the most frequently occurring value. Returns #N/A if no duplicates.
   */
  public mode(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata(ast.procedureName === 'MODE.SNGL' ? 'MODE.SNGL' : 'MODE'),
      (...args: number[]) => {
        const modes = this.computeModes(args)
        if (modes.length === 0) {
          return new CellError(ErrorType.NA, ErrorMessage.ValueNotFound)
        }
        return modes[0]
      }
    )
  }

  /**
   * MODE.MULT(value1, ...)
   *
   * Returns all modes as a vertical array. Returns #N/A if no duplicates.
   */
  public modeMult(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MODE.MULT'),
      (...args: number[]) => {
        const modes = this.computeModes(args)
        if (modes.length === 0) {
          return new CellError(ErrorType.NA, ErrorMessage.ValueNotFound)
        }
        return SimpleRangeValue.onlyNumbers(modes.map(m => [m]))
      }
    )
  }

  /**
   * Predicts the result array size for MODE.MULT.
   *
   * Falls back to the number of literal arguments as the upper bound.
   * If ranges are passed, uses their total cell count.
   */
  public modeMultArraySize(ast: ProcedureAst, state: InterpreterState): ArraySize {
    // Upper bound: all args could be modes
    let count = 0
    for (const arg of ast.args) {
      const size = this.arraySizeForAst(arg, state)
      count += size.width * size.height
    }
    return new ArraySize(1, Math.max(1, count))
  }

  /**
   * AVERAGE.WEIGHTED(values, weights, [values2, weights2, ...])
   *
   * Returns the weighted average: Sum(value * weight) / Sum(weight).
   */
  public averageWeighted(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('AVERAGE.WEIGHTED'),
      (...args: unknown[]) => {
        let weightedSum = 0
        let totalWeight = 0

        for (let i = 0; i < args.length; i += 2) {
          const values = extractNumbersFromRange(args[i] as SimpleRangeValue)
          const weights = extractNumbersFromRange(args[i + 1] as SimpleRangeValue)

          if (values.length !== weights.length) {
            return new CellError(ErrorType.VALUE, ErrorMessage.EqualLength)
          }

          for (let j = 0; j < values.length; j++) {
            if (weights[j] < 0) {
              return new CellError(ErrorType.VALUE, ErrorMessage.ValueSmall)
            }
            weightedSum += values[j] * weights[j]
            totalWeight += weights[j]
          }
        }

        if (totalWeight === 0) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }
        return weightedSum / totalWeight
      }
    )
  }

  /**
   * MARGINOFERROR(range, confidence)
   *
   * Returns the margin of error: confidence * stddev / sqrt(n).
   */
  public marginOfError(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MARGINOFERROR'),
      (range: SimpleRangeValue, confidence: number) => {
        const nums = extractNumbersFromRange(range)
        if (nums.length < 2) {
          return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.TwoValues)
        }
        const {stddev} = computeMeanAndStdDev(nums)
        if (stddev === undefined) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }
        return confidence * stddev / Math.sqrt(nums.length)
      }
    )
  }

  /**
   * ERF.PRECISE(x)
   *
   * Returns the error function integrated from 0 to x. Single-argument variant of ERF.
   * Uses the high-precision Chebyshev-series implementation from jstat (error < 1e-14).
   */
  public erfPrecise(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ERF.PRECISE'),
      (x: number) => jstatErf(x)
    )
  }

  /**
   * ERFC.PRECISE(x)
   *
   * Returns the complementary error function: 1 - ERF(x).
   * Uses the high-precision Chebyshev-series implementation from jstat (error < 1e-14).
   */
  public erfcPrecise(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ERFC.PRECISE'),
      (x: number) => 1 - jstatErf(x)
    )
  }

  // ---------------------------------------------------------------------------
  // Private helpers
  // ---------------------------------------------------------------------------

  /**
   * Computes slope and intercept of the least-squares linear regression line
   * through the (knownX[i], knownY[i]) data points.
   */
  private linearRegression(knownY: number[], knownX: number[]): { slope: number, intercept: number } {
    const n = knownY.length
    const sumX = knownX.reduce((a, b) => a + b, 0)
    const sumY = knownY.reduce((a, b) => a + b, 0)
    const sumXY = knownX.reduce((sum, x, i) => sum + x * knownY[i], 0)
    const sumX2 = knownX.reduce((sum, x) => sum + x * x, 0)
    const denom = n * sumX2 - sumX * sumX
    const slope = denom !== 0 ? (n * sumXY - sumX * sumY) / denom : Infinity
    const intercept = (sumY - slope * sumX) / n
    return {slope, intercept}
  }

  /**
   * Returns the falling factorial n * (n-1) * ... * (n-k+1), i.e. n! / (n-k)!
   */
  private fallingFactorial(n: number, k: number): number {
    let result = 1
    for (let i = 0; i < k; i++) {
      result *= n - i
    }
    return result
  }

  /**
   * Returns the 0-based fractional rank of x in a sorted array using linear interpolation.
   */
  private computePercentRank(sorted: number[], x: number): number {
    const exactIdx = sorted.indexOf(x)
    if (exactIdx !== -1) {
      return exactIdx
    }
    // x is between two values — find fractional position
    for (let i = 0; i < sorted.length - 1; i++) {
      if (sorted[i] <= x && x < sorted[i + 1]) {
        return i + (x - sorted[i]) / (sorted[i + 1] - sorted[i])
      }
    }
    return sorted.length - 1
  }

  /**
   * Returns all modes (most frequent values) from an array of numbers.
   * Returns an empty array if all values are unique.
   */
  private computeModes(nums: number[]): number[] {
    const freq = new Map<number, number>()
    for (const n of nums) {
      freq.set(n, (freq.get(n) ?? 0) + 1)
    }

    const maxFreq = Math.max(...freq.values())
    if (maxFreq < 2) {
      return []
    }

    return [...freq.entries()]
      .filter(([, count]) => count === maxFreq)
      .map(([value]) => value)
      .sort((a, b) => a - b)
  }
}
