/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue} from '../../InterpreterValue'
import {
  hypgeom,
  lognormal,
  negbin,
  normal,
} from '../3rdparty/jstat/jstat'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible statistical function overrides.
 *
 * Overrides distribution functions that require a `cumulative` boolean argument
 * to default it to `true` when omitted, matching Google Sheets behavior.
 *
 * Affected functions:
 * - HYPGEOM.DIST / HYPGEOMDIST
 * - LOGNORM.DIST / LOGNORMDIST
 * - NEGBINOM.DIST / NEGBINOMDIST
 * - NORM.S.DIST / NORMSDIST
 */
export class GoogleSheetsStatisticalFixesPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsStatisticalFixesPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'HYPGEOM.DIST': {
      method: 'hypgeomdist',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
    },
    'LOGNORM.DIST': {
      method: 'lognormdist',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER, greaterThan: 0},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
    },
    'NEGBINOM.DIST': {
      method: 'negbinomdist',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 1},
        {argumentType: FunctionArgumentType.NUMBER, minValue: 0, maxValue: 1},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
    },
    'NORM.S.DIST': {
      method: 'normsdist',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
    },
  }

  public static aliases = {
    HYPGEOMDIST: 'HYPGEOM.DIST',
    LOGNORMDIST: 'LOGNORM.DIST',
    NEGBINOMDIST: 'NEGBINOM.DIST',
    NORMSDIST: 'NORM.S.DIST',
  }

  /**
   * HYPGEOM.DIST — hypergeometric distribution with cumulative defaulting to true.
   */
  public hypgeomdist(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('HYPGEOM.DIST'),
      (s: number, numberS: number, populationS: number, numberPop: number, cumulative: boolean) => {
        if (s > numberS || s > populationS || numberS > numberPop || populationS > numberPop) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueLarge)
        }
        if (s + numberPop < populationS + numberS) {
          return new CellError(ErrorType.NUM, ErrorMessage.ValueLarge)
        }
        s = Math.trunc(s)
        numberS = Math.trunc(numberS)
        populationS = Math.trunc(populationS)
        numberPop = Math.trunc(numberPop)

        if (cumulative) {
          return hypgeom.cdf(s, numberPop, populationS, numberS)
        } else {
          return hypgeom.pdf(s, numberPop, populationS, numberS)
        }
      }
    )
  }

  /**
   * LOGNORM.DIST — lognormal distribution with cumulative defaulting to true.
   */
  public lognormdist(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LOGNORM.DIST'),
      (x: number, mean: number, stddev: number, cumulative: boolean) => {
        if (cumulative) {
          return lognormal.cdf(x, mean, stddev)
        } else {
          return lognormal.pdf(x, mean, stddev)
        }
      }
    )
  }

  /**
   * NEGBINOM.DIST — negative binomial distribution with cumulative defaulting to true.
   */
  public negbinomdist(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('NEGBINOM.DIST'),
      (nf: number, ns: number, p: number, cumulative: boolean) => {
        nf = Math.trunc(nf)
        ns = Math.trunc(ns)
        if (cumulative) {
          return negbin.cdf(nf, ns, p)
        } else {
          return negbin.pdf(nf, ns, p)
        }
      }
    )
  }

  /**
   * NORM.S.DIST — standard normal distribution with cumulative defaulting to true.
   */
  public normsdist(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('NORM.S.DIST'),
      (x: number, cumulative: boolean) => {
        if (cumulative) {
          return normal.cdf(x, 0, 1)
        } else {
          return normal.pdf(x, 0, 1)
        }
      }
    )
  }
}
