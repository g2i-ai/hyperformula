/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {ExtendedNumber, InternalNoErrorScalarValue, InterpreterValue} from '../../InterpreterValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible operator functions.
 *
 * Implements 16 operator functions matching Google Sheets behavior:
 * - Arithmetic: ADD, MINUS, MULTIPLY, DIVIDE, POW
 * - Comparison: GT, GTE, LT, LTE, EQ, NE
 * - Text: CONCAT
 * - Unary: UMINUS, UPLUS, UNARY_PERCENT
 * - Utility: ISBETWEEN
 *
 * All arithmetic operations delegate to ArithmeticHelper to ensure
 * epsilon rounding and ExtendedNumber type preservation (dates, percents,
 * currencies) consistent with the rest of the engine.
 *
 * All comparison operations delegate to ArithmeticHelper.compare() to
 * ensure locale-aware string collation and correct cross-type ordering.
 */
export class GoogleSheetsOperatorPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsOperatorPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'ADD': {
      method: 'add',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'MINUS': {
      method: 'minus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'MULTIPLY': {
      method: 'multiply',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'DIVIDE': {
      method: 'divide',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'POW': {
      method: 'pow',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'GT': {
      method: 'gt',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'GTE': {
      method: 'gte',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'LT': {
      method: 'lt',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'LTE': {
      method: 'lte',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'EQ': {
      method: 'eq',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'NE': {
      method: 'ne',
      parameters: [
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
        {argumentType: FunctionArgumentType.NOERROR, passSubtype: true},
      ],
    },
    'CONCAT': {
      method: 'concat',
      parameters: [
        {argumentType: FunctionArgumentType.STRING, passSubtype: true},
        {argumentType: FunctionArgumentType.STRING, passSubtype: true},
      ],
    },
    'UMINUS': {
      method: 'uminus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'UPLUS': {
      method: 'uplus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'UNARY_PERCENT': {
      method: 'unaryPercent',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER, passSubtype: true},
      ],
    },
    'ISBETWEEN': {
      method: 'isbetween',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true},
      ],
    },
  }

  /**
   * ADD(a, b) - Returns a + b with epsilon rounding and type preservation.
   */
  public add(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ADD'),
      this.arithmeticHelper.addWithEpsilon
    )
  }

  /**
   * MINUS(a, b) - Returns a - b with epsilon rounding and type preservation.
   */
  public minus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MINUS'),
      this.arithmeticHelper.subtract
    )
  }

  /**
   * MULTIPLY(a, b) - Returns a * b with type preservation.
   */
  public multiply(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MULTIPLY'),
      this.arithmeticHelper.multiply
    )
  }

  /**
   * DIVIDE(a, b) - Returns a / b with type preservation, or DIV_BY_ZERO error if b = 0.
   */
  public divide(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DIVIDE'),
      this.arithmeticHelper.divide
    )
  }

  /**
   * POW(a, b) - Returns a raised to the power of b.
   */
  public pow(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('POW'),
      this.arithmeticHelper.pow
    )
  }

  /**
   * GT(a, b) - Returns true if a > b using locale-aware comparison.
   */
  public gt(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('GT'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.gt(a, b)
    )
  }

  /**
   * GTE(a, b) - Returns true if a >= b using locale-aware comparison.
   */
  public gte(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('GTE'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.geq(a, b)
    )
  }

  /**
   * LT(a, b) - Returns true if a < b using locale-aware comparison.
   */
  public lt(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LT'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.lt(a, b)
    )
  }

  /**
   * LTE(a, b) - Returns true if a <= b using locale-aware comparison.
   */
  public lte(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LTE'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.leq(a, b)
    )
  }

  /**
   * EQ(a, b) - Returns true if a == b using locale-aware comparison.
   */
  public eq(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('EQ'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.eq(a, b)
    )
  }

  /**
   * NE(a, b) - Returns true if a != b using locale-aware comparison.
   */
  public ne(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('NE'),
      (a: InternalNoErrorScalarValue, b: InternalNoErrorScalarValue) => this.arithmeticHelper.neq(a, b)
    )
  }

  /**
   * CONCAT(a, b) - Returns String(a) + String(b)
   */
  public concat(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'),
      this.arithmeticHelper.concat
    )
  }

  /**
   * UMINUS(a) - Returns -a (negation) with type preservation.
   */
  public uminus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UMINUS'),
      this.arithmeticHelper.unaryMinus
    )
  }

  /**
   * UPLUS(a) - Returns +a (identity).
   */
  public uplus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UPLUS'),
      (a: ExtendedNumber) => a
    )
  }

  /**
   * UNARY_PERCENT(a) - Returns a / 100 as a PercentNumber.
   */
  public unaryPercent(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UNARY_PERCENT'),
      this.arithmeticHelper.unaryPercent
    )
  }

  /**
   * ISBETWEEN(val, lo, hi, loInc, hiInc) - Returns true if val is in [lo, hi].
   *
   * Uses epsilon-aware floatCmp for all comparisons to stay consistent with
   * GTE/LTE and avoid false negatives at floating-point boundaries (e.g. 0.1+0.2 vs 0.3).
   *
   * Returns #NUM! error when lo > hi (epsilon-aware), matching Google Sheets behavior.
   */
  public isbetween(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ISBETWEEN'),
      (val: number, lo: number, hi: number, loInc: boolean, hiInc: boolean) => {
        const cmpLoHi = this.arithmeticHelper.floatCmp(lo, hi)
        if (cmpLoHi > 0) {
          return new CellError(ErrorType.NUM, ErrorMessage.NaN)
        }
        const cmpValLo = this.arithmeticHelper.floatCmp(val, lo)
        const cmpValHi = this.arithmeticHelper.floatCmp(val, hi)
        const lowerOk = loInc ? cmpValLo >= 0 : cmpValLo > 0
        const upperOk = hiInc ? cmpValHi <= 0 : cmpValHi < 0
        return lowerOk && upperOk
      }
    )
  }
}
