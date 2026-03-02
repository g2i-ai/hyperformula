/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue} from '../../InterpreterValue'
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
 */
export class GoogleSheetsOperatorPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsOperatorPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'ADD': {
      method: 'add',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'MINUS': {
      method: 'minus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'MULTIPLY': {
      method: 'multiply',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'DIVIDE': {
      method: 'divide',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'POW': {
      method: 'pow',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'GT': {
      method: 'gt',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'GTE': {
      method: 'gte',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'LT': {
      method: 'lt',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'LTE': {
      method: 'lte',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'EQ': {
      method: 'eq',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'NE': {
      method: 'ne',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
    },
    'CONCAT': {
      method: 'concat',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'UMINUS': {
      method: 'uminus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'UPLUS': {
      method: 'uplus',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'UNARY_PERCENT': {
      method: 'unaryPercent',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
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
   * ADD(a, b) - Returns a + b
   */
  public add(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ADD'),
      (a: number, b: number) => a + b
    )
  }

  /**
   * MINUS(a, b) - Returns a - b
   */
  public minus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MINUS'),
      (a: number, b: number) => a - b
    )
  }

  /**
   * MULTIPLY(a, b) - Returns a * b
   */
  public multiply(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MULTIPLY'),
      (a: number, b: number) => a * b
    )
  }

  /**
   * DIVIDE(a, b) - Returns a / b, or DIV_BY_ZERO error if b = 0
   */
  public divide(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DIVIDE'),
      (a: number, b: number) => {
        if (b === 0) {
          return new CellError(ErrorType.DIV_BY_ZERO)
        }
        return a / b
      }
    )
  }

  /**
   * POW(a, b) - Returns a raised to the power of b
   */
  public pow(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('POW'),
      (a: number, b: number) => Math.pow(a, b)
    )
  }

  /**
   * GT(a, b) - Returns true if a > b
   */
  public gt(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('GT'),
      (a: any, b: any) => a > b
    )
  }

  /**
   * GTE(a, b) - Returns true if a >= b
   */
  public gte(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('GTE'),
      (a: any, b: any) => a >= b
    )
  }

  /**
   * LT(a, b) - Returns true if a < b
   */
  public lt(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LT'),
      (a: any, b: any) => a < b
    )
  }

  /**
   * LTE(a, b) - Returns true if a <= b
   */
  public lte(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LTE'),
      (a: any, b: any) => a <= b
    )
  }

  /**
   * EQ(a, b) - Returns true if a === b
   */
  public eq(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('EQ'),
      (a: any, b: any) => a === b
    )
  }

  /**
   * NE(a, b) - Returns true if a !== b
   */
  public ne(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('NE'),
      (a: any, b: any) => a !== b
    )
  }

  /**
   * CONCAT(a, b) - Returns String(a) + String(b)
   */
  public concat(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'),
      (a: string, b: string) => String(a) + String(b)
    )
  }

  /**
   * UMINUS(a) - Returns -a (negation)
   */
  public uminus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UMINUS'),
      (a: number) => -a
    )
  }

  /**
   * UPLUS(a) - Returns +a (identity)
   */
  public uplus(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UPLUS'),
      (a: number) => +a
    )
  }

  /**
   * UNARY_PERCENT(a) - Returns a / 100
   */
  public unaryPercent(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('UNARY_PERCENT'),
      (a: number) => a / 100
    )
  }

  /**
   * ISBETWEEN(val, lo, hi, loInc, hiInc) - Returns true if lo <= val <= hi (or < / > based on inclusive flags)
   */
  public isbetween(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ISBETWEEN'),
      (val: number, lo: number, hi: number, loInc: boolean, hiInc: boolean) => {
        const lowerOk = loInc ? val >= lo : val > lo
        const upperOk = hiInc ? val <= hi : val < hi
        return lowerOk && upperOk
      }
    )
  }
}
