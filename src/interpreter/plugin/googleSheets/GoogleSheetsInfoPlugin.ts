/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {AstNodeType, ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue, isExtendedNumber, EmptyValue} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {ArrayFormulaVertex} from '../../../DependencyGraph'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible info functions.
 *
 * Implements functions for type checking and validation:
 * - ERROR.TYPE: Returns numeric code for error types
 * - TYPE: Returns numeric code for value types
 * - ISEMAIL: Validates email addresses
 * - ISURL: Validates URLs
 */
export class GoogleSheetsInfoPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsInfoPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'ERROR.TYPE': {
      method: 'errorType',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR}
      ],
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'TYPE': {
      method: 'type',
      parameters: [
        {argumentType: FunctionArgumentType.SCALAR}
      ],
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'ISEMAIL': {
      method: 'isemail',
      parameters: [
        {argumentType: FunctionArgumentType.STRING}
      ]
    },
    'ISURL': {
      method: 'isurl',
      parameters: [
        {argumentType: FunctionArgumentType.STRING}
      ]
    },
  }

  /**
   * Corresponds to ERROR.TYPE(value)
   *
   * Returns a number representing the type of error.
   * Returns #N/A if the value is not an error.
   *
   * Error type mappings:
   * - NULL! → 1
   * - DIV/0! → 2
   * - VALUE! → 3
   * - REF! → 4
   * - NAME? → 5
   * - NUM! → 6
   * - N/A → 7
   *
   * @param {ProcedureAst} ast - The procedure AST node
   * @param {InterpreterState} state - The interpreter state
   * @returns {InterpreterValue} Numeric error code or #N/A error
   */
  public errorType(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length !== 1) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    const arg = this.evaluateAst(ast.args[0], state)
    let val = arg

    if (arg instanceof SimpleRangeValue) {
      val = arg.data[0]?.[0] ?? EmptyValue
    }

    if (val instanceof CellError) {
      const mapping: Record<string, number> = {
        'ERROR': 1,
        'DIV_BY_ZERO': 2,
        'VALUE': 3,
        'REF': 4,
        'NAME': 5,
        'NUM': 6,
        'NA': 7,
      }
      return mapping[val.type] ?? new CellError(ErrorType.NA)
    }

    return new CellError(ErrorType.NA)
  }

  /**
   * Corresponds to TYPE(value)
   *
   * Returns a number representing the type of the value.
   * Returns #N/A if the value type is not recognized.
   *
   * Type mappings:
   * - number → 1
   * - text → 2
   * - boolean → 4
   * - error → 16
   * - array → 64
   *
   * @param {ProcedureAst} ast - The procedure AST node
   * @param {InterpreterState} state - The interpreter state
   * @returns {InterpreterValue} Numeric type code or #N/A error
   */
  public type(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length !== 1) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    // Detect if argument is a cell reference pointing to an array formula vertex
    const argNode = ast.args[0]
    if (argNode.type === AstNodeType.CELL_REFERENCE) {
      const address = argNode.reference.toSimpleCellAddress(state.formulaAddress)
      const vertex = this.dependencyGraph.getCell(address)
      if (vertex instanceof ArrayFormulaVertex) {
        return 64 // Array type
      }
    }

    const val = this.evaluateAst(argNode, state)

    if (val instanceof SimpleRangeValue) {
      return 64 // Array type (inline array literal)
    }

    if (val instanceof CellError) {
      return 16 // Error type
    }

    if (isExtendedNumber(val)) {
      return 1 // Number type
    }

    if (typeof val === 'string') {
      // Empty string is treated as empty (number type 1, like 0)
      if (val === '') {
        return 1
      }
      return 2 // Text type
    }

    if (typeof val === 'boolean') {
      return 4 // Boolean type
    }

    if (val === EmptyValue) {
      return 1 // Empty is treated as number type (0)
    }

    return new CellError(ErrorType.NA)
  }

  /**
   * Corresponds to ISEMAIL(text)
   *
   * Checks whether a value is a valid email address.
   * Uses basic email validation pattern.
   *
   * @param {ProcedureAst} ast - The procedure AST node
   * @param {InterpreterState} state - The interpreter state
   * @returns {InterpreterValue} True if email is valid, false otherwise
   */
  public isemail(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ISEMAIL'), (email: string) => {
      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/
      return emailRegex.test(email)
    })
  }

  /**
   * Corresponds to ISURL(text)
   *
   * Checks whether a value is a valid URL.
   * Validates basic HTTP/HTTPS URL format.
   *
   * @param {ProcedureAst} ast - The procedure AST node
   * @param {InterpreterState} state - The interpreter state
   * @returns {InterpreterValue} True if URL is valid, false otherwise
   */
  public isurl(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ISURL'), (url: string) => {
      const urlRegex = /^https?:\/\/.+/i
      return urlRegex.test(url)
    })
  }
}
