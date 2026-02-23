/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType, SimpleCellAddress} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {AstNodeType, ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InternalScalarValue, InterpreterValue, isExtendedNumber, getRawValue, EmptyValue, NumberType, DateNumber} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible miscellaneous functions.
 *
 * Includes: GESTEP, LOOKUP, CELL, ENCODEURL, EPOCHTODATE, ISDATE
 */
export class GoogleSheetsMiscPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsMiscPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'GESTEP': {
      method: 'gestep',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 0},
      ],
    },
    'LOOKUP': {
      method: 'lookup',
      doesNotNeedArgumentsToBeComputed: true,
    },
    'CELL': {
      method: 'cell',
      doesNotNeedArgumentsToBeComputed: true,
    },
    'ENCODEURL': {
      method: 'encodeurl',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'EPOCHTODATE': {
      method: 'epochtodate',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, defaultValue: 1},
      ],
      returnNumberType: NumberType.NUMBER_DATE,
    },
    'ISDATE': {
      method: 'isdate',
      doesNotNeedArgumentsToBeComputed: true,
    },
  }

  /**
   * GESTEP(number, [step])
   *
   * Returns 1 if number >= step, else 0.
   */
  public gestep(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('GESTEP'), (number: number, step: number) => {
      return number >= step ? 1 : 0
    })
  }

  /**
   * LOOKUP(search_key, search_range, [result_range])
   *
   * Performs binary search in a sorted range (ascending order).
   * If 2 arguments: search_range is a 2-column array, search first column and return second.
   * If 3 arguments: search in search_range and return from result_range.
   */
  public lookup(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length < 2 || ast.args.length > 3) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    const key = this.evaluateAst(ast.args[0], state)
    if (key instanceof SimpleRangeValue || key instanceof CellError) {
      return key instanceof CellError ? key : new CellError(ErrorType.VALUE)
    }

    const searchRange = this.evaluateAst(ast.args[1], state)
    if (!(searchRange instanceof SimpleRangeValue)) {
      return new CellError(ErrorType.VALUE)
    }

    if (ast.args.length === 2) {
      // 2-argument form: search_range is 2-column array
      if (searchRange.width() < 2) {
        return new CellError(ErrorType.VALUE)
      }
      return this.lookupTwoArgument(key, searchRange)
    } else {
      // 3-argument form: separate search and result ranges
      const resultRange = this.evaluateAst(ast.args[2], state)
      if (!(resultRange instanceof SimpleRangeValue)) {
        return new CellError(ErrorType.VALUE)
      }
      return this.lookupThreeArgument(key, searchRange, resultRange)
    }
  }

  private lookupTwoArgument(key: InternalScalarValue, searchRange: SimpleRangeValue): InternalScalarValue {
    const keyNum = this.coerceScalarToNumberOrError(key)
    if (keyNum instanceof CellError) {
      return keyNum
    }

    const keyValue = getRawValue(keyNum)
    const values = searchRange.data

    // Binary search in first column for largest value <= key
    let bestRowIndex = -1
    for (let i = 0; i < values.length; i++) {
      const cellValue = values[i][0]
      const cellNum = this.coerceScalarToNumberOrError(cellValue)
      if (cellNum instanceof CellError) {
        continue
      }
      const cellValue_ = getRawValue(cellNum)
      if (cellValue_ <= keyValue) {
        bestRowIndex = i
      } else {
        break
      }
    }

    if (bestRowIndex === -1) {
      return new CellError(ErrorType.NA)
    }

    return values[bestRowIndex][1] ?? EmptyValue
  }

  private lookupThreeArgument(key: InternalScalarValue, searchRange: SimpleRangeValue, resultRange: SimpleRangeValue): InternalScalarValue {
    const keyNum = this.coerceScalarToNumberOrError(key)
    if (keyNum instanceof CellError) {
      return keyNum
    }

    const keyValue = getRawValue(keyNum)
    const searchValues = searchRange.data
    const resultValues = resultRange.data

    // Binary search in search range for largest value <= key
    let bestIndex = -1
    const flatSearchValues = searchValues.flat()
    for (let i = 0; i < flatSearchValues.length; i++) {
      const cellValue = flatSearchValues[i]
      const cellNum = this.coerceScalarToNumberOrError(cellValue)
      if (cellNum instanceof CellError) {
        continue
      }
      const cellValue_ = getRawValue(cellNum)
      if (cellValue_ <= keyValue) {
        bestIndex = i
      } else {
        break
      }
    }

    if (bestIndex === -1) {
      return new CellError(ErrorType.NA)
    }

    // Map to result range
    const resultFlat = resultValues.flat()
    return resultFlat[bestIndex] ?? EmptyValue
  }

  /**
   * CELL(info_type, reference)
   *
   * Returns metadata about a cell.
   * Supported info_type values: "contents", "address", "col", "row", "type"
   */
  public cell(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length !== 2) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    const infoTypeArg = ast.args[0]
    let infoType: string | undefined

    // Extract info_type as string
    if (infoTypeArg.type === AstNodeType.STRING) {
      infoType = infoTypeArg.value
    } else {
      const evaluated = this.evaluateAst(infoTypeArg, state)
      if (typeof evaluated === 'string') {
        infoType = evaluated
      } else {
        return new CellError(ErrorType.VALUE)
      }
    }

    // Extract reference
    const refArg = ast.args[1]
    let cellRef: SimpleCellAddress | undefined

    if (refArg.type === AstNodeType.CELL_REFERENCE) {
      cellRef = refArg.reference.toSimpleCellAddress(state.formulaAddress)
    } else {
      return new CellError(ErrorType.VALUE)
    }

    if (!cellRef) {
      return new CellError(ErrorType.VALUE)
    }

    // Return info based on type
    switch (infoType?.toLowerCase()) {
      case 'contents': {
        const cellValue = this.dependencyGraph.getCellValue(cellRef)
        return cellValue ?? EmptyValue
      }
      case 'address': {
        return `$${String.fromCharCode(65 + cellRef.col)}$${cellRef.row + 1}`
      }
      case 'col': {
        return cellRef.col + 1
      }
      case 'row': {
        return cellRef.row + 1
      }
      case 'type': {
        const cellValue = this.dependencyGraph.getCellValue(cellRef)
        if (cellValue === undefined || cellValue === EmptyValue) {
          return 'b'
        }
        if (typeof cellValue === 'string') {
          return 't'
        }
        if (typeof cellValue === 'boolean') {
          return 'l'
        }
        if (typeof cellValue === 'number' || isExtendedNumber(cellValue)) {
          return 'v'
        }
        if (cellValue instanceof CellError) {
          return 'e'
        }
        return 'v'
      }
      default:
        return new CellError(ErrorType.VALUE)
    }
  }

  /**
   * ENCODEURL(text)
   *
   * URL-encodes a string using encodeURIComponent.
   */
  public encodeurl(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ENCODEURL'), (text: string) => {
      return encodeURIComponent(text)
    })
  }

  /**
   * EPOCHTODATE(timestamp, [unit])
   *
   * Converts Unix timestamp to HyperFormula date serial number.
   * unit: 1=seconds (default), 2=milliseconds, 3=microseconds
   * HyperFormula epoch: 1899-12-30, Unix epoch: 1970-01-01 (difference: 25569 days)
   */
  public epochtodate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('EPOCHTODATE'), (timestamp: number, unit: number) => {
      // Determine divisor based on unit
      let divisor = 1
      if (unit === 2) {
        divisor = 1000
      } else if (unit === 3) {
        divisor = 1000000
      }

      // Convert to seconds, then to days, then add HF epoch offset
      const days = (timestamp / divisor) / 86400 + 25569
      return new DateNumber(days)
    })
  }

  /**
   * ISDATE(value)
   *
   * Returns TRUE if value is a date, FALSE otherwise.
   */
  public isdate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length !== 1) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    const valueArg = ast.args[0]
    const value = this.evaluateAst(valueArg, state)

    // Check if value is a DateNumber or similar date-typed number
    if (isExtendedNumber(value)) {
      const numberType = value instanceof DateNumber ? value.getDetailedType() : NumberType.NUMBER_RAW

      return numberType === NumberType.NUMBER_DATE || numberType === NumberType.NUMBER_DATETIME
    }

    // Try to parse string as date
    if (typeof value === 'string') {
      const parsed = this.dateTimeHelper.dateStringToDateNumber(value)
      return parsed !== undefined
    }

    return false
  }
}
