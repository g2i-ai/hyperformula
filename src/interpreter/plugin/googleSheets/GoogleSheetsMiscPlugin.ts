/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType, SimpleCellAddress} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {AstNodeType, ProcedureAst} from '../../../parser'
import {columnIndexToLabel} from '../../../parser/addressRepresentationConverters'
import {InterpreterState} from '../../InterpreterState'
import {InternalScalarValue, InterpreterValue, isExtendedNumber, EmptyValue, NumberType, DateNumber, RichNumber, getTypeOfExtendedNumber} from '../../InterpreterValue'
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
   * If 2 arguments: search_range is a 2D array; searches first column (or first
   *   row if cols > rows) and returns from the last column (or last row).
   * If 3 arguments: search in search_range and return from result_range.
   *
   * Supports both numeric and string keys.
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
      // 2-argument form: search_range is 2D array
      // If cols > rows: search first row, return last row
      // Otherwise: search first column, return last column
      const isHorizontal = searchRange.width() > searchRange.height()
      if (!isHorizontal && searchRange.width() < 2) {
        return new CellError(ErrorType.VALUE)
      }
      if (isHorizontal && searchRange.height() < 2) {
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

  /**
   * Returns true when cellVal <= key using natural type ordering.
   * CellError values are treated as non-matching (search continues past them).
   * EmptyValue is delegated to ArithmeticHelper which coerces it appropriately.
   */
  private lookupValueLessThanOrEqual(cellVal: InternalScalarValue, key: InternalScalarValue): boolean {
    if (cellVal instanceof CellError || key instanceof CellError) {
      return false
    }
    return this.arithmeticHelper.leq(cellVal, key)
  }

  private lookupTwoArgument(key: InternalScalarValue, searchRange: SimpleRangeValue): InternalScalarValue {
    if (key instanceof CellError) {
      return key
    }

    const values = searchRange.data
    const isHorizontal = searchRange.width() > searchRange.height()

    if (isHorizontal) {
      // Search across first row, return corresponding value from last row
      const firstRow = values[0]
      const lastRow = values[values.length - 1]
      let bestColIndex = -1
      for (let i = 0; i < firstRow.length; i++) {
        if (this.lookupValueLessThanOrEqual(firstRow[i], key)) {
          bestColIndex = i
        } else {
          break
        }
      }
      if (bestColIndex === -1) {
        return new CellError(ErrorType.NA)
      }
      return lastRow[bestColIndex] ?? EmptyValue
    }

    // Search down first column, return corresponding value from last column
    const lastColIndex = values[0].length - 1
    let bestRowIndex = -1
    for (let i = 0; i < values.length; i++) {
      if (this.lookupValueLessThanOrEqual(values[i][0], key)) {
        bestRowIndex = i
      } else {
        break
      }
    }

    if (bestRowIndex === -1) {
      return new CellError(ErrorType.NA)
    }

    return values[bestRowIndex][lastColIndex] ?? EmptyValue
  }

  private lookupThreeArgument(key: InternalScalarValue, searchRange: SimpleRangeValue, resultRange: SimpleRangeValue): InternalScalarValue {
    if (key instanceof CellError) {
      return key
    }

    const searchValues = searchRange.data
    const resultValues = resultRange.data

    // Linear scan (ascending order assumed) for largest value <= key
    let bestIndex = -1
    const flatSearchValues = searchValues.flat()
    for (let i = 0; i < flatSearchValues.length; i++) {
      const cellValue = flatSearchValues[i]
      if (this.lookupValueLessThanOrEqual(cellValue, key)) {
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
   *
   * CELL("type") returns Google Sheets codes:
   *   'b' — blank (empty cell)
   *   'l' — label (text string)
   *   'v' — value (number, boolean, date, or formula error)
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
        const colLabel = columnIndexToLabel(cellRef.col)
        return `$${colLabel}$${cellRef.row + 1}`
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
          return 'l'
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
   *
   * Returns VALUE error for unit values outside {1, 2, 3}.
   */
  public epochtodate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('EPOCHTODATE'), (timestamp: number, unit: number) => {
      const divisors: Record<number, number> = {1: 1, 2: 1000, 3: 1000000}
      const divisor = divisors[unit]

      if (divisor === undefined) {
        return new CellError(ErrorType.VALUE, ErrorMessage.BadMode)
      }

      // Convert to seconds, then to days, then add HF epoch offset
      const days = (timestamp / divisor) / 86400 + 25569
      return new DateNumber(days)
    })
  }

  /**
   * ISDATE(value)
   *
   * Returns TRUE if value is a date or datetime typed number, FALSE otherwise.
   * Only typed date values (DateNumber, DateTimeNumber) return TRUE — strings
   * are NOT parsed, matching Google Sheets behavior.
   */
  public isdate(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    if (ast.args.length !== 1) {
      return new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)
    }

    const value = this.evaluateAst(ast.args[0], state)

    if (!(value instanceof RichNumber)) {
      return false
    }

    const numberType = getTypeOfExtendedNumber(value)
    return numberType === NumberType.NUMBER_DATE || numberType === NumberType.NUMBER_DATETIME
  }
}
