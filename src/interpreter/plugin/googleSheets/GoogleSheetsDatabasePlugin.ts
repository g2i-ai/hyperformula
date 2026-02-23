/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {EmptyValue, InternalScalarValue, InterpreterValue} from '../../InterpreterValue'
import {SimpleRangeValue} from '../../../SimpleRangeValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/** Parsed criterion with operator and comparison value. */
interface ParsedCriterion {
  operator: '=' | '<>' | '<' | '<=' | '>' | '>='
  value: string | number
  isWildcard: boolean
}

/** Result of extracting field values from the database matching criteria. */
interface DatabaseResult {
  values: InternalScalarValue[]
  error?: CellError
}

/** The database function parameter definition — shared across all 12 functions. */
const DATABASE_PARAMETERS = [
  {argumentType: FunctionArgumentType.RANGE},
  {argumentType: FunctionArgumentType.ANY},
  {argumentType: FunctionArgumentType.RANGE},
]

/**
 * Parses a criterion string into a structured form with operator and value.
 *
 * Handles comparison operators (>, <, >=, <=, <>) and wildcard patterns (* and ?).
 */
function parseCriterion(raw: string): ParsedCriterion {
  const comparisonMatch = /^(>=|<=|<>|>|<|=)(.*)$/.exec(raw)

  if (comparisonMatch) {
    const operator = comparisonMatch[1] as ParsedCriterion['operator']
    const valueStr = comparisonMatch[2]
    const numericValue = Number(valueStr)
    const value = isNaN(numericValue) ? valueStr : numericValue
    return {operator, value, isWildcard: false}
  }

  const isWildcard = /[*?]/.test(raw)
  const numericValue = Number(raw)
  const value = !isWildcard && !isNaN(numericValue) && raw.trim() !== '' ? numericValue : raw
  return {operator: '=', value, isWildcard}
}

/**
 * Converts a wildcard pattern (using * and ?) into a RegExp for matching.
 */
function wildcardToRegExp(pattern: string): RegExp {
  const escaped = pattern
    .replace(/[.+^${}()|[\]\\]/g, '\\$&')
    .replace(/\*/g, '.*')
    .replace(/\?/g, '.')
  return new RegExp(`^${escaped}$`, 'i')
}

/**
 * Tests whether a cell value matches a parsed criterion.
 */
function matchesCriterion(cellValue: InternalScalarValue, criterion: ParsedCriterion): boolean {
  if (cellValue === null || cellValue === undefined) {
    return false
  }

  const rawCellValue = cellValue instanceof Object && 'val' in cellValue
    ? (cellValue as {val: number}).val
    : cellValue

  const {operator, value, isWildcard} = criterion

  if (operator === '=') {
    if (isWildcard && typeof value === 'string' && (typeof rawCellValue === 'string')) {
      return wildcardToRegExp(value).test(rawCellValue)
    }
    if (typeof rawCellValue === 'string' && typeof value === 'string') {
      return rawCellValue.toLowerCase() === value.toLowerCase()
    }
    return rawCellValue === value
  }

  // Handle <> for strings (not-equal comparison)
  if (operator === '<>') {
    if (typeof rawCellValue === 'string' && typeof value === 'string') {
      return rawCellValue.toLowerCase() !== value.toLowerCase()
    }
    return rawCellValue !== value
  }

  const numCell = Number(rawCellValue)
  const numValue = Number(value)

  if (isNaN(numCell) || isNaN(numValue)) {
    return false
  }

  switch (operator) {
    case '<': return numCell < numValue
    case '<=': return numCell <= numValue
    case '>': return numCell > numValue
    case '>=': return numCell >= numValue
    default: return false
  }
}

/**
 * Finds the column index in `headers` that matches the given `fieldName` (case-insensitive).
 * Returns -1 if not found.
 */
function findColumnIndexByName(headers: InternalScalarValue[], fieldName: string): number {
  return headers.findIndex(h => {
    if (h === null || h === undefined || h === EmptyValue) return false
    const rawH = h instanceof Object && 'val' in h ? (h as {val: number}).val : h
    return typeof rawH === 'string' && rawH.toLowerCase() === fieldName.toLowerCase()
  })
}

/**
 * Resolves the 0-based column index from a field argument (either a 1-based number or a header name string).
 */
function resolveFieldColumnIndex(fieldArg: InterpreterValue, headers: InternalScalarValue[]): number | CellError {
  if (fieldArg instanceof SimpleRangeValue) {
    return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
  }

  if (fieldArg instanceof CellError) {
    return fieldArg
  }

  const rawField = fieldArg instanceof Object && 'val' in fieldArg
    ? (fieldArg as {val: number}).val
    : fieldArg

  if (typeof rawField === 'number') {
    const colIndex = Math.floor(rawField) - 1
    if (colIndex < 0 || colIndex >= headers.length) {
      return new CellError(ErrorType.VALUE, ErrorMessage.IndexBounds)
    }
    return colIndex
  }

  if (typeof rawField === 'string') {
    const idx = findColumnIndexByName(headers, rawField)
    if (idx === -1) {
      return new CellError(ErrorType.VALUE, ErrorMessage.ValueNotFound)
    }
    return idx
  }

  return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
}

/**
 * Determines which database row indices match the given criteria range.
 *
 * Criteria range: first row is headers, subsequent rows are criteria values.
 * Multiple rows = OR logic; multiple columns in same row = AND logic.
 */
function findMatchingRowIndices(
  dbHeaders: InternalScalarValue[],
  dbDataRows: InternalScalarValue[][],
  criteriaData: InternalScalarValue[][]
): number[] {
  if (criteriaData.length < 2) {
    // No criteria rows — return all rows
    return dbDataRows.map((_, i) => i)
  }

  const criteriaHeaders = criteriaData[0]
  const criteriaRows = criteriaData.slice(1)

  // Map criteria column indices to database column indices
  const criteriaColumnMap = criteriaHeaders.map(header => {
    if (header === null || header === undefined || header === EmptyValue) return -1
    const rawHeader = header instanceof Object && 'val' in header
      ? (header as {val: number}).val
      : header
    if (typeof rawHeader !== 'string') return -1
    return findColumnIndexByName(dbHeaders, rawHeader)
  })

  return dbDataRows.reduce<number[]>((matchingIndices, row, rowIdx) => {
    // OR logic across criteria rows
    const rowMatchesAnyCriterion = criteriaRows.some(criteriaRow => {
      // AND logic across columns in same criteria row
      return criteriaColumnMap.every((dbColIdx, criteriaColIdx) => {
        if (dbColIdx === -1) return true // Unknown header, skip

        const criteriaCell = criteriaRow[criteriaColIdx]
        if (criteriaCell === null || criteriaCell === undefined || criteriaCell === '' || criteriaCell === EmptyValue) {
          return true // Empty criterion matches everything
        }

        const rawCriteria = criteriaCell instanceof Object && 'val' in criteriaCell
          ? String((criteriaCell as {val: number}).val)
          : String(criteriaCell)

        const parsedCriterion = parseCriterion(rawCriteria)
        return matchesCriterion(row[dbColIdx], parsedCriterion)
      })
    })

    if (rowMatchesAnyCriterion) {
      matchingIndices.push(rowIdx)
    }

    return matchingIndices
  }, [])
}

/**
 * Google Sheets-compatible database function implementations.
 *
 * All functions use the signature: DFUNC(database, field, criteria)
 * - database: a range where the first row contains column headers
 * - field: column to aggregate (1-based index or header name)
 * - criteria: a range where the first row contains headers and remaining rows are criteria values
 */
export class GoogleSheetsDatabasePlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsDatabasePlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'DAVERAGE': {
      method: 'daverage',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DCOUNT': {
      method: 'dcount',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DCOUNTA': {
      method: 'dcounta',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DGET': {
      method: 'dget',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DMAX': {
      method: 'dmax',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DMIN': {
      method: 'dmin',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DPRODUCT': {
      method: 'dproduct',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DSTDEV': {
      method: 'dstdev',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DSTDEVP': {
      method: 'dstdevp',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DSUM': {
      method: 'dsum',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DVAR': {
      method: 'dvar',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
    'DVARP': {
      method: 'dvarp',
      parameters: DATABASE_PARAMETERS,
      doesNotNeedArgumentsToBeComputed: true,
      vectorizationForbidden: true,
    },
  }

  /**
   * Extracts numeric values from the field column of all database rows that match the criteria.
   *
   * This is the shared core logic for all database aggregate functions.
   */
  private getDatabaseValues(ast: ProcedureAst, state: InterpreterState): DatabaseResult {
    if (ast.args.length !== 3) {
      return {values: [], error: new CellError(ErrorType.NA, ErrorMessage.WrongArgNumber)}
    }

    const dbArg = this.evaluateAst(ast.args[0], state)
    const fieldArg = this.evaluateAst(ast.args[1], state)
    const criteriaArg = this.evaluateAst(ast.args[2], state)

    if (!(dbArg instanceof SimpleRangeValue)) {
      return {values: [], error: new CellError(ErrorType.VALUE, ErrorMessage.CellRangeExpected)}
    }

    if (!(criteriaArg instanceof SimpleRangeValue)) {
      return {values: [], error: new CellError(ErrorType.VALUE, ErrorMessage.CellRangeExpected)}
    }

    const dbData = dbArg.data
    if (dbData.length < 2) {
      // No data rows (only headers or empty)
      return {values: []}
    }

    const dbHeaders = dbData[0]
    const dbDataRows = dbData.slice(1)

    const fieldColIndexOrError = resolveFieldColumnIndex(fieldArg, dbHeaders)
    if (fieldColIndexOrError instanceof CellError) {
      return {values: [], error: fieldColIndexOrError}
    }

    const criteriaData = criteriaArg.data
    const matchingIndices = findMatchingRowIndices(dbHeaders, dbDataRows, criteriaData)

    const values = matchingIndices.map(idx => dbDataRows[idx][fieldColIndexOrError])

    return {values}
  }

  /**
   * Extracts only numeric values from the database results (skips non-numeric entries).
   */
  private getNumericValues(ast: ProcedureAst, state: InterpreterState): {nums: number[], error?: CellError} {
    const result = this.getDatabaseValues(ast, state)
    if (result.error) {
      return {nums: [], error: result.error}
    }

    const nums = result.values.reduce<number[]>((acc, val) => {
      const raw = val instanceof Object && 'val' in val ? (val as {val: number}).val : val
      if (typeof raw === 'number') {
        acc.push(raw)
      }
      return acc
    }, [])

    return {nums}
  }

  /** DSUM — sum of numeric values in the field column for matching rows. */
  public dsum(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    return nums.reduce((sum, n) => sum + n, 0)
  }

  /** DAVERAGE — arithmetic mean of numeric values in the field column for matching rows. */
  public daverage(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.EmptyRange)
    return nums.reduce((sum, n) => sum + n, 0) / nums.length
  }

  /** DCOUNT — count of rows with numeric values in the field column that match the criteria. */
  public dcount(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    return nums.length
  }

  /** DCOUNTA — count of rows with non-empty values in the field column that match the criteria. */
  public dcounta(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const result = this.getDatabaseValues(ast, state)
    if (result.error) return result.error

    const nonEmptyCount = result.values.filter(val => {
      const raw = val instanceof Object && 'val' in val ? (val as {val: number}).val : val
      return raw !== null && raw !== undefined && raw !== '' && typeof raw !== 'symbol'
    }).length

    return nonEmptyCount
  }

  /**
   * DGET — returns the single value from the field column of matching rows.
   * Returns #NUM! if no rows match; returns #NUM! if multiple rows match.
   */
  public dget(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const result = this.getDatabaseValues(ast, state)
    if (result.error) return result.error

    if (result.values.length === 0) {
      return new CellError(ErrorType.NUM, ErrorMessage.ValueNotFound)
    }

    if (result.values.length > 1) {
      return new CellError(ErrorType.NUM, ErrorMessage.OneValue)
    }

    const val = result.values[0]
    const raw = val instanceof Object && 'val' in val ? (val as {val: number}).val : val
    return raw as InternalScalarValue
  }

  /** DMAX — maximum numeric value in the field column for matching rows. */
  public dmax(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return 0
    return Math.max(...nums)
  }

  /** DMIN — minimum numeric value in the field column for matching rows. */
  public dmin(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return 0
    return Math.min(...nums)
  }

  /** DPRODUCT — product of numeric values in the field column for matching rows. */
  public dproduct(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return 0
    return nums.reduce((product, n) => product * n, 1)
  }

  /** DSTDEV — sample standard deviation of numeric values in the field column for matching rows. */
  public dstdev(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length < 2) return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.TwoValues)
    return Math.sqrt(sampleVariance(nums))
  }

  /** DSTDEVP — population standard deviation of numeric values in the field column for matching rows. */
  public dstdevp(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.EmptyRange)
    return Math.sqrt(populationVariance(nums))
  }

  /** DVAR — sample variance of numeric values in the field column for matching rows. */
  public dvar(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length < 2) return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.TwoValues)
    return sampleVariance(nums)
  }

  /** DVARP — population variance of numeric values in the field column for matching rows. */
  public dvarp(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    const {nums, error} = this.getNumericValues(ast, state)
    if (error) return error
    if (nums.length === 0) return new CellError(ErrorType.DIV_BY_ZERO, ErrorMessage.EmptyRange)
    return populationVariance(nums)
  }
}

/**
 * Computes the sample variance (divides by n-1) of a numeric array.
 */
function sampleVariance(nums: number[]): number {
  const mean = nums.reduce((sum, n) => sum + n, 0) / nums.length
  const squaredDiffs = nums.map(n => (n - mean) ** 2)
  return squaredDiffs.reduce((sum, d) => sum + d, 0) / (nums.length - 1)
}

/**
 * Computes the population variance (divides by n) of a numeric array.
 */
function populationVariance(nums: number[]): number {
  const mean = nums.reduce((sum, n) => sum + n, 0) / nums.length
  const squaredDiffs = nums.map(n => (n - mean) ** 2)
  return squaredDiffs.reduce((sum, d) => sum + d, 0) / nums.length
}
