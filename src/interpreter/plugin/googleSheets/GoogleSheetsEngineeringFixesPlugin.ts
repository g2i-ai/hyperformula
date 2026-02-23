/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {padLeft} from '../../../format/format'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue} from '../../InterpreterValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

const MAX_LENGTH = 10

/**
 * Google Sheets-compatible engineering function overrides.
 *
 * Overrides HEX2BIN, HEX2DEC, HEX2OCT to accept lowercase hex strings,
 * matching Google Sheets behavior where hex input is case-insensitive.
 */
export class GoogleSheetsEngineeringFixesPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsEngineeringFixesPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'HEX2BIN': {
      method: 'hex2bin',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, optionalArg: true, minValue: 0, maxValue: MAX_LENGTH},
      ],
    },
    'HEX2DEC': {
      method: 'hex2dec',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'HEX2OCT': {
      method: 'hex2oct',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, optionalArg: true, minValue: 0, maxValue: MAX_LENGTH},
      ],
    },
  }

  /**
   * HEX2BIN — converts a hexadecimal string to binary, accepting lowercase input.
   */
  public hex2bin(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('HEX2BIN'),
      (hexadecimal: string, places: number | undefined) => {
        const upper = hexadecimal.toUpperCase()
        const hexadecimalWithSign = coerceHexString(upper)
        if (hexadecimalWithSign === undefined) {
          return new CellError(ErrorType.NUM, ErrorMessage.NotHex)
        }
        return decimalToBaseWithExactPadding(twoComplementToDecimal(hexadecimalWithSign, 16), 2, places)
      }
    )
  }

  /**
   * HEX2DEC — converts a hexadecimal string to decimal, accepting lowercase input.
   */
  public hex2dec(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('HEX2DEC'),
      (hexadecimal: string) => {
        const upper = hexadecimal.toUpperCase()
        const hexadecimalWithSign = coerceHexString(upper)
        if (hexadecimalWithSign === undefined) {
          return new CellError(ErrorType.NUM, ErrorMessage.NotHex)
        }
        return twoComplementToDecimal(hexadecimalWithSign, 16)
      }
    )
  }

  /**
   * HEX2OCT — converts a hexadecimal string to octal, accepting lowercase input.
   */
  public hex2oct(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('HEX2OCT'),
      (hexadecimal: string, places: number | undefined) => {
        const upper = hexadecimal.toUpperCase()
        const hexadecimalWithSign = coerceHexString(upper)
        if (hexadecimalWithSign === undefined) {
          return new CellError(ErrorType.NUM, ErrorMessage.NotHex)
        }
        return decimalToBaseWithExactPadding(twoComplementToDecimal(hexadecimalWithSign, 16), 8, places)
      }
    )
  }
}

/** Validates a hex string (uppercase) and returns it if valid, undefined otherwise. */
function coerceHexString(value: string): string | undefined {
  if (value.length > MAX_LENGTH || !/^[0-9A-F]+$/.test(value)) {
    return undefined
  }
  return value
}

function decimalToBaseWithExactPadding(value: number, base: number, places?: number): string | CellError {
  if (value > maxValFromBase(base)) {
    return new CellError(ErrorType.NUM, ErrorMessage.ValueBaseLarge)
  }
  if (value < minValFromBase(base)) {
    return new CellError(ErrorType.NUM, ErrorMessage.ValueBaseSmall)
  }
  const result = decimalToRadixComplement(value, base)
  if (places === undefined || value < 0) {
    return result
  } else if (result.length > places) {
    return new CellError(ErrorType.NUM, ErrorMessage.ValueBaseLong)
  } else {
    return padLeft(result, places)
  }
}

function minValFromBase(base: number): number {
  return -Math.pow(base, MAX_LENGTH) / 2
}

function maxValFromBase(base: number): number {
  return -minValFromBase(base) - 1
}

function decimalToRadixComplement(value: number, base: number): string {
  const offset = value < 0 ? Math.pow(base, MAX_LENGTH) : 0
  return (value + offset).toString(base).toUpperCase()
}

function twoComplementToDecimal(value: string, base: number): number {
  const parsed = parseInt(value, base)
  const offset = Math.pow(base, MAX_LENGTH)
  return (parsed >= offset / 2) ? parsed - offset : parsed
}
