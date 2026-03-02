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
import {
  coerceStringToBase,
  decimalToBaseWithExactPadding,
  twoComplementToDecimal,
} from '../RadixConversionPlugin'

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
        const hexadecimalWithSign = coerceStringToBase(upper, 16, MAX_LENGTH)
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
        const hexadecimalWithSign = coerceStringToBase(upper, 16, MAX_LENGTH)
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
        const hexadecimalWithSign = coerceStringToBase(upper, 16, MAX_LENGTH)
        if (hexadecimalWithSign === undefined) {
          return new CellError(ErrorType.NUM, ErrorMessage.NotHex)
        }
        return decimalToBaseWithExactPadding(twoComplementToDecimal(hexadecimalWithSign, 16), 8, places)
      }
    )
  }
}
