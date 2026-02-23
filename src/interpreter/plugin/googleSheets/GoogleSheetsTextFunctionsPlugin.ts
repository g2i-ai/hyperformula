/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {CellError, ErrorType} from '../../../Cell'
import {ErrorMessage} from '../../../error-message'
import {ProcedureAst} from '../../../parser'
import {InterpreterState} from '../../InterpreterState'
import {InterpreterValue, RawScalarValue} from '../../InterpreterValue'
import {FunctionArgumentType, FunctionPlugin, FunctionPluginTypecheck, ImplementedFunctions} from '../FunctionPlugin'

/**
 * Google Sheets-compatible text functions plugin.
 *
 * Implements functions available in Google Sheets but not in standard HyperFormula:
 * REGEXMATCH, REGEXEXTRACT, REGEXREPLACE, JOIN, TEXTJOIN, DOLLAR, FIXED,
 * ASC, FINDB, LEFTB, LENB, MIDB, RIGHTB, REPLACEB, SEARCHB.
 */
export class GoogleSheetsTextFunctionsPlugin extends FunctionPlugin implements FunctionPluginTypecheck<GoogleSheetsTextFunctionsPlugin> {
  public static implementedFunctions: ImplementedFunctions = {
    'REGEXMATCH': {
      method: 'regexmatch',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'REGEXEXTRACT': {
      method: 'regexextract',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'REGEXREPLACE': {
      method: 'regexreplace',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'JOIN': {
      method: 'join',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'TEXTJOIN': {
      method: 'textjoin',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.BOOLEAN},
        {argumentType: FunctionArgumentType.SCALAR},
      ],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    'DOLLAR': {
      method: 'dollar',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, defaultValue: 2},
      ],
    },
    'FIXED': {
      method: 'fixed',
      parameters: [
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.INTEGER, defaultValue: 2},
        {argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false},
      ],
    },
    'ASC': {
      method: 'asc',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'FINDB': {
      method: 'findb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 1},
      ],
    },
    'LEFTB': {
      method: 'leftb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 1},
      ],
    },
    'LENB': {
      method: 'lenb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'MIDB': {
      method: 'midb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
      ],
    },
    'RIGHTB': {
      method: 'rightb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 1},
      ],
    },
    'REPLACEB': {
      method: 'replaceb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.NUMBER},
        {argumentType: FunctionArgumentType.STRING},
      ],
    },
    'SEARCHB': {
      method: 'searchb',
      parameters: [
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.STRING},
        {argumentType: FunctionArgumentType.NUMBER, defaultValue: 1},
      ],
    },
  }

  /**
   * REGEXMATCH(text, regex)
   *
   * Returns true if text matches the regular expression.
   */
  public regexmatch(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('REGEXMATCH'), (text: string, regex: string) => {
      try {
        return new RegExp(regex).test(text)
      } catch {
        return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
      }
    })
  }

  /**
   * REGEXEXTRACT(text, regex)
   *
   * Returns the first match of the regular expression in text.
   * Returns the first capture group if one exists, otherwise the full match.
   * Returns #N/A if no match is found.
   */
  public regexextract(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('REGEXEXTRACT'), (text: string, regex: string) => {
      let compiledRegex: RegExp
      try {
        compiledRegex = new RegExp(regex)
      } catch {
        return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
      }

      const match = text.match(compiledRegex)
      if (match === null) {
        return new CellError(ErrorType.NA, ErrorMessage.ValueNotFound)
      }

      // Return first capture group if present, otherwise the full match
      return match[1] !== undefined ? match[1] : match[0]
    })
  }

  /**
   * REGEXREPLACE(text, regex, replacement)
   *
   * Replaces all occurrences of regex in text with replacement.
   * Matches Google Sheets behavior which replaces all occurrences.
   */
  public regexreplace(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('REGEXREPLACE'), (text: string, regex: string, replacement: string) => {
      try {
        return text.replace(new RegExp(regex, 'g'), replacement)
      } catch {
        return new CellError(ErrorType.VALUE, ErrorMessage.WrongType)
      }
    })
  }

  /**
   * JOIN(delimiter, value1, [value2, ...])
   *
   * Joins elements with the specified delimiter.
   */
  public join(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('JOIN'), (delimiter: string, ...values: RawScalarValue[]) => {
      return values
        .filter(v => v !== null && v !== undefined)
        .map(v => String(v))
        .join(delimiter)
    })
  }

  /**
   * TEXTJOIN(delimiter, ignore_empty, text1, [text2, ...])
   *
   * Joins text strings with a delimiter, optionally ignoring empty values.
   */
  public textjoin(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('TEXTJOIN'), (delimiter: string, ignoreEmpty: boolean, ...texts: RawScalarValue[]) => {
      const parts = texts
        .filter(v => v !== null && v !== undefined)
        .map(v => String(v))
        .filter(s => !ignoreEmpty || s !== '')

      return parts.join(delimiter)
    })
  }

  /**
   * DOLLAR(number, [decimals])
   *
   * Formats a number as a USD currency string with the specified decimal places.
   */
  public dollar(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('DOLLAR'), (number: number, decimals: number) => {
      const fractionDigits = Math.max(0, decimals)
      return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD',
        minimumFractionDigits: fractionDigits,
        maximumFractionDigits: fractionDigits,
      }).format(number)
    })
  }

  /**
   * FIXED(number, [decimals], [no_commas])
   *
   * Formats a number with fixed decimal places.
   * When no_commas is true, suppresses the thousands separator.
   */
  public fixed(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('FIXED'), (number: number, decimals: number, noCommas: boolean) => {
      const fractionDigits = Math.max(0, decimals)
      return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: fractionDigits,
        maximumFractionDigits: fractionDigits,
        useGrouping: !noCommas,
      }).format(number)
    })
  }

  /**
   * ASC(text)
   *
   * Converts full-width ASCII and katakana characters to half-width equivalents.
   * Converts characters in U+FF01-U+FF5E to U+0021-U+007E.
   * Also converts the full-width space U+3000 to a regular space.
   */
  public asc(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('ASC'), (text: string) => {
      return text
        .replace(/[\uFF01-\uFF5E]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0xFEE0))
        .replace(/\u3000/g, ' ')
    })
  }

  /**
   * FINDB(find_text, within_text, [start_num])
   *
   * Returns the byte position of find_text within within_text.
   * Case-sensitive. 1-indexed. Returns #VALUE! if not found.
   */
  public findb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('FINDB'), (findText: string, withinText: string, startNum: number) => {
      if (startNum < 1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.LessThanOne)
      }

      const encoder = new TextEncoder()
      const withinBytes = encoder.encode(withinText)
      const findBytes = encoder.encode(findText)
      const startByteIndex = this.getByteIndices(withinText)[startNum - 1]

      if (startByteIndex === undefined) {
        return new CellError(ErrorType.VALUE, ErrorMessage.IndexBounds)
      }

      const searchSlice = withinBytes.slice(startByteIndex)
      const foundByteOffset = this.findBytesIndex(searchSlice, findBytes)

      if (foundByteOffset === -1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.PatternNotFound)
      }

      return startByteIndex + foundByteOffset + 1
    })
  }

  /**
   * LEFTB(text, [num_bytes])
   *
   * Returns the leftmost num_bytes bytes of text. Default is 1.
   */
  public leftb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LEFTB'), (text: string, numBytes: number) => {
      if (numBytes < 0) {
        return new CellError(ErrorType.VALUE, ErrorMessage.NegativeLength)
      }

      return this.sliceByBytes(text, 0, numBytes)
    })
  }

  /**
   * LENB(text)
   *
   * Returns the byte length of text using UTF-8 encoding.
   */
  public lenb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('LENB'), (text: string) => {
      return new TextEncoder().encode(text).length
    })
  }

  /**
   * MIDB(text, start_num, num_bytes)
   *
   * Returns a substring from text starting at byte position start_num, for num_bytes bytes.
   * start_num is 1-indexed.
   */
  public midb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('MIDB'), (text: string, startNum: number, numBytes: number) => {
      if (startNum < 1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.LessThanOne)
      }
      if (numBytes < 0) {
        return new CellError(ErrorType.VALUE, ErrorMessage.NegativeLength)
      }

      // startNum is a 1-indexed byte position; pass as byteOffset directly
      return this.sliceByBytes(text, 0, numBytes, startNum - 1)
    })
  }

  /**
   * RIGHTB(text, [num_bytes])
   *
   * Returns the rightmost num_bytes bytes of text. Default is 1.
   */
  public rightb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('RIGHTB'), (text: string, numBytes: number) => {
      if (numBytes < 0) {
        return new CellError(ErrorType.VALUE, ErrorMessage.NegativeLength)
      }

      const encoder = new TextEncoder()
      const totalBytes = encoder.encode(text).length
      const startByte = Math.max(0, totalBytes - numBytes)

      return this.sliceByBytes(text, 0, totalBytes, startByte)
    })
  }

  /**
   * REPLACEB(old_text, start_num, num_bytes, new_text)
   *
   * Replaces part of a text string based on byte position.
   * start_num is 1-indexed.
   */
  public replaceb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('REPLACEB'), (oldText: string, startNum: number, numBytes: number, newText: string) => {
      if (startNum < 1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.LessThanOne)
      }
      if (numBytes < 0) {
        return new CellError(ErrorType.VALUE, ErrorMessage.NegativeLength)
      }

      const encoder = new TextEncoder()
      const decoder = new TextDecoder()
      const bytes = encoder.encode(oldText)
      const startByte = startNum - 1
      const endByte = startByte + numBytes

      const before = decoder.decode(bytes.slice(0, startByte))
      const after = decoder.decode(bytes.slice(endByte))

      return before + newText + after
    })
  }

  /**
   * SEARCHB(find_text, within_text, [start_num])
   *
   * Returns the byte position of find_text within within_text.
   * Case-insensitive. 1-indexed. Returns #VALUE! if not found.
   */
  public searchb(ast: ProcedureAst, state: InterpreterState): InterpreterValue {
    return this.runFunction(ast.args, state, this.metadata('SEARCHB'), (findText: string, withinText: string, startNum: number) => {
      if (startNum < 1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.LessThanOne)
      }

      const lowerFind = findText.toLowerCase()
      const lowerWithin = withinText.toLowerCase()

      const encoder = new TextEncoder()
      const withinBytes = encoder.encode(lowerWithin)
      const findBytes = encoder.encode(lowerFind)
      const startByteIndex = this.getByteIndices(lowerWithin)[startNum - 1]

      if (startByteIndex === undefined) {
        return new CellError(ErrorType.VALUE, ErrorMessage.IndexBounds)
      }

      const searchSlice = withinBytes.slice(startByteIndex)
      const foundByteOffset = this.findBytesIndex(searchSlice, findBytes)

      if (foundByteOffset === -1) {
        return new CellError(ErrorType.VALUE, ErrorMessage.PatternNotFound)
      }

      return startByteIndex + foundByteOffset + 1
    })
  }

  /**
   * Returns an array mapping character index to byte offset.
   * Index 0 = byte offset of first char, index 1 = byte offset of second char, etc.
   * Array length is text.length + 1 (last entry is total byte length).
   */
  private getByteIndices(text: string): number[] {
    const encoder = new TextEncoder()
    const indices: number[] = [0]
    for (let i = 0; i < text.length; i++) {
      const charBytes = encoder.encode(text[i]).length
      indices.push(indices[indices.length - 1] + charBytes)
    }
    return indices
  }

  /**
   * Finds the byte index of needle within haystack, or -1 if not found.
   */
  private findBytesIndex(haystack: Uint8Array, needle: Uint8Array): number {
    if (needle.length === 0) {
      return 0
    }
    outer: for (let i = 0; i <= haystack.length - needle.length; i++) {
      for (let j = 0; j < needle.length; j++) {
        if (haystack[i + j] !== needle[j]) {
          continue outer
        }
      }
      return i
    }
    return -1
  }

  /**
   * Extracts a substring from text using byte-based slicing.
   * startCharIndex: character index to start from (0-based)
   * byteCount: number of bytes to include
   * byteOffset: optional absolute byte offset to start from (overrides startCharIndex)
   */
  private sliceByBytes(text: string, startCharIndex: number, byteCount: number, byteOffset?: number): string {
    const encoder = new TextEncoder()
    const decoder = new TextDecoder()
    const bytes = encoder.encode(text)
    const byteIndices = this.getByteIndices(text)

    const startByte = byteOffset !== undefined ? byteOffset : (byteIndices[startCharIndex] ?? bytes.length)
    const endByte = Math.min(startByte + byteCount, bytes.length)

    return decoder.decode(bytes.slice(startByte, endByte))
  }
}
