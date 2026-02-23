/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

import {
  ABSOLUTE_OPERATOR, ALL_DIGITS_ARRAY, ALL_UNICODE_LETTERS_ARRAY, CELL_REFERENCE_WITH_NEXT_CHARACTER_PATTERN
} from './parser-consts'
import {columnLabelToIndex} from './addressRepresentationConverters'

/**
 * Helper class for recognizing CellReference token in text
 */
export class CellReferenceMatcher {
  readonly POSSIBLE_START_CHARACTERS = [
    ...ALL_UNICODE_LETTERS_ARRAY,
    ...ALL_DIGITS_ARRAY,
    ABSOLUTE_OPERATOR,
    "'",
    '_',
  ]

  private cellReferenceRegexp = new RegExp(CELL_REFERENCE_WITH_NEXT_CHARACTER_PATTERN, 'y')
  private compatibilityMode: 'default' | 'googleSheets' = 'default'
  private maxColumns: number = 18278

  /** Configures bounds checking for GSheets mode */
  configure(compatibilityMode: 'default' | 'googleSheets', maxColumns: number): void {
    this.compatibilityMode = compatibilityMode
    this.maxColumns = maxColumns
  }

  /**
   * Method used by the lexer to recognize CellReference token in text.
   * In GSheets mode, rejects references whose column exceeds maxColumns,
   * allowing them to fall through to NamedExpression.
   */
  match(text: string, startOffset: number): RegExpExecArray | null {
    this.cellReferenceRegexp.lastIndex = startOffset

    const execResult = this.cellReferenceRegexp.exec(text+'@')

    if (execResult == null || execResult[1] == null) {
      return null
    }

    if (this.compatibilityMode === 'googleSheets') {
      const matched = execResult[1]
      const colMatch = matched.match(/\$?([A-Za-z]+)\$?[0-9]+$/)
      if (colMatch) {
        const colIndex = columnLabelToIndex(colMatch[1])
        if (colIndex >= this.maxColumns) {
          return null
        }
      }
    }

    execResult[0] = execResult[1]
    return execResult
  }
}
