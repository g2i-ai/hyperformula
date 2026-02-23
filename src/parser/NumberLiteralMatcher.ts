/**
 * @license
 * Copyright (c) 2025 Handsoncode. All rights reserved.
 */

/**
 * Custom lexer matcher for NumberLiteral tokens.
 *
 * When thousandSeparator collides with functionArgSeparator (e.g., both are comma
 * in Google Sheets en-US locale), this matcher uses a 3-digit-group heuristic to
 * disambiguate at lex time: `1,000` is a thousands-grouped number, while `1,00`
 * falls through to the argument separator token.
 */
export class NumberLiteralMatcher {
  readonly POSSIBLE_START_CHARACTERS = [
    '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '.', ',',
  ]

  private decimalSeparator: '.' | ',' = '.'
  private thousandSeparator: '' | ',' | ' ' | '.' = ''
  private thousandSeparatorEnabled = false
  private fallbackRegex: RegExp = this.buildFallbackRegex('.')

  /** Configures the matcher for the current locale settings. */
  configure(decimalSep: '.' | ',', thousandSep: '' | ',' | ' ' | '.'): void {
    this.decimalSeparator = decimalSep
    this.thousandSeparator = thousandSep
    this.thousandSeparatorEnabled = thousandSep !== ''
    this.fallbackRegex = this.buildFallbackRegex(decimalSep)
  }

  /**
   * Chevrotain custom token matcher.
   * Returns a RegExpExecArray-like result or null.
   */
  match(text: string, startOffset: number): RegExpExecArray | null {
    if (this.thousandSeparatorEnabled) {
      return this.matchWithThousands(text, startOffset)
    }
    return this.matchFallback(text, startOffset)
  }

  private buildFallbackRegex(decimalSep: '.' | ','): RegExp {
    const sep = decimalSep === '.' ? '\\.' : ','
    return new RegExp(`(${sep}\\d+|\\d+(${sep}\\d*)?)(e[+-]?\\d+)?`, 'y')
  }

  /** Fast path: standard regex when no thousands separator collision. */
  private matchFallback(text: string, startOffset: number): RegExpExecArray | null {
    this.fallbackRegex.lastIndex = startOffset
    return this.fallbackRegex.exec(text)
  }

  /**
   * Thousands-aware matching using character-by-character parsing.
   *
   * Algorithm:
   * 1. If starts with decimalSep → consume decimalSep + digits → scientific → done
   * 2. If starts with digit → consume leading digits, count them
   * 3. If leadingDigitCount > 3 → skip thousands grouping
   * 4. If leadingDigitCount <= 3 AND next char is thousandSep:
   *    a. Lookahead: exactly 3 digits after sep AND char after those NOT a digit?
   *    b. Yes → consume sep + 3 digits, repeat from 4
   *    c. No → stop
   * 5. Optional decimal part (decimalSep + digits)
   * 6. Optional scientific notation (e/E ± digits)
   * 7. Build RegExpExecArray result
   */
  private matchWithThousands(text: string, startOffset: number): RegExpExecArray | null {
    let pos = startOffset
    const len = text.length

    if (pos >= len) {
      return null
    }

    const ch = text[pos]

    // Case 1: Starts with decimal separator (e.g., ".5" or ",5")
    if (ch === this.decimalSeparator) {
      if (pos + 1 >= len || !isDigit(text[pos + 1])) {
        return null
      }
      pos++ // consume decimal sep
      pos = consumeDigits(text, pos, len)
      pos = this.consumeScientific(text, pos, len)
      return buildResult(text, startOffset, pos)
    }

    // Case 2: Must start with a digit
    if (!isDigit(ch)) {
      return null
    }

    // Consume leading digit group
    const leadingStart = pos
    pos = consumeDigits(text, pos, len)
    const leadingCount = pos - leadingStart

    // Thousands grouping: only if leading group is 1-3 digits
    if (leadingCount <= 3) {
      pos = this.consumeThousandsGroups(text, pos, len)
    }

    // Optional decimal part
    if (pos < len && text[pos] === this.decimalSeparator) {
      pos++ // consume decimal sep
      pos = consumeDigits(text, pos, len)
    }

    // Optional scientific notation
    pos = this.consumeScientific(text, pos, len)

    // Must have consumed at least something
    if (pos === startOffset) {
      return null
    }

    return buildResult(text, startOffset, pos)
  }

  /**
   * Greedily consumes thousands-separated groups (e.g., `,000,000`).
   * Each group must be exactly 3 digits, and the character after those 3 digits
   * must NOT be a digit (to avoid matching `1,0000`).
   */
  private consumeThousandsGroups(text: string, pos: number, len: number): number {
    while (pos < len && text[pos] === this.thousandSeparator) {
      // Fewer than 3 characters remain after the separator — cannot form a valid group
      if (pos + 3 >= len) {
        break
      }

      // Check that next 3 chars are digits
      if (!allDigits(text, pos + 1, pos + 4)) {
        break
      }

      // Check that char after the 3 digits is NOT a digit
      if (pos + 4 < len && isDigit(text[pos + 4])) {
        break
      }

      // Consume separator + 3 digits
      pos += 4
    }
    return pos
  }

  /** Consumes optional scientific notation (e.g., `e5`, `E-3`, `e+10`). */
  private consumeScientific(text: string, pos: number, len: number): number {
    if (pos < len && (text[pos] === 'e' || text[pos] === 'E')) {
      let nextPos = pos + 1
      if (nextPos < len && (text[nextPos] === '+' || text[nextPos] === '-')) {
        nextPos++
      }
      if (nextPos < len && isDigit(text[nextPos])) {
        nextPos = consumeDigits(text, nextPos, len)
        return nextPos
      }
    }
    return pos
  }
}

function isDigit(ch: string): boolean {
  return ch >= '0' && ch <= '9'
}

function consumeDigits(text: string, pos: number, len: number): number {
  while (pos < len && isDigit(text[pos])) {
    pos++
  }
  return pos
}

function allDigits(text: string, from: number, to: number): boolean {
  for (let i = from; i < to; i++) {
    if (!isDigit(text[i])) {
      return false
    }
  }
  return true
}

/** Builds a RegExpExecArray-compatible result for Chevrotain. */
function buildResult(text: string, startOffset: number, endOffset: number): RegExpExecArray | null {
  if (endOffset <= startOffset) {
    return null
  }
  const matched = text.substring(startOffset, endOffset)
  const result = [matched] as RegExpExecArray
  result.index = startOffset
  result.input = text
  return result
}
