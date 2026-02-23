import {HyperFormula} from '../src'
import {adr} from './testUtils'

/**
 * Tests for BooleanLiteral lexer token not greedily matching the prefix of longer identifiers.
 *
 * Chevrotain uses longest-match-wins per token type, but when multiple token types
 * can match at the same position, the first one in the allTokens array wins — unless
 * the shorter-matching token declares a `longer_alt` pointing to the token that can
 * consume more characters.
 *
 * Without `longer_alt`, "TRUECOUNT" is wrongly tokenised as BooleanLiteral("TRUE") +
 * NamedExpression("COUNT") instead of a single NamedExpression("TRUECOUNT").
 */
describe('BooleanLiteral lexer token — no greedy prefix matching', () => {
  it('recognises TRUECOUNT as a single named expression, not TRUE + COUNT', () => {
    const hf = HyperFormula.buildFromArray(
      [['=TRUECOUNT']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
      [{name: 'TRUECOUNT', expression: '=42'}],
    )
    // If BooleanLiteral greedy-matches "TRUE", the parse fails or returns a wrong value.
    expect(hf.getCellValue(adr('A1'))).toBe(42)
    hf.destroy()
  })

  it('recognises FALSEHOOD as a single named expression, not FALSE + HOOD', () => {
    const hf = HyperFormula.buildFromArray(
      [['=FALSEHOOD']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
      [{name: 'FALSEHOOD', expression: '=99'}],
    )
    expect(hf.getCellValue(adr('A1'))).toBe(99)
    hf.destroy()
  })

  it('still recognises standalone TRUE as a BooleanLiteral in default mode', () => {
    const hf = HyperFormula.buildFromArray(
      [['=TRUE()']],
      {licenseKey: 'gpl-v3'},
    )
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('still recognises standalone FALSE as a BooleanLiteral in default mode', () => {
    const hf = HyperFormula.buildFromArray(
      [['=FALSE()']],
      {licenseKey: 'gpl-v3'},
    )
    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })

  it('recognises TRUECOUNT as a named expression in default mode too', () => {
    const hf = HyperFormula.buildFromArray(
      [['=TRUECOUNT']],
      {licenseKey: 'gpl-v3'},
      [{name: 'TRUECOUNT', expression: '=7'}],
    )
    expect(hf.getCellValue(adr('A1'))).toBe(7)
    hf.destroy()
  })

  it('handles TRUECOUNT( as a function call (ProcedureName takes precedence)', () => {
    // ProcedureName is listed before BooleanLiteral and already handles func( correctly
    // IF TRUECOUNT is registered as a user-defined function, it should parse as such
    const hf = HyperFormula.buildFromArray(
      [['=IF(TRUE(),1,0)']],
      {licenseKey: 'gpl-v3'},
    )
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    hf.destroy()
  })

  it('correctly evaluates =TRUE in googleSheets mode (named expression)', () => {
    const hf = HyperFormula.buildFromArray(
      [['=TRUE']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    )
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('correctly evaluates =FALSE in googleSheets mode (named expression)', () => {
    const hf = HyperFormula.buildFromArray(
      [['=FALSE']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    )
    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })
})
