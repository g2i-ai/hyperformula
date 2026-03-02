import {HyperFormula} from '../src'
import {adr} from './testUtils'

/**
 * Tests for per-instance CellReferenceMatcher isolation.
 *
 * Previously, a module-level `cellReferenceMatcher` singleton was reconfigured
 * (compatibilityMode, maxColumns) every time `buildLexerConfig` was called. Because the
 * CellReference token's `pattern` function closed over this singleton, building a second
 * HyperFormula instance overwrote the configuration used by the first instance's lexer.
 *
 * This caused two concrete bugs:
 * 1. After building a googleSheets(maxColumns=5) instance, a default-mode instance could
 *    no longer reference columns beyond J (index 9+) in newly-parsed formulas — those
 *    column labels were incorrectly rejected as NamedExpressions.
 * 2. After building a default-mode instance, a googleSheets(maxColumns=5) instance would
 *    accept out-of-bounds column references that it should reject.
 *
 * The fix creates a new CellReferenceMatcher (and a new CellReference token) per
 * `buildLexerConfig` call, ensuring each HyperFormula instance owns its own state.
 */
describe('CellReferenceMatcher — no shared mutable state between instances', () => {
  it('googleSheets instance with maxColumns=5 still rejects wide refs after a default instance is built', () => {
    // Build the googleSheets instance first (singleton: googleSheets, maxColumns=5)
    const hfGs5 = HyperFormula.buildFromArray(
      [[1, 2, 3, 4]],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', maxColumns: 5},
    )

    // Build default instance second (singleton now: default, maxColumns=18278)
    const hfDefault = HyperFormula.buildFromArray(
      [[1, 2, 3, 4, 5, 6]],
      {licenseKey: 'gpl-v3', compatibilityMode: 'default'},
    )

    // Now add a new formula to hfGs5. The formula is parsed AFTER the singleton has been
    // reconfigured to default/18278. If the singleton is shared, F1 would be parsed as a
    // valid cell reference. If each instance owns its own matcher, F1 is still rejected.
    hfGs5.addRows(0, [1, 1])
    hfGs5.setCellContents(adr('A2'), '=F1')

    const gsVal = hfGs5.getCellValue(adr('A2'))
    // F is column index 5, which equals maxColumns=5, so it should be rejected
    // and treated as a named expression → results in a #NAME? error (not null or a number).
    // Note: null (empty cell) also has typeof === 'object', so we check for the error type.
    expect(gsVal).not.toBeNull()
    expect(typeof gsVal).toBe('object') // DetailedCellError

    hfDefault.destroy()
    hfGs5.destroy()
  })

  it('googleSheets instance with small maxColumns rejects wide column references at build time', () => {
    // Build the googleSheets instance first.
    const hfGs = HyperFormula.buildFromArray(
      [[1, 2, 3, 4, 5], ['=F1']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', maxColumns: 5},
    )

    // Build a default instance after; this reconfigures the singleton to default mode / maxColumns=18278.
    const hfDefault = HyperFormula.buildFromArray(
      [[1, 2, 3, 4, 5, 6], ['=F1']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'default'},
    )

    // In default mode, =F1 is a valid cell reference (col F = index 5)
    expect(hfDefault.getCellValue(adr('A2'))).toBe(6)

    // In googleSheets mode with maxColumns=5, column F (index 5) meets or exceeds the
    // limit so the lexer should NOT emit it as a CellReference. It falls through to
    // NamedExpression, which resolves to a #NAME? error since no named expression "F1" exists.
    const gsVal = hfGs.getCellValue(adr('A2'))
    expect(typeof gsVal).toBe('object') // should be a DetailedCellError, not a number

    hfGs.destroy()
    hfDefault.destroy()
  })

  it('default-mode instance does not lose wide-column references after a narrow googleSheets instance is built', () => {
    // Build default instance first
    const hfDefault = HyperFormula.buildFromArray(
      [[1, 2, 3, 4, 5, 6, 7, 8, 9, 10]],
      {licenseKey: 'gpl-v3', compatibilityMode: 'default'},
    )

    // Build a googleSheets instance with maxColumns=5 second.
    // Before the fix, this reconfigured the shared singleton to googleSheets+maxColumns=5,
    // making the default instance's lexer reject column references beyond E (index 4).
    const hfGs5 = HyperFormula.buildFromArray(
      [[1, 2, 3, 4]],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', maxColumns: 5},
    )

    // The default instance should still be able to reference column J (index 9) after
    // the googleSheets instance is built.
    hfDefault.addRows(0, [1, 1])
    hfDefault.setCellContents(adr('A2'), '=J1')
    expect(hfDefault.getCellValue(adr('A2'))).toBe(10)

    hfDefault.destroy()
    hfGs5.destroy()
  })

  it('second googleSheets instance with different maxColumns is configured independently', () => {
    // Build a googleSheets instance with maxColumns=10 (F1 is within bounds)
    const hfGs10 = HyperFormula.buildFromArray(
      [[1, 2, 3, 4, 5, 6], ['=F1']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', maxColumns: 10},
    )

    // Build a googleSheets instance with maxColumns=5 (F1 exceeds bounds).
    // This instance only has 4 columns of data to stay within the limit.
    const hfGs5 = HyperFormula.buildFromArray(
      [[1, 2, 3, 4], ['=F1']],
      {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', maxColumns: 5},
    )

    // hfGs10 was built when singleton had maxColumns=10 → F1 is valid
    expect(hfGs10.getCellValue(adr('A2'))).toBe(6)

    // hfGs5 now owns maxColumns=5 → F1 (col index 5) meets limit → parse as named expression → error
    const gsVal5 = hfGs5.getCellValue(adr('A2'))
    expect(typeof gsVal5).toBe('object') // DetailedCellError

    hfGs10.destroy()
    hfGs5.destroy()
  })
})
