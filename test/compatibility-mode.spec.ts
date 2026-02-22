import {HyperFormula} from '../src'
import {Config} from '../src/Config'
import {adr} from './testUtils'

describe('compatibilityMode config option', () => {
  it('should default to "default"', () => {
    expect(Config.defaultConfig.compatibilityMode).toBe('default')
  })

  it('should accept "googleSheets" value', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.compatibilityMode).toBe('googleSheets')
  })

  it('should accept "default" value', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'default'})
    expect(config.compatibilityMode).toBe('default')
  })

  it('should be accessible on the engine config', () => {
    const hf = HyperFormula.buildEmpty({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(hf.getConfig().compatibilityMode).toBe('googleSheets')
    hf.destroy()
  })
})

describe('Google Sheets config preset', () => {
  it('should use GSheets date formats when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.dateFormats).toEqual(['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'])
  })

  it('should use GSheets locale when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.localeLang).toBe('en-US')
  })

  it('should use GSheets currency symbols when not explicitly provided', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})
    expect(config.currencySymbol).toEqual(['$', 'USD'])
  })

  it('should allow user to override GSheets preset values', () => {
    const config = new Config({
      licenseKey: 'gpl-v3',
      compatibilityMode: 'googleSheets',
      dateFormats: ['YYYY-MM-DD'],
      localeLang: 'de-DE',
    })
    expect(config.dateFormats).toEqual(['YYYY-MM-DD'])
    expect(config.localeLang).toBe('de-DE')
    expect(config.currencySymbol).toEqual(['$', 'USD'])
  })

  it('should use standard defaults when compatibilityMode is "default"', () => {
    const config = new Config({licenseKey: 'gpl-v3', compatibilityMode: 'default'})
    expect(config.dateFormats).toEqual(['DD/MM/YYYY', 'DD/MM/YY'])
    expect(config.localeLang).toBe('en')
    expect(config.currencySymbol).toEqual(['$'])
  })
})

describe('Google Sheets named expression auto-registration', () => {
  it('should auto-register TRUE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('should auto-register FALSE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=FALSE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'})

    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })

  it('should NOT auto-register TRUE/FALSE in default mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'default'})

    // In default mode, TRUE without () is a named expression reference,
    // which doesn't exist — should be a NAME error
    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(Object) // DetailedCellError
    hf.destroy()
  })

  it('should not conflict with user-provided named expressions', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MyExpr'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    [{name: 'MyExpr', expression: '=42'}])

    expect(hf.getCellValue(adr('A1'))).toBe(42)
    hf.destroy()
  })

  it('should not overwrite user-defined TRUE named expression', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    [{name: 'TRUE', expression: '=42'}])

    expect(hf.getCellValue(adr('A1'))).toBe(42)
    hf.destroy()
  })

  it('should still register global TRUE when only a sheet-scoped TRUE exists', () => {
    const hf = HyperFormula.buildFromSheets({
      Sheet1: [['=TRUE']],
      Sheet2: [['=TRUE']],
    }, {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    [{name: 'TRUE', expression: '=42', scope: 0}])

    expect(hf.getCellValue(adr('A1', 0))).toBe(42)
    expect(hf.getCellValue(adr('A1', 1))).toBe(true)
    hf.destroy()
  })

  it('should not overwrite user-defined FALSE named expression', () => {
    const hf = HyperFormula.buildFromArray([
      ['=FALSE'],
    ], {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'},
    [{name: 'FALSE', expression: '=99'}])

    expect(hf.getCellValue(adr('A1'))).toBe(99)
    hf.destroy()
  })
})
