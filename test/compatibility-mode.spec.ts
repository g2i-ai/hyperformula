import { HyperFormula } from '../src'
import { Config } from '../src/Config'
import { deDE } from '../src/i18n/languages'
import { adr } from './testUtils'

describe('compatibilityMode config option', () => {
  it('should default to "default"', () => {
    expect(Config.defaultConfig.compatibilityMode).toBe('default')
  })

  it('should accept "googleSheets" value', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    expect(config.compatibilityMode).toBe('googleSheets')
  })

  it('should accept "default" value', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'default' })
    expect(config.compatibilityMode).toBe('default')
  })

  it('should be accessible on the engine config', () => {
    const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    expect(hf.getConfig().compatibilityMode).toBe('googleSheets')
    hf.destroy()
  })
})

describe('Google Sheets config preset', () => {
  it('should use GSheets date formats when not explicitly provided', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    expect(config.dateFormats).toEqual(['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'])
  })

  it('should use GSheets locale when not explicitly provided', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    expect(config.localeLang).toBe('en-US')
  })

  it('should use GSheets currency symbols when not explicitly provided', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
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
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'default' })
    expect(config.dateFormats).toEqual(['DD/MM/YYYY', 'DD/MM/YY'])
    expect(config.localeLang).toBe('en')
    expect(config.currencySymbol).toEqual(['$'])
  })

  it('should apply GSheets defaults when switching mode via mergeConfig', () => {
    const config = new Config({ licenseKey: 'gpl-v3' })
    const mergedConfig = config.mergeConfig({ compatibilityMode: 'googleSheets' })

    expect(mergedConfig.dateFormats).toEqual(['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'])
    expect(mergedConfig.localeLang).toBe('en-US')
    expect(mergedConfig.currencySymbol).toEqual(['$', 'USD'])
  })

  it('should restore default preset when switching back via mergeConfig', () => {
    const config = new Config({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })
    const mergedConfig = config.mergeConfig({ compatibilityMode: 'default' })

    expect(mergedConfig.dateFormats).toEqual(['DD/MM/YYYY', 'DD/MM/YY'])
    expect(mergedConfig.localeLang).toBe('en')
    expect(mergedConfig.currencySymbol).toEqual(['$'])
  })

  it('should apply GSheets defaults after switching mode with updateConfig', () => {
    const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' })

    hf.updateConfig({ compatibilityMode: 'googleSheets' })

    const updatedConfig = hf.getConfig()
    expect(updatedConfig.dateFormats).toEqual(['MM/DD/YYYY', 'MM/DD/YY', 'YYYY/MM/DD'])
    expect(updatedConfig.localeLang).toBe('en-US')
    expect(updatedConfig.currencySymbol).toEqual(['$', 'USD'])
    hf.destroy()
  })

  it('should restore standard defaults after switching from googleSheets to default with updateConfig', () => {
    const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

    hf.updateConfig({ compatibilityMode: 'default' })

    const updatedConfig = hf.getConfig()
    expect(updatedConfig.dateFormats).toEqual(['DD/MM/YYYY', 'DD/MM/YY'])
    expect(updatedConfig.localeLang).toBe('en')
    expect(updatedConfig.currencySymbol).toEqual(['$'])
    hf.destroy()
  })
})

describe('Google Sheets named expression auto-registration', () => {
  it('should auto-register TRUE/FALSE in configured language for non-English mode', () => {
    HyperFormula.registerLanguage('deDE', deDE)
    const hf = HyperFormula.buildFromArray([
      ['=TRUE', '=FALSE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets', language: 'deDE' })

    expect(hf.getCellValue(adr('A1'))).toBe(true)
    expect(hf.getCellValue(adr('B1'))).toBe(false)
    hf.destroy()
  })

  it('should auto-register TRUE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('should auto-register FALSE named expression in googleSheets mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=FALSE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })

  it('should NOT auto-register TRUE/FALSE in default mode', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'default' })

    // In default mode, TRUE without () is a named expression reference,
    // which doesn't exist — should be a NAME error
    const val = hf.getCellValue(adr('A1'))
    expect(val).toBeInstanceOf(Object) // DetailedCellError
    hf.destroy()
  })

  it('should not conflict with user-provided named expressions', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MyExpr'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' },
      [{ name: 'MyExpr', expression: '=42' }])

    expect(hf.getCellValue(adr('A1'))).toBe(42)
    hf.destroy()
  })

  it('should treat TRUE/FALSE as parser keywords that cannot be shadowed by named expressions', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' },
      [{ name: 'TRUE', expression: '=42' }])

    // In GSheets mode, TRUE is a parser keyword — named expressions cannot shadow it
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    hf.destroy()
  })

  it('should treat TRUE as parser keyword regardless of sheet-scoped named expressions', () => {
    const hf = HyperFormula.buildFromSheets({
      Sheet1: [['=TRUE']],
      Sheet2: [['=TRUE']],
    }, { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' },
      [{ name: 'TRUE', expression: '=42', scope: 0 }])

    // TRUE is a keyword in both sheets
    expect(hf.getCellValue(adr('A1', 0))).toBe(true)
    expect(hf.getCellValue(adr('A1', 1))).toBe(true)
    hf.destroy()
  })

  it('should treat FALSE as parser keyword that cannot be shadowed', () => {
    const hf = HyperFormula.buildFromArray([
      ['=FALSE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' },
      [{ name: 'FALSE', expression: '=99' }])

    expect(hf.getCellValue(adr('A1'))).toBe(false)
    hf.destroy()
  })

  it('should normalize TRUE/FALSE to function calls after switching from googleSheets to default', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE', '=FALSE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

    expect(hf.getCellValue(adr('A1'))).toBe(true)
    expect(hf.getCellValue(adr('B1'))).toBe(false)

    hf.updateConfig({ compatibilityMode: 'default' })

    // After switching modes, formulas get serialized and re-parsed.
    // GSheets mode parsed =TRUE as ProcedureAst('TRUE', []), which serializes as =TRUE().
    // =TRUE() is valid in default mode too, so the value is still correct.
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    expect(hf.getCellValue(adr('B1'))).toBe(false)
    hf.destroy()
  })

  it('should normalize TRUE/FALSE to function calls even when user named expressions exist', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TRUE', '=FALSE'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' }, [
      { name: 'TRUE', expression: '=42' },
      { name: 'FALSE', expression: '=99' },
    ])

    // In GSheets mode, parser keywords take precedence over named expressions
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    expect(hf.getCellValue(adr('B1'))).toBe(false)

    hf.updateConfig({ compatibilityMode: 'default' })

    // After switching, formulas were normalized to =TRUE()/=FALSE() (function form),
    // so they still evaluate as boolean functions, not named expressions
    expect(hf.getCellValue(adr('A1'))).toBe(true)
    expect(hf.getCellValue(adr('B1'))).toBe(false)
    hf.destroy()
  })
})

describe('Google Sheets plugin registration', () => {
  it('should load GSheets plugins when compatibilityMode is googleSheets', () => {
    // Validates the wiring — specific function overrides tested in plugin test files
    const hf = HyperFormula.buildFromArray([
      [1, 2, '=SUM(A1,B1)'],
    ], { licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets' })

    expect(hf.getCellValue(adr('C1'))).toBe(3)
    hf.destroy()
  })
})
