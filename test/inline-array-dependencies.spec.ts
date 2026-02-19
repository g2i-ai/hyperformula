import {HyperFormula} from '../src'
import {adr} from './testUtils'

describe('Inline array dependency tracking', () => {
  it('should update cell when inline array references change', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2],
      ['={A1, B1}', null],
    ], {licenseKey: 'gpl-v3'})

    // Initial value should be computed
    expect(hf.getCellValue(adr('A2'))).toBe(1)

    // Change a referenced cell
    hf.setCellContents(adr('A1'), 10)

    // The inline array cell should update
    expect(hf.getCellValue(adr('A2'))).toBe(10)

    hf.destroy()
  })

  it('should update cell when nested inline array references change', () => {
    const hf = HyperFormula.buildFromArray([
      [1, 2, 3],
      ['={A1, B1; C1, A1}', null],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A2'))).toBe(1)

    hf.setCellContents(adr('C1'), 99)
    expect(hf.getCellValue(adr('A2'))).toBe(1)

    hf.setCellContents(adr('A1'), 50)
    expect(hf.getCellValue(adr('A2'))).toBe(50)

    hf.destroy()
  })

  it('should track cell reference dependencies in inline arrays', () => {
    const hf = HyperFormula.buildFromArray([
      ['={A2, A3}', null],
      [10, null],
      [20, null],
    ], {licenseKey: 'gpl-v3'})

    expect(hf.getCellValue(adr('A1'))).toBe(10)

    hf.setCellContents(adr('A2'), 42)
    expect(hf.getCellValue(adr('A1'))).toBe(42)

    hf.setCellContents(adr('A3'), 99)
    expect(hf.getCellValue(adr('B1'))).toBe(99)

    hf.destroy()
  })
})
