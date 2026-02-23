import {DetailedCellError, HyperFormula} from '../src'
import {ErrorType} from '../src/Cell'
import {adr} from './testUtils'

describe('GoogleSheetsDatabasePlugin', () => {
  const buildWithPlugin = (data: (string | number | null | undefined)[][]) =>
    HyperFormula.buildFromArray(data, {
      licenseKey: 'gpl-v3',
      compatibilityMode: 'googleSheets',
    })

  /**
   * Database layout:
   *   A1:D1  — headers: Name, Department, Salary, Age
   *   A2:D6  — data rows (5 employees)
   *   A8:A9  — criteria block (header + value) — overwritten per test via spread
   */
  const database: (string | number)[][] = [
    ['Name', 'Department', 'Salary', 'Age'],  // row 1
    ['Alice', 'Engineering', 90000, 30],       // row 2
    ['Bob', 'Engineering', 80000, 25],         // row 3
    ['Carol', 'Sales', 85000, 35],             // row 4
    ['Dave', 'Engineering', 95000, 28],        // row 5
    ['Eve', 'Sales', 70000, 32],               // row 6
  ]

  const spacer: (string | number | null)[][] = [[]]  // row 7

  // ---------------------------------------------------------------------------
  // DSUM
  // ---------------------------------------------------------------------------

  it('DSUM sums matching values by header name', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],                        // row 8
      ['Engineering'],                       // row 9
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:A9)'],     // row 11
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(265000)
    hf.destroy()
  })

  it('DSUM with numeric field index (1-based)', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DSUM(A1:D6, 3, A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(265000)
    hf.destroy()
  })

  it('DSUM returns 0 when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Marketing'],
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(0)
    hf.destroy()
  })

  it('DSUM with comparison operator criterion (>80000)', () => {
    const data = [
      ...database,
      ...spacer,
      ['Salary'],
      ['>80000'],
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Alice=90000, Carol=85000, Dave=95000 match >80000
    expect(hf.getCellValue(adr('A11'))).toBe(270000)
    hf.destroy()
  })

  it('DSUM with >= operator', () => {
    const data = [
      ...database,
      ...spacer,
      ['Salary'],
      ['>=85000'],
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Alice=90000, Carol=85000, Dave=95000
    expect(hf.getCellValue(adr('A11'))).toBe(270000)
    hf.destroy()
  })

  it('DSUM with <> operator', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['<>Engineering'],
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Carol=85000, Eve=70000
    expect(hf.getCellValue(adr('A11'))).toBe(155000)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DAVERAGE
  // ---------------------------------------------------------------------------

  it('DAVERAGE computes average of matching values', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DAVERAGE(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // (90000 + 80000 + 95000) / 3 = 88333.333...
    expect(hf.getCellValue(adr('A11'))).toBeCloseTo(88333.333, 2)
    hf.destroy()
  })

  it('DAVERAGE returns DIV_BY_ZERO error when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Marketing'],
      spacer[0],
      ['=DAVERAGE(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    expect((val as DetailedCellError).type).toBe(ErrorType.DIV_BY_ZERO)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DCOUNT
  // ---------------------------------------------------------------------------

  it('DCOUNT counts rows with numeric field values matching criteria', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DCOUNT(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(3)
    hf.destroy()
  })

  it('DCOUNT counts 0 when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Marketing'],
      spacer[0],
      ['=DCOUNT(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(0)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DCOUNTA
  // ---------------------------------------------------------------------------

  it('DCOUNTA counts non-empty values in field column for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Sales'],
      spacer[0],
      ['=DCOUNTA(A1:D6, "Name", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Carol and Eve are in Sales
    expect(hf.getCellValue(adr('A11'))).toBe(2)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DGET
  // ---------------------------------------------------------------------------

  it('DGET returns single matching value', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DGET(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(90000)
    hf.destroy()
  })

  it('DGET returns VALUE error when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Nobody'],
      spacer[0],
      ['=DGET(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    expect((val as DetailedCellError).type).toBe(ErrorType.VALUE)
    hf.destroy()
  })

  it('DGET returns NUM error when multiple rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DGET(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    expect((val as DetailedCellError).type).toBe(ErrorType.NUM)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DMAX / DMIN
  // ---------------------------------------------------------------------------

  it('DMAX returns maximum numeric value for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DMAX(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(95000)
    hf.destroy()
  })

  it('DMIN returns minimum numeric value for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DMIN(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(80000)
    hf.destroy()
  })

  it('DMAX returns 0 when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Marketing'],
      spacer[0],
      ['=DMAX(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(0)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DPRODUCT
  // ---------------------------------------------------------------------------

  it('DPRODUCT returns product of matching values', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Sales'],
      spacer[0],
      ['=DPRODUCT(A1:D6, "Age", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Carol age=35, Eve age=32 → 35 * 32 = 1120
    expect(hf.getCellValue(adr('A11'))).toBe(1120)
    hf.destroy()
  })

  it('DPRODUCT returns 0 when no rows match', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Marketing'],
      spacer[0],
      ['=DPRODUCT(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(0)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DSTDEV / DSTDEVP
  // ---------------------------------------------------------------------------

  it('DSTDEV returns sample standard deviation for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DSTDEV(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Engineering salaries: 90000, 80000, 95000
    // mean = 88333.33, sample stdev ≈ 7637.63
    const val = hf.getCellValue(adr('A11')) as number
    expect(val).toBeCloseTo(7637.626, 2)
    hf.destroy()
  })

  it('DSTDEV returns DIV_BY_ZERO error with fewer than 2 matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DSTDEV(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    expect((val as DetailedCellError).type).toBe(ErrorType.DIV_BY_ZERO)
    hf.destroy()
  })

  it('DSTDEVP returns population standard deviation for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DSTDEVP(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Engineering salaries: 90000, 80000, 95000
    // mean = 88333.33, population stdev ≈ 6236.09
    const val = hf.getCellValue(adr('A11')) as number
    expect(val).toBeCloseTo(6236.09, 1)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // DVAR / DVARP
  // ---------------------------------------------------------------------------

  it('DVAR returns sample variance for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DVAR(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Engineering salaries: 90000, 80000, 95000 — sample variance
    const val = hf.getCellValue(adr('A11')) as number
    expect(val).toBeCloseTo(58333333.333, 0)
    hf.destroy()
  })

  it('DVAR returns DIV_BY_ZERO error with fewer than 2 matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DVAR(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    expect((val as DetailedCellError).type).toBe(ErrorType.DIV_BY_ZERO)
    hf.destroy()
  })

  it('DVARP returns population variance for matching rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      ['Engineering'],
      spacer[0],
      ['=DVARP(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11')) as number
    expect(val).toBeCloseTo(38888888.888, 0)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // AND logic (multiple criteria columns in same criteria row)
  // ---------------------------------------------------------------------------

  it('multiple criteria columns in same row apply AND logic', () => {
    // Criteria: Department=Engineering AND Age>28
    const data = [
      ...database,
      ...spacer,
      ['Department', 'Age'],   // row 8 — two criteria headers
      ['Engineering', '>28'],  // row 9 — AND: both must match
      spacer[0],
      ['=DSUM(A1:D6, "Salary", A8:B9)'],
    ]
    const hf = buildWithPlugin(data)
    // Alice (Eng, 30) and Dave would not match because Dave is 28, not >28
    // Alice=90000 only
    expect(hf.getCellValue(adr('A11'))).toBe(90000)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // OR logic (multiple criteria rows)
  // ---------------------------------------------------------------------------

  it('multiple criteria rows apply OR logic', () => {
    // Criteria: Department=Engineering OR Department=Sales
    const data = [
      ...database,
      ...spacer,
      ['Department'],     // row 8 — one criteria header
      ['Engineering'],    // row 9 — first criteria row
      ['Sales'],          // row 10 — second criteria row (OR)
      ['=DCOUNT(A1:D6, "Salary", A8:A10)'],
    ]
    const hf = buildWithPlugin(data)
    // All 5 employees are in Engineering or Sales
    expect(hf.getCellValue(adr('A11'))).toBe(5)
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // Wildcard matching
  // ---------------------------------------------------------------------------

  it('wildcard * matches any sequence of characters in criteria', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['A*'],            // matches names starting with A
      spacer[0],
      ['=DCOUNT(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    // Only Alice starts with A
    expect(hf.getCellValue(adr('A11'))).toBe(1)
    hf.destroy()
  })

  it('wildcard ? matches any single character in criteria', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['?ob'],           // matches "Bob"
      spacer[0],
      ['=DGET(A1:D6, "Name", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe('Bob')
    hf.destroy()
  })

  // ---------------------------------------------------------------------------
  // Edge cases
  // ---------------------------------------------------------------------------

  it('field referenced by column number 1 works', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DGET(A1:D6, 1, A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe('Alice')
    hf.destroy()
  })

  it('field index out of bounds returns VALUE error', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DSUM(A1:D6, 99, A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('unknown field name returns VALUE error', () => {
    const data = [
      ...database,
      ...spacer,
      ['Name'],
      ['Alice'],
      spacer[0],
      ['=DSUM(A1:D6, "NoSuchColumn", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    const val = hf.getCellValue(adr('A11'))
    expect(val).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('empty criteria value matches all rows', () => {
    const data = [
      ...database,
      ...spacer,
      ['Department'],
      [''],              // empty criterion = match all
      spacer[0],
      ['=DCOUNT(A1:D6, "Salary", A8:A9)'],
    ]
    const hf = buildWithPlugin(data)
    expect(hf.getCellValue(adr('A11'))).toBe(5)
    hf.destroy()
  })
})
