/**
 * Tests for GoogleSheetsArrayPlugin — SORT, UNIQUE, FLATTEN, CHOOSECOLS, CHOOSEROWS,
 * HSTACK, VSTACK, WRAPCOLS, WRAPROWS, TOCOL, TOROW, SEQUENCE, FREQUENCY,
 * MDETERM, MINVERSE, MUNIT, GROWTH, TREND, LINEST, LOGEST.
 *
 * Layout convention: input data lives in columns D+ so that the formula result
 * in column A has empty cells to spill into.
 */

import {DetailedCellError, HyperFormula} from '../src'
import {adr} from './testUtils'

const gsOptions = {licenseKey: 'gpl-v3', compatibilityMode: 'googleSheets'} as const

// ─── SORT ────────────────────────────────────────────────────────────────────

describe('SORT', () => {
  it('sorts a column range ascending by default', () => {
    // D1:D3 has data; formula spills into A1:A3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3)', null, null, 3],
      [null, null, null, 1],
      [null, null, null, 2],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    hf.destroy()
  })

  it('sorts 2D range by first column ascending', () => {
    // D1:E3 has data; formula spills into A1:B3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:E3)', null, null, 3, 'c'],
      [null, null, null, 1, 'a'],
      [null, null, null, 2, 'b'],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe('a')
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    hf.destroy()
  })

  it('sorts by second column descending', () => {
    // D1:E3 has data; formula spills into A1:B3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:E3,2,FALSE())', null, null, 1, 'c'],
      [null, null, null, 2, 'a'],
      [null, null, null, 3, 'b'],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe('c')
    expect(hf.getCellValue(adr('A2'))).toBe(3)
    expect(hf.getCellValue(adr('B2'))).toBe('b')
    hf.destroy()
  })
})

// ─── UNIQUE ───────────────────────────────────────────────────────────────────

describe('UNIQUE', () => {
  it('removes duplicate rows', () => {
    // D1:E4 has data; formula spills into A1:B4 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=UNIQUE(D1:E4)', null, null, 1, 'a'],
      [null, null, null, 2, 'b'],
      [null, null, null, 1, 'a'],
      [null, null, null, 3, 'c'],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe('a')
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('B2'))).toBe('b')
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    hf.destroy()
  })

  it('preserves order of first occurrence', () => {
    // D1:D4 has data; formula spills into A1:A3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=UNIQUE(D1:D4)', null, null, 'b'],
      [null, null, null, 'a'],
      [null, null, null, 'b'],
      [null, null, null, 'c'],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe('b')
    expect(hf.getCellValue(adr('A2'))).toBe('a')
    expect(hf.getCellValue(adr('A3'))).toBe('c')
    hf.destroy()
  })
})

// ─── FLATTEN ─────────────────────────────────────────────────────────────────

describe('FLATTEN', () => {
  it('flattens a 2D range into a single column', () => {
    // D1:E2 has data; formula spills into A1:A4 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=FLATTEN(D1:E2)', null, null, 1, 2],
      [null, null, null, 3, 4],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    expect(hf.getCellValue(adr('A4'))).toBe(4)
    hf.destroy()
  })

  it('flattens multiple ranges into one column', () => {
    // D1:D2 and E1:E2 have data; formula spills into A1:A4 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=FLATTEN(D1:D2,E1:E2)', null, null, 1, 3],
      [null, null, null, 2, 4],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    expect(hf.getCellValue(adr('A4'))).toBe(4)
    hf.destroy()
  })
})

// ─── CHOOSECOLS ───────────────────────────────────────────────────────────────

describe('CHOOSECOLS', () => {
  it('selects specific columns by positive index', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSECOLS(A3:C4,1,3)', null, null, null, null],
      [null, null, null, null, null],
      [1, 2, 3, null, null],
      [4, 5, 6, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(3)
    expect(hf.getCellValue(adr('A2'))).toBe(4)
    expect(hf.getCellValue(adr('B2'))).toBe(6)
    hf.destroy()
  })

  it('selects columns using negative indices from the end', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSECOLS(A3:C3,-1)', null, null, null],
      [null, null, null, null],
      [10, 20, 30, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(30)
    hf.destroy()
  })

  it('returns VALUE error for out-of-bounds column index', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSECOLS(A3:B3,5)', null, null],
      [null, null, null],
      [1, 2, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

// ─── CHOOSEROWS ───────────────────────────────────────────────────────────────

describe('CHOOSEROWS', () => {
  it('selects specific rows by positive index', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSEROWS(A3:B5,1,3)', null, null, null],
      [null, null, null, null],
      [1, 2, null, null],
      [3, 4, null, null],
      [5, 6, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('A2'))).toBe(5)
    expect(hf.getCellValue(adr('B2'))).toBe(6)
    hf.destroy()
  })

  it('selects rows using negative indices from the end', () => {
    const hf = HyperFormula.buildFromArray([
      ['=CHOOSEROWS(A3:B5,-1)', null, null, null],
      [null, null, null, null],
      [1, 2, null, null],
      [3, 4, null, null],
      [5, 6, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(5)
    expect(hf.getCellValue(adr('B1'))).toBe(6)
    hf.destroy()
  })
})

// ─── HSTACK ───────────────────────────────────────────────────────────────────

describe('HSTACK', () => {
  it('horizontally concatenates two ranges of equal height', () => {
    const hf = HyperFormula.buildFromArray([
      ['=HSTACK(A3:A4,B3:B4)', null, null, null],
      [null, null, null, null],
      [1, 3, null, null],
      [2, 4, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(3)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('B2'))).toBe(4)
    hf.destroy()
  })

  it('pads shorter ranges with empty values when heights differ', () => {
    // Data in D/E cols; result spills into A1:B3
    const hf = HyperFormula.buildFromArray([
      ['=HSTACK(D1:D3,E1:E2)', null, null, 1, 10],
      [null, null, null, 2, 20],
      [null, null, null, 3, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(10)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    // EmptyValue (padding) renders as null
    expect(hf.getCellValue(adr('B3'))).toBeNull()
    hf.destroy()
  })
})

// ─── VSTACK ───────────────────────────────────────────────────────────────────

describe('VSTACK', () => {
  it('vertically stacks two ranges', () => {
    const hf = HyperFormula.buildFromArray([
      ['=VSTACK(A3:B3,A4:B4)', null, null, null],
      [null, null, null, null],
      [1, 2, null, null],
      [3, 4, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('A2'))).toBe(3)
    expect(hf.getCellValue(adr('B2'))).toBe(4)
    hf.destroy()
  })

  it('pads narrower ranges with empty values', () => {
    // Data in D/E cols; result spills into A1:B2
    const hf = HyperFormula.buildFromArray([
      ['=VSTACK(D1:E1,D2:D2)', null, null, 1, 2],
      [null, null, null, 3, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('A2'))).toBe(3)
    // EmptyValue (padding) renders as null
    expect(hf.getCellValue(adr('B2'))).toBeNull()
    hf.destroy()
  })
})

// ─── WRAPCOLS ─────────────────────────────────────────────────────────────────

describe('WRAPCOLS', () => {
  it('wraps a 6-element range into 2-row columns', () => {
    const hf = HyperFormula.buildFromArray([
      ['=WRAPCOLS(A3:F3,2)', null, null, null, null, null, null],
      [null, null, null, null, null, null, null],
      [1, 2, 3, 4, 5, 6, null],
    ], gsOptions)

    // Columns: [1,2], [3,4], [5,6]
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('B1'))).toBe(3)
    expect(hf.getCellValue(adr('C1'))).toBe(5)
    hf.destroy()
  })

  it('pads last column with pad_with value when not evenly divisible', () => {
    // Data in row 1 of D:H; result spills into A1:B3 (3 rows x 2 cols)
    const hf = HyperFormula.buildFromArray([
      ['=WRAPCOLS(D1:H1,3,0)', null, null, 1, 2, 3, 4, 5],
      [null, null, null, null, null, null, null, null],
      [null, null, null, null, null, null, null, null],
    ], gsOptions)

    // Columns: [1,2,3], [4,5,0]
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(4)
    expect(hf.getCellValue(adr('B3'))).toBe(0)
    hf.destroy()
  })
})

// ─── WRAPROWS ─────────────────────────────────────────────────────────────────

describe('WRAPROWS', () => {
  it('wraps a 6-element range into 3-column rows', () => {
    const hf = HyperFormula.buildFromArray([
      ['=WRAPROWS(A3:F3,3)', null, null, null, null, null, null],
      [null, null, null, null, null, null, null],
      [1, 2, 3, 4, 5, 6, null],
    ], gsOptions)

    // Rows: [1,2,3], [4,5,6]
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('C1'))).toBe(3)
    expect(hf.getCellValue(adr('A2'))).toBe(4)
    expect(hf.getCellValue(adr('C2'))).toBe(6)
    hf.destroy()
  })

  it('pads last row with pad_with value', () => {
    const hf = HyperFormula.buildFromArray([
      ['=WRAPROWS(A3:E3,3,-1)', null, null, null, null, null],
      [null, null, null, null, null, null],
      [1, 2, 3, 4, 5, null],
    ], gsOptions)

    // Rows: [1,2,3], [4,5,-1]
    expect(hf.getCellValue(adr('C2'))).toBe(-1)
    hf.destroy()
  })
})

// ─── TOCOL ────────────────────────────────────────────────────────────────────

describe('TOCOL', () => {
  it('reshapes a 2D range into a single column (row-major)', () => {
    // D1:E2 has data; formula spills into A1:A4 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=TOCOL(D1:E2)', null, null, 1, 2],
      [null, null, null, 3, 4],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    expect(hf.getCellValue(adr('A4'))).toBe(4)
    hf.destroy()
  })

  it('ignores blanks when ignore=1', () => {
    // D1:E2 has data with blank; formula spills into A1:A3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=TOCOL(D1:E2,1)', null, null, 1, null],
      [null, null, null, 2, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    hf.destroy()
  })
})

// ─── TOROW ────────────────────────────────────────────────────────────────────

describe('TOROW', () => {
  it('reshapes a 2D range into a single row (row-major)', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TOROW(A3:B4)', null, null, null, null],
      [null, null, null, null, null],
      [1, 2, null, null, null],
      [3, 4, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('C1'))).toBe(3)
    expect(hf.getCellValue(adr('D1'))).toBe(4)
    hf.destroy()
  })
})

// ─── SEQUENCE ─────────────────────────────────────────────────────────────────

describe('SEQUENCE', () => {
  it('generates a 2x2 array starting at 1 with step 1', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SEQUENCE(2,2,1,1)', null, null],
      [null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('A2'))).toBe(3)
    expect(hf.getCellValue(adr('B2'))).toBe(4)
    hf.destroy()
  })

  it('generates a 1x5 row sequence with step 2', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SEQUENCE(1,5,0,2)', null, null, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(0)
    expect(hf.getCellValue(adr('B1'))).toBe(2)
    expect(hf.getCellValue(adr('E1'))).toBe(8)
    hf.destroy()
  })

  it('generates a 3x1 column with negative step', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SEQUENCE(3,1,10,-5)', null],
      [null, null],
      [null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(10)
    expect(hf.getCellValue(adr('A2'))).toBe(5)
    expect(hf.getCellValue(adr('A3'))).toBe(0)
    hf.destroy()
  })
})

// ─── FREQUENCY ────────────────────────────────────────────────────────────────

describe('FREQUENCY', () => {
  it('returns frequency distribution with one extra bin for overflow', () => {
    // Data in D/E; result spills into A1:A4 (4 rows = 3 bins + overflow)
    const hf = HyperFormula.buildFromArray([
      ['=FREQUENCY(D1:D5,E1:E3)', null, null, 1, 2],
      [null, null, null, 3, 5],
      [null, null, null, 4, 8],
      [null, null, null, 7, null],
      [null, null, null, 9, null],
    ], gsOptions)

    // Bins: [2, 5, 8] (sorted)
    // <=2: [1] → count=1
    // <=5: [3,4] → count=2
    // <=8: [7] → count=1
    // >8: [9] → count=1
    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(1)
    expect(hf.getCellValue(adr('A4'))).toBe(1)
    hf.destroy()
  })
})

// ─── MDETERM ─────────────────────────────────────────────────────────────────

describe('MDETERM', () => {
  it('computes determinant of 2x2 matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MDETERM(A3:B4)', null, null],
      [null, null, null],
      [1, 2, null],
      [3, 4, null],
    ], gsOptions)

    // det([[1,2],[3,4]]) = 1*4 - 2*3 = -2
    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(-2)
    hf.destroy()
  })

  it('computes determinant of identity matrix as 1', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MDETERM(A3:B4)', null, null],
      [null, null, null],
      [1, 0, null],
      [0, 1, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1)
    hf.destroy()
  })

  it('returns 0 for singular matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MDETERM(A3:B4)', null, null],
      [null, null, null],
      [1, 2, null],
      [2, 4, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0)
    hf.destroy()
  })

  it('returns VALUE error for non-square matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MDETERM(A3:C4)', null, null, null],
      [null, null, null, null],
      [1, 2, 3, null],
      [4, 5, 6, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

// ─── MINVERSE ────────────────────────────────────────────────────────────────

describe('MINVERSE', () => {
  it('computes inverse of 2x2 matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MINVERSE(A3:B4)', null, null, null],
      [null, null, null, null],
      [1, 2, null, null],
      [3, 4, null, null],
    ], gsOptions)

    // inv([[1,2],[3,4]]) = 1/(-2) * [[4,-2],[-3,1]] = [[-2,1],[1.5,-0.5]]
    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(-2)
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(1.5)
    expect(hf.getCellValue(adr('B2'))).toBeCloseTo(-0.5)
    hf.destroy()
  })

  it('returns NUM error for singular matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MINVERSE(A3:B4)', null, null],
      [null, null, null],
      [1, 2, null],
      [2, 4, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

// ─── MUNIT ───────────────────────────────────────────────────────────────────

describe('MUNIT', () => {
  it('generates a 2x2 identity matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MUNIT(2)', null, null],
      [null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(0)
    expect(hf.getCellValue(adr('A2'))).toBe(0)
    expect(hf.getCellValue(adr('B2'))).toBe(1)
    hf.destroy()
  })

  it('generates a 3x3 identity matrix', () => {
    const hf = HyperFormula.buildFromArray([
      ['=MUNIT(3)', null, null, null],
      [null, null, null, null],
      [null, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('B1'))).toBe(0)
    expect(hf.getCellValue(adr('A2'))).toBe(0)
    expect(hf.getCellValue(adr('B2'))).toBe(1)
    expect(hf.getCellValue(adr('C3'))).toBe(1)
    hf.destroy()
  })
})

// ─── TREND ────────────────────────────────────────────────────────────────────

describe('TREND', () => {
  it('predicts y values using linear regression', () => {
    // y = 2x: slope=2, intercept=0
    const hf = HyperFormula.buildFromArray([
      ['=TREND(A3:A5,B3:B5,C3:C4)', null, null, null, null],
      [null, null, null, null, null],
      [2, 1, 4, null, null],
      [4, 2, 5, null, null],
      [6, 3, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(8)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(10)
    hf.destroy()
  })

  it('uses default x values (1,2,...,n) when known_x omitted', () => {
    // D1:D3 has data; formula spills into A1:A3 (no overlap)
    const hf = HyperFormula.buildFromArray([
      ['=TREND(D1:D3)', null, null, 1],
      [null, null, null, 3],
      [null, null, null, 5],
    ], gsOptions)

    // y = 2x - 1: at x=1→1, x=2→3, x=3→5
    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(1)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(3)
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(5)
    hf.destroy()
  })
})

// ─── GROWTH ───────────────────────────────────────────────────────────────────

describe('GROWTH', () => {
  it('predicts exponential growth values', () => {
    // y = 1 * 2^x: at x=1→2, x=2→4, x=3→8
    const hf = HyperFormula.buildFromArray([
      ['=GROWTH(A3:A5,B3:B5,C3:C4)', null, null, null, null],
      [null, null, null, null, null],
      [2, 1, 4, null, null],
      [4, 2, 5, null, null],
      [8, 3, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(16)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(32)
    hf.destroy()
  })

  it('returns NUM error when y values contain non-positive numbers', () => {
    const hf = HyperFormula.buildFromArray([
      ['=GROWTH(A3:A4)', null, null],
      [null, null, null],
      [2, null, null],
      [-4, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

// ─── LINEST ───────────────────────────────────────────────────────────────────

describe('LINEST', () => {
  it('returns [slope, intercept] for basic linear regression', () => {
    // y = 2x + 1
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A3:A5,B3:B5)', null, null, null],
      [null, null, null, null],
      [3, 1, null, null],
      [5, 2, null, null],
      [7, 3, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // slope
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)  // intercept
    hf.destroy()
  })

  it('returns 5-row statistics array when stats=TRUE', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A8:A10,B8:B10,TRUE(),TRUE())', null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [null, null, null, null, null, null],
      [3, 1, null, null, null, null],
      [5, 2, null, null, null, null],
      [7, 3, null, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)   // slope
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)   // intercept
    // R² should be 1 for perfect fit
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(1)   // r²
    hf.destroy()
  })
})

// ─── LOGEST ───────────────────────────────────────────────────────────────────

describe('LOGEST', () => {
  it('returns [m, b] for basic exponential regression', () => {
    // y = 1 * 2^x: m=2, b=1
    const hf = HyperFormula.buildFromArray([
      ['=LOGEST(A3:A5,B3:B5)', null, null, null],
      [null, null, null, null],
      [2, 1, null, null],
      [4, 2, null, null],
      [8, 3, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // m
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)  // b
    hf.destroy()
  })

  it('returns NUM error when y values include non-positive numbers', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LOGEST(A3:A4)', null, null],
      [null, null, null],
      [0, null, null],
      [4, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

// ─── SORT error propagation and numeric ascending ─────────────────────────────

describe('SORT (error handling and numeric ascending)', () => {
  it('propagates error from sort column argument', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3,NA())', null, null, 3],
      [null, null, null, 1],
      [null, null, null, 2],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('propagates error from ascending argument', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3,1,NA())', null, null, 3],
      [null, null, null, 1],
      [null, null, null, 2],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })

  it('treats numeric 0 as descending (ascending=FALSE)', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3,1,0)', null, null, 1],
      [null, null, null, 2],
      [null, null, null, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(3)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(1)
    hf.destroy()
  })

  it('treats numeric 1 as ascending', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3,1,1)', null, null, 3],
      [null, null, null, 1],
      [null, null, null, 2],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)
    expect(hf.getCellValue(adr('A2'))).toBe(2)
    expect(hf.getCellValue(adr('A3'))).toBe(3)
    hf.destroy()
  })
})

// ─── LINEST/LOGEST static size prediction ────────────────────────────────────

describe('LINEST size prediction with literal FALSE stats argument', () => {
  it('does not occupy rows 2-5 when stats=FALSE() is literal', () => {
    // With stats=FALSE(), result is 1 row: [slope, intercept]
    // If size prediction wrongly returns 5 rows, A2-A5 become part of the array
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A8:A10,B8:B10,TRUE(),FALSE())', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [3, 1],
      [5, 2],
      [7, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // slope
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)  // intercept
    // A2 must NOT be part of the array when stats=FALSE() (size should be 1 row)
    expect(hf.isCellPartOfArray(adr('A2'))).toBe(false)
    hf.destroy()
  })

  it('does not occupy rows 2-5 when stats=0 (numeric false) is literal', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A8:A10,B8:B10,TRUE(),0)', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [3, 1],
      [5, 2],
      [7, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)
    expect(hf.isCellPartOfArray(adr('A2'))).toBe(false)
    hf.destroy()
  })
})

describe('LOGEST size prediction with literal FALSE stats argument', () => {
  it('does not occupy rows 2-5 when stats=FALSE() is literal', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LOGEST(A8:A10,B8:B10,TRUE(),FALSE())', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [2, 1],
      [4, 2],
      [8, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // m
    expect(hf.getCellValue(adr('B1'))).toBeCloseTo(1)  // b
    expect(hf.isCellPartOfArray(adr('A2'))).toBe(false)
    hf.destroy()
  })
})

// ─── Regression function error propagation and coercion ──────────────────────

describe('LINEST known_x error propagation', () => {
  it('propagates CellError from known_x instead of silently using defaults', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A3:A5,NA())', null],
      [null, null],
      [3, null],
      [5, null],
      [7, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('LOGEST known_x error propagation', () => {
  it('propagates CellError from known_x instead of silently using defaults', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LOGEST(A3:A5,NA())', null],
      [null, null],
      [2, null],
      [4, null],
      [8, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('LINEST numeric boolean coercion', () => {
  it('treats numeric 1 as TRUE for stats argument (5-row result)', () => {
    // stats=1 should produce the 5-row result, not 1-row
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A8:A10,B8:B10,1,1)', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [3, 1],
      [5, 2],
      [7, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // slope
    // A3 should have r² when stats=1 (TRUE); if returnStats is wrongly false, A3 is null
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(1)  // r²
    hf.destroy()
  })

  it('treats numeric 0 as FALSE for stats argument (1-row result)', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LINEST(A8:A10,B8:B10,1,0)', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [3, 1],
      [5, 2],
      [7, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)
    expect(hf.isCellPartOfArray(adr('A2'))).toBe(false)
    hf.destroy()
  })
})

describe('LOGEST numeric boolean coercion', () => {
  it('treats numeric 1 as TRUE for stats argument (5-row result)', () => {
    const hf = HyperFormula.buildFromArray([
      ['=LOGEST(A8:A10,B8:B10,1,1)', null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      [2, 1],
      [4, 2],
      [8, 3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(2)  // m
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(1)  // r²
    hf.destroy()
  })
})

describe('GROWTH empty known_x argument handling', () => {
  it('uses default x values when known_x is omitted (empty arg)', () => {
    // =GROWTH(y,,new_x) — known_x skipped with ","
    // D1:D3 = [2,4,8], F1:F3 = [1,2,3]
    const hf = HyperFormula.buildFromArray([
      ['=GROWTH(D1:D3,,F1:F3)', null, null, 2, null, 1],
      [null, null, null, 4, null, 2],
      [null, null, null, 8, null, 3],
    ], gsOptions)

    // y = 1 * 2^x, so growth at x=1,2,3 should be 2,4,8
    // If empty arg causes error, A1 would be a CellError
    expect(hf.getCellValue(adr('A1'))).not.toBeInstanceOf(DetailedCellError)
    hf.destroy()
  })
})

describe('GROWTH/TREND size prediction with empty new_x argument', () => {
  it('GROWTH uses knownY size when new_x is omitted (empty arg)', () => {
    // =GROWTH(y,x,,TRUE) — new_x skipped; size should match knownY
    // D1:D3 = [2,4,8], F1:F3 = [1,2,3]
    const hf = HyperFormula.buildFromArray([
      ['=GROWTH(D1:D3,F1:F3,,TRUE())', null, null, 2, null, 1],
      [null, null, null, 4, null, 2],
      [null, null, null, 8, null, 3],
    ], gsOptions)

    // The result should have 3 rows (same as knownY), not 1x1
    expect(hf.isCellPartOfArray(adr('A2'))).toBe(true)
    expect(hf.isCellPartOfArray(adr('A3'))).toBe(true)
    hf.destroy()
  })

  it('TREND uses knownY size when new_x is omitted (empty arg)', () => {
    const hf = HyperFormula.buildFromArray([
      ['=TREND(D1:D3,F1:F3,,TRUE())', null, null, 3, null, 1],
      [null, null, null, 5, null, 2],
      [null, null, null, 7, null, 3],
    ], gsOptions)

    expect(hf.isCellPartOfArray(adr('A2'))).toBe(true)
    expect(hf.isCellPartOfArray(adr('A3'))).toBe(true)
    hf.destroy()
  })
})

// ─── SORT: RichNumber handling ───────────────────────────────────────────────

describe('SORT RichNumber handling', () => {
  it('sorts percentage values numerically', () => {
    // '30%' is stored internally as PercentNumber(0.30) — a RichNumber
    // typeOrder must use isExtendedNumber() or percentages sort as objects (last)
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3)', null, null, '30%'],
      [null, null, null, '10%'],
      [null, null, null, '20%'],
    ], gsOptions)

    // getCellValue exports raw number: 0.1, 0.2, 0.3
    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.1)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(0.2)
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(0.3)
    hf.destroy()
  })

  it('sorts percentage values descending', () => {
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3,1,FALSE())', null, null, '10%'],
      [null, null, null, '30%'],
      [null, null, null, '20%'],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.3)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(0.2)
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(0.1)
    hf.destroy()
  })

  it('sorts mixed plain numbers and percentages together numerically', () => {
    // 50% = 0.5, which sits between 0.3 and 0.7
    const hf = HyperFormula.buildFromArray([
      ['=SORT(D1:D3)', null, null, 0.7],
      [null, null, null, '50%'],
      [null, null, null, 0.3],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBeCloseTo(0.3)
    expect(hf.getCellValue(adr('A2'))).toBeCloseTo(0.5)
    expect(hf.getCellValue(adr('A3'))).toBeCloseTo(0.7)
    hf.destroy()
  })
})

// ─── FREQUENCY: RichNumber handling ──────────────────────────────────────────

describe('FREQUENCY RichNumber handling', () => {
  it('counts percentage values in bins (they are RichNumber, not plain number)', () => {
    // '10%'=0.1, '30%'=0.3, '50%'=0.5
    // Bins: [0.2, 0.4] → expected counts: [1 (≤0.2), 1 (≤0.4), 1 (>0.4), overflow]
    const hf = HyperFormula.buildFromArray([
      ['=FREQUENCY(D1:D3,E1:E2)', null, null, '10%', 0.2],
      [null, null, null, '30%', 0.4],
      [null, null, null, '50%', null],
      [null, null, null, null, null],
    ], gsOptions)

    expect(hf.getCellValue(adr('A1'))).toBe(1)  // 0.1 ≤ 0.2
    expect(hf.getCellValue(adr('A2'))).toBe(1)  // 0.3 ≤ 0.4
    expect(hf.getCellValue(adr('A3'))).toBe(1)  // 0.5 > 0.4
    hf.destroy()
  })
})
