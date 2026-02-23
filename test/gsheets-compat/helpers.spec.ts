/**
 * Tests for valuesMatch in helpers.ts.
 *
 * Verifies that numeric comparison uses strict parsing to avoid
 * false-positive matches from strings with partial numeric prefixes.
 */

import {valuesMatch} from './helpers'

describe('valuesMatch', () => {
  describe('exact matches', () => {
    it('matches identical strings', () => {
      expect(valuesMatch('hello', 'hello')).toBe(true)
    })

    it('matches identical numbers as strings', () => {
      expect(valuesMatch('42', '42')).toBe(true)
    })

    it('matches error strings', () => {
      expect(valuesMatch('#N/A', '#N/A')).toBe(true)
    })
  })

  describe('boolean matching', () => {
    it('matches TRUE case-insensitively', () => {
      expect(valuesMatch('TRUE', 'true')).toBe(true)
    })

    it('matches FALSE case-insensitively', () => {
      expect(valuesMatch('FALSE', 'false')).toBe(true)
    })
  })

  describe('numeric comparison with tolerance', () => {
    it('matches numbers within relative tolerance', () => {
      expect(valuesMatch('3.14159265', '3.14159266')).toBe(true)
    })

    it('does not match numbers outside tolerance', () => {
      expect(valuesMatch('1.0', '2.0')).toBe(false)
    })

    it('matches zero values', () => {
      expect(valuesMatch('0', '0.0')).toBe(true)
    })
  })

  describe('complex number strings (parseFloat false-positive prevention)', () => {
    it('does not match different complex numbers with same real part', () => {
      // parseFloat("4-9i") === parseFloat("4+9i") === 4, but they are different
      expect(valuesMatch('4-9i', '4+9i')).toBe(false)
    })

    it('does not match different complex numbers', () => {
      expect(valuesMatch('3+4i', '3-4i')).toBe(false)
    })

    it('does not match a complex number and its real part', () => {
      expect(valuesMatch('4+9i', '4')).toBe(false)
    })
  })

  describe('date string false-positive prevention', () => {
    it('does not match different date strings with same leading number', () => {
      // parseFloat("7/20/1969") === parseFloat("7/21/1969") === 7
      expect(valuesMatch('7/20/1969', '7/21/1969')).toBe(false)
    })

    it('does not match a date string and a plain number', () => {
      expect(valuesMatch('7/20/1969', '7')).toBe(false)
    })
  })

  describe('mismatched values', () => {
    it('does not match different error strings', () => {
      expect(valuesMatch('#N/A', '#VALUE!')).toBe(false)
    })

    it('does not match string and number', () => {
      expect(valuesMatch('hello', '42')).toBe(false)
    })
  })
})
