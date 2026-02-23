import {DetailedCellError, HyperFormula, RawCellContent} from '../src'
import {adr} from './testUtils'
import {GoogleSheetsTextFunctionsPlugin} from '../src/interpreter/plugin/googleSheets/GoogleSheetsTextFunctionsPlugin'

describe('GoogleSheetsTextFunctionsPlugin', () => {
  beforeEach(() => {
    HyperFormula.registerFunctionPlugin(GoogleSheetsTextFunctionsPlugin)
  })

  const buildWithPlugin = (data: RawCellContent[][]) =>
    HyperFormula.buildFromArray(data, {
      licenseKey: 'gpl-v3',
    })

  describe('REGEXMATCH', () => {
    it('returns true when text matches regex', () => {
      const hf = buildWithPlugin([['=REGEXMATCH("hello world", "world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns false when text does not match regex', () => {
      const hf = buildWithPlugin([['=REGEXMATCH("hello world", "xyz")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('is case sensitive by default', () => {
      const hf = buildWithPlugin([['=REGEXMATCH("Hello", "hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe(false)
      hf.destroy()
    })

    it('supports regex patterns', () => {
      const hf = buildWithPlugin([['=REGEXMATCH("abc123", "\\d+")']])
      expect(hf.getCellValue(adr('A1'))).toBe(true)
      hf.destroy()
    })

    it('returns #VALUE! for invalid regex', () => {
      const hf = buildWithPlugin([['=REGEXMATCH("hello", "[")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('REGEXEXTRACT', () => {
    it('returns full match when no capture group', () => {
      const hf = buildWithPlugin([['=REGEXEXTRACT("hello world", "\\w+")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello')
      hf.destroy()
    })

    it('returns first capture group when one exists', () => {
      const hf = buildWithPlugin([['=REGEXEXTRACT("price: $12.50", "\\$(\\d+\\.\\d+)")']])
      expect(hf.getCellValue(adr('A1'))).toBe('12.50')
      hf.destroy()
    })

    it('returns #N/A when no match found', () => {
      const hf = buildWithPlugin([['=REGEXEXTRACT("hello", "\\d+")']])
      const result = hf.getCellValue(adr('A1'))
      expect(result).toBeInstanceOf(DetailedCellError)
      expect((result as DetailedCellError).type).toBe('NA')
      hf.destroy()
    })

    it('returns #VALUE! for invalid regex', () => {
      const hf = buildWithPlugin([['=REGEXEXTRACT("hello", "[")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('REGEXREPLACE', () => {
    it('replaces all occurrences', () => {
      const hf = buildWithPlugin([['=REGEXREPLACE("foo bar foo", "foo", "baz")']])
      expect(hf.getCellValue(adr('A1'))).toBe('baz bar baz')
      hf.destroy()
    })

    it('replaces with regex pattern', () => {
      const hf = buildWithPlugin([['=REGEXREPLACE("aabbcc", "[abc]", "x")']])
      expect(hf.getCellValue(adr('A1'))).toBe('xxxxxx')
      hf.destroy()
    })

    it('supports regex groups in replacement', () => {
      const hf = buildWithPlugin([['=REGEXREPLACE("2024-01-15", "(\\d{4})-(\\d{2})-(\\d{2})", "$3/$2/$1")']])
      expect(hf.getCellValue(adr('A1'))).toBe('15/01/2024')
      hf.destroy()
    })

    it('returns #VALUE! for invalid regex', () => {
      const hf = buildWithPlugin([['=REGEXREPLACE("hello", "[", "x")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('JOIN', () => {
    it('joins values with delimiter', () => {
      const hf = buildWithPlugin([['=JOIN("-", "a", "b", "c")']])
      expect(hf.getCellValue(adr('A1'))).toBe('a-b-c')
      hf.destroy()
    })

    it('joins two values', () => {
      const hf = buildWithPlugin([['=JOIN(", ", "hello", "world")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello, world')
      hf.destroy()
    })

    it('handles numeric values', () => {
      const hf = buildWithPlugin([['=JOIN("+", 1, 2, 3)']])
      expect(hf.getCellValue(adr('A1'))).toBe('1+2+3')
      hf.destroy()
    })

    it('joins single value', () => {
      const hf = buildWithPlugin([['=JOIN("-", "only")']])
      expect(hf.getCellValue(adr('A1'))).toBe('only')
      hf.destroy()
    })
  })

  describe('TEXTJOIN', () => {
    it('joins texts with delimiter ignoring empty strings', () => {
      const hf = buildWithPlugin([['=TEXTJOIN("-", TRUE(), "a", "", "b")']])
      expect(hf.getCellValue(adr('A1'))).toBe('a-b')
      hf.destroy()
    })

    it('keeps empty strings when ignore_empty is false', () => {
      const hf = buildWithPlugin([['=TEXTJOIN("-", FALSE(), "a", "", "b")']])
      expect(hf.getCellValue(adr('A1'))).toBe('a--b')
      hf.destroy()
    })

    it('handles multiple texts', () => {
      const hf = buildWithPlugin([['=TEXTJOIN(", ", TRUE(), "one", "two", "three")']])
      expect(hf.getCellValue(adr('A1'))).toBe('one, two, three')
      hf.destroy()
    })

    it('returns empty string when all values are empty and ignore_empty is true', () => {
      const hf = buildWithPlugin([['=TEXTJOIN(",", TRUE(), "", "")']])
      expect(hf.getCellValue(adr('A1'))).toBe('')
      hf.destroy()
    })
  })

  describe('DOLLAR', () => {
    it('formats with 2 decimal places by default', () => {
      const hf = buildWithPlugin([['=DOLLAR(1234.567)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1,234.57')
      hf.destroy()
    })

    it('formats with specified decimal places', () => {
      const hf = buildWithPlugin([['=DOLLAR(1.2351, 4)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1.2351')
      hf.destroy()
    })

    it('formats large numbers with commas', () => {
      const hf = buildWithPlugin([['=DOLLAR(1234567.89, 2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1,234,567.89')
      hf.destroy()
    })

    it('formats negative numbers', () => {
      const hf = buildWithPlugin([['=DOLLAR(-42.5, 2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('-$42.50')
      hf.destroy()
    })

    it('rounds to specified decimal places', () => {
      const hf = buildWithPlugin([['=DOLLAR(1234.567, 2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1,234.57')
      hf.destroy()
    })

    it('rounds to the left of decimal with negative decimals', () => {
      const hf = buildWithPlugin([['=DOLLAR(1234, -2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1,200')
      hf.destroy()
    })

    it('rounds to nearest 10 with decimals=-1', () => {
      const hf = buildWithPlugin([['=DOLLAR(1234, -1)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$1,230')
      hf.destroy()
    })

    it('rounds to nearest 1000 with decimals=-3', () => {
      const hf = buildWithPlugin([['=DOLLAR(1750, -3)']])
      expect(hf.getCellValue(adr('A1'))).toBe('$2,000')
      hf.destroy()
    })

    it('returns #VALUE! for decimals exceeding engine limit (> 20)', () => {
      const hf = buildWithPlugin([['=DOLLAR(1.5, 21)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('FIXED', () => {
    it('formats with 2 decimal places by default', () => {
      const hf = buildWithPlugin([['=FIXED(1234.5678)']])
      expect(hf.getCellValue(adr('A1'))).toBe('1,234.57')
      hf.destroy()
    })

    it('formats with specified decimals and no_commas=true', () => {
      const hf = buildWithPlugin([['=FIXED(966364281, 4, TRUE())']])
      expect(hf.getCellValue(adr('A1'))).toBe('966364281.0000')
      hf.destroy()
    })

    it('includes commas when no_commas is false', () => {
      const hf = buildWithPlugin([['=FIXED(966364281, 4, FALSE())']])
      expect(hf.getCellValue(adr('A1'))).toBe('966,364,281.0000')
      hf.destroy()
    })

    it('rounds to specified decimals', () => {
      const hf = buildWithPlugin([['=FIXED(3.14159, 2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('3.14')
      hf.destroy()
    })

    it('handles negative numbers', () => {
      const hf = buildWithPlugin([['=FIXED(-1234.5, 1, TRUE())']])
      expect(hf.getCellValue(adr('A1'))).toBe('-1234.5')
      hf.destroy()
    })

    it('rounds to the left of decimal with negative decimals', () => {
      const hf = buildWithPlugin([['=FIXED(1234, -2)']])
      expect(hf.getCellValue(adr('A1'))).toBe('1,200')
      hf.destroy()
    })

    it('rounds to nearest 10 with decimals=-1', () => {
      const hf = buildWithPlugin([['=FIXED(1234, -1)']])
      expect(hf.getCellValue(adr('A1'))).toBe('1,230')
      hf.destroy()
    })

    it('rounds to nearest 1000 with decimals=-3 and no_commas=true', () => {
      const hf = buildWithPlugin([['=FIXED(1750, -3, TRUE())']])
      expect(hf.getCellValue(adr('A1'))).toBe('2000')
      hf.destroy()
    })

    it('returns #VALUE! for decimals exceeding engine limit (> 20)', () => {
      const hf = buildWithPlugin([['=FIXED(1.5, 21)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('ASC', () => {
    it('converts full-width ASCII to half-width', () => {
      const hf = buildWithPlugin([[`=ASC("\uFF21\uFF22\uFF23")`]])
      expect(hf.getCellValue(adr('A1'))).toBe('ABC')
      hf.destroy()
    })

    it('converts full-width digits to half-width', () => {
      const hf = buildWithPlugin([[`=ASC("\uFF11\uFF12\uFF13")`]])
      expect(hf.getCellValue(adr('A1'))).toBe('123')
      hf.destroy()
    })

    it('converts full-width space to regular space', () => {
      const hf = buildWithPlugin([[`=ASC("\u3000")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(' ')
      hf.destroy()
    })

    it('leaves regular ASCII unchanged', () => {
      const hf = buildWithPlugin([['=ASC("hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello')
      hf.destroy()
    })
  })

  describe('LENB', () => {
    it('returns byte length for ASCII text', () => {
      const hf = buildWithPlugin([['=LENB("hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe(5)
      hf.destroy()
    })

    it('returns byte length for multi-byte characters', () => {
      const hf = buildWithPlugin([[`=LENB("\u65E5")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(3)
      hf.destroy()
    })

    it('returns 0 for empty string', () => {
      const hf = buildWithPlugin([['=LENB("")']])
      expect(hf.getCellValue(adr('A1'))).toBe(0)
      hf.destroy()
    })

    it('counts bytes for mixed ASCII and multi-byte', () => {
      // "a日b" = 1 + 3 + 1 = 5 bytes
      const hf = buildWithPlugin([[`=LENB("a\u65E5b")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(5)
      hf.destroy()
    })

    it('counts 4 bytes for a surrogate-pair emoji (not 6)', () => {
      // 😀 (U+1F600) is 4 bytes in UTF-8, but 2 UTF-16 code units
      const hf = buildWithPlugin([[`=LENB("\uD83D\uDE00")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })
  })

  describe('LEFTB', () => {
    it('returns leftmost bytes for ASCII text', () => {
      const hf = buildWithPlugin([['=LEFTB("hello", 3)']])
      expect(hf.getCellValue(adr('A1'))).toBe('hel')
      hf.destroy()
    })

    it('handles multi-byte characters - requests partial character bytes', () => {
      // "日本語" each char is 3 bytes. Requesting 3 bytes = first char "日"
      const hf = buildWithPlugin([[`=LEFTB("\u65E5\u672C\u8A9E", 3)`]])
      expect(hf.getCellValue(adr('A1'))).toBe('\u65E5')
      hf.destroy()
    })

    it('returns default 1 byte when no count given', () => {
      const hf = buildWithPlugin([['=LEFTB("hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe('h')
      hf.destroy()
    })

    it('returns empty string for 0 bytes', () => {
      const hf = buildWithPlugin([['=LEFTB("hello", 0)']])
      expect(hf.getCellValue(adr('A1'))).toBe('')
      hf.destroy()
    })

    it('returns #VALUE! for negative byte count', () => {
      const hf = buildWithPlugin([['=LEFTB("hello", -1)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('RIGHTB', () => {
    it('returns rightmost bytes for ASCII text', () => {
      const hf = buildWithPlugin([['=RIGHTB("hello", 3)']])
      expect(hf.getCellValue(adr('A1'))).toBe('llo')
      hf.destroy()
    })

    it('handles multi-byte characters', () => {
      // "日本語" each char is 3 bytes. Requesting last 3 bytes = "語"
      const hf = buildWithPlugin([[`=RIGHTB("\u65E5\u672C\u8A9E", 3)`]])
      expect(hf.getCellValue(adr('A1'))).toBe('\u8A9E')
      hf.destroy()
    })

    it('returns default 1 byte when no count given', () => {
      const hf = buildWithPlugin([['=RIGHTB("hello")']])
      expect(hf.getCellValue(adr('A1'))).toBe('o')
      hf.destroy()
    })

    it('returns #VALUE! for negative byte count', () => {
      const hf = buildWithPlugin([['=RIGHTB("hello", -1)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })

  describe('MIDB', () => {
    it('extracts bytes from middle of ASCII text', () => {
      const hf = buildWithPlugin([['=MIDB("hello world", 7, 5)']])
      expect(hf.getCellValue(adr('A1'))).toBe('world')
      hf.destroy()
    })

    it('handles multi-byte characters', () => {
      // "日本語" each char is 3 bytes. Start at byte 4 (2nd char), take 3 bytes
      const hf = buildWithPlugin([[`=MIDB("\u65E5\u672C\u8A9E", 4, 3)`]])
      expect(hf.getCellValue(adr('A1'))).toBe('\u672C')
      hf.destroy()
    })

    it('returns #VALUE! when start_num is less than 1', () => {
      const hf = buildWithPlugin([['=MIDB("hello", 0, 3)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('returns #VALUE! for negative byte count', () => {
      const hf = buildWithPlugin([['=MIDB("hello", 1, -1)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('starts at byte position 1 (first byte)', () => {
      const hf = buildWithPlugin([['=MIDB("hello", 1, 3)']])
      expect(hf.getCellValue(adr('A1'))).toBe('hel')
      hf.destroy()
    })
  })

  describe('FINDB', () => {
    it('finds byte position in ASCII text', () => {
      const hf = buildWithPlugin([['=FINDB("world", "hello world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(7)
      hf.destroy()
    })

    it('is case-sensitive', () => {
      const hf = buildWithPlugin([['=FINDB("World", "hello world")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('finds byte position with multi-byte characters', () => {
      // "日本語" - each char is 3 bytes. "語" starts at byte 7
      const hf = buildWithPlugin([[`=FINDB("\u8A9E", "\u65E5\u672C\u8A9E")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(7)
      hf.destroy()
    })

    it('returns #VALUE! when not found', () => {
      const hf = buildWithPlugin([['=FINDB("z", "hello")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('respects start position', () => {
      const hf = buildWithPlugin([['=FINDB("l", "hello", 4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })

    it('returns #VALUE! when start_num is less than 1', () => {
      const hf = buildWithPlugin([['=FINDB("h", "hello", 0)']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('treats start_num as a byte position not a character index', () => {
      // "日本語" each char is 3 bytes; "本" starts at byte 4
      // FINDB("本", "日本語", 4) should find "本" starting from byte 4
      const hf = buildWithPlugin([[`=FINDB("\u672C", "\u65E5\u672C\u8A9E", 4)`]])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })
  })

  describe('SEARCHB', () => {
    it('finds byte position case-insensitively', () => {
      const hf = buildWithPlugin([['=SEARCHB("WORLD", "hello world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(7)
      hf.destroy()
    })

    it('is case-insensitive', () => {
      const hf = buildWithPlugin([['=SEARCHB("Hello", "hello world")']])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })

    it('finds byte position with multi-byte characters', () => {
      // "日本語" - each char is 3 bytes. "語" starts at byte 7
      const hf = buildWithPlugin([[`=SEARCHB("\u8A9E", "\u65E5\u672C\u8A9E")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(7)
      hf.destroy()
    })

    it('returns #VALUE! when not found', () => {
      const hf = buildWithPlugin([['=SEARCHB("z", "hello")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })

    it('respects start position', () => {
      const hf = buildWithPlugin([['=SEARCHB("l", "hello", 4)']])
      expect(hf.getCellValue(adr('A1'))).toBe(4)
      hf.destroy()
    })

    it('returns byte position from original text when lowercase expands byte length', () => {
      // Turkish İ (U+0130, 2 bytes) lowercases to i + combining dot (3 bytes)
      // 's' in 'İstanbul' is at byte 3 in original (İ=2 bytes + s)
      // but would be at byte 4 if computed from lowercased string
      const hf = buildWithPlugin([[`=SEARCHB("s", "\u0130stanbul")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(3)
      hf.destroy()
    })

    it('finds a character whose toLowerCase expands to multiple chars', () => {
      // SEARCHB("İ", "İstanbul") - searching for İ within İstanbul
      // İ is at byte position 1 (first char, 2 bytes in UTF-8)
      const hf = buildWithPlugin([[`=SEARCHB("\u0130", "\u0130stanbul")`]])
      expect(hf.getCellValue(adr('A1'))).toBe(1)
      hf.destroy()
    })
  })

  describe('REPLACEB', () => {
    it('replaces bytes in ASCII text', () => {
      const hf = buildWithPlugin([['=REPLACEB("hello world", 7, 5, "earth")']])
      expect(hf.getCellValue(adr('A1'))).toBe('hello earth')
      hf.destroy()
    })

    it('replaces bytes in text with multi-byte characters', () => {
      // "日本語" - replace 2nd char (bytes 4-6) with "x"
      const hf = buildWithPlugin([[`=REPLACEB("\u65E5\u672C\u8A9E", 4, 3, "x")`]])
      expect(hf.getCellValue(adr('A1'))).toBe('\u65E5x\u8A9E')
      hf.destroy()
    })

    it('can insert text without removing existing', () => {
      const hf = buildWithPlugin([['=REPLACEB("hello", 3, 0, "XX")']])
      expect(hf.getCellValue(adr('A1'))).toBe('heXXllo')
      hf.destroy()
    })

    it('can delete bytes by replacing with empty string', () => {
      const hf = buildWithPlugin([['=REPLACEB("hello", 2, 3, "")']])
      expect(hf.getCellValue(adr('A1'))).toBe('ho')
      hf.destroy()
    })

    it('returns #VALUE! when start_num is less than 1', () => {
      const hf = buildWithPlugin([['=REPLACEB("hello", 0, 1, "x")']])
      expect(hf.getCellValue(adr('A1'))).toBeInstanceOf(DetailedCellError)
      hf.destroy()
    })
  })
})
