import { describe, test, expect } from 'vitest';
import { parseSpecialCodes, hasSpecialCodes } from '../../src/taskpane/commands/special-codes';

describe('hasSpecialCodes', () => {
  test('returns false for plain text', () => {
    expect(hasSpecialCodes('hello world')).toBe(false);
    expect(hasSpecialCodes('')).toBe(false);
    expect(hasSpecialCodes('no codes here')).toBe(false);
  });

  test('returns true for ^p', () => {
    expect(hasSpecialCodes('^p')).toBe(true);
    expect(hasSpecialCodes('before^pafter')).toBe(true);
    expect(hasSpecialCodes('^P')).toBe(true);
  });

  test('returns true for ^l', () => {
    expect(hasSpecialCodes('^l')).toBe(true);
    expect(hasSpecialCodes('^L')).toBe(true);
  });

  test('returns true for ^t', () => {
    expect(hasSpecialCodes('^t')).toBe(true);
    expect(hasSpecialCodes('^T')).toBe(true);
  });

  test('returns true for ^s', () => {
    expect(hasSpecialCodes('^s')).toBe(true);
    expect(hasSpecialCodes('^S')).toBe(true);
  });

  test('returns true for ^m (page break)', () => {
    expect(hasSpecialCodes('^m')).toBe(true);
    expect(hasSpecialCodes('^M')).toBe(true);
  });

  test('returns true for ^n (column break)', () => {
    expect(hasSpecialCodes('^n')).toBe(true);
    expect(hasSpecialCodes('^N')).toBe(true);
  });

  test('returns true for ^~ (non-breaking hyphen)', () => {
    expect(hasSpecialCodes('^~')).toBe(true);
  });

  test('returns true for ^- (optional hyphen)', () => {
    expect(hasSpecialCodes('^-')).toBe(true);
  });

  test('returns true for ^+ (em dash)', () => {
    expect(hasSpecialCodes('^+')).toBe(true);
  });

  test('returns true for ^= (en dash)', () => {
    expect(hasSpecialCodes('^=')).toBe(true);
  });

  test('returns true for ^^ (escaped caret)', () => {
    expect(hasSpecialCodes('^^')).toBe(true);
    expect(hasSpecialCodes('a^^b')).toBe(true);
  });

  test('returns false for ^ followed by non-code letter', () => {
    expect(hasSpecialCodes('^x')).toBe(false);
    expect(hasSpecialCodes('^z')).toBe(false);
    expect(hasSpecialCodes('^1')).toBe(false);
  });

  test('returns false for find-only codes (^?, ^#, ^$, ^w, etc.)', () => {
    expect(hasSpecialCodes('^?')).toBe(false);
    expect(hasSpecialCodes('^#')).toBe(false);
    expect(hasSpecialCodes('^$')).toBe(false);
    expect(hasSpecialCodes('^w')).toBe(false);
    expect(hasSpecialCodes('^f')).toBe(false);
    expect(hasSpecialCodes('^e')).toBe(false);
    expect(hasSpecialCodes('^d')).toBe(false);
    expect(hasSpecialCodes('^g')).toBe(false);
    expect(hasSpecialCodes('^b')).toBe(false);
    expect(hasSpecialCodes('^v')).toBe(false);
  });
});

describe('parseSpecialCodes', () => {
  test('plain text returns single text segment', () => {
    const result = parseSpecialCodes('hello world');
    expect(result).toEqual([{ type: 'text', content: 'hello world' }]);
  });

  test('empty string returns empty text segment', () => {
    const result = parseSpecialCodes('');
    expect(result).toEqual([{ type: 'text', content: '' }]);
  });

  test('^p produces paragraph segment', () => {
    const result = parseSpecialCodes('^p');
    expect(result).toEqual([{ type: 'paragraph' }]);
  });

  test('^l produces lineBreak segment', () => {
    const result = parseSpecialCodes('^l');
    expect(result).toEqual([{ type: 'lineBreak' }]);
  });

  test('^t produces tab character in text segment', () => {
    const result = parseSpecialCodes('^t');
    expect(result).toEqual([{ type: 'text', content: '\t' }]);
  });

  test('^s produces non-breaking space in text segment', () => {
    const result = parseSpecialCodes('^s');
    expect(result).toEqual([{ type: 'text', content: '\u00A0' }]);
  });

  test('^m produces pageBreak segment', () => {
    const result = parseSpecialCodes('^m');
    expect(result).toEqual([{ type: 'pageBreak' }]);
  });

  test('^n produces columnBreak segment', () => {
    const result = parseSpecialCodes('^n');
    expect(result).toEqual([{ type: 'columnBreak' }]);
  });

  test('^~ produces non-breaking hyphen', () => {
    const result = parseSpecialCodes('^~');
    expect(result).toEqual([{ type: 'text', content: '\u2011' }]);
  });

  test('^- produces optional/soft hyphen', () => {
    const result = parseSpecialCodes('^-');
    expect(result).toEqual([{ type: 'text', content: '\u00AD' }]);
  });

  test('^+ produces em dash', () => {
    const result = parseSpecialCodes('^+');
    expect(result).toEqual([{ type: 'text', content: '\u2014' }]);
  });

  test('^= produces en dash', () => {
    const result = parseSpecialCodes('^=');
    expect(result).toEqual([{ type: 'text', content: '\u2013' }]);
  });

  test('text before and after ^p', () => {
    const result = parseSpecialCodes('Hello^pWorld');
    expect(result).toEqual([
      { type: 'text', content: 'Hello' },
      { type: 'paragraph' },
      { type: 'text', content: 'World' },
    ]);
  });

  test('text with ^t in the middle', () => {
    const result = parseSpecialCodes('col1^tcol2^tcol3');
    expect(result).toEqual([
      { type: 'text', content: 'col1' },
      { type: 'text', content: '\t' },
      { type: 'text', content: 'col2' },
      { type: 'text', content: '\t' },
      { type: 'text', content: 'col3' },
    ]);
  });

  test('multiple ^p creates multiple paragraph breaks', () => {
    const result = parseSpecialCodes('^p^p');
    expect(result).toEqual([
      { type: 'paragraph' },
      { type: 'paragraph' },
    ]);
  });

  test('^p at end of text', () => {
    const result = parseSpecialCodes('text^p');
    expect(result).toEqual([
      { type: 'text', content: 'text' },
      { type: 'paragraph' },
    ]);
  });

  test('^p at start of text', () => {
    const result = parseSpecialCodes('^ptext');
    expect(result).toEqual([
      { type: 'paragraph' },
      { type: 'text', content: 'text' },
    ]);
  });

  test('^^ produces literal caret', () => {
    const result = parseSpecialCodes('^^');
    expect(result).toEqual([{ type: 'text', content: '^' }]);
  });

  test('^^p produces literal ^p (not a paragraph break)', () => {
    const result = parseSpecialCodes('^^p');
    expect(result).toEqual([{ type: 'text', content: '^p' }]);
  });

  test('^^^p produces literal ^ followed by paragraph break', () => {
    const result = parseSpecialCodes('^^^p');
    expect(result).toEqual([
      { type: 'text', content: '^' },
      { type: 'paragraph' },
    ]);
  });

  test('mixed codes in sequence', () => {
    const result = parseSpecialCodes('A^tB^pC^lD');
    expect(result).toEqual([
      { type: 'text', content: 'A' },
      { type: 'text', content: '\t' },
      { type: 'text', content: 'B' },
      { type: 'paragraph' },
      { type: 'text', content: 'C' },
      { type: 'lineBreak' },
      { type: 'text', content: 'D' },
    ]);
  });

  test('codes are case-insensitive (letter codes)', () => {
    const lower = parseSpecialCodes('^p^t^l^s^m^n');
    const upper = parseSpecialCodes('^P^T^L^S^M^N');
    expect(lower).toEqual(upper);
  });

  test('unknown codes are passed through as literal text', () => {
    const result = parseSpecialCodes('a^xb');
    expect(result).toEqual([{ type: 'text', content: 'a^xb' }]);
  });

  test('find-only codes (^?, ^#, ^$, ^w, ^f, ^e, ^d, ^g) are NOT expanded', () => {
    // These codes only have meaning in search queries, not in replacement text.
    // They should pass through as literal text if someone puts them in a replacement.
    expect(parseSpecialCodes('^?')).toEqual([{ type: 'text', content: '^?' }]);
    expect(parseSpecialCodes('^#')).toEqual([{ type: 'text', content: '^#' }]);
    expect(parseSpecialCodes('^$')).toEqual([{ type: 'text', content: '^$' }]);
    expect(parseSpecialCodes('^w')).toEqual([{ type: 'text', content: '^w' }]);
    expect(parseSpecialCodes('^f')).toEqual([{ type: 'text', content: '^f' }]);
    expect(parseSpecialCodes('^e')).toEqual([{ type: 'text', content: '^e' }]);
    expect(parseSpecialCodes('^d')).toEqual([{ type: 'text', content: '^d' }]);
    expect(parseSpecialCodes('^g')).toEqual([{ type: 'text', content: '^g' }]);
    expect(parseSpecialCodes('^b')).toEqual([{ type: 'text', content: '^b' }]);
    expect(parseSpecialCodes('^v')).toEqual([{ type: 'text', content: '^v' }]);
    expect(parseSpecialCodes('^%')).toEqual([{ type: 'text', content: '^%' }]);
  });

  test('mixed letter and symbol codes', () => {
    const result = parseSpecialCodes('A^+B^=C^~D');
    expect(result).toEqual([
      { type: 'text', content: 'A' },
      { type: 'text', content: '\u2014' },
      { type: 'text', content: 'B' },
      { type: 'text', content: '\u2013' },
      { type: 'text', content: 'C' },
      { type: 'text', content: '\u2011' },
      { type: 'text', content: 'D' },
    ]);
  });

  test('^^ in middle of text', () => {
    const result = parseSpecialCodes('price is 5^^10');
    expect(result).toEqual([{ type: 'text', content: 'price is 5^10' }]);
  });

  test('complex: escaped caret before code', () => {
    const result = parseSpecialCodes('a^^^^p');
    // ^^ = ^, ^^ = ^, then nothing left for ^p pattern
    // Actually: ^^^^ -> two placeholders, then ^p is a code
    // Wait: "a^^^^p" -> replace ^^ -> "a\0\0p" -> no ^[plts] found -> "a^^p"
    // Hmm, let me think. Input: a^^^^p
    // Pass 1: replace all ^^ with placeholder: a + \0 + \0 + p => "a\u0000\u0000p"
    // Pass 2: scan for ^[plts] — none found (no ^ left)
    // Result: single text segment "a^^p" (after restore)
    expect(result).toEqual([{ type: 'text', content: 'a^^p' }]);
  });

  test('escaped caret followed by code: ^^^t', () => {
    // Input: ^^^t
    // Pass 1: first ^^ -> placeholder, leaving ^t
    // "placeholder + ^t"
    // Pass 2: ^t found at position 1
    // Before ^t: placeholder -> restored to "^"
    // ^t -> tab
    const result = parseSpecialCodes('^^^t');
    expect(result).toEqual([
      { type: 'text', content: '^' },
      { type: 'text', content: '\t' },
    ]);
  });
});
