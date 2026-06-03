import { describe, test, expect } from 'vitest';
import {
  checkSearchLength,
  checkNonEmpty,
  checkNonNegative,
  checkBounds,
  checkOccurrence,
  normalizeAlignment,
  checkHexColor,
  checkUrl,
} from '../../src/server/validation';

describe('checkSearchLength', () => {
  test('accepts text within limit', () => {
    expect(() => checkSearchLength('short')).not.toThrow();
    expect(() => checkSearchLength('x'.repeat(255))).not.toThrow();
  });

  test('rejects text exceeding 255 chars', () => {
    expect(() => checkSearchLength('x'.repeat(256))).toThrow('too long');
    expect(() => checkSearchLength('x'.repeat(1000))).toThrow('too long');
  });
});

describe('checkNonEmpty', () => {
  test('accepts non-empty strings', () => {
    expect(() => checkNonEmpty('hello', 'field')).not.toThrow();
  });

  test('rejects empty string', () => {
    expect(() => checkNonEmpty('', 'field')).toThrow('non-empty');
  });

  test('rejects whitespace-only', () => {
    expect(() => checkNonEmpty('   ', 'field')).toThrow('non-empty');
  });

  test('rejects null/undefined', () => {
    expect(() => checkNonEmpty(null, 'field')).toThrow('non-empty');
    expect(() => checkNonEmpty(undefined, 'field')).toThrow('non-empty');
  });
});

describe('checkNonNegative', () => {
  test('accepts zero and positive integers', () => {
    expect(() => checkNonNegative(0, 'idx')).not.toThrow();
    expect(() => checkNonNegative(5, 'idx')).not.toThrow();
    expect(() => checkNonNegative(100, 'idx')).not.toThrow();
  });

  test('rejects negative', () => {
    expect(() => checkNonNegative(-1, 'idx')).toThrow('non-negative integer');
  });

  test('rejects float values', () => {
    expect(() => checkNonNegative(1.5, 'idx')).toThrow('non-negative integer');
    expect(() => checkNonNegative(0.1, 'idx')).toThrow('non-negative integer');
    expect(() => checkNonNegative(1.999, 'idx')).toThrow('non-negative integer');
  });

  test('rejects NaN and Infinity', () => {
    expect(() => checkNonNegative(NaN, 'idx')).toThrow('non-negative integer');
    expect(() => checkNonNegative(Infinity, 'idx')).toThrow('non-negative integer');
  });

  test('rejects non-number types', () => {
    expect(() => checkNonNegative('5' as any, 'idx')).toThrow('non-negative integer');
    expect(() => checkNonNegative(null as any, 'idx')).toThrow('non-negative integer');
    expect(() => checkNonNegative(undefined as any, 'idx')).toThrow('non-negative integer');
  });
});

describe('checkBounds', () => {
  test('accepts in-range index', () => {
    expect(() => checkBounds(0, 5, 'idx')).not.toThrow();
    expect(() => checkBounds(4, 5, 'idx')).not.toThrow();
  });

  test('rejects out-of-range index', () => {
    expect(() => checkBounds(5, 5, 'idx')).toThrow('out of range');
    expect(() => checkBounds(10, 3, 'idx')).toThrow('out of range');
  });
});

describe('checkOccurrence', () => {
  test('returns 0 when undefined', () => {
    expect(checkOccurrence(undefined, 3)).toBe(0);
  });

  test('returns value when in range', () => {
    expect(checkOccurrence(2, 5)).toBe(2);
  });

  test('rejects negative', () => {
    expect(() => checkOccurrence(-1, 5)).toThrow('non-negative');
  });

  test('rejects out of range', () => {
    expect(() => checkOccurrence(3, 3)).toThrow('not found');
  });
});

describe('normalizeAlignment', () => {
  test('normalizes aliases (case-insensitive)', () => {
    expect(normalizeAlignment('Left')).toBe('Left');
    expect(normalizeAlignment('left')).toBe('Left');
    expect(normalizeAlignment('LEFT')).toBe('Left');
    expect(normalizeAlignment('Center')).toBe('Centered');
    expect(normalizeAlignment('center')).toBe('Centered');
    expect(normalizeAlignment('Centered')).toBe('Centered');
    expect(normalizeAlignment('Right')).toBe('Right');
    expect(normalizeAlignment('right')).toBe('Right');
    expect(normalizeAlignment('Justify')).toBe('Justified');
    expect(normalizeAlignment('justify')).toBe('Justified');
    expect(normalizeAlignment('Justified')).toBe('Justified');
    expect(normalizeAlignment('JUSTIFIED')).toBe('Justified');
  });

  test('returns null for undefined', () => {
    expect(normalizeAlignment(undefined)).toBeNull();
  });

  test('rejects invalid values', () => {
    expect(() => normalizeAlignment('Middle')).toThrow('Invalid alignment');
    expect(() => normalizeAlignment('diagonal')).toThrow('Invalid alignment');
  });
});

describe('checkHexColor', () => {
  test('accepts valid hex colors', () => {
    expect(() => checkHexColor('#FF0000', 'color')).not.toThrow();
    expect(() => checkHexColor('#000000', 'color')).not.toThrow();
    expect(() => checkHexColor('#abcdef', 'color')).not.toThrow();
  });

  test('rejects invalid hex colors', () => {
    expect(() => checkHexColor('red', 'color')).toThrow('hex color');
    expect(() => checkHexColor('#FFF', 'color')).toThrow('hex color');
    expect(() => checkHexColor('FF0000', 'color')).toThrow('hex color');
  });
});

describe('checkUrl', () => {
  test('accepts valid URLs', () => {
    expect(() => checkUrl('https://example.com')).not.toThrow();
    expect(() => checkUrl('http://example.com/path?q=1')).not.toThrow();
  });

  test('rejects non-http URLs', () => {
    expect(() => checkUrl('ftp://example.com')).toThrow('valid HTTP or HTTPS URL');
    expect(() => checkUrl('bad')).toThrow('valid HTTP or HTTPS URL');
  });

  test('rejects URLs with unsafe characters', () => {
    expect(() => checkUrl('https://evil.com/"><script>')).toThrow('Malformed URL');
    expect(() => checkUrl('https://evil.com/path|cmd')).toThrow('Malformed URL');
  });
});
