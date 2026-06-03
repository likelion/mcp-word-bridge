import { describe, test, expect } from 'vitest';
import { MAX_SEARCH_LENGTH, MAX_BATCH_OPERATIONS, TIMEOUT_DEFAULT, TIMEOUT_HEAVY, HEAVY_OPERATIONS, MAX_PAYLOAD } from '../../src/shared/constants';

describe('Constants', () => {
  test('MAX_SEARCH_LENGTH is 255', () => {
    expect(MAX_SEARCH_LENGTH).toBe(255);
  });

  test('MAX_BATCH_OPERATIONS is 50', () => {
    expect(MAX_BATCH_OPERATIONS).toBe(50);
  });

  test('TIMEOUT_DEFAULT is 30 seconds', () => {
    expect(TIMEOUT_DEFAULT).toBe(30_000);
  });

  test('TIMEOUT_HEAVY is 60 seconds', () => {
    expect(TIMEOUT_HEAVY).toBe(60_000);
  });

  test('HEAVY_OPERATIONS includes expected actions', () => {
    expect(HEAVY_OPERATIONS.has('insertHtml')).toBe(true);
    expect(HEAVY_OPERATIONS.has('insertOoxml')).toBe(true);
    expect(HEAVY_OPERATIONS.has('getStyles')).toBe(true);
    expect(HEAVY_OPERATIONS.has('insertTableOfContents')).toBe(true);
    expect(HEAVY_OPERATIONS.has('batchExecute')).toBe(true);
  });

  test('MAX_PAYLOAD is 10MB', () => {
    expect(MAX_PAYLOAD).toBe(10 * 1024 * 1024);
  });
});
