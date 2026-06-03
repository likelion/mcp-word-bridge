/**
 * Live regression tests for input validation and edge cases.
 * Requires Word with the MCP Word Bridge add-in loaded.
 */
import { describe, test, expect, beforeAll, afterAll, beforeEach } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

// =============================================================================
// search_and_replace special code rejection
// =============================================================================
describe('search_and_replace special code rejection', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Test paragraph one' });
    await client.call('word_insert_paragraph', { text: 'Test paragraph two' });
  });

  test('rejects ^p in find', async () => {
    if (skip()) return;
    const err = await client.expectError('word_search_and_replace', { find: '^p', replace: 'X' });
    expect(err).toContain('special code');
  });

  test('rejects ^13 in find', async () => {
    if (skip()) return;
    const err = await client.expectError('word_search_and_replace', { find: '^13', replace: 'X' });
    expect(err).toContain('special code');
  });

  test('rejects ^w in find', async () => {
    if (skip()) return;
    const err = await client.expectError('word_search_and_replace', { find: '^w', replace: 'X' });
    expect(err).toContain('special code');
  });

  test('rejects ^p in replace', async () => {
    if (skip()) return;
    const err = await client.expectError('word_search_and_replace', { find: 'Test', replace: '^p' });
    expect(err).toContain('special code');
  });

  test('normal replacement still works', async () => {
    if (skip()) return;
    const result = await client.call('word_search_and_replace', { find: 'one', replace: '1' });
    expect(result.replacements).toBe(1);
  });
});

// =============================================================================
// Bookmark persistence after search
// =============================================================================
describe('Bookmark persistence after search', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Test paragraph one' });
    await client.call('word_insert_bookmark', { name: 'bm_test', anchorText: 'one' });
  });

  test('bookmark exists before replacement', async () => {
    if (skip()) return;
    const bookmarks = await client.call('word_get_bookmarks');
    expect(bookmarks.bookmarks).toContain('bm_test');
  });
});

// =============================================================================
// Content control duplicate tag rejection
// =============================================================================
describe('Content control duplicate tag rejection', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'First control text' });
    await client.call('word_insert_paragraph', { text: 'Second control text' });
    await client.call('word_insert_content_control', { anchorText: 'First', tag: 'dup_tag' });
    await client.call('word_insert_content_control', { anchorText: 'Second', tag: 'dup_tag' });
  });

  test('rejects set_content_control_text with duplicate tag', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_content_control_text', { tag: 'dup_tag', text: 'new' });
    expect(err).toContain('Multiple content controls');
    expect(err).toContain('dup_tag');
  });
});

// =============================================================================
// Non-integer index rejection
// =============================================================================
describe('Non-integer index rejection', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Para 0' });
    await client.call('word_insert_paragraph', { text: 'Para 1' });
    await client.call('word_insert_paragraph', { text: 'Para 2' });
  });

  test('get_paragraph_by_index rejects float', async () => {
    if (skip()) return;
    const err = await client.expectError('word_get_paragraph_by_index', { index: 1.5 });
    expect(err).toContain('non-negative integer');
  });

  test('insert_paragraph_at_index rejects float', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_paragraph_at_index', { index: 1.999, text: 'x' });
    expect(err).toContain('non-negative integer');
  });

  test('copy_paragraph rejects float count', async () => {
    if (skip()) return;
    const err = await client.expectError('word_copy_paragraph', { fromIndex: 0, toIndex: 2, count: 1.7 });
    expect(err).toContain('integer');
  });
});

// =============================================================================
// Spacing bounds validation
// =============================================================================
describe('Spacing bounds validation', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Spacing test' });
  });

  test('rejects leftIndent of 99999', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, leftIndent: 99999 });
    expect(err).toContain('exceeds maximum');
  });

  test('rejects rightIndent of 99999', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, rightIndent: 99999 });
    expect(err).toContain('exceeds maximum');
  });

  test('rejects firstLineIndent of 99999', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, firstLineIndent: 99999 });
    expect(err).toContain('exceeds maximum');
  });

  test('accepts normal spacing values', async () => {
    if (skip()) return;
    const result = await client.call('word_set_paragraph_spacing', { index: 0, leftIndent: 36, spaceBefore: 12 });
    expect(result.success).toBe(true);
  });
});

// =============================================================================
// add_table_row partial values warning
// =============================================================================
describe('add_table_row partial values warning', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_table', { rows: 2, cols: 3, data: [['A', 'B', 'C'], ['1', '2', '3']] });
  });

  test('warns when fewer values than columns', async () => {
    if (skip()) return;
    const result = await client.call('word_add_table_row', { tableIndex: 0, values: ['one'] });
    expect(result.success).toBe(true);
    expect(result.warning).toContain('1 of 3 cells populated');
  });

  test('no warning when all values provided', async () => {
    if (skip()) return;
    const result = await client.call('word_add_table_row', { tableIndex: 0, values: ['x', 'y', 'z'] });
    expect(result.success).toBe(true);
    expect(result.warning).toBeUndefined();
  });

  test('still errors when too many values', async () => {
    if (skip()) return;
    const err = await client.expectError('word_add_table_row', { tableIndex: 0, values: ['a', 'b', 'c', 'd'] });
    expect(err).toContain('only has 3 columns');
  });
});

// =============================================================================
// Custom property key length validation
// =============================================================================
describe('Custom property key length validation', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
  });

  test('rejects key longer than 255 chars', async () => {
    if (skip()) return;
    const longKey = 'k'.repeat(280);
    const err = await client.expectError('word_set_custom_property', { key: longKey, value: 'val' });
    expect(err).toContain('255 characters or fewer');
  });

  test('accepts key at 255 chars', async () => {
    if (skip()) return;
    const key = 'k'.repeat(255);
    const result = await client.call('word_set_custom_property', { key, value: 'val' });
    expect(result.success).toBe(true);
  });
});

// =============================================================================
// insert_hyperlink on already-linked text
// =============================================================================
describe('insert_hyperlink replaces existing link with warning', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Click here for link' });
    await client.call('word_insert_hyperlink', { anchorText: 'here', url: 'https://first.com' });
  });

  test('second hyperlink replaces first with warning', async () => {
    if (skip()) return;
    const result = await client.call('word_insert_hyperlink', { anchorText: 'here', url: 'https://second.com' });
    expect(result.success).toBe(true);
    expect(result.warning).toContain('Replaced existing hyperlink');
    expect(result.warning).toContain('https://first.com');
  });
});
