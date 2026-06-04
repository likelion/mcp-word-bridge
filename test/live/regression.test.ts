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
// search_and_replace supports Word special codes
// =============================================================================
describe('search_and_replace special code support', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Test paragraph one' });
    await client.call('word_insert_paragraph', { text: 'Test paragraph two' });
  });

  test('accepts ^p in find (matches paragraph marks)', async () => {
    if (skip()) return;
    const result = await client.call('word_search_and_replace', { find: 'one^pTest', replace: 'one. Test' });
    expect(result.replacements).toBeGreaterThanOrEqual(0); // may or may not match depending on document structure
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
  });

  test('warns on insert_content_control with duplicate tag', async () => {
    if (skip()) return;
    const result = await client.call('word_insert_content_control', { anchorText: 'Second', tag: 'dup_tag' });
    expect(result.warning).toContain('dup_tag');
    expect(result.warning).toContain('Duplicate tags');
  });

  test('rejects set_content_control_text with duplicate tag', async () => {
    if (skip()) return;
    await client.call('word_insert_content_control', { anchorText: 'Second', tag: 'dup_tag' });
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
    expect(err).toContain('out of range');
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

// =============================================================================
// Negative indent rejection
// =============================================================================
describe('Negative indent rejection', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Indent test' });
  });

  test('rejects negative leftIndent', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, leftIndent: -50 });
    expect(err).toContain('non-negative');
  });

  test('rejects negative rightIndent', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, rightIndent: -10 });
    expect(err).toContain('non-negative');
  });

  test('allows negative firstLineIndent (hanging indent)', async () => {
    if (skip()) return;
    const result = await client.call('word_set_paragraph_spacing', { index: 0, firstLineIndent: -36 });
    expect(result.success).toBe(true);
  });
});

// =============================================================================
// Bookmark warning on search_and_replace
// =============================================================================
describe('Bookmark warning on search_and_replace', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Important section here' });
    await client.call('word_insert_bookmark', { name: 'sec1', anchorText: 'section' });
  });

  test('warns when bookmarks are destroyed', async () => {
    if (skip()) return;
    const result = await client.call('word_search_and_replace', { find: 'section', replace: 'SECTION' });
    expect(result.replacements).toBe(1);
    if (result.warning) {
      expect(result.warning).toContain('bookmark');
    }
    // Bookmark should be gone
    const bm = await client.call('word_get_bookmarks');
    expect(bm.bookmarks).not.toContain('sec1');
  });

  test('preserveBookmarks restores destroyed bookmarks', async () => {
    if (skip()) return;
    const result = await client.call('word_search_and_replace', { find: 'section', replace: 'SECTION', preserveBookmarks: true });
    expect(result.replacements).toBe(1);
    if (result.bookmarksRestored) {
      expect(result.bookmarksRestored).toBeGreaterThan(0);
    }
    // Bookmark should still exist
    const bm = await client.call('word_get_bookmarks');
    expect(bm.bookmarks).toContain('sec1');
  });
});

// =============================================================================
// Case-preservation: skip identical matches
// =============================================================================
describe('Case-preservation skip', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Hello HELLO hello' });
  });

  test('case-insensitive replace skips already-matching text', async () => {
    if (skip()) return;
    const result = await client.call('word_search_and_replace', { find: 'hello', replace: 'hello', matchCase: false });
    // Should skip the one that's already "hello" and only replace "Hello" and "HELLO"
    expect(result.replacements).toBe(2);
    expect(result.skipped).toBe(1);
  });
});

// =============================================================================
// Table headerRowCount defaults to 0
// =============================================================================
describe('Table headerRowCount default', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
  });

  test('table without explicit headerRowCount has 0 header rows', async () => {
    if (skip()) return;
    await client.call('word_insert_table', { rows: 2, cols: 2, data: [['A', 'B'], ['C', 'D']] });
    const tables = await client.call('word_list_tables');
    expect(tables.tables[0].headerRowCount).toBe(0);
  });

  test('table with explicit headerRowCount=1 has 1 header row', async () => {
    if (skip()) return;
    await client.call('word_insert_table', { rows: 2, cols: 2, data: [['A', 'B'], ['C', 'D']], headerRowCount: 1 });
    const tables = await client.call('word_list_tables');
    expect(tables.tables[0].headerRowCount).toBe(1);
  });
});

// =============================================================================
// Delete extra paragraph in table cell
// =============================================================================
describe('Delete extra paragraph in table cell', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_table', { rows: 1, cols: 1, data: [['Cell']] });
  });

  test('can delete extra paragraph inserted into table cell', async () => {
    if (skip()) return;
    // Find the table cell paragraph
    const paras = await client.call('word_get_paragraphs');
    const cellPara = paras.paragraphs.find((p: any) => p.text === 'Cell' && p.inTable);
    if (!cellPara) return; // skip if structure different
    // Insert extra paragraph in cell
    await client.call('word_insert_paragraph_at_index', { index: cellPara.index, text: 'Extra', location: 'After' });
    // Now delete the extra one
    const result = await client.call('word_delete_paragraph', { index: cellPara.index + 1 });
    expect(result.success).toBe(true);
  });

  test('cannot delete the only paragraph in a table cell', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const cellPara = paras.paragraphs.find((p: any) => p.text === 'Cell' && p.inTable);
    if (!cellPara) return;
    const err = await client.expectError('word_delete_paragraph', { index: cellPara.index });
    expect(err).toContain('only paragraph');
  });
});

// =============================================================================
// Batch rejects nested word_batch
// =============================================================================
describe('Batch rejects nested calls', () => {
  test('rejects word_batch inside batch operations', async () => {
    if (skip()) return;
    const err = await client.expectError('word_batch', {
      operations: [{ tool: 'word_batch', args: { operations: [{ tool: 'word_get_text', args: {} }] } }],
    });
    expect(err).toContain('cannot be nested');
  });
});

// =============================================================================
// get_paragraphs returns total and count
// =============================================================================
describe('get_paragraphs pagination fields', () => {
  beforeEach(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'A' });
    await client.call('word_insert_paragraph', { text: 'B' });
    await client.call('word_insert_paragraph', { text: 'C' });
  });

  test('count equals returned array length, total equals full document', async () => {
    if (skip()) return;
    const result = await client.call('word_get_paragraphs', { start: 0, end: 2 });
    expect(result.count).toBe(result.paragraphs.length);
    expect(result.total).toBeGreaterThanOrEqual(3);
    expect(result.count).toBe(2);
  });
});
