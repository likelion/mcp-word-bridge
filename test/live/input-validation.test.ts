import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Input Validation', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'validation content' });
    await client.call('word_insert_table', { rows: 3, cols: 3 });
  });

  test('negative start index rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_get_paragraphs', { start: -1, end: 0 });
    expect(err).toContain('non-negative');
  });

  test('negative margin rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_page_layout', { topMargin: -10 });
    expect(err).toContain('non-negative');
  });

  test('invalid URL rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_hyperlink', { anchorText: 'validation', url: 'bad' });
    expect(err).toContain('valid HTTP or HTTPS URL');
  });

  test('backwards merge range error', async () => {
    if (skip()) return;
    const err = await client.expectError('word_merge_table_cells', { tableIndex: 0, topRow: 2, firstCell: 0, bottomRow: 0, lastCell: 0 });
    expect(err).toContain('less than or equal to bottomRow');
  });

  test('invalid alignment rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_paragraph_at_index', { index: 0, text: 'Bad', alignment: 'Middle' });
    expect(err).toContain('Invalid alignment');
  });

  test('invalid style gives friendly error', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_paragraph_at_index', { index: 0, text: 'Bad', style: 'FAKE_STYLE' });
    expect(err).toContain('Style not found');
  });

  test('empty anchorText gives friendly error', async () => {
    if (skip()) return;
    const err = await client.expectError('word_format_text', { text: '', bold: true });
    expect(err).toContain('non-empty');
  });

  test('size 0 rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_format_text', { text: 'validation', size: 0 });
    expect(err).toContain('size must be positive');
  });

  test('change tracking invalid mode', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_change_tracking', { mode: 'BadMode' });
    expect(err).toContain('Invalid mode');
  });

  test('negative footnote index', async () => {
    if (skip()) return;
    const err = await client.expectError('word_delete_footnote', { index: -1 });
    expect(err).toContain('non-negative');
  });

  test('format_text with no formatting properties rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_format_text', { text: 'validation' });
    expect(err).toContain('At least one formatting property');
  });

  test('merge single cell rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_merge_table_cells', { tableIndex: 0, topRow: 0, firstCell: 0, bottomRow: 0, lastCell: 0 });
    expect(err).toContain('single cell');
  });

  test('bookmark name exceeding 40 chars rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_bookmark', { name: 'a_very_long_bookmark_name_that_exceeds_the_limit', anchorText: 'validation' });
    expect(err).toContain('40-character maximum');
  });

  test('lineSpacing below 1 point rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_set_paragraph_spacing', { index: 0, lineSpacing: 0.5 });
    expect(err).toContain('below the minimum');
  });

  test('get_paragraph_by_index includes inTable field', async () => {
    if (skip()) return;
    const normal = await client.call('word_get_paragraph_by_index', { index: 0 });
    expect(normal).toHaveProperty('inTable');
    expect(normal.inTable).toBe(false);
  });

  test('accept_all_tracked_changes returns count', async () => {
    if (skip()) return;
    const result = await client.call('word_accept_all_tracked_changes');
    expect(result).toHaveProperty('count');
    expect(result.count).toBe(0);
  });

  test('insert_content_control without anchorText rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_content_control', {});
    expect(err).toContain('anchorText');
    expect(err).toContain('non-empty');
  });
});
