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
    expect(err).toContain('cannot be empty');
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
});
