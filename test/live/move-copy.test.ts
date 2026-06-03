import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Move paragraph preserves rich content', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Para zero' });
    await client.call('word_insert_paragraph', { text: 'Para one with footnote' });
    await client.call('word_insert_footnote', { anchorText: 'footnote', text: 'Important reference' });
    await client.call('word_format_text', { text: 'one', bold: true });
    await client.call('word_insert_hyperlink', { anchorText: 'footnote', url: 'https://example.com' });
    await client.call('word_insert_paragraph', { text: 'Para two' });
    await client.call('word_insert_paragraph', { text: 'Para three destination' });
  });

  test('footnote preserved after move', async () => {
    if (skip()) return;
    await client.call('word_move_paragraph', { fromIndex: 1, toIndex: 3 });
    const fn = await client.call('word_get_footnotes');
    expect(fn.count).toBe(1);
    expect(fn.footnotes[0].text).toBe('Important reference');
  });

  test('bold formatting preserved after move', async () => {
    if (skip()) return;
    const info = await client.call('word_get_font_info', { text: 'one' });
    expect(info.bold).toBe(true);
  });

  test('hyperlink preserved after move', async () => {
    if (skip()) return;
    const links = await client.call('word_get_hyperlinks');
    expect(links.count).toBeGreaterThan(0);
    expect(links.hyperlinks[0].url).toBe('https://example.com');
  });

  test('response includes toIndexRequested', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'A' });
    await client.call('word_insert_paragraph', { text: 'B' });
    await client.call('word_insert_paragraph', { text: 'C' });
    await client.call('word_insert_paragraph', { text: 'D' });
    const result = await client.call('word_move_paragraph', { fromIndex: 0, toIndex: 3 });
    expect(result.moved.toIndexRequested).toBe(3);
    expect(typeof result.moved.to).toBe('number');
  });
});

describe('Copy paragraph preserves rich content', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Source paragraph with note' });
    await client.call('word_insert_footnote', { anchorText: 'note', text: 'Copied footnote' });
    await client.call('word_format_text', { text: 'Source', bold: true });
    await client.call('word_insert_paragraph', { text: 'Destination marker' });
  });

  test('copy creates duplicate with footnote', async () => {
    if (skip()) return;
    await client.call('word_copy_paragraph', { fromIndex: 0, toIndex: 1 });
    const fn = await client.call('word_get_footnotes');
    expect(fn.count).toBe(2);
  });

  test('source unchanged after copy', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs[0].text).toContain('Source paragraph');
  });
});
