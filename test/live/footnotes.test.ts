import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Footnotes & Endnotes', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'A claim needs citation here' });
    await client.call('word_insert_paragraph', { text: 'Another reference point' });
    await client.call('word_insert_footnote', { anchorText: 'claim', text: 'Source: 2024 report' });
    await client.call('word_insert_endnote', { anchorText: 'reference', text: 'See appendix B' });
  });

  test('footnote created', async () => {
    if (skip()) return;
    const fn = await client.call('word_get_footnotes');
    expect(fn.count).toBe(1);
    expect(fn.footnotes[0].text).toBe('Source: 2024 report');
  });

  test('endnote created', async () => {
    if (skip()) return;
    const en = await client.call('word_get_endnotes');
    expect(en.count).toBe(1);
    expect(en.endnotes[0].text).toBe('See appendix B');
  });

  test('footnote at index', async () => {
    if (skip()) return;
    await client.call('word_insert_footnote_at_index', { paragraphIndex: 1, text: 'Index note' });
    const fn = await client.call('word_get_footnotes');
    expect(fn.count).toBe(2);
  });

  test('delete footnote and endnote', async () => {
    if (skip()) return;
    await client.call('word_delete_footnote', { index: 0 });
    const fn = await client.call('word_get_footnotes');
    expect(fn.count).toBe(1);
    await client.call('word_delete_endnote', { index: 0 });
    const en = await client.call('word_get_endnotes');
    expect(en.count).toBe(0);
  });
});
