import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Search, Text & Formatting', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'The quick brown fox jumps over the lazy dog' });
    await client.call('word_insert_paragraph', { text: 'Hello HELLO hello world' });
    await client.call('word_insert_paragraph', { text: 'cat concatenate category' });
    await client.call('word_insert_paragraph', { text: 'Format bold italic red big courier' });
  });

  test('search finds text', async () => {
    if (skip()) return;
    const r = await client.call('word_search', { query: 'quick' });
    expect(r.count).toBe(1);
  });

  test('search case-insensitive', async () => {
    if (skip()) return;
    const r = await client.call('word_search', { query: 'hello' });
    expect(r.count).toBe(3);
  });

  test('search case-sensitive', async () => {
    if (skip()) return;
    const r = await client.call('word_search', { query: 'Hello', matchCase: true });
    expect(r.count).toBe(1);
  });

  test('search whole word', async () => {
    if (skip()) return;
    const r = await client.call('word_search', { query: 'cat', matchWholeWord: true });
    expect(r.count).toBe(1);
  });

  test('insert text after anchor', async () => {
    if (skip()) return;
    await client.call('word_insert_text_at_match', { text: ' INSERTED', after: 'quick' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('quick INSERTED');
  });

  test('search and replace', async () => {
    if (skip()) return;
    await client.call('word_search_and_replace', { find: 'lazy', replace: 'energetic' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('energetic');
  });

  test('apply bold and verify', async () => {
    if (skip()) return;
    await client.call('word_format_text', { text: 'bold', bold: true });
    const info = await client.call('word_get_font_info', { text: 'bold' });
    expect(info.bold).toBe(true);
  });

  test('apply color and verify', async () => {
    if (skip()) return;
    await client.call('word_format_text', { text: 'red', color: '#FF0000' });
    const info = await client.call('word_get_font_info', { text: 'red' });
    expect(info.color).toBe('#FF0000');
  });

  test('clear formatting resets', async () => {
    if (skip()) return;
    await client.call('word_format_text', { text: 'italic', italic: true });
    await client.call('word_clear_formatting', { text: 'italic' });
    const info = await client.call('word_get_font_info', { text: 'italic' });
    expect(info.italic).toBe(false);
  });

  test('long search string rejected gracefully', async () => {
    if (skip()) return;
    const err = await client.expectError('word_format_text', { text: 'x'.repeat(260), bold: true });
    expect(err).toContain('too long');
    expect(err).not.toContain('SearchStringInvalidOrTooLong');
  });
});
