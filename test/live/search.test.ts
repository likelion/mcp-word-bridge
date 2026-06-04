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

  test('search and replace with ^p creates paragraph break', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'AAA BBB CCC' });
    await client.call('word_search_and_replace', { find: 'BBB', replace: 'X^pY' });
    const paras = await client.call('word_get_paragraphs');
    const texts = paras.paragraphs.map((p: any) => p.text);
    // ^p creates a paragraph break: "X" ends one paragraph, "Y" starts the next
    expect(texts).toContain('AAA X');
    expect(texts.some((t: string) => t.startsWith('Y'))).toBe(true);
  });

  test('search and replace with ^t inserts tab', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'left right' });
    await client.call('word_search_and_replace', { find: 'left right', replace: 'left^tright' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('left\tright');
  });

  test('search and replace with ^^ inserts literal caret', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'replace me' });
    await client.call('word_search_and_replace', { find: 'replace me', replace: '^^p is not a break' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('^p is not a break');
  });

  test('insert text at match with ^p creates paragraph break', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Hello World' });
    await client.call('word_insert_text_at_match', { text: '^p', after: 'Hello' });
    const paras = await client.call('word_get_paragraphs');
    const texts = paras.paragraphs.map((p: any) => p.text);
    expect(texts.some((t: string) => t.includes('Hello'))).toBe(true);
    expect(texts.some((t: string) => t.includes('World'))).toBe(true);
    // They should be in separate paragraphs
    const helloIdx = texts.findIndex((t: string) => t.includes('Hello'));
    const worldIdx = texts.findIndex((t: string) => t.includes('World'));
    expect(worldIdx).toBeGreaterThan(helloIdx);
  });

  test('insert text at match with ^t inserts tab', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'before after' });
    await client.call('word_insert_text_at_match', { text: '^t', after: 'before' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('before\t');
  });

  test('plain text replacement still works (no codes)', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'The old text here' });
    await client.call('word_search_and_replace', { find: 'old', replace: 'new' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('The new text here');
  });

  test('wildcard search finds pattern matches', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'cat bat hat mat' });
    const r = await client.call('word_search', { query: '[cbhm]at', matchWildcards: true });
    expect(r.count).toBe(4);
  });

  test('search and replace with ^+ inserts em dash', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'A -- B' });
    await client.call('word_search_and_replace', { find: ' -- ', replace: '^+' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('A\u2014B');
  });

  test('search and replace with ^= inserts en dash', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'pages 1-10' });
    await client.call('word_search_and_replace', { find: '1-10', replace: '1^=10' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('1\u201310');
  });

  test('search and replace with ^m inserts page break', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Chapter end marker next chapter' });
    await client.call('word_search_and_replace', { find: 'marker', replace: '^m' });
    const paras = await client.call('word_get_paragraphs');
    // A page break splits into separate paragraphs
    expect(paras.total).toBeGreaterThanOrEqual(2);
  });

  test('search with matchPrefix finds word beginnings', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'international intercept splintered' });
    const r = await client.call('word_search', { query: 'inter', matchPrefix: true });
    expect(r.count).toBe(2); // international, intercept — not splintered
  });

  test('search with matchSuffix finds word endings', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'within begin interesting' });
    const r = await client.call('word_search', { query: 'in', matchSuffix: true });
    expect(r.count).toBe(2); // within, begin — not interesting
  });

  test('search with ignorePunct matches across punctuation', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'video, you are great' });
    const r = await client.call('word_search', { query: 'video you', ignorePunct: true });
    expect(r.count).toBe(1);
  });

  test('search with ignoreSpace matches across extra whitespace', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'hello     world' });
    const r = await client.call('word_search', { query: 'hello world', ignoreSpace: true });
    expect(r.count).toBe(1);
  });

  test('search and replace with ^~ inserts non-breaking hyphen', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'self control' });
    await client.call('word_search_and_replace', { find: 'self control', replace: 'self^~control' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('self\u2011control');
  });

  test('search and replace with ^= and ^+ inserts dashes', async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'A-B--C' });
    await client.call('word_search_and_replace', { find: 'A-B--C', replace: 'A^=B^+C' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('A\u2013B\u2014C');
  });
});
