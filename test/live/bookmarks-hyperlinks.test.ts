import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Bookmarks & Hyperlinks', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Bookmark target text here' });
    await client.call('word_insert_paragraph', { text: 'Visit example site today' });
    await client.call('word_insert_bookmark', { name: 'TestBM', anchorText: 'target text' });
    await client.call('word_insert_hyperlink', { anchorText: 'example', url: 'https://example.com' });
  });

  test('bookmark exists and has text', async () => {
    if (skip()) return;
    const bm = await client.call('word_get_bookmarks');
    expect(bm.bookmarks).toContain('TestBM');
    const r = await client.call('word_get_bookmark_text', { name: 'TestBM' });
    expect(r.text).toBe('target text');
  });

  test('hyperlink created', async () => {
    if (skip()) return;
    const links = await client.call('word_get_hyperlinks');
    expect(links.count).toBeGreaterThan(0);
    expect(links.hyperlinks[0].url).toBe('https://example.com');
  });

  test('remove hyperlink keeps text', async () => {
    if (skip()) return;
    await client.call('word_remove_hyperlink', { anchorText: 'example' });
    const text = await client.call('word_get_text');
    expect(text.text).toContain('example');
  });

  test('delete bookmark', async () => {
    if (skip()) return;
    await client.call('word_delete_bookmark', { name: 'TestBM' });
    const bm = await client.call('word_get_bookmarks');
    expect(bm.bookmarks).not.toContain('TestBM');
  });

  test('invalid bookmark name rejected', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_bookmark', { name: 'has spaces', anchorText: 'Bookmark' });
    expect(err).toContain('Invalid bookmark name');
  });

  test('unsafe URL characters rejected', async () => {
    if (skip()) return;
    await client.call('word_insert_paragraph', { text: 'link anchor' });
    const err = await client.expectError('word_insert_hyperlink', { anchorText: 'link', url: 'https://evil.com/"><script>' });
    expect(err).toContain('Malformed URL');
  });
});
