import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Lists', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_list', { items: ['Apple', 'Banana', 'Cherry'], numbered: true });
  });

  test('list items created', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs.filter((p: any) => p.isListItem).length).toBe(3);
  });

  test('list info returns level', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const idx = paras.paragraphs.findIndex((p: any) => p.isListItem);
    const info = await client.call('word_get_list_info', { index: idx });
    expect(info.isListItem).toBe(true);
    expect(info.level).toBe(0);
  });

  test('change list level', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const idx = paras.paragraphs.findIndex((p: any) => p.text === 'Banana');
    await client.call('word_set_list_level', { index: idx, level: 1 });
    const info = await client.call('word_get_list_info', { index: idx });
    expect(info.level).toBe(1);
  });
});
