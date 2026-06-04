import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Document & Paragraphs', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_set_document_properties', { title: 'Live Test', subject: 'Testing', category: 'QA' });
    await client.call('word_insert_paragraph', { text: 'Heading One', style: 'Heading 1' });
    await client.call('word_insert_paragraph', { text: 'Normal paragraph with words one two three' });
    await client.call('word_insert_paragraph', { text: 'Centered text', alignment: 'Center' });
    await client.call('word_insert_paragraph', { text: 'Right aligned', alignment: 'Right' });
    await client.call('word_insert_paragraph', { text: 'Paragraph to delete' });
  });

  test('document properties set correctly', async () => {
    if (skip()) return;
    const props = await client.call('word_get_document_properties');
    expect(props.title).toBe('Live Test');
    expect(props.subject).toBe('Testing');
  });

  test('word count reflects content', async () => {
    if (skip()) return;
    const wc = await client.call('word_get_word_count');
    expect(wc.words).toBeGreaterThan(10);
  });

  test('get_text returns full content', async () => {
    if (skip()) return;
    const result = await client.call('word_get_text');
    expect(result.text).toContain('Heading One');
    expect(result.text).toContain('Normal paragraph');
  });

  test('Heading 1 style applied', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const h = paras.paragraphs.find((p: any) => p.text === 'Heading One');
    expect(h.style).toBe('Heading 1');
  });

  test('alignments applied', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs.find((p: any) => p.text === 'Centered text').alignment).toBe('Center');
    expect(paras.paragraphs.find((p: any) => p.text === 'Right aligned').alignment).toBe('Right');
  });

  test('paragraph details include font info', async () => {
    if (skip()) return;
    const para = await client.call('word_get_paragraph_by_index', { index: 1 });
    expect(para).toHaveProperty('font');
    expect(para.font).toHaveProperty('name');
  });

  test('pagination with start/end', async () => {
    if (skip()) return;
    const page = await client.call('word_get_paragraphs', { start: 1, end: 3 });
    expect(page.paragraphs.length).toBe(2);
    expect(page.count).toBe(2);
    expect(page.total).toBeGreaterThan(2);
  });

  test('out-of-range start returns warning', async () => {
    if (skip()) return;
    const result = await client.call('word_get_paragraphs', { start: 100, end: 200 });
    expect(result.paragraphs.length).toBe(0);
    expect(result.warning).toContain('beyond');
  });

  test('delete paragraph removes it', async () => {
    if (skip()) return;
    const before = await client.call('word_get_paragraphs');
    const idx = before.paragraphs.findIndex((p: any) => p.text === 'Paragraph to delete');
    await client.call('word_delete_paragraph', { index: idx });
    const after = await client.call('word_get_paragraphs');
    expect(after.paragraphs.find((p: any) => p.text === 'Paragraph to delete')).toBeUndefined();
  });

  test('set paragraph spacing', async () => {
    if (skip()) return;
    await client.call('word_set_paragraph_spacing', { index: 1, lineSpacing: 24, spaceBefore: 12 });
    const para = await client.call('word_get_paragraph_by_index', { index: 1 });
    expect(para.lineSpacing).toBe(24);
    expect(para.spaceBefore).toBe(12);
  });

  test('word_clear resets document', async () => {
    if (skip()) return;
    await client.call('word_insert_paragraph', { text: 'Temp' });
    await client.call('word_clear');
    const paras = await client.call('word_get_paragraphs');
    expect(paras.count).toBe(1);
  });
});
