import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Document Outline', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Chapter One', style: 'Heading 1' });
    await client.call('word_insert_paragraph', { text: 'Section A', style: 'Heading 2' });
    await client.call('word_insert_paragraph', { text: 'Body text' });
    await client.call('word_insert_paragraph', { text: 'Chapter Two', style: 'Heading 1' });
  });

  test('returns headings up to maxLevel', async () => {
    if (skip()) return;
    const outline = await client.call('word_get_document_outline', { maxLevel: 1 });
    expect(outline.outline.length).toBe(2);
    expect(outline.outline[0].text).toBe('Chapter One');
    expect(outline.outline[1].text).toBe('Chapter Two');
  });

  test('includes lower-level headings when maxLevel allows', async () => {
    if (skip()) return;
    const outline = await client.call('word_get_document_outline', { maxLevel: 3 });
    expect(outline.outline.length).toBe(3);
    expect(outline.outline.find((h: any) => h.text === 'Section A')).toBeDefined();
  });

  test('includes paragraph indices', async () => {
    if (skip()) return;
    const outline = await client.call('word_get_document_outline');
    for (const h of outline.outline) {
      expect(typeof h.index).toBe('number');
    }
  });
});

describe('Outline after TOC insertion', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Chapter One', style: 'Heading 1' });
    await client.call('word_insert_paragraph', { text: 'Body of chapter one' });
    await client.call('word_insert_paragraph', { text: 'Chapter Two', style: 'Heading 1' });
    await client.call('word_insert_paragraph', { text: 'Body of chapter two' });
    await client.call('word_insert_table_of_contents', { location: 'Start' });
  });

  test('outline includes all headings', async () => {
    if (skip()) return;
    const outline = await client.call('word_get_document_outline', { maxLevel: 3 });
    const texts = outline.outline.map((h: any) => h.text);
    expect(texts).toContain('Chapter One');
    expect(texts).toContain('Chapter Two');
  });

  test('no duplicates from TOC entries', async () => {
    if (skip()) return;
    const outline = await client.call('word_get_document_outline', { maxLevel: 3 });
    const chapterOnes = outline.outline.filter((h: any) => h.text === 'Chapter One');
    expect(chapterOnes.length).toBe(1);
  });
});
