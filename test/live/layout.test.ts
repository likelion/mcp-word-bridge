import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Headers, Images & Layout', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_set_header_footer', { type: 'header', text: 'Test Header' });
    await client.call('word_set_header_footer', { type: 'footer', text: 'Page Footer' });
    const PIXEL = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    await client.call('word_insert_image', { base64: PIXEL, altText: 'TestPixel', width: 50 });
    await client.call('word_insert_paragraph', { text: 'After image' });
  });

  test('header and footer set', async () => {
    if (skip()) return;
    const h = await client.call('word_get_header_footer', { type: 'header' });
    expect(h.text).toBe('Test Header');
    const f = await client.call('word_get_header_footer', { type: 'footer' });
    expect(f.text).toBe('Page Footer');
  });

  test('image exists with alt text', async () => {
    if (skip()) return;
    const images = await client.call('word_get_images');
    expect(images.count).toBe(1);
    expect(images.images[0].altText).toBe('TestPixel');
  });

  test('page layout readable and settable', async () => {
    if (skip()) return;
    await client.call('word_set_page_layout', { orientation: 'Landscape' });
    const layout = await client.call('word_get_page_layout');
    expect(layout.orientation).toBe('Landscape');
    await client.call('word_set_page_layout', { orientation: 'Portrait' });
  });

  test('section break', async () => {
    if (skip()) return;
    await client.call('word_insert_section_break', { paragraphIndex: 0 });
    const sections = await client.call('word_get_sections');
    expect(sections.count).toBe(2);
  });

  test('delete image', async () => {
    if (skip()) return;
    await client.call('word_delete_image', { index: 0 });
    const images = await client.call('word_get_images');
    expect(images.count).toBe(0);
  });
});

describe('Figure caption and deletion', () => {
  const PIXEL = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_image', { base64: PIXEL, altText: 'Diagram', caption: 'System architecture' });
  });

  test('image paragraph is centered', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const capIdx = paras.paragraphs.findIndex((p: any) => p.style === 'Caption');
    expect(capIdx).toBeGreaterThanOrEqual(1);
    // The image sits in the paragraph directly above its caption.
    expect(paras.paragraphs[capIdx - 1].alignment).toBe('Center');
  });

  test('figure caption is centered and directly below the image', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const caption = paras.paragraphs.find((p: any) => p.style === 'Caption');
    expect(caption).toBeDefined();
    expect(caption.text).toBe('Figure 1: System architecture');
    expect(caption.alignment).toBe('Center');
  });

  test('trailing spacer after caption resets to Normal (not Caption)', async () => {
    if (skip()) return;
    const paras = await client.call('word_get_paragraphs');
    const capIdx = paras.paragraphs.findIndex((p: any) => p.style === 'Caption');
    const after = paras.paragraphs[capIdx + 1];
    expect(after).toBeDefined();
    expect(after.style).toBe('Normal');
    expect(after.text).toBe('');
  });

  test('delete image removes the figure and its caption', async () => {
    if (skip()) return;
    const result = await client.call('word_delete_image', { index: 0 });
    expect(result.success).toBe(true);
    expect(result.captionDeleted).toBe(true);
    const images = await client.call('word_get_images');
    expect(images.count).toBe(0);
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs.some((p: any) => p.style === 'Caption')).toBe(false);
  });
});
