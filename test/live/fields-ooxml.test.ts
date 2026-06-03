import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Fields return text result', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
    await client.call('word_insert_paragraph', { text: 'Link text here' });
    await client.call('word_insert_hyperlink', { anchorText: 'Link text', url: 'https://example.com' });
  });

  test('field result is a string', async () => {
    if (skip()) return;
    const fields = await client.call('word_get_fields');
    expect(fields.count).toBeGreaterThan(0);
    expect(typeof fields.fields[0].result).toBe('string');
    expect(fields.fields[0].result).not.toContain('storyType');
  });

  test('field result contains display text', async () => {
    if (skip()) return;
    const fields = await client.call('word_get_fields');
    const hyperlink = fields.fields.find((f: any) => f.type === 'Hyperlink');
    expect(hyperlink.result).toContain('Link text');
  });
});

describe('OOXML insertion validates structure', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
  });

  test('rejects arbitrary XML', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_ooxml', { ooxml: '<invalid>not real</invalid>' });
    expect(err).toContain('pkg:package');
  });

  test('rejects empty string', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_ooxml', { ooxml: '' });
    expect(err).toContain('non-empty');
  });

  test('rejects plain text', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_ooxml', { ooxml: 'just plain text' });
    expect(err).toContain('pkg:package');
  });
});
