import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Batch Operations', () => {
  test('batch inserts multiple paragraphs', async () => {
    if (skip()) return;
    await client.resetDocument();
    const r = await client.call('word_batch', {
      operations: [
        { tool: 'word_insert_paragraph', args: { text: 'Batch 1' } },
        { tool: 'word_insert_paragraph', args: { text: 'Batch 2' } },
        { tool: 'word_insert_paragraph', args: { text: 'Batch 3' } },
      ],
    });
    expect(r.completed).toBe(3);
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs.filter((p: any) => p.text.startsWith('Batch')).length).toBe(3);
  });

  test('batch stops on first error', async () => {
    if (skip()) return;
    await client.resetDocument();
    const r = await client.call('word_batch', {
      operations: [
        { tool: 'word_insert_paragraph', args: { text: 'Good' } },
        { tool: 'word_delete_paragraph', args: { index: 999 } },
        { tool: 'word_insert_paragraph', args: { text: 'Should not run' } },
      ],
    });
    expect(r.completed).toBe(1);
    expect(r.failed).toBe(1);
    const paras = await client.call('word_get_paragraphs');
    expect(paras.paragraphs.some((p: any) => p.text === 'Should not run')).toBe(false);
  });

  test('batch with empty operations gives error', async () => {
    if (skip()) return;
    const err = await client.expectError('word_batch', { operations: [] });
    expect(err).toContain('non-empty');
  });

  test('batch supports mixed operations', async () => {
    if (skip()) return;
    await client.resetDocument();
    const r = await client.call('word_batch', {
      operations: [
        { tool: 'word_insert_paragraph', args: { text: 'Hello World' } },
        { tool: 'word_search', args: { query: 'Hello' } },
        { tool: 'word_save', args: {} },
      ],
    });
    expect(r.completed).toBe(3);
    expect(r.results[1].result.count).toBe(1);
  });

  test('batch can include equation insertion', async () => {
    if (skip()) return;
    await client.resetDocument();
    const r = await client.call('word_batch', {
      operations: [
        { tool: 'word_insert_paragraph', args: { text: 'Before equation' } },
        { tool: 'word_insert_equation', args: { latex: 'x^2 + y^2 = z^2' } },
      ],
    });
    expect(r.completed).toBe(2);
  });
});
