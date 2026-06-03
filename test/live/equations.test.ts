import { describe, test, expect, beforeAll, afterAll } from 'vitest';
import { getClient, skip, cleanup } from './setup';
import type { LiveTestClient } from './client';

let client: LiveTestClient;
beforeAll(async () => { client = await getClient(); }, 60000);
afterAll(cleanup);

describe('Equations', () => {
  beforeAll(async () => {
    if (skip()) return;
    await client.resetDocument();
  });

  test('insert display equation', async () => {
    if (skip()) return;
    const r = await client.call('word_insert_equation', { latex: '\\frac{a}{b}' });
    expect(r.success).toBe(true);
    expect(r.displayMode).toBe(true);
  });

  test('insert inline equation with anchor', async () => {
    if (skip()) return;
    await client.call('word_insert_paragraph', { text: 'Inline here: ' });
    const r = await client.call('word_insert_equation', { latex: 'E=mc^2', displayMode: false, anchorText: 'Inline here:' });
    expect(r.success).toBe(true);
    expect(r.displayMode).toBe(false);
  });

  test('invalid LaTeX errors', async () => {
    if (skip()) return;
    const err = await client.expectError('word_insert_equation', { latex: '\\frac{' });
    expect(err).toContain('LaTeX parse error');
  });
});
