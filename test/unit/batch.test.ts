import { describe, test, expect } from 'vitest';
import { buildToolRegistry } from '../../src/server/tools';

describe('Batch tool argument validation', () => {
  const { handlers } = buildToolRegistry();
  const batchHandler = handlers.get('word_batch')!;

  // Mock bridge that returns a valid batch result
  const mockBridge: any = {
    send: async () => ({ results: [] }),
  };

  test('rejects empty operations array', async () => {
    await expect(batchHandler({ operations: [] }, mockBridge)).rejects.toThrow('non-empty');
  });

  test('rejects missing operations', async () => {
    await expect(batchHandler({}, mockBridge)).rejects.toThrow('non-empty');
  });

  test('rejects > 50 operations', async () => {
    const ops = Array.from({ length: 51 }, () => ({ tool: 'word_get_text', args: {} }));
    await expect(batchHandler({ operations: ops }, mockBridge)).rejects.toThrow('50');
  });

  test('reports error for operation without tool name', async () => {
    const result = await batchHandler({ operations: [{ args: {} }] }, mockBridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.results[0].success).toBe(false);
    expect(parsed.results[0].error).toContain('Missing tool name');
  });

  test('reports error for unknown tool', async () => {
    const result = await batchHandler({ operations: [{ tool: 'word_nonexistent', args: {} }] }, mockBridge);
    const parsed = JSON.parse(result.content[0].text);
    expect(parsed.results[0].success).toBe(false);
    expect(parsed.results[0].error).toContain('Unknown tool');
  });

  test('rejects nested word_batch calls', async () => {
    await expect(batchHandler({
      operations: [{ tool: 'word_batch', args: { operations: [{ tool: 'word_get_text', args: {} }] } }],
    }, mockBridge)).rejects.toThrow('cannot be nested');
  });
});
