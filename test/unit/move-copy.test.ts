import { describe, test, expect } from 'vitest';
import { buildToolRegistry } from '../../src/server/tools';

describe('Move paragraph argument validation', () => {
  const { handlers } = buildToolRegistry();
  const moveHandler = handlers.get('word_move_paragraph')!;
  const copyHandler = handlers.get('word_copy_paragraph')!;

  const mockBridge: any = { send: async () => ({ total: 5, count: 5, paragraphs: [] }) };

  test('rejects negative fromIndex', async () => {
    await expect(moveHandler({ fromIndex: -1, toIndex: 2 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('rejects negative toIndex', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: -1 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('rejects float fromIndex', async () => {
    await expect(moveHandler({ fromIndex: 1.5, toIndex: 4 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('rejects float toIndex', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 2.7 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('rejects same fromIndex and toIndex', async () => {
    await expect(moveHandler({ fromIndex: 2, toIndex: 2 }, mockBridge)).rejects.toThrow('must be different');
  });

  test('rejects count < 1', async () => {
    await expect(moveHandler({ fromIndex: 0, toIndex: 2, count: 0 }, mockBridge)).rejects.toThrow('at least 1');
  });

  test('rejects toIndex inside source range', async () => {
    await expect(moveHandler({ fromIndex: 1, toIndex: 2, count: 3 }, mockBridge)).rejects.toThrow('inside the source range');
  });

  test('rejects fromIndex out of bounds', async () => {
    await expect(moveHandler({ fromIndex: 10, toIndex: 2 }, mockBridge)).rejects.toThrow('out of range');
  });

  test('copy rejects negative fromIndex', async () => {
    await expect(copyHandler({ fromIndex: -1, toIndex: 2 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('copy rejects float fromIndex', async () => {
    await expect(copyHandler({ fromIndex: 0.5, toIndex: 2 }, mockBridge)).rejects.toThrow('non-negative integer');
  });

  test('copy rejects count < 1', async () => {
    await expect(copyHandler({ fromIndex: 0, toIndex: 2, count: 0 }, mockBridge)).rejects.toThrow('at least 1');
  });

  test('copy rejects out of bounds', async () => {
    await expect(copyHandler({ fromIndex: 10, toIndex: 2 }, mockBridge)).rejects.toThrow('out of range');
  });
});
