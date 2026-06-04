import { describe, test, expect } from 'vitest';
import { createMutex } from '../../src/server/mutex';

describe('createMutex', () => {
  test('executes a single function and returns its result', async () => {
    const mutex = createMutex();
    const result = await mutex.run(async () => 42);
    expect(result).toBe(42);
  });

  test('propagates rejections', async () => {
    const mutex = createMutex();
    await expect(mutex.run(async () => { throw new Error('boom'); })).rejects.toThrow('boom');
  });

  test('continues after a rejection (does not deadlock)', async () => {
    const mutex = createMutex();
    await expect(mutex.run(async () => { throw new Error('fail'); })).rejects.toThrow('fail');
    const result = await mutex.run(async () => 'recovered');
    expect(result).toBe('recovered');
  });

  test('serializes concurrent calls in FIFO order', async () => {
    const mutex = createMutex();
    const order: number[] = [];

    // Launch 5 concurrent tasks with varying internal delays
    const tasks = [0, 1, 2, 3, 4].map(i =>
      mutex.run(async () => {
        // Simulate async work — later entries have shorter delays
        // Without serialization, they'd resolve out of order
        await delay(10 - i * 2);
        order.push(i);
        return i;
      }),
    );

    const results = await Promise.all(tasks);
    expect(results).toEqual([0, 1, 2, 3, 4]);
    expect(order).toEqual([0, 1, 2, 3, 4]);
  });

  test('does not interleave multi-step operations', async () => {
    const mutex = createMutex();
    const log: string[] = [];

    const taskA = mutex.run(async () => {
      log.push('A-start');
      await delay(10);
      log.push('A-end');
    });

    const taskB = mutex.run(async () => {
      log.push('B-start');
      await delay(5);
      log.push('B-end');
    });

    await Promise.all([taskA, taskB]);
    // B must not start until A finishes
    expect(log).toEqual(['A-start', 'A-end', 'B-start', 'B-end']);
  });

  test('a rejection does not skip subsequent entries', async () => {
    const mutex = createMutex();
    const order: string[] = [];

    const t1 = mutex.run(async () => { order.push('1'); });
    const t2 = mutex.run(async () => { order.push('2'); throw new Error('x'); });
    const t3 = mutex.run(async () => { order.push('3'); });

    await t1;
    await expect(t2).rejects.toThrow('x');
    await t3;
    expect(order).toEqual(['1', '2', '3']);
  });

  test('handles high concurrency without deadlock', async () => {
    const mutex = createMutex();
    let counter = 0;

    const tasks = Array.from({ length: 100 }, () =>
      mutex.run(async () => {
        const before = counter;
        await delay(0); // yield to event loop
        counter = before + 1;
        return counter;
      }),
    );

    const results = await Promise.all(tasks);
    // Without mutex, counter would be corrupted by concurrent increments.
    // With mutex, each increment is atomic → final counter = 100.
    expect(counter).toBe(100);
    expect(results).toEqual(Array.from({ length: 100 }, (_, i) => i + 1));
  });

  test('preserves return types (generic correctness)', async () => {
    const mutex = createMutex();
    const str: string = await mutex.run(async () => 'hello');
    const num: number = await mutex.run(async () => 123);
    const obj: { x: number } = await mutex.run(async () => ({ x: 1 }));
    expect(str).toBe('hello');
    expect(num).toBe(123);
    expect(obj).toEqual({ x: 1 });
  });
});

function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}
