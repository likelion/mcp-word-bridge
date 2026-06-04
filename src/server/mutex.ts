/**
 * Simple async mutex — queues callers so only one executes at a time.
 * Used to serialize MCP tool calls and prevent concurrent document mutations.
 */
export interface Mutex {
  /** Execute fn exclusively. If another fn is running, wait in FIFO order. */
  run<T>(fn: () => Promise<T>): Promise<T>;
}

export function createMutex(): Mutex {
  let chain: Promise<unknown> = Promise.resolve();

  return {
    run<T>(fn: () => Promise<T>): Promise<T> {
      // Chain this fn after the current tail. Whether the previous entry
      // resolved or rejected, we always proceed to the next.
      const next = chain.then(fn, fn);
      // Update the chain tail, swallowing the result/error so the chain
      // itself never rejects (which would skip entries).
      chain = next.then(noop, noop);
      return next;
    },
  };
}

function noop(): void {}
