/**
 * Vitest reporter for the live suite. After all tests finish, it writes a
 * completion marker into the connected Word document (e.g.
 * "✅ Live tests complete — 148 passed, 0 failed") so the run leaves a clear
 * end-of-suite stamp instead of the residue of whichever suite happened to run
 * last. No-ops silently when Word is not connected.
 */
import type { Reporter } from 'vitest/node';
import { LiveTestClient } from './client';

/** Recursively count test results by state across the task tree. */
function countTests(tasks: any[], acc: { passed: number; failed: number }): void {
  for (const task of tasks) {
    if (task.type === 'test' || task.type === 'custom') {
      const state = task.result?.state;
      if (state === 'pass') acc.passed++;
      else if (state === 'fail') acc.failed++;
    }
    if (task.tasks) countTests(task.tasks, acc);
  }
}

export default class CompletionReporter implements Reporter {
  async onFinished(files: any[] = []): Promise<void> {
    const acc = { passed: 0, failed: 0 };
    countTests(files, acc);

    const status = acc.failed === 0 ? '✅' : '❌';
    const marker = `${status} Live tests complete — ${acc.passed} passed, ${acc.failed} failed`;

    const client = new LiveTestClient();
    try {
      await client.connect();
      const available = await client.waitForWord(10);
      if (!available) return; // Word not connected — nothing to stamp.
      // Clear the residue of the last suite so the marker is the only content.
      await client.call('word_clear');
      await client.call('word_insert_paragraph', { text: marker, location: 'End' });
    } catch {
      // Reporter side-effect only; never fail the run because the stamp couldn't be written.
    } finally {
      await client.disconnect();
    }
  }
}
