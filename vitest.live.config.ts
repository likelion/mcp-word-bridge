import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/live/**/*.test.ts'],
    testTimeout: 60000,
    globals: true,
    sequence: { concurrent: false },
    fileParallelism: false,
    // 'default' keeps console output; the custom reporter stamps a completion
    // marker into the Word document after the whole suite finishes.
    reporters: ['default', './test/live/completion-reporter.ts'],
  },
});
