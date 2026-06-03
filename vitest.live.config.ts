import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/live/**/*.test.ts'],
    testTimeout: 60000,
    globals: true,
    sequence: { concurrent: false },
    fileParallelism: false,
  },
  resolve: {
    alias: {
      '@shared': '/src/shared',
    },
  },
});
