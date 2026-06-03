import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/unit/**/*.test.ts'],
    globals: true,
    coverage: {
      provider: 'v8',
      include: ['src/**/*.ts'],
      exclude: ['src/taskpane/**'],
      reporter: ['text', 'json', 'lcov'],
    },
  },
  resolve: {
    alias: {
      '@shared': '/src/shared',
    },
  },
});
