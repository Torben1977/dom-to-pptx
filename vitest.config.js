import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    environment: 'jsdom',
    globals: true,
    // Several suites launch a real browser. Under the full run they compete for
    // the machine and regularly needed more than the 5 s default, which showed
    // up as tests failing in a different combination on every run.
    testTimeout: 30_000,
  },
});
