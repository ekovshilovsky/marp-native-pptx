import { defineConfig } from 'vitest/config'

// Visual-regression config: ONLY the test/visual suite, with a long timeout
// (each case shells out to LibreOffice to render the deck to PNG).
export default defineConfig({
  test: {
    environment: 'node',
    include: ['test/visual/**/*.test.ts'],
    testTimeout: 180000,
    hookTimeout: 180000,
  },
})
