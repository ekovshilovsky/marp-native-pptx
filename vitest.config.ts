import { defineConfig, configDefaults } from 'vitest/config'

export default defineConfig({
  test: {
    environment: 'node',
    testTimeout: 60000, // browser-based layout tests need headroom
    // visual-regression tests render via LibreOffice and are opt-in (npm run
    // test:visual) so the fast suite stays light and dependency-free.
    exclude: [...configDefaults.exclude, 'test/visual/**'],
  },
})
