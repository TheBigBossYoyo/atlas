import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests/e2e',
  timeout: 60_000,
  fullyParallel: false,
  // Atlas takes an Electron single-instance lock: a second app launched by a
  // parallel worker quits immediately, so e2e files must run one at a time.
  workers: 1,
  retries: 0,
  use: {
    headless: true,
  },
  webServer: {
    command: 'npm run dev',
    port: 5173,
    reuseExistingServer: true,
    timeout: 120_000,
  },
})
