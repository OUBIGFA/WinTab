import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests',
  fullyParallel: true,
  workers: 2,
  retries: 0,
  reporter: 'list',
  use: { baseURL: 'http://127.0.0.1:4173', viewport: { width: 1224, height: 1100 } },
  webServer: {
    command: 'npm exec vite -- preview --host 127.0.0.1 --port 4173 --strictPort',
    url: 'http://127.0.0.1:4173',
    reuseExistingServer: false,
  },
})
