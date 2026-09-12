import { defineConfig } from "@playwright/test";
export default defineConfig({
  testDir: "./e2e",
  outputDir: "/tmp/calendar-poc-playwright-results",
  use: { baseURL: "http://localhost:5173", headless: true, channel: "chrome" },
  webServer: {
    command: "npm run dev",
    url: "http://localhost:5173/calendar/",
    reuseExistingServer: true,
    timeout: 30000,
  },
});
