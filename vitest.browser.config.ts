import { playwright } from "@vitest/browser-playwright";
import { defineConfig } from "vitest/config";

import { preferBrowserFilesPlugin } from "./src/utils/browser";

export default defineConfig({
  plugins: [preferBrowserFilesPlugin()],
  resolve: {
    tsconfigPaths: true
  },
  define: {
    global: "globalThis"
  },
  test: {
    globals: true,
    testTimeout: 30000,
    setupFiles: ["./src/test/browser/setup.ts"],
    browser: {
      enabled: true,
      provider: playwright(),
      headless: true,
      instances: [
        {
          browser: "chromium"
        }
      ]
    },
    include: [
      // Platform/runtime browser tests (IIFE smoke, archive/stream native APIs).
      "src/modules/**/__tests__/browser/**/*.test.ts",
      "src/utils/__tests__/browser/**/*.test.ts",
      // Pure-computation modules: their node test suites are free of Node-only
      // APIs and disk fixtures, so we re-run them in a real browser to prove
      // the shipped logic behaves identically there. Node-stream / fixture
      // dependent files are excluded below.
      "src/modules/xml/__tests__/**/*.test.ts",
      "src/modules/markdown/__tests__/**/*.test.ts",
      "src/modules/formula/**/__tests__/**/*.test.ts",
      "src/modules/word/**/__tests__/**/*.test.ts"
    ],
    exclude: [
      // Formula's Excel-oracle reference suite reads fixture files from disk
      // (Node `fs`), so it cannot run in the browser.
      "src/modules/formula/__tests__/reference/**"
    ]
  }
});
