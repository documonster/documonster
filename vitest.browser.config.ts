import { defineConfig } from "vitest/config";
import { playwright } from "@vitest/browser-playwright";
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
      "src/modules/**/__tests__/browser/**/*.test.ts",
      "src/utils/__tests__/browser/**/*.test.ts"
    ]
  }
});
