import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    tsconfigPaths: true,
    extensions: [".ts", ".tsx", ".js", ".jsx", ".json"]
  },
  test: {
    globals: true,
    testTimeout: 30000,
    include: ["src/**/__tests__/**/*.test.ts"],
    exclude: ["src/**/__tests__/browser/**/*.test.ts"],
    setupFiles: ["src/test/setup.ts"],
    isolate: false
  }
});
