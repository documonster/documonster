import { defineConfig } from "rolldown";

import { preferBrowserFilesPlugin } from "./src/utils/browser";

export default defineConfig({
  input: "./src/modules/csv/worker/worker.entry.ts",
  platform: "browser",
  tsconfig: "./tsconfig.json",
  plugins: [preferBrowserFilesPlugin()],
  output: {
    dir: "./tmp/csv-worker-bundle",
    format: "iife",
    sourcemap: false,
    exports: "none",
    entryFileNames: "csv-worker.iife.js"
  }
});
