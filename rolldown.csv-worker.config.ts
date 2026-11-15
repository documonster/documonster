import { defineConfig } from "rolldown";

import { browserTargets } from "./scripts/lib/browser-targets";
import { preferBrowserFilesPlugin } from "./src/utils/browser";

export default defineConfig({
  input: "./src/modules/csv/worker/worker.entry.ts",
  platform: "browser",
  tsconfig: "./tsconfig.json",
  transform: { target: browserTargets() },
  plugins: [preferBrowserFilesPlugin()],
  output: {
    dir: "./tmp/csv-worker-bundle",
    format: "iife",
    sourcemap: false,
    exports: "none",
    entryFileNames: "csv-worker.iife.js"
  }
});
