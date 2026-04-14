import fs from "node:fs";

import { defineConfig } from "rolldown";
import { visualizer } from "rollup-plugin-visualizer";

import { preferBrowserFilesPlugin } from "./src/utils/browser";

const pkg = JSON.parse(fs.readFileSync("./package.json", "utf-8"));
const banner = `/*!
 * ${pkg.name} v${pkg.version}
 * ${pkg.description}
 * (c) ${new Date().getFullYear()} ${pkg.author.name}
 * Released under the ${pkg.license} License
 */`;

// Common config shared by both builds
// Browser version now has NO Node.js polyfills - pure browser code

const commonConfig = {
  input: "./src/index.browser.ts",
  platform: "browser" as const,
  tsconfig: "./tsconfig.json",
  plugins: [preferBrowserFilesPlugin()]
};

const copyLicensePlugin = {
  name: "copy-license",
  writeBundle() {
    if (!fs.existsSync("./dist/iife")) {
      fs.mkdirSync("./dist/iife", { recursive: true });
    }
    fs.copyFileSync("./LICENSE", "./dist/iife/LICENSE");
    fs.copyFileSync("./THIRD_PARTY_NOTICES.md", "./dist/iife/THIRD_PARTY_NOTICES.md");
  }
};

const analyzePlugins =
  process.env.ANALYZE === "true"
    ? [
        visualizer({
          filename: "./dist/stats.html",
          open: false,
          gzipSize: true,
          brotliSize: true,
          template: "treemap"
        })
      ]
    : [];

export default defineConfig([
  {
    ...commonConfig,
    output: {
      dir: "./dist/iife",
      format: "iife",
      name: "ExcelTS",
      sourcemap: true,
      banner,
      exports: "named",
      entryFileNames: "excelts.iife.js"
    },
    plugins: [...commonConfig.plugins, copyLicensePlugin, ...analyzePlugins]
  },
  {
    ...commonConfig,
    output: {
      dir: "./dist/iife",
      format: "iife",
      name: "ExcelTS",
      sourcemap: false,
      banner,
      exports: "named",
      minify: true,
      entryFileNames: "excelts.iife.min.js"
    },
    plugins: [...commonConfig.plugins, copyLicensePlugin]
  }
]);
