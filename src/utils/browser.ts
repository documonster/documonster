import fs from "node:fs";
import path from "node:path";

import type { Plugin } from "vitest/config";

/**
 * Prefer sibling `.browser.*` implementation when it exists.
 *
 * Example:
 * - resolved: `/abs/path/to/streams.ts`
 * - if `/abs/path/to/streams.browser.ts` exists => use it
 */
export function preferBrowserFilesPlugin(): Plugin {
  return {
    name: "excelts-prefer-browser-files",
    enforce: "pre",
    async resolveId(source, importer) {
      if (source.startsWith("\0") || source.startsWith("virtual:") || source.startsWith("node:")) {
        return null;
      }

      const resolved = await this.resolve(source, importer, {
        skipSelf: true
      });

      if (!resolved) {
        return null;
      }
      if (typeof resolved === "object" && resolved.external) {
        return resolved;
      }

      const resolvedId = typeof resolved === "string" ? resolved : resolved.id;
      if (typeof resolvedId !== "string") {
        return resolved;
      }

      const cleanId = resolvedId.split("?")[0].split("#")[0];
      if (!path.isAbsolute(cleanId)) {
        return resolved;
      }
      if (cleanId.includes(".browser.")) {
        return resolved;
      }

      const ext = path.extname(cleanId);
      if (!ext || cleanId.endsWith(".d.ts")) {
        return resolved;
      }

      const candidate = cleanId.slice(0, -ext.length) + ".browser" + ext;
      if (fs.existsSync(candidate)) {
        return candidate;
      }

      return resolved;
    }
  };
}
