import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

/**
 * Node/browser `Workbook` surface parity.
 *
 * `documonster/excel` ships two entry points. The `Workbook` namespace is
 * assembled from a platform-specific surface module:
 *   - Node:    `surface/workbook.ts`
 *   - browser: `surface/workbook.browser.ts`
 *
 * Regression (#185 follow-up): the named-cell-style functions were added to the
 * Node surface only, so `Workbook.defineCellStyle` was `undefined` at runtime in
 * browser/bundler builds even though the types resolved. This test locks the two
 * surfaces together so they can never silently drift again.
 *
 * It is a `.node.test.ts` because it must read BOTH source files directly — the
 * browser test runner rewrites `surface/workbook` → `surface/workbook.browser`,
 * which would make the two imports collapse onto the same module.
 *
 * We parse the re-exported *names* statically (rather than importing at runtime)
 * so the check reflects the source contract regardless of module resolution.
 */

const NODE_SURFACE = fileURLToPath(new URL("../workbook.ts", import.meta.url));
const BROWSER_SURFACE = fileURLToPath(new URL("../workbook.browser.ts", import.meta.url));

/**
 * Extract the value (non-type) names a surface module re-exports, using each
 * member's exported alias (`foo as bar` → `bar`). Type-only members
 * (`type X`, `export type { … }`) are ignored — they don't exist at runtime.
 */
function exportedValueNames(filePath: string): Set<string> {
  const src = readFileSync(filePath, "utf8");
  const names = new Set<string>();

  // Match every `export { ... } from "..."` block (multi-line).
  const blockRe = /export\s*\{([^}]*)\}\s*from\s*["'][^"']+["']/g;
  let block: RegExpExecArray | null;
  while ((block = blockRe.exec(src)) !== null) {
    for (const raw of block[1].split(",")) {
      const member = raw.trim();
      if (!member) {
        continue;
      }
      // Skip type-only members: `type Foo` or `type Foo as Bar`.
      if (/^type\s/.test(member)) {
        continue;
      }
      // `foo as bar` → exported name is `bar`; otherwise the name itself.
      const asMatch = member.match(/\bas\s+([A-Za-z_$][\w$]*)\s*$/);
      names.add(asMatch ? asMatch[1] : member);
    }
  }
  return names;
}

describe("Workbook surface node/browser parity", () => {
  const nodeNames = exportedValueNames(NODE_SURFACE);
  const browserNames = exportedValueNames(BROWSER_SURFACE);

  it("exposes the full named-cell-style API on both entries", () => {
    for (const fn of [
      "defineCellStyle",
      "getCellStyle",
      "listCellStyles",
      "removeCellStyle",
      "useBuiltinCellStyle"
    ]) {
      expect(nodeNames.has(fn), `Node Workbook surface missing "${fn}"`).toBe(true);
      expect(browserNames.has(fn), `browser Workbook surface missing "${fn}"`).toBe(true);
    }
  });

  it("differs only by the documented platform-specific IO members", () => {
    // The only legitimate divergence: Node adds file-path IO (`readFile` /
    // `writeFile`). Everything else — including the streaming helpers and
    // cross-platform `toBuffer`/`read` — is shared by both surfaces.
    const NODE_ONLY = new Set(["readFile", "writeFile"]);

    const nodeExtra = [...nodeNames].filter(n => !browserNames.has(n) && !NODE_ONLY.has(n)).sort();
    const browserExtra = [...browserNames].filter(n => !nodeNames.has(n)).sort();

    expect(nodeExtra, "Node-only members not in the documented allow-list").toEqual([]);
    expect(browserExtra, "browser surface must not add members Node lacks").toEqual([]);
  });
});
