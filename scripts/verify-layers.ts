/**
 * Module dependency-layer verification.
 *
 * Encodes the layer rules documented in AGENTS.md and fails (exit 1) if any
 * production source file imports across a forbidden module boundary. This is
 * the machine-enforced counterpart to the prose rules — oxlint cannot express
 * path-based import boundaries (its `no-restricted-imports` is a no-op), so we
 * scan imports ourselves.
 *
 * Scope: production `.ts` under `src/modules/<m>/` and `src/utils/`. Test
 * (`__tests__/`) and example (`examples/`) files are exempt — they may reach
 * across layers freely.
 *
 * Usage: node scripts/verify-layers.ts
 */

import fs from "node:fs";
import path from "node:path";

const ROOT = path.resolve(import.meta.dirname, "..");
const MODULES_DIR = path.join(ROOT, "src", "modules");
const UTILS_DIR = path.join(ROOT, "src", "utils");

const MODULES = [
  "excel",
  "word",
  "formula",
  "pdf",
  "csv",
  "markdown",
  "xml",
  "archive",
  "stream"
] as const;
type ModuleName = (typeof MODULES)[number] | "utils";

/**
 * For each module, the set of OTHER modules it is allowed to import from.
 * A module may always import from itself; that is handled separately. `utils`
 * may import nothing. The lists mirror the layer diagram in AGENTS.md.
 */
const ALLOWED: Record<ModuleName, ReadonlySet<ModuleName>> = {
  utils: new Set([]),
  xml: new Set(["utils"]),
  markdown: new Set(["utils"]),
  stream: new Set(["utils"]),
  csv: new Set(["stream", "utils"]),
  archive: new Set(["stream", "utils"]),
  formula: new Set(["utils"]),
  excel: new Set(["formula", "archive", "xml", "csv", "markdown", "stream", "utils"]),
  word: new Set(["formula", "archive", "xml", "csv", "markdown", "stream", "utils"]),
  // pdf may reach excel/word ONLY via the bridge files (see EXCEPTIONS); the
  // base allow-set covers the unconditional dependencies.
  pdf: new Set(["archive", "utils"])
};

/**
 * Per-file exceptions: a bridge file may import from a module that its host
 * module's base allow-set forbids. Keyed by repo-relative POSIX path.
 */
const EXCEPTIONS: Record<string, ReadonlySet<ModuleName>> = {
  "src/modules/pdf/excel-bridge.ts": new Set(["excel"]),
  "src/modules/pdf/word-chart-bridge.ts": new Set(["excel", "word"]),
  "src/modules/pdf/word-bridge.ts": new Set(["word"]),
  "src/modules/pdf/word-layout-to-pdf.ts": new Set(["word"]),
  "src/modules/word/bridge/excel-bridge.ts": new Set(["excel"])
};

const ALIAS_RE = /@(excel|word|formula|pdf|csv|markdown|xml|archive|stream|utils)\b/;

/** Recursively collect production `.ts` files (skip tests & examples). */
function collect(dir: string, out: string[]): void {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (entry.name === "__tests__" || entry.name === "examples") {
        continue;
      }
      collect(full, out);
    } else if (entry.isFile() && entry.name.endsWith(".ts") && !entry.name.endsWith(".d.ts")) {
      out.push(full);
    }
  }
}

/**
 * Extract the module aliases (`@excel`, `@word`, …) referenced by `import` /
 * `export ... from` / dynamic `import()` statements in a source file. Returns
 * a map of target module → first line number where it appears.
 */
function importedModules(source: string): Map<ModuleName, number> {
  const found = new Map<ModuleName, number>();
  // Match the specifier of any static or dynamic import/export-from across the
  // WHOLE source (not line-by-line) so multi-line statements like
  //   import {\n  a,\n  b\n} from "@excel/...";
  // are matched on their closing `from "..."`. The `[^"']` runs may span
  // newlines, which is what lets the regex cross lines. Module specifiers are
  // string literals (`"`/`'`) — never template literals — so backticks are
  // deliberately excluded to avoid matching `@alias` mentions inside doc
  // comments such as "must NOT import from `@excel/...`".
  const stmtRe =
    /(?:^|[\s;])(?:import|export)[^"'`]*?\bfrom\s*["']([^"']+)["']|\bimport\s*\(\s*["']([^"']+)["']\s*\)/g;
  let m: RegExpExecArray | null;
  while ((m = stmtRe.exec(source)) !== null) {
    const spec = m[1] ?? m[2];
    if (!spec) {
      continue;
    }
    const alias = spec.match(ALIAS_RE);
    if (alias) {
      const mod = alias[1] as ModuleName;
      if (!found.has(mod)) {
        // Line number of the matched specifier (count newlines up to the match).
        const idx = m.index + m[0].length;
        const line = source.slice(0, idx).split("\n").length;
        found.set(mod, line);
      }
    }
  }
  return found;
}

interface Violation {
  file: string;
  line: number;
  target: ModuleName;
  reason: string;
}

function checkFile(absPath: string, owner: ModuleName, violations: Violation[]): void {
  const rel = path.relative(ROOT, absPath).split(path.sep).join("/");
  const source = fs.readFileSync(absPath, "utf-8");
  const allowed = ALLOWED[owner];
  const extra = EXCEPTIONS[rel];

  for (const [target, line] of importedModules(source)) {
    if (target === owner) {
      continue; // same-module imports are always fine
    }
    if (allowed.has(target)) {
      continue;
    }
    if (extra?.has(target)) {
      continue;
    }
    violations.push({
      file: rel,
      line,
      target,
      reason:
        owner === "utils"
          ? `utils must not import from any module (found @${target})`
          : `${owner} may not import from @${target} (layer/boundary rule)`
    });
  }
}

function main(): void {
  const violations: Violation[] = [];

  for (const mod of MODULES) {
    const dir = path.join(MODULES_DIR, mod);
    if (!fs.existsSync(dir)) {
      continue;
    }
    const files: string[] = [];
    collect(dir, files);
    for (const f of files) {
      checkFile(f, mod, violations);
    }
  }

  // utils (Layer 0): may import nothing from modules. Its internal files use
  // relative paths, so alias references here can only be cross-module.
  const utilsFiles: string[] = [];
  collect(UTILS_DIR, utilsFiles);
  for (const f of utilsFiles) {
    checkFile(f, "utils", violations);
  }

  if (violations.length === 0) {
    console.log(`✓ Module layer check passed — no forbidden cross-module imports.`);
    return;
  }

  console.error(`✗ Module layer check failed — ${violations.length} violation(s):\n`);
  for (const v of violations) {
    console.error(`  ${v.file}:${v.line}\n    → ${v.reason}`);
  }
  console.error(
    `\nSee the "Module Dependency Layers" section in AGENTS.md. If a new bridge` +
      ` file legitimately needs a cross-module import, register it in the` +
      ` EXCEPTIONS map of scripts/verify-layers.ts (and document it in AGENTS.md).`
  );
  process.exit(1);
}

main();
