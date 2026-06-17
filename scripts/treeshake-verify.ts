/**
 * Tree-shaking verification script.
 *
 * Verifies tree-shaking correctness across three bundlers (esbuild, rolldown,
 * rspack) and two platforms (node + browser) by:
 * 1. Creating a temporary consumer project that depends on excelts via symlink
 * 2. Bundling minimal import scenarios
 * 3. Inspecting output to verify unused modules are eliminated
 *
 * Usage: npx tsx scripts/treeshake-verify.ts
 */

import fs from "node:fs";
import path from "node:path";

import { rspack } from "@rspack/core";
import { buildSync } from "esbuild";
import { build as rolldownBuild } from "rolldown";

// =============================================================================
// Configuration
// =============================================================================

const ROOT = path.resolve(import.meta.dirname, "..");
const PKG = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf-8"));
const PKG_NAME: string = PKG.name;
const TMP_DIR = path.join(ROOT, "tmp");

// Common mustNotInclude patterns
const ALL_MODULES = [
  "modules/excel/",
  "modules/formula/",
  "modules/pdf/",
  "modules/csv/",
  "modules/archive/",
  "modules/stream/",
  "modules/xml/",
  "modules/markdown/"
];
const NOT_EXCEL_PDF_CSV = ["modules/excel/", "modules/formula/", "modules/pdf/", "modules/csv/"];

/** Exclude all modules except the listed ones */
function allModulesExcept(...keep: string[]): string[] {
  return ALL_MODULES.filter(m => !keep.some(k => m.includes(k)));
}

// =============================================================================
// Scenarios
// =============================================================================

interface Scenario {
  name: string;
  importFrom: string;
  imports: string[];
  mustNotInclude: string[];
  platform?: "browser" | "node";
  /** Bundlers to skip for this scenario (known tool limitations). */
  excludeBundlers?: string[];
  /**
   * When true, bundle with code-splitting and assert only against the ENTRY
   * chunk. Use for namespaces with lazy `import()` boundaries (e.g. `Pdf`'s
   * cross-module converters): the heavy dependency must live in a separate
   * on-demand chunk, not the entry a consumer pays for upfront.
   */
  lazySplit?: boolean;
}

/** Shorthand for creating a scenario */
function s(
  name: string,
  importFrom: string,
  imports: string[],
  mustNotInclude: string[],
  platform?: "browser" | "node",
  excludeBundlers?: string[]
): Scenario {
  return { name, importFrom, imports, mustNotInclude, platform, excludeBundlers };
}

// =============================================================================
// Full-coverage namespace matrix
//
// Every public domain namespace on every subpath is verified individually,
// using code-splitting + ENTRY-CHUNK isolation (lazySplit) so the assertion
// reflects what a consumer eagerly pays for under the real target bundlers
// (rolldown / rspack). For each namespace we declare the set of LOWER-LAYER
// modules it is legitimately allowed to reach (per AGENTS.md layer rules);
// every other module tree must be absent from the entry chunk.
//
// `allowed` lists module-path fragments that may legitimately appear. The
// resulting `mustNotInclude` is "all module trees except self + allowed".
// =============================================================================

const ALL_MODULE_TREES = [
  "modules/excel/",
  "modules/word/",
  "modules/pdf/",
  "modules/formula/",
  "modules/csv/",
  "modules/markdown/",
  "modules/xml/",
  "modules/archive/",
  "modules/stream/"
];

/** Build a mustNotInclude list = every module tree except `self` and `allowed`. */
function exclude(self: string, allowed: string[]): string[] {
  const keep = new Set([`modules/${self}/`, ...allowed]);
  return ALL_MODULE_TREES.filter(m => !keep.has(m));
}

/**
 * Generate a full-coverage, entry-chunk-isolated scenario for one namespace.
 * `allowed` = lower-layer module trees this namespace may legitimately pull.
 */
function ns(
  mod: string,
  nsName: string,
  allowed: string[],
  platform?: "browser" | "node",
  excludeBundlers?: string[]
): Scenario {
  const tag = platform === "browser" ? "browser " : "";
  const allowedNote = allowed.length
    ? `allows ${allowed.map(a => a.replace("modules/", "").replace(/\/$/, "")).join("+")}`
    : "isolated";
  return {
    name: `${tag}/${mod}: ${nsName} (${allowedNote})`,
    importFrom: `${PKG_NAME}/${mod}`,
    imports: [nsName],
    mustNotInclude: exclude(mod, allowed),
    platform,
    lazySplit: true,
    excludeBundlers
  };
}

const scenarios: Scenario[] = [
  // ===========================================================================
  // /excel subpath — ALL 20 namespaces. Per the layer rules, excel may reach
  // formula / archive / xml / csv / markdown / stream, but NEVER pdf or word.
  // Measured legitimate lower-layer reach is encoded per-namespace below.
  // ===========================================================================
  ns("excel", "Address", []),
  ns("excel", "Anchor", []),
  ns("excel", "Cell", []),
  ns("excel", "Chart", ["modules/xml/"]), // xml/encode for chart XML
  ns("excel", "Chartsheet", []),
  ns("excel", "Column", []),
  ns("excel", "DataValidation", []),
  ns("excel", "DefinedNames", ["modules/formula/"]), // named ranges may hold formulas
  ns("excel", "Form", []),
  ns("excel", "Image", []),
  ns("excel", "Note", []),
  ns("excel", "Pivot", []),
  ns("excel", "Range", []),
  ns("excel", "Row", []),
  ns("excel", "Sparkline", []),
  // xlsx streaming = zip + xml writer/sax + stream primitives
  ns("excel", "Stream", ["modules/archive/", "modules/xml/", "modules/stream/"]),
  ns("excel", "Table", []),
  ns("excel", "Watermark", []),
  // save xlsx = zip + xml + stream; defined-names pull the formula syntax probe
  ns("excel", "Workbook", [
    "modules/archive/",
    "modules/xml/",
    "modules/stream/",
    "modules/formula/"
  ]),
  ns("excel", "Worksheet", []),

  // ===========================================================================
  // /word subpath — ALL 19 namespaces. word may reach formula / archive / xml /
  // csv / markdown / stream, but NEVER pdf or excel (excel-bridge is lazy).
  // ===========================================================================
  ns("word", "Build", ["modules/xml/"]), // xml/encode for content XML
  ns("word", "Convert", ["modules/archive/", "modules/xml/", "modules/stream/"]), // ODT/docx IO
  ns("word", "Diff", []),
  ns("word", "Document", []),
  ns("word", "Font", []),
  ns("word", "Glossary", []),
  ns("word", "Io", ["modules/archive/", "modules/xml/", "modules/stream/"]), // docx read/write
  ns("word", "Layout", ["modules/xml/"]), // xml/encode
  ns("word", "Ole", ["modules/xml/"]), // xml/encode
  ns("word", "Query", ["modules/xml/"]), // parses docx content (dom/sax)
  ns("word", "RenderContext", []),
  ns("word", "Security", []),
  ns("word", "Streaming", ["modules/archive/", "modules/xml/", "modules/stream/"]),
  ns("word", "Styles", []),
  ns("word", "Template", ["modules/xml/"]), // parses template content (dom/sax)
  ns("word", "Theme", []),
  ns("word", "Units", []),
  ns("word", "Validation", []),
  ns("word", "Vba", []),

  // ===========================================================================
  // Small modules — single namespace each.
  // ===========================================================================
  ns("csv", "Csv", ["modules/stream/"]), // streaming CSV
  ns("xml", "Xml", []),
  ns("markdown", "Markdown", []),
  ns("formula", "Formula", []),
  ns("pdf", "Pdf", ["modules/archive/", "modules/xml/"]), // zlib + PDF metadata XML

  // ===========================================================================
  // Browser platform — ALL namespaces re-verified on the browser entries.
  // ===========================================================================
  ns("excel", "Address", [], "browser"),
  ns("excel", "Anchor", [], "browser"),
  ns("excel", "Cell", [], "browser"),
  ns("excel", "Chart", ["modules/xml/"], "browser"),
  ns("excel", "Chartsheet", [], "browser"),
  ns("excel", "Column", [], "browser"),
  ns("excel", "DataValidation", [], "browser"),
  ns("excel", "DefinedNames", ["modules/formula/"], "browser"),
  ns("excel", "Form", [], "browser"),
  ns("excel", "Image", [], "browser"),
  ns("excel", "Note", [], "browser"),
  ns("excel", "Pivot", [], "browser"),
  ns("excel", "Range", [], "browser"),
  ns("excel", "Row", [], "browser"),
  ns("excel", "Sparkline", [], "browser"),
  ns("excel", "Stream", ["modules/archive/", "modules/xml/", "modules/stream/"], "browser"),
  ns("excel", "Table", [], "browser"),
  ns("excel", "Watermark", [], "browser"),
  ns(
    "excel",
    "Workbook",
    ["modules/archive/", "modules/xml/", "modules/stream/", "modules/formula/"],
    "browser"
  ),
  ns("excel", "Worksheet", [], "browser"),

  ns("word", "Build", ["modules/xml/"], "browser"),
  ns("word", "Convert", ["modules/archive/", "modules/xml/", "modules/stream/"], "browser"),
  ns("word", "Diff", [], "browser"),
  ns("word", "Document", [], "browser"),
  ns("word", "Font", [], "browser"),
  ns("word", "Glossary", [], "browser"),
  ns("word", "Io", ["modules/archive/", "modules/xml/", "modules/stream/"], "browser"),
  ns("word", "Layout", ["modules/xml/"], "browser"),
  ns("word", "Ole", ["modules/xml/"], "browser"),
  ns("word", "Query", ["modules/xml/"], "browser"),
  ns("word", "RenderContext", [], "browser"),
  ns("word", "Security", [], "browser"),
  ns("word", "Streaming", ["modules/archive/", "modules/xml/", "modules/stream/"], "browser"),
  ns("word", "Styles", [], "browser"),
  ns("word", "Template", ["modules/xml/"], "browser"),
  ns("word", "Theme", [], "browser"),
  ns("word", "Units", [], "browser"),
  ns("word", "Validation", [], "browser"),
  ns("word", "Vba", [], "browser"),

  ns("csv", "Csv", ["modules/stream/"], "browser"),
  ns("xml", "Xml", [], "browser"),
  ns("markdown", "Markdown", [], "browser"),
  ns("formula", "Formula", [], "browser"),
  ns("pdf", "Pdf", ["modules/archive/", "modules/xml/"], "browser"),

  // ===========================================================================
  // Infrastructure modules — intentionally flat exports (not namespaced).
  // ===========================================================================
  s(
    "/archive: crc32 (minimal)",
    `${PKG_NAME}/archive`,
    ["crc32"],
    [...NOT_EXCEL_PDF_CSV, "modules/archive/zip/", "modules/archive/unzip/", "modules/archive/tar/"]
  ),
  s("/stream: pipeline", `${PKG_NAME}/stream`, ["pipeline"], allModulesExcept("stream"))
];

// =============================================================================
// Setup
// =============================================================================

function setupTmpProject(): void {
  if (fs.existsSync(TMP_DIR)) {
    fs.rmSync(TMP_DIR, { recursive: true });
  }
  fs.mkdirSync(TMP_DIR, { recursive: true });

  fs.writeFileSync(
    path.join(TMP_DIR, "package.json"),
    JSON.stringify({ name: "treeshake-test", type: "module", private: true }, null, 2)
  );

  // Symlink excelts into node_modules (handles scoped packages)
  const nmDir = path.join(TMP_DIR, "node_modules");
  const scope = PKG_NAME.startsWith("@") ? PKG_NAME.split("/")[0] : null;
  const parentDir = scope ? path.join(nmDir, scope) : nmDir;
  fs.mkdirSync(parentDir, { recursive: true });
  fs.symlinkSync(ROOT, path.join(nmDir, PKG_NAME), "dir");
}

// =============================================================================
// Shared types & helpers
// =============================================================================

interface ModuleEntry {
  path: string;
  bytes: number;
}

interface ScenarioResult {
  name: string;
  bundler: string;
  bundleSize: number;
  contributingModules: ModuleEntry[];
  parsedModuleCount: number;
  violations: { pattern: string; matchedModules: ModuleEntry[] }[];
  passed: boolean;
}

function makeEntryCode(scenario: Scenario): string {
  const names = scenario.imports.join(", ");
  return `import { ${names} } from "${scenario.importFrom}";\nconsole.log(${names});`;
}

function normalizePath(filePath: string): string {
  const idx = filePath.indexOf("dist/");
  return idx >= 0 ? filePath.substring(idx) : filePath;
}

function checkViolations(
  contributing: ModuleEntry[],
  mustNotInclude: string[]
): ScenarioResult["violations"] {
  const violations: ScenarioResult["violations"] = [];
  for (const pattern of mustNotInclude) {
    const matched = contributing.filter(m => m.path.includes(pattern));
    if (matched.length > 0) {
      violations.push({ pattern, matchedModules: matched });
    }
  }
  return violations;
}

function makeResult(
  scenario: Scenario,
  bundler: string,
  bundleSize: number,
  contributing: ModuleEntry[],
  parsedCount: number
): ScenarioResult {
  const violations = checkViolations(contributing, scenario.mustNotInclude);
  return {
    name: scenario.name,
    bundler,
    bundleSize,
    contributingModules: contributing.sort((a, b) => b.bytes - a.bytes),
    parsedModuleCount: parsedCount,
    violations,
    passed: violations.length === 0
  };
}

function makeError(name: string, bundler: string, message: string): ScenarioResult {
  return {
    name,
    bundler,
    bundleSize: 0,
    contributingModules: [],
    parsedModuleCount: 0,
    violations: [{ pattern: "BUILD_FAILED", matchedModules: [{ path: message, bytes: 0 }] }],
    passed: false
  };
}

function writeEntry(bundler: string, scenario: Scenario): { entryFile: string; slug: string } {
  const slug = scenario.name.replace(/\W+/g, "-");
  const entryFile = path.join(TMP_DIR, `${bundler}-${slug}.mjs`);
  fs.writeFileSync(entryFile, makeEntryCode(scenario), "utf-8");
  return { entryFile, slug };
}

// =============================================================================
// Bundler runners
// =============================================================================

function runEsbuild(scenario: Scenario): ScenarioResult {
  const { entryFile, slug } = writeEntry("esbuild", scenario);

  try {
    if (scenario.lazySplit) {
      // Code-split build: the entry chunk must be free of the excluded
      // modules; lazy `import()` targets land in separate on-demand chunks.
      const outDir = path.join(TMP_DIR, `esbuild-${slug}-split`);
      const result = buildSync({
        entryPoints: [entryFile],
        bundle: true,
        format: "esm",
        platform: scenario.platform === "browser" ? "browser" : "node",
        outdir: outDir,
        metafile: true,
        treeShaking: true,
        splitting: true,
        minify: false,
        write: true,
        external: ["node:*"]
      });
      const meta = result.metafile!;
      const entryOut = Object.entries(meta.outputs).find(([, o]) =>
        (o.entryPoint ?? "").includes(path.basename(entryFile))
      );
      if (!entryOut) {
        return makeError(scenario.name, "esbuild", "entry chunk not found");
      }
      const [entryKey, entryMeta] = entryOut;
      const contributing: ModuleEntry[] = [];
      for (const [fp, info] of Object.entries(entryMeta.inputs)) {
        if (fp.includes("dist/") && info.bytesInOutput > 0) {
          contributing.push({ path: normalizePath(fp), bytes: info.bytesInOutput });
        }
      }
      const parsedCount = Object.keys(entryMeta.inputs).filter(m => m.includes("dist/")).length;
      return makeResult(scenario, "esbuild", fs.statSync(entryKey).size, contributing, parsedCount);
    }

    const outFile = path.join(TMP_DIR, `esbuild-${slug}.out.mjs`);
    const result = buildSync({
      entryPoints: [entryFile],
      bundle: true,
      format: "esm",
      platform: scenario.platform === "browser" ? "browser" : "node",
      outfile: outFile,
      metafile: true,
      treeShaking: true,
      minify: false,
      write: true,
      external: ["node:*"]
    });

    const meta = result.metafile!;
    const outKey = Object.keys(meta.outputs).find(k => k.endsWith(".mjs"))!;
    const outMeta = meta.outputs[outKey];

    const contributing: ModuleEntry[] = [];
    for (const [fp, info] of Object.entries(outMeta.inputs)) {
      if (fp.includes("dist/") && info.bytesInOutput > 0) {
        contributing.push({ path: normalizePath(fp), bytes: info.bytesInOutput });
      }
    }

    const parsedCount = Object.keys(meta.inputs).filter(m => m.includes("dist/")).length;
    return makeResult(scenario, "esbuild", fs.statSync(outFile).size, contributing, parsedCount);
  } catch (err: any) {
    return makeError(scenario.name, "esbuild", err.message);
  }
}

async function runRolldown(scenario: Scenario): Promise<ScenarioResult> {
  const { entryFile, slug } = writeEntry("rolldown", scenario);
  const outDir = path.join(TMP_DIR, `rolldown-${slug}-out`);

  try {
    const output = await rolldownBuild({
      input: entryFile,
      platform: scenario.platform === "browser" ? "browser" : "node",
      resolve: {
        conditionNames:
          scenario.platform === "browser" ? ["browser", "import", "default"] : ["import", "default"]
      },
      treeshake: true,
      external: [/^node:/],
      output: { dir: outDir, format: "esm", entryFileNames: "out.mjs" }
    });

    const outFile = path.join(outDir, "out.mjs");
    const bundleSize = fs.existsSync(outFile) ? fs.statSync(outFile).size : 0;
    const contributing: ModuleEntry[] = [];
    let parsedCount = 0;

    // Only inspect the ENTRY chunk. Lazy `import()` targets are emitted as
    // separate on-demand chunks (rolldown splits them by default), so a
    // dynamic cross-module boundary must NOT appear in the entry a consumer
    // pays for upfront. Counting every chunk would defeat that contract.
    for (const chunk of output.output ?? []) {
      if (chunk.type !== "chunk" || !chunk.isEntry || !chunk.modules) {
        continue;
      }
      for (const [modId, modInfo] of Object.entries(chunk.modules)) {
        if (!modId.includes("dist/")) {
          continue;
        }
        parsedCount++;
        const info = modInfo as { renderedLength?: number; code?: string };
        const len = info.renderedLength ?? info.code?.length ?? 0;
        if (len > 0) {
          contributing.push({ path: normalizePath(modId), bytes: len });
        }
      }
    }

    return makeResult(scenario, "rolldown", bundleSize, contributing, parsedCount);
  } catch (err: any) {
    return makeError(scenario.name, "rolldown", err.message);
  }
}

function runRspack(scenario: Scenario): Promise<ScenarioResult> {
  const { entryFile, slug } = writeEntry("rspack", scenario);
  const outDir = path.join(TMP_DIR, `rspack-${slug}-out`);

  return new Promise(resolve => {
    const compiler = rspack({
      mode: "production",
      entry: entryFile,
      output: {
        path: outDir,
        filename: "out.mjs",
        chunkFilename: "[name].chunk.mjs",
        module: true,
        library: { type: "module" }
      },
      target: scenario.platform === "browser" ? "web" : "node",
      externals: [/^node:/],
      resolve: {
        conditionNames:
          scenario.platform === "browser"
            ? ["browser", "import", "default"]
            : ["import", "default"],
        symlinks: true
      },
      optimization: { usedExports: true, sideEffects: true, minimize: false, innerGraph: true },
      stats: {
        all: false,
        modules: true,
        chunks: true,
        chunkModules: true,
        ids: true,
        modulesSpace: Infinity
      }
    });

    compiler.run((err, stats) => {
      const close = () => new Promise<void>(r => compiler.close(() => r()));

      if (err || !stats) {
        close().then(() => resolve(makeError(scenario.name, "rspack", err?.message ?? "no stats")));
        return;
      }
      if (stats.hasErrors()) {
        const errors = stats.toJson({ errors: true }).errors ?? [];
        close().then(() =>
          resolve(makeError(scenario.name, "rspack", errors.map((e: any) => e.message).join("\n")))
        );
        return;
      }

      const outFile = path.join(outDir, "out.mjs");
      const bundleSize = fs.existsSync(outFile) ? fs.statSync(outFile).size : 0;
      const jsonStats = stats.toJson({ modules: true, chunks: true, reasons: false });
      const contributing: ModuleEntry[] = [];
      let parsedCount = 0;

      // Collect the ids of modules that belong to an INITIAL (entry) chunk.
      // Lazy `import()` targets land in non-initial chunks, so restricting to
      // initial chunks reflects what the consumer eagerly pays for.
      const initialModuleIds = new Set<string | number>();
      for (const ch of jsonStats.chunks ?? []) {
        if (!ch.initial) {
          continue;
        }
        for (const m of ch.modules ?? []) {
          if (m.id != null) {
            initialModuleIds.add(m.id);
          }
          const mn = m.name ?? m.identifier ?? "";
          if (mn) {
            initialModuleIds.add(mn);
          }
        }
      }

      for (const mod of jsonStats.modules ?? []) {
        const modName: string = mod.name ?? mod.identifier ?? "";
        if (!modName.includes("dist/")) {
          continue;
        }
        // Skip modules that only live in lazy (non-initial) chunks.
        const inInitial =
          (mod.id != null && initialModuleIds.has(mod.id)) || initialModuleIds.has(modName);
        if (!inInitial) {
          continue;
        }
        parsedCount++;
        const size = mod.size ?? 0;
        if (size > 0) {
          contributing.push({ path: normalizePath(modName), bytes: size });
        }
      }

      close().then(() =>
        resolve(makeResult(scenario, "rspack", bundleSize, contributing, parsedCount))
      );
    });
  });
}

// =============================================================================
// Report
// =============================================================================

function report(results: ScenarioResult[]): boolean {
  let allPassed = true;

  const byBundler = new Map<string, ScenarioResult[]>();
  for (const r of results) {
    const list = byBundler.get(r.bundler) ?? [];
    list.push(r);
    byBundler.set(r.bundler, list);
  }

  for (const [bundler, bundlerResults] of byBundler) {
    console.log("\n" + "=".repeat(80));
    console.log(`  TREE-SHAKING VERIFICATION — ${bundler.toUpperCase()}`);
    console.log("=".repeat(80) + "\n");

    for (const r of bundlerResults) {
      const status = r.passed ? "PASS" : "FAIL";
      console.log(`  [${status}] ${r.name}`);
      console.log(
        `         Bundle: ${(r.bundleSize / 1024).toFixed(1)} KB | ` +
          `Contributing: ${r.contributingModules.length} / ${r.parsedModuleCount} parsed`
      );

      if (r.violations.length > 0) {
        allPassed = false;
        for (const v of r.violations) {
          const totalBytes = v.matchedModules.reduce((s, m) => s + m.bytes, 0);
          console.log(
            `         VIOLATION: "${v.pattern}" — ${v.matchedModules.length} files, ` +
              `${(totalBytes / 1024).toFixed(1)} KB leaked:`
          );
          for (const m of v.matchedModules.slice(0, 5)) {
            console.log(`           - ${m.path} (${m.bytes} B)`);
          }
          if (v.matchedModules.length > 5) {
            console.log(`           ... and ${v.matchedModules.length - 5} more`);
          }
        }
      }
      console.log();
    }

    const passed = bundlerResults.filter(r => r.passed).length;
    console.log(`  ${bundler}: ${passed}/${bundlerResults.length} passed`);
  }

  console.log("\n" + "=".repeat(80));
  console.log(
    allPassed
      ? "  ALL SCENARIOS PASSED across all bundlers"
      : "  SOME SCENARIOS FAILED — tree-shaking issues detected"
  );
  console.log("=".repeat(80) + "\n");

  return allPassed;
}

// =============================================================================
// Main
// =============================================================================

async function main(): Promise<void> {
  setupTmpProject();

  const results: ScenarioResult[] = [];
  for (const s of scenarios) {
    if (!s.excludeBundlers?.includes("esbuild")) {
      results.push(runEsbuild(s));
    }
  }
  for (const s of scenarios) {
    if (!s.excludeBundlers?.includes("rolldown")) {
      results.push(await runRolldown(s));
    }
  }
  for (const s of scenarios) {
    if (!s.excludeBundlers?.includes("rspack")) {
      results.push(await runRspack(s));
    }
  }

  const allPassed = report(results);
  fs.rmSync(TMP_DIR, { recursive: true });

  if (!allPassed) {
    process.exit(1);
  }
}

main();
