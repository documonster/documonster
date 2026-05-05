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
}

/** Shorthand for creating a scenario */
function s(
  name: string,
  importFrom: string,
  imports: string[],
  mustNotInclude: string[],
  platform?: "browser" | "node"
): Scenario {
  return { name, importFrom, imports, mustNotInclude, platform };
}

const scenarios: Scenario[] = [
  // Root entry (Node)
  s("root: dateToExcel", PKG_NAME, ["dateToExcel"], ALL_MODULES),
  s("root: BaseError", PKG_NAME, ["BaseError"], ALL_MODULES),
  s(
    "root: encodeCell",
    PKG_NAME,
    ["encodeCell"],
    [...allModulesExcept("excel"), "modules/excel/workbook", "modules/excel/worksheet"]
  ),
  s(
    "root: Workbook (no pdf/formula leak)",
    PKG_NAME,
    ["Workbook"],
    // A `Workbook` import *must* pull the two glue files from the
    // formula module — they declare the registry that
    // `Workbook.calculateFormulas()` and `defined-names` dispatch
    // through. That is ~3.9 KB total, orders of magnitude smaller than
    // the engine itself (~200 KB), which stays out of the bundle unless
    // `installFormulaEngine` is imported. Assert on the engine source
    // tree, not the glue.
    [
      "modules/pdf/",
      "modules/formula/syntax/",
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/compile/",
      "modules/formula/materialize/",
      "modules/formula/integration/"
    ]
  ),
  s(
    "root: excelToPdf (no csv/formula leak)",
    PKG_NAME,
    ["excelToPdf"],
    [
      "modules/csv/",
      "modules/formula/syntax/",
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/compile/",
      "modules/formula/materialize/",
      "modules/formula/integration/"
    ]
  ),
  s(
    "root: CsvParserStream (no pdf/archive leak)",
    PKG_NAME,
    ["CsvParserStream"],
    ["modules/pdf/", "modules/archive/"]
  ),

  // /archive subpath (Node)
  s("/archive: ZipArchive", `${PKG_NAME}/archive`, ["ZipArchive"], NOT_EXCEL_PDF_CSV),
  s("/archive: ZipReader", `${PKG_NAME}/archive`, ["ZipReader"], NOT_EXCEL_PDF_CSV),
  s(
    "/archive: crc32 (minimal)",
    `${PKG_NAME}/archive`,
    ["crc32"],
    [...NOT_EXCEL_PDF_CSV, "modules/archive/zip/", "modules/archive/unzip/", "modules/archive/tar/"]
  ),
  s(
    "/archive: compress (minimal)",
    `${PKG_NAME}/archive`,
    ["compress"],
    [...NOT_EXCEL_PDF_CSV, "modules/archive/zip/", "modules/archive/unzip/", "modules/archive/tar/"]
  ),
  s("/archive: TarArchive", `${PKG_NAME}/archive`, ["TarArchive"], NOT_EXCEL_PDF_CSV),

  // /csv subpath (Node)
  s("/csv: formatCsv", `${PKG_NAME}/csv`, ["formatCsv"], allModulesExcept("csv")),
  s("/csv: parseCsv", `${PKG_NAME}/csv`, ["parseCsv"], allModulesExcept("csv")),
  s(
    "/csv: CsvParserStream",
    `${PKG_NAME}/csv`,
    ["CsvParserStream"],
    ["modules/excel/", "modules/pdf/", "modules/archive/"]
  ),
  s(
    "/csv: detectDelimiter (minimal)",
    `${PKG_NAME}/csv`,
    ["detectDelimiter"],
    [...allModulesExcept("csv"), "modules/csv/parse/", "modules/csv/format/"]
  ),

  // /stream subpath (Node)
  s("/stream: pipeline", `${PKG_NAME}/stream`, ["pipeline"], allModulesExcept("stream")),
  s(
    "/stream: createTransform",
    `${PKG_NAME}/stream`,
    ["createTransform"],
    allModulesExcept("stream")
  ),
  s(
    "/stream: ChunkedBuilder",
    `${PKG_NAME}/stream`,
    ["ChunkedBuilder"],
    allModulesExcept("stream")
  ),
  s("/stream: collect (minimal)", `${PKG_NAME}/stream`, ["collect"], allModulesExcept("stream")),

  // /xml subpath (Node)
  s("/xml: xmlEncode", `${PKG_NAME}/xml`, ["xmlEncode"], allModulesExcept("xml")),
  s("/xml: XmlWriter", `${PKG_NAME}/xml`, ["XmlWriter"], allModulesExcept("xml")),
  s("/xml: parseXml", `${PKG_NAME}/xml`, ["parseXml"], allModulesExcept("xml")),
  s("/xml: SaxParser", `${PKG_NAME}/xml`, ["SaxParser"], allModulesExcept("xml")),

  // /pdf subpath (Node)
  // PDF legitimately depends on archive/compression for zlib/deflate.
  // `excelToPdf` legitimately references the formula registry glue
  // (host-registry.ts) so that an app that also imports
  // `@cj-tech-master/excelts/formula` gets automatic recalc before
  // render — see the root Workbook scenario for the rationale.
  s("/pdf: pdf", `${PKG_NAME}/pdf`, ["pdf"], allModulesExcept("pdf", "archive")),
  s(
    "/pdf: excelToPdf",
    `${PKG_NAME}/pdf`,
    ["excelToPdf"],
    [
      ...allModulesExcept("pdf", "excel", "archive", "formula"),
      "modules/formula/syntax/",
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/compile/",
      "modules/formula/materialize/",
      "modules/formula/integration/"
    ]
  ),
  s("/pdf: readPdf", `${PKG_NAME}/pdf`, ["readPdf"], allModulesExcept("pdf", "archive")),
  s("/pdf: PageSizes", `${PKG_NAME}/pdf`, ["PageSizes"], allModulesExcept("pdf")),

  // /formula subpath (Node)
  //
  // The functional form must be fully tree-shakeable: importing just
  // `tokenize` (or any other leaf export) must not pull the evaluator,
  // function registry, materialiser, the excel module, the engine
  // installer, or the host glue. The scenarios below lock that
  // contract in.
  s(
    "/formula: tokenize (no engine / excel leak)",
    `${PKG_NAME}/formula`,
    ["tokenize"],
    [
      ...allModulesExcept("formula"),
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/integration/",
      "modules/formula/materialize/",
      "modules/formula/install.js",
      "modules/formula/host-registry.js",
      "modules/formula/default-syntax-probe.js"
    ]
  ),
  s(
    "/formula: parse (no engine / excel leak)",
    `${PKG_NAME}/formula`,
    ["parse"],
    [
      ...allModulesExcept("formula"),
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/integration/",
      "modules/formula/materialize/",
      "modules/formula/install.js",
      "modules/formula/host-registry.js",
      "modules/formula/default-syntax-probe.js"
    ]
  ),
  s(
    "/formula: calculateFormulas (no excel, no installer)",
    `${PKG_NAME}/formula`,
    ["calculateFormulas"],
    [
      ...allModulesExcept("formula"),
      "modules/formula/install.js",
      "modules/formula/host-registry.js",
      "modules/formula/default-syntax-probe.js"
    ]
  ),
  s(
    "/formula: installFormulaEngine (pulls full engine, no other modules)",
    `${PKG_NAME}/formula`,
    ["installFormulaEngine"],
    allModulesExcept("formula")
  ),

  // /markdown subpath (Node)
  s(
    "/markdown: parseMarkdown",
    `${PKG_NAME}/markdown`,
    ["parseMarkdown"],
    allModulesExcept("markdown")
  ),
  s(
    "/markdown: formatMarkdown",
    `${PKG_NAME}/markdown`,
    ["formatMarkdown"],
    allModulesExcept("markdown")
  ),
  s(
    "/markdown: parseMarkdownAll",
    `${PKG_NAME}/markdown`,
    ["parseMarkdownAll"],
    allModulesExcept("markdown")
  ),

  // Browser platform
  s(
    "browser root: Workbook (no pdf/formula leak)",
    PKG_NAME,
    ["Workbook"],
    // Same 3.9 KB glue exemption as the Node scenario above — see the
    // comment there for the rationale.
    [
      "modules/pdf/",
      "modules/formula/syntax/",
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/compile/",
      "modules/formula/materialize/",
      "modules/formula/integration/"
    ],
    "browser"
  ),
  s(
    "browser root: encodeCell",
    PKG_NAME,
    ["encodeCell"],
    [...allModulesExcept("excel"), "modules/excel/workbook", "modules/excel/worksheet"],
    "browser"
  ),
  s("browser root: BaseError", PKG_NAME, ["BaseError"], ALL_MODULES, "browser"),
  s(
    "browser /archive: ZipArchive",
    `${PKG_NAME}/archive`,
    ["ZipArchive"],
    NOT_EXCEL_PDF_CSV,
    "browser"
  ),
  s(
    "browser /csv: formatCsv",
    `${PKG_NAME}/csv`,
    ["formatCsv"],
    allModulesExcept("csv"),
    "browser"
  ),
  s(
    "browser /stream: ChunkedBuilder",
    `${PKG_NAME}/stream`,
    ["ChunkedBuilder"],
    allModulesExcept("stream"),
    "browser"
  ),
  s(
    "browser /xml: xmlEncode",
    `${PKG_NAME}/xml`,
    ["xmlEncode"],
    allModulesExcept("xml"),
    "browser"
  ),
  s("browser /pdf: pdf", `${PKG_NAME}/pdf`, ["pdf"], allModulesExcept("pdf", "archive"), "browser"),
  s(
    "browser /formula: tokenize (no engine leak)",
    `${PKG_NAME}/formula`,
    ["tokenize"],
    [
      ...allModulesExcept("formula"),
      "modules/formula/runtime/",
      "modules/formula/functions/",
      "modules/formula/integration/",
      "modules/formula/materialize/",
      "modules/formula/install.js",
      "modules/formula/host-registry.js",
      "modules/formula/default-syntax-probe.js"
    ],
    "browser"
  ),
  s(
    "browser /formula: calculateFormulas (no installer leak)",
    `${PKG_NAME}/formula`,
    ["calculateFormulas"],
    [
      ...allModulesExcept("formula"),
      "modules/formula/install.js",
      "modules/formula/host-registry.js",
      "modules/formula/default-syntax-probe.js"
    ],
    "browser"
  ),
  s(
    "browser /markdown: parseMarkdown",
    `${PKG_NAME}/markdown`,
    ["parseMarkdown"],
    allModulesExcept("markdown"),
    "browser"
  )
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
  const outFile = path.join(TMP_DIR, `esbuild-${slug}.out.mjs`);

  try {
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

    for (const chunk of output.output ?? []) {
      if (chunk.type !== "chunk" || !chunk.modules) {
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
      output: { path: outDir, filename: "out.mjs", library: { type: "module" } },
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
      stats: { all: false, modules: true, modulesSpace: Infinity }
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
      const jsonStats = stats.toJson({ modules: true, reasons: false });
      const contributing: ModuleEntry[] = [];
      let parsedCount = 0;

      for (const mod of jsonStats.modules ?? []) {
        const modName: string = mod.name ?? mod.identifier ?? "";
        if (!modName.includes("dist/")) {
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
    results.push(runEsbuild(s));
  }
  for (const s of scenarios) {
    results.push(await runRolldown(s));
  }
  for (const s of scenarios) {
    results.push(await runRspack(s));
  }

  const allPassed = report(results);
  fs.rmSync(TMP_DIR, { recursive: true });

  if (!allPassed) {
    process.exit(1);
  }
}

main();
