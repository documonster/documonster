/**
 * The published JavaScript must parse in every browser the README supports: Chrome 98, Firefox 102,
 * Safari 15.4 (`browserslist` in `package.json`).
 *
 * A file an engine cannot parse is not one broken function: the engine rejects it at load, so the module —
 * and every chunk a bundler put it in — is gone. Two ways that has happened, and why each needs this check:
 *
 * - **Regular expressions are never lowered.** `tsc` and rolldown lower syntax to a target but ship a pattern
 *   as written. 0.25.1 had one lookbehind in `excel/utils/future-functions.ts`, and every bundle reaching it
 *   failed to load on Safari 16.0–16.3.
 * - **The IIFE bundles were not lowered at all.** rolldown's transform defaults to `esnext`, so TypeScript's
 *   static class fields came out as class static blocks, which Safari parses only from 16.4. Eight of the
 *   twelve bundles failed to load on Safari 15.4–16.3. `rolldown.config.ts` now takes its target from
 *   `browserslist`; this check is what notices if that stops being true.
 *
 * Refused, with the first release of each engine that parses it:
 *
 * - lookbehind `(?<=` / `(?<!` — Safari 16.4
 * - the `v` (unicode sets) flag — Chrome 112, Firefox 116, Safari 17
 * - inline modifiers `(?i:` / `(?-i:` — Chrome 125, Firefox 132, Safari 26
 * - a group name used twice — Chrome 125, Firefox 129, Safari 17
 * - a class static block — Safari 16.4
 * - `using` / `await using` — Chrome 134, Firefox 141, no Safari
 * - import attributes `with { … }` — Chrome 123, Firefox 138, Safari 17.2
 * - decorators and `accessor` fields — no engine without a flag
 *
 * The tree is parsed, not grepped: text cannot tell a regex from a division or from a string that mentions
 * `(?<=`. Regex literals are checked whole. For `RegExp(…)` / `new RegExp(…)` the pattern and flags are
 * evaluated as far as the file allows — string literals, templates, `+`, `String.raw`, and a `const` declared
 * once in the same file — and every computed part is checked as an atom that adds no syntax; a pattern that is
 * the caller's own (a search, a formula's `REGEX` argument) is the caller's.
 *
 * Worker sources reach a browser as *strings* handed to `new Blob([…])`, which no parse of the module that
 * holds them can see inside. A Blob whose text the file itself spells out is parsed as a program of its own;
 * one whose text comes from another module is read from the built tree — both through `WORKER_SPAWNERS` in
 * `lib/worker-scripts.ts`, which `verify-browser-floor.ts` reads too. Every file that starts a worker must be
 * listed there, so a new one cannot ship unchecked.
 *
 * Usage: `node scripts/verify-browser-syntax.ts [--root <dir>] [<dir> …]`. With no directories it checks
 * `dist/esm`, `dist/iife` and the worker scripts; given directories (`src`, say) it checks the files in them
 * that ship — `.ts` included, tests and examples left out as `tsconfig.esm.json` leaves them out. `--root`
 * exists so `src/test/__tests__/browser-syntax-gate.node.test.ts` can point it at trees built to break it.
 */
import fs from "node:fs";
import path from "node:path";

import ts from "typescript-api";

import { relPosix } from "./lib/paths.ts";
import {
  WORKER_SPAWNERS,
  blobScript,
  readWorkerScripts,
  singleConsts,
  staticText
} from "./lib/worker-scripts.ts";

function argValue(flag: string): string | undefined {
  const at = process.argv.indexOf(flag);
  return at === -1 ? undefined : process.argv[at + 1];
}

const ROOT = path.resolve(argValue("--root") ?? process.cwd());
const DIRS = process.argv
  .slice(2)
  .filter((arg, i, all) => arg !== "--root" && all[i - 1] !== "--root");
const DEFAULT_DIRS = ["dist/esm", "dist/iife"];

const SCANNED_EXTENSIONS = [".js", ".mjs", ".cjs", ".ts", ".mts", ".cts"];

/** What `pattern` with `flags` uses that the floor cannot parse. */
function regexProblems(pattern: string, flags: string): string[] {
  const found: string[] = [];
  if (flags.includes("v")) {
    found.push("the v flag");
  }
  const names = new Set<string>();
  let inClass = false;
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === "\\") {
      i++;
    } else if (inClass) {
      inClass = ch !== "]";
    } else if (ch === "[") {
      inClass = true;
    } else if (ch === "(" && pattern[i + 1] === "?") {
      const rest = pattern.slice(i + 2);
      const named = /^<([^=!>][^>]*)>/.exec(rest);
      if (rest.startsWith("<=") || rest.startsWith("<!")) {
        found.push("lookbehind");
      } else if (named) {
        if (names.has(named[1]!)) {
          found.push(`duplicate group name "${named[1]}"`);
        }
        names.add(named[1]!);
      } else if (/^[ims]*(?:-[ims]*)?:/.test(rest) && !rest.startsWith(":")) {
        found.push("inline modifiers");
      }
    }
  }
  return found;
}

function isRegExpCallee(node: ts.Expression): boolean {
  return (
    (ts.isIdentifier(node) && node.text === "RegExp") ||
    (ts.isPropertyAccessExpression(node) && node.name.text === "RegExp")
  );
}

/** The syntax `node` itself uses that the floor cannot parse, other than inside a regular expression. */
function syntaxProblems(node: ts.Node): string[] {
  if (ts.isClassStaticBlockDeclaration(node)) {
    return ["a class static block"];
  }
  if (ts.isVariableDeclarationList(node) && (node.flags & ts.NodeFlags.Using) !== 0) {
    return [
      (node.flags & ts.NodeFlags.AwaitUsing) === ts.NodeFlags.AwaitUsing ? "await using" : "using"
    ];
  }
  if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.attributes) {
    return ["import attributes"];
  }
  if (
    ts.isCallExpression(node) &&
    node.expression.kind === ts.SyntaxKind.ImportKeyword &&
    node.arguments.length > 1
  ) {
    return ["import attributes"];
  }
  if (ts.isDecorator(node)) {
    return ["a decorator"];
  }
  if (node.kind === ts.SyntaxKind.AccessorKeyword) {
    return ["an accessor field"];
  }
  return [];
}

/** Every problem in `source`, as `label:line:column — what in code`. */
function scanSource(label: string, source: string, kind: ts.ScriptKind): string[] {
  const file = ts.createSourceFile(label, source, ts.ScriptTarget.Latest, true, kind);
  const consts = singleConsts(file);
  const problems: string[] = [];
  const report = (node: ts.Node, what: string[]): void => {
    if (what.length === 0) {
      return;
    }
    const { line, character } = file.getLineAndCharacterOfPosition(node.getStart(file));
    const code = node.getText(file).replace(/\s+/g, " ");
    const shown = code.length > 100 ? `${code.slice(0, 100)}…` : code;
    problems.push(`${label}:${line + 1}:${character + 1} — ${what.join(", ")} in ${shown}`);
  };
  const visit = (node: ts.Node): void => {
    report(node, syntaxProblems(node));
    if (ts.isRegularExpressionLiteral(node)) {
      const slash = node.text.lastIndexOf("/");
      report(node, regexProblems(node.text.slice(1, slash), node.text.slice(slash + 1)));
    } else if (
      (ts.isNewExpression(node) || ts.isCallExpression(node)) &&
      isRegExpCallee(node.expression) &&
      node.arguments &&
      node.arguments.length > 0
    ) {
      const [patternArg, flagsArg] = node.arguments;
      const pattern = ts.isRegularExpressionLiteral(patternArg!)
        ? ""
        : (staticText(patternArg!, "x", consts) ?? "");
      const flags = flagsArg ? (staticText(flagsArg, "", consts) ?? "") : "";
      report(node, regexProblems(pattern, flags));
    } else {
      // A script assembled at run time cannot be read here; one the file spells out can.
      const script = blobScript(node, consts);
      if (script !== undefined) {
        const { line } = file.getLineAndCharacterOfPosition(node.getStart(file));
        problems.push(
          ...scanSource(`${label}:${line + 1} (Blob script)`, script, ts.ScriptKind.JS)
        );
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return problems;
}

/** Whether `source` starts a worker: `new Worker(…)` or `new SharedWorker(…)`. */
function startsWorker(source: string): boolean {
  const file = ts.createSourceFile("x.js", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  let found = false;
  const visit = (node: ts.Node): void => {
    if (
      ts.isNewExpression(node) &&
      ts.isIdentifier(node.expression) &&
      (node.expression.text === "Worker" || node.expression.text === "SharedWorker")
    ) {
      found = true;
    }
    if (!found) {
      ts.forEachChild(node, visit);
    }
  };
  visit(file);
  return found;
}

/** What `tsconfig.esm.json` leaves out of `dist/esm`, so a source scan reads what ships. */
const UNSHIPPED_DIRS = new Set(["node_modules", "__tests__", "examples"]);

function* files(dir: string): Generator<string> {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (!UNSHIPPED_DIRS.has(entry.name) && relPosix(ROOT, full) !== "src/test") {
        yield* files(full);
      }
    } else if (
      SCANNED_EXTENSIONS.includes(path.extname(entry.name)) &&
      !/\.d\.[cm]?ts$|\.test\.[cm]?ts$/.test(entry.name)
    ) {
      yield full;
    }
  }
}

function scriptKind(file: string): ts.ScriptKind {
  return /\.[cm]?ts$/.test(file) ? ts.ScriptKind.TS : ts.ScriptKind.JS;
}

const problems: string[] = [];
let scanned = 0;
const checkWorkers = DIRS.length === 0;
const spawners = new Set(WORKER_SPAWNERS.map(spawner => `dist/${spawner.file}`));

for (const dir of DIRS.length > 0 ? DIRS : DEFAULT_DIRS) {
  const full = path.resolve(ROOT, dir);
  if (!fs.existsSync(full)) {
    problems.push(`${dir}: does not exist — build first`);
    continue;
  }
  for (const file of files(full)) {
    scanned++;
    const label = relPosix(ROOT, file);
    const source = fs.readFileSync(file, "utf8");
    problems.push(...scanSource(label, source, scriptKind(file)));
    if (
      checkWorkers &&
      label.startsWith("dist/esm/") &&
      !spawners.has(label) &&
      /\b(?:Shared)?Worker\b/.test(source) &&
      startsWorker(source)
    ) {
      problems.push(
        `${label}: starts a worker but is not in WORKER_SPAWNERS, so its script is unchecked`
      );
    }
  }
}

if (checkWorkers) {
  // An inline script was checked above, where its spawner was parsed; one a module exports is checked here.
  const workers = await readWorkerScripts(path.resolve(ROOT, "dist"));
  problems.push(...workers.problems.map(problem => `dist/${problem}`));
  for (const { source, text, inline } of workers.scripts) {
    if (!inline) {
      scanned++;
      problems.push(...scanSource(`dist/${source} (worker script)`, text, ts.ScriptKind.JS));
    }
  }
}

if (problems.length > 0) {
  console.error(
    `✗ verify:browser-syntax — ${problems.length} problem(s) for the browser floor (Chrome 98, Firefox 102, Safari 15.4):`
  );
  for (const problem of problems) {
    console.error(`  ${problem}`);
  }
  process.exit(1);
}

console.log(
  `✓ verify:browser-syntax — ${scanned} file(s) use only syntax Chrome 98, Firefox 102 and Safari 15.4 parse`
);
