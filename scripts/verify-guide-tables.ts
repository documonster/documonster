/**
 * The replacements the migration guides' tables give have to compile.
 *
 * A guide's tables are where a reader copies from — `Cell.setValue(ws, "A1", 42)`, `Chart.place(ws, range, opts)` —
 * and they are inline code, which `verify-doc-examples.ts` deliberately does not compile (it reads fenced blocks; it
 * checks that a table's names exist, not that a call is well formed). So a replacement could name a real function
 * with its arguments in the wrong number, or call something that is not a function, and nothing would notice.
 *
 * This takes every inline code span in the right-hand columns of the guides' tables that calls a `documonster`
 * namespace (`Range.expand(r, t, l, b, right)`, `new Stream.WorkbookWriter(opts)`, `await Workbook.read(buf)`),
 * imports each namespace from the entry that publishes the member it calls, and compiles the span as a statement.
 * The placeholders a table uses (`ws`, `addr`, `opts`) are declared `any`: what is checked is the call — that the
 * member exists and accepts that many arguments — not the placeholders' types, which a table never states.
 *
 * Spans that are not code (`Row.*(ws, n)`, a `…`, a parameter marked optional with `?`, a type union) are skipped.
 *
 * Needs `dist/types`; `pnpm check` builds it first.
 *
 * Usage: node scripts/verify-guide-tables.ts
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import ts from "typescript-api";

import { dumpApi } from "./api-changes/dump.ts";

const ROOT = path.resolve(import.meta.dirname, "..");
const GUIDES = [
  "MIGRATION.md",
  "MIGRATION_EXCELJS_TO_DOCUMONSTER.md",
  "MIGRATION_EXCELTS_TO_DOCUMONSTER.md"
];

interface Snippet {
  guide: string;
  line: number;
  code: string;
}

/** The code spans in the right-hand columns of `text`'s tables that call something. */
function snippets(guide: string, text: string): Snippet[] {
  const out: Snippet[] = [];
  let fence = false;
  text.split("\n").forEach((row, i) => {
    if (row.startsWith("```")) {
      fence = !fence;
    }
    if (fence || !row.startsWith("|") || /^\|\s*-/.test(row)) {
      return;
    }
    for (const cell of row.split(/(?<!\\)\|/).slice(2, -1)) {
      for (const m of cell.matchAll(/`([^`]+)`/g)) {
        const code = m[1]!.replace(/\\\|/g, "|").trim();
        if (
          !/[A-Z][\w$]*\.[A-Za-z_$][\w$]*\s*\(/.test(code) ||
          /^import\b|;\s*\S/.test(code) ||
          /…|\*\(|\?\s*[),}]|\?$|\(\.\.\.\)|, \.\.\.\)|\s\|\s|^\(|→/.test(code)
        ) {
          continue;
        }
        out.push({ guide, line: i + 1, code });
      }
    }
  });
  return out;
}

/** Globals a snippet may use as themselves; every other free name in it is a placeholder. */
const GLOBALS = new Set([
  "Object",
  "Array",
  "Number",
  "String",
  "Boolean",
  "Date",
  "Math",
  "JSON",
  "Promise",
  "Map",
  "Set",
  "Symbol",
  "Uint8Array",
  "ArrayBuffer",
  "Buffer",
  "TextEncoder",
  "TextDecoder",
  "Blob",
  "crypto",
  "globalThis",
  "undefined",
  "NaN",
  "Infinity",
  "console",
  "fetch",
  "Error",
  "RegExp",
  "BigInt"
]);

/** The names `code` uses without declaring them, which are not globals: the table's placeholders. */
function freeNames(code: string): Set<string> {
  const sf = ts.createSourceFile("s.ts", `async () => { ${code}; }`, ts.ScriptTarget.ESNext, true);
  const declared = new Set<string>();
  const used = new Set<string>();
  const visit = (n: ts.Node): void => {
    if (ts.isParameter(n) || ts.isVariableDeclaration(n) || ts.isBindingElement(n)) {
      const names = (b: ts.BindingName): void => {
        if (ts.isIdentifier(b)) {
          declared.add(b.text);
        } else {
          b.elements.forEach(e => !ts.isOmittedExpression(e) && names(e.name));
        }
      };
      names(n.name);
    }
    if (ts.isIdentifier(n)) {
      const p = n.parent;
      const isMemberName =
        (ts.isPropertyAccessExpression(p) && p.name === n) ||
        (ts.isPropertyAssignment(p) && p.name === n) ||
        (ts.isParameter(p) && p.name === n) ||
        (ts.isVariableDeclaration(p) && p.name === n) ||
        (ts.isBindingElement(p) && p.name === n);
      if (!isMemberName) {
        used.add(n.text);
      }
    }
    n.forEachChild(visit);
  };
  visit(sf);
  return new Set([...used].filter(n => !declared.has(n) && !GLOBALS.has(n)));
}

function main(): void {
  if (!fs.existsSync(path.join(ROOT, "dist/types"))) {
    console.error("✗ verify:guide-tables — dist/types is missing. Run pnpm build:esm.");
    process.exit(2);
  }
  const api = dumpApi(ROOT, "node", "documonster", [path.join(ROOT, "node_modules/@types")]);
  /** Namespace (or class) → the entries that publish it, and the members each has. */
  const publishers = new Map<string, Map<string, Set<string>>>();
  for (const [k, v] of Object.entries(api)) {
    const [entry, name] = k.split(" :: ") as [string, string];
    const [top, ...rest] = name.replace(" (type)", "").split(/[.#]/);
    if (v.deprecated && rest.length > 0) {
      continue;
    }
    const byEntry = publishers.get(top!) ?? publishers.set(top!, new Map()).get(top!)!;
    const members = byEntry.get(entry) ?? byEntry.set(entry, new Set()).get(entry)!;
    if (rest.length > 0) {
      members.add(rest[0]!.replace(/\?$/, ""));
    }
  }

  const all = GUIDES.flatMap(g => snippets(g, fs.readFileSync(path.join(ROOT, g), "utf8")));
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-guide-tables-"));
  try {
    const imports = new Map<string, string>(); // `${entry}::${ns}` → local alias
    const bodies: string[] = [];
    const lineOf: { snippet: Snippet; start: number; end: number }[] = [];
    let source = "";
    const unresolved: string[] = [];
    all.forEach((snippet, i) => {
      let code = snippet.code;
      for (const m of [...snippet.code.matchAll(/(?<![\w$.])([A-Z][\w$]*)\.([A-Za-z_$][\w$]*)/g)]) {
        const [whole, ns, member] = m as unknown as [string, string, string];
        const byEntry = publishers.get(ns);
        if (!byEntry) {
          continue; // not a documonster namespace (`Object.entries`, `Buffer.from`)
        }
        const entry = [...byEntry].find(([, members]) => members.has(member))?.[0];
        if (!entry) {
          unresolved.push(`${snippet.guide}:${snippet.line}: \`${ns}.${member}\` is not published`);
          continue;
        }
        const key = `${entry}::${ns}`;
        const alias = imports.get(key) ?? `${ns}__${imports.size}`;
        imports.set(key, alias);
        code = code.replace(whole, `${alias}.${member}`);
      }
      if (code === snippet.code) {
        bodies.push("");
        return; // calls nothing documonster publishes
      }
      const placeholders = [...freeNames(code)].filter(n => !/__\d+$/.test(n));
      const declare = placeholders.length
        ? `let ${placeholders.map(n => `${n}: any`).join(", ")};\n`
        : "";
      bodies.push(`async function snippet${i}(): Promise<void> {\n${declare}${code};\n}`);
    });
    for (const [key, alias] of imports) {
      const [entry, ns] = key.split("::") as [string, string];
      source += `import { ${ns} as ${alias} } from "${entry}";\n`;
    }
    source += "/* declarations */\n";
    for (let i = 0; i < bodies.length; i++) {
      const start = source.split("\n").length;
      source += bodies[i] + "\n";
      lineOf.push({ snippet: all[i]!, start, end: source.split("\n").length - 1 });
    }

    const file = path.join(work, "tables.ts");
    fs.writeFileSync(file, source);
    fs.mkdirSync(path.join(work, "node_modules"));
    fs.symlinkSync(ROOT, path.join(work, "node_modules", "documonster"));
    const program = ts.createProgram([file], {
      strict: true,
      noEmit: true,
      target: ts.ScriptTarget.ESNext,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      skipLibCheck: true,
      lib: ["lib.esnext.d.ts", "lib.dom.d.ts"],
      types: ["node"],
      typeRoots: [path.join(ROOT, "node_modules/@types")]
    });
    const diagnostics = ts.getPreEmitDiagnostics(program).filter(d => d.file?.fileName === file);

    const problems = [...unresolved];
    for (const d of diagnostics) {
      if (d.start === undefined) {
        continue;
      }
      const line = d.file!.getLineAndCharacterOfPosition(d.start).line + 1;
      const at = lineOf.find(l => line >= l.start && line <= l.end);
      const message = ts.flattenDiagnosticMessageText(d.messageText, " ").slice(0, 200);
      problems.push(
        at
          ? `${at.snippet.guide}:${at.snippet.line}: \`${at.snippet.code}\` — ${message}`
          : `generated line ${line}: ${message}`
      );
    }
    if (problems.length) {
      console.error(`✗ verify:guide-tables — ${problems.length} problem(s):`);
      for (const p of problems) {
        console.error(`  ${p}`);
      }
      process.exit(1);
    }
    console.log(
      `✓ verify:guide-tables — ${all.length} table replacement(s) across ${GUIDES.length} guide(s) compile.`
    );
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

main();
