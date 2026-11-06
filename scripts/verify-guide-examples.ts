/**
 * Every TypeScript example in the migration guides has to compile — as a whole block, the way a reader pastes it.
 *
 * `verify-doc-examples.ts` checks that a documented import and each `Namespace.member` it names exist, and
 * `verify-guide-tables.ts` that a table's one-line replacement is a well-formed call. Neither compiles a fenced
 * block, so a block could declare the same `const` twice (a "before" and an "after" pasted together), use a name it
 * never imported, or hand a browser result to a DOM API that refuses its type — and every gate stayed green. All
 * three shipped in these guides; this is the check that would have caught them.
 *
 * Each ```` ```ts ```` / ```` ```typescript ```` fence is compiled as its own module, strictly. The info string after
 * the language selects how:
 *
 * - (nothing) — the Node surface (`types: ["node"]`, the default export conditions).
 * - `browser` — the browser surface: `customConditions: ["browser"]`, `lib: DOM`, and no Node types, so a Node-only
 *   member (`Workbook.readFile`) or a Node-only global (`Buffer`) is an error, as it is for a reader there.
 * - `before` — code of the library being migrated *from* (exceljs, excelts), which is not installed here. It is not
 *   compiled, and it is counted in the result line so the number of uncompiled blocks is visible, not silent.
 *
 * A block may use the placeholders declared in {@link PRELUDE} (`wb`, `ws`, `bytes`, …) without declaring them —
 * they stand for "the workbook you already have" — typed as what they stand for, never `any`. Any other free name
 * is an error: a block that needs one declares it.
 *
 * `js`, `json`, `bash` and `html` fences are configuration and shell, not API use, and are out of scope.
 *
 * Needs `dist/types`; `pnpm check` builds it first.
 *
 * Usage: node scripts/verify-guide-examples.ts
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import ts from "typescript-api";

const ROOT = path.resolve(import.meta.dirname, "..");
const GUIDES = [
  "MIGRATION.md",
  "MIGRATION_EXCELJS_TO_DOCUMONSTER.md",
  "MIGRATION_EXCELTS_TO_DOCUMONSTER.md"
];

type Mode = "node" | "browser" | "before";

interface Block {
  guide: string;
  /** 1-based line of the first code line (the line after the opening fence). */
  line: number;
  mode: Mode;
  code: string;
}

/** The placeholders a block may use undeclared, typed as what they stand for. */
const PRELUDE = `
declare const wb: import("documonster/excel").Workbook.Handle;
declare const ws: import("documonster/excel").Worksheet.Handle;
declare const bytes: Uint8Array;
declare const pdfBytes: Uint8Array;
declare const doc: import("documonster/word").DocxDocument;
`;

function blocks(guide: string, text: string): Block[] {
  const out: Block[] = [];
  const lines = text.split("\n");
  for (let i = 0; i < lines.length; i++) {
    const open = /^```(ts|typescript)\b\s*(.*)$/.exec(lines[i]!);
    if (!open) {
      continue;
    }
    const flags = open[2]!.trim().split(/\s+/).filter(Boolean);
    const unknown = flags.filter(f => f !== "browser" && f !== "before");
    if (unknown.length) {
      throw new Error(`${guide}:${i + 1}: unknown fence flag(s) ${unknown.join(", ")}`);
    }
    const mode: Mode = flags.includes("before")
      ? "before"
      : flags.includes("browser")
        ? "browser"
        : "node";
    const body: string[] = [];
    let j = i + 1;
    while (j < lines.length && !lines[j]!.startsWith("```")) {
      body.push(lines[j]!);
      j++;
    }
    out.push({ guide, line: i + 2, mode, code: body.join("\n") });
    i = j;
  }
  return out;
}

function compile(work: string, list: Block[], browser: boolean): string[] {
  if (list.length === 0) {
    return [];
  }
  const dir = path.join(work, browser ? "browser" : "node");
  fs.mkdirSync(dir);
  fs.writeFileSync(path.join(dir, "prelude.d.ts"), PRELUDE);
  const files = list.map((b, i) => {
    const file = path.join(dir, `block${i}.ts`);
    fs.writeFileSync(file, `${b.code}\nexport {};\n`);
    return file;
  });
  const program = ts.createProgram([path.join(dir, "prelude.d.ts"), ...files], {
    strict: true,
    noEmit: true,
    target: ts.ScriptTarget.ESNext,
    module: ts.ModuleKind.ESNext,
    moduleResolution: ts.ModuleResolutionKind.Bundler,
    skipLibCheck: true,
    ...(browser
      ? { customConditions: ["browser"], lib: ["lib.esnext.d.ts", "lib.dom.d.ts"], types: [] }
      : {
          lib: ["lib.esnext.d.ts"],
          types: ["node"],
          typeRoots: [path.join(ROOT, "node_modules/@types")]
        })
  });
  const problems: string[] = [];
  for (const d of ts.getPreEmitDiagnostics(program)) {
    const index = d.file ? files.indexOf(d.file.fileName) : -1;
    const message = ts.flattenDiagnosticMessageText(d.messageText, " ").slice(0, 240);
    if (index === -1 || d.start === undefined) {
      problems.push(`${browser ? "[browser] " : ""}${d.file?.fileName ?? "?"}: ${message}`);
      continue;
    }
    const block = list[index]!;
    const line = d.file!.getLineAndCharacterOfPosition(d.start).line;
    problems.push(`${block.guide}:${block.line + line}${browser ? " [browser]" : ""}: ${message}`);
  }
  return problems;
}

function main(): void {
  if (!fs.existsSync(path.join(ROOT, "dist/types"))) {
    console.error("✗ verify:guide-examples — dist/types is missing. Run pnpm build:esm.");
    process.exit(2);
  }
  const all = GUIDES.flatMap(g => blocks(g, fs.readFileSync(path.join(ROOT, g), "utf8")));
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-guide-examples-"));
  try {
    fs.mkdirSync(path.join(work, "node_modules"));
    fs.symlinkSync(ROOT, path.join(work, "node_modules", "documonster"));
    const problems = [
      ...compile(
        work,
        all.filter(b => b.mode === "node"),
        false
      ),
      ...compile(
        work,
        all.filter(b => b.mode === "browser"),
        true
      )
    ];
    if (problems.length) {
      console.error(`✗ verify:guide-examples — ${problems.length} problem(s):`);
      for (const p of problems) {
        console.error(`  ${p}`);
      }
      process.exit(1);
    }
    const count = (m: Mode): number => all.filter(b => b.mode === m).length;
    console.log(
      `✓ verify:guide-examples — ${count("node")} Node and ${count("browser")} browser example(s) across ` +
        `${GUIDES.length} guide(s) compile; ${count("before")} marked \`before\` (another library's code) are not compiled.`
    );
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

main();
