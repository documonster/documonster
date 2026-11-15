/**
 * `verify-browser-syntax.ts` has to fire on the syntax it refuses, and only on that.
 *
 * It runs as a subprocess against fixture trees, through the `--root` flag it takes for the purpose, so the
 * argument handling and the exit code that fails `pnpm build` are covered too. The two real failures it was
 * written for come first: a lookbehind in a regex literal, which kept 0.25.1 from loading on Safari 16.0–16.3,
 * and the class static blocks the unlowered IIFE bundles shipped, which kept them from loading before 16.4.
 */

import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { afterEach, beforeEach, describe, expect, it } from "vitest";

const SCRIPTS = path.resolve(import.meta.dirname, "../../../scripts");
const SCRIPT = path.join(SCRIPTS, "verify-browser-syntax.ts");

let root: string;

beforeEach(() => {
  root = mkdtempSync(path.join(tmpdir(), "browser-syntax-"));
});

afterEach(() => {
  rmSync(root, { recursive: true, force: true });
});

function write(relative: string, contents: string): void {
  const full = path.join(root, relative);
  mkdirSync(path.dirname(full), { recursive: true });
  writeFileSync(full, contents);
}

function run(args: readonly string[] = []): { status: number; output: string } {
  try {
    const stdout = execFileSync(process.execPath, [SCRIPT, "--root", root, ...args], {
      encoding: "utf8",
      stdio: "pipe"
    });
    return { status: 0, output: stdout };
  } catch (error) {
    const failed = error as { status?: number; stdout?: string; stderr?: string };
    return { status: failed.status ?? 1, output: `${failed.stdout ?? ""}${failed.stderr ?? ""}` };
  }
}

const INFLATE_SPAWNER = "dist/esm/modules/archive/unzip/stream.browser.js";

/** A built tree whose only content is `code`, with the worker spawners and scripts the gate expects. */
function writeDist(
  code: string,
  { workerScript = "self.onmessage = () => {};", inlineScript = "onmessage = () => {};" } = {}
): void {
  write("dist/esm/index.js", code);
  write("dist/iife/documonster.iife.min.js", "var a=1;");
  write(
    INFLATE_SPAWNER,
    `const code = ${JSON.stringify(inlineScript)};\n` +
      'export const w = () => new Worker(URL.createObjectURL(new Blob([code], { type: "text/javascript" })));\n'
  );
  write("dist/esm/modules/csv/worker/pool.browser.js", "export const w = u => new Worker(u);\n");
  write(
    "dist/esm/modules/csv/worker/worker-script.generated.js",
    `export const CSV_WORKER_SCRIPT = ${JSON.stringify(workerScript)};\n`
  );
  write(
    "dist/esm/modules/archive/compression/worker-pool/pool.browser.js",
    "export const w = u => new Worker(u);\n"
  );
  write(
    "dist/esm/modules/archive/compression/worker-pool/worker-script.js",
    "export function generateWorkerScript() { return 'self.onmessage = () => {};'; }\n"
  );
}

describe("verify-browser-syntax", () => {
  it.each([
    ["a lookbehind literal", "export const r = /(?<![\\p{L}_.])_xlpm\\.\\w+/giu;", "lookbehind"],
    ["a positive lookbehind", "export const r = /(?<=>)x/;", "lookbehind"],
    ["the v flag", "export const r = /[\\p{L}--\\p{Lu}]/v;", "the v flag"],
    ["inline modifiers", "export const r = /(?i:a)b/;", "inline modifiers"],
    ["removed modifiers", "export const r = /(?-i:a)b/i;", "inline modifiers"],
    ["a duplicate group name", "export const r = /(?<y>a)|(?<y>b)/;", 'duplicate group name "y"'],
    ["a lookbehind in new RegExp", 'export const r = new RegExp("(?<!a)b");', "lookbehind"],
    ["a lookbehind in RegExp()", `export const r = RegExp(\`(?<=\${"a"})b\`);`, "lookbehind"],
    ["a concatenated pattern", 'export const r = new RegExp("(?<" + "=a)b");', "lookbehind"],
    [
      "a pattern held in a const",
      `const BEHIND = "(?<=";\nexport const r = new RegExp(\`\${BEHIND}a)b\`);`,
      "lookbehind"
    ],
    ["a String.raw pattern", "export const r = new RegExp(String.raw`(?<=\\d)x`);", "lookbehind"],
    ["a flag passed to new RegExp", 'export const r = new RegExp(s, "gv");', "the v flag"],
    ["a flag held in a const", 'const F = "v";\nexport const r = new RegExp(s, F);', "the v flag"],
    ["a flag added to a literal", 'export const r = new RegExp(/a/, "v");', "the v flag"],
    ["a class static block", "export class A { static { this.x = 1; } }", "a class static block"],
    ["a minified static block", "export var A=class{static{this.x=1}};", "a class static block"],
    ["using", "export function f() { using r = g(); }", "using"],
    ["await using", "export async function f() { await using r = g(); }", "await using"],
    ["import attributes", 'import d from "./d.json" with { type: "json" };', "import attributes"],
    [
      "dynamic import attributes",
      'export const d = import("./d.json", { with: { type: "json" } });',
      "import attributes"
    ]
  ])("refuses %s", (_, code, what) => {
    writeDist(code);
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain("dist/esm/index.js:");
    expect(result.output).toContain(`— ${what}`);
  });

  it.each([
    ["the consumed-prefix form", "export const r = /(?:^|[^\\p{L}])(_xlpm\\.\\w+)/giu;"],
    // Inside the floor: Chrome 90, Firefox 88, Safari 15.
    ["the d flag", "export const r = /a/d;"],
    ["lookahead and named groups", "export const r = /(?=a)(?!b)(?<x>c)\\k<x>/su;"],
    ["an escaped parenthesis", "export const r = /\\(?<=a/;"],
    ["a parenthesis in a class", "export const r = /[(?<=]a/;"],
    ["a string that mentions the syntax", 'export const s = "/(?<=a)/d static { }";'],
    ["a comment that mentions it", "// /(?<=a)/v using x\nexport const x = 1;"],
    ["division that reads like a regex", "export const f = (a, b, d) => a /(b)/ d;"],
    ["an escaped string pattern", 'export const r = new RegExp("\\\\(?<=a)");'],
    ["a computed pattern", "export const r = new RegExp(source, flags);"],
    [
      "a name declared twice, which cannot be known",
      'export function f() { const p = "(?<=a)"; return p; }\n' +
        'export function g() { const p = "a"; return new RegExp(p); }'
    ],
    ["a parameter that shadows nothing", "export const f = p => new RegExp(p);"],
    ["class fields", "export class A { static x = 1; #y = 2; z = this.#y; }"],
    ["a variable named using", "let using = 1; using = 2; export { using };"],
    ["a Blob of text", 'export const b = new Blob(["/(?<=a)/"], { type: "text/plain" });']
  ])("accepts %s", (_, code) => {
    writeDist(code);
    expect(run()).toMatchObject({ status: 0 });
  });

  it("checks the minified bundles", () => {
    writeDist("export {};");
    write("dist/iife/documonster.iife.min.js", "var a=/(?<!x)y/g,b=1/2/3;");
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain("dist/iife/documonster.iife.min.js:1:7 — lookbehind");
  });

  it("checks a worker script a module exports", () => {
    writeDist("export {};", { workerScript: "class W { static { } }\nconst r = /(?<=a)b/;" });
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain(
      "worker-script.generated.js (worker script):1:11 — a class static block"
    );
    expect(result.output).toContain("worker-script.generated.js (worker script):2:11 — lookbehind");
  });

  it("checks a worker script spelled out in a Blob", () => {
    writeDist("export {};", { inlineScript: "const r = /(?<=a)b/;" });
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain(`${INFLATE_SPAWNER}:2 (Blob script):1:11 — lookbehind`);
  });

  it("fails when an inline worker script can no longer be read", () => {
    writeDist("export {};");
    write(
      INFLATE_SPAWNER,
      "export const w = code => new Worker(URL.createObjectURL(new Blob([code])));\n"
    );
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain(
      `${INFLATE_SPAWNER}: WORKER_SPAWNERS says it spells out its worker script in a Blob, but 0 were read`
    );
  });

  it("fails on a file that starts a worker nobody listed", () => {
    writeDist("export const w = u => new SharedWorker(u);");
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain(
      "dist/esm/index.js: starts a worker but is not in WORKER_SPAWNERS"
    );
  });

  it("does not take a mention of Worker for starting one", () => {
    writeDist('export const hasWorkers = typeof Worker !== "undefined";');
    expect(run()).toMatchObject({ status: 0 });
  });

  it("fails when a listed worker script is gone, rather than passing without it", () => {
    writeDist("export {};");
    rmSync(path.join(root, "dist/esm/modules/archive/compression/worker-pool/worker-script.js"));
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain("listed as holding a worker script, but does not exist");
  });

  it("fails when the tree has not been built", () => {
    const result = run();
    expect(result.status).toBe(1);
    expect(result.output).toContain("dist/esm: does not exist");
  });

  it("checks the TypeScript that ships when pointed at source, and only that", () => {
    write("src/a.ts", "const r: RegExp = /(?<=a)b/;\nexport { r };\n");
    write("src/a.d.ts", "export declare const r: RegExp; // /(?<=a)b/\n");
    write("src/a.test.ts", "await using x = f();\n");
    write("src/m/__tests__/b.ts", "const r = /(?<=a)b/;\n");
    write("src/m/examples/c.ts", "const r = /(?<=a)b/;\n");
    write("src/test/d.ts", "const r = /(?<=a)b/;\n");
    const result = run(["src"]);
    expect(result.status).toBe(1);
    expect(result.output).toContain("src/a.ts:1:19 — lookbehind");
    expect(result.output).toContain("1 problem(s)");
  });
});

describe("browserTargets", () => {
  function targets(browserslist: unknown): { status: number; output: string } {
    write("package.json", JSON.stringify({ browserslist }));
    try {
      const stdout = execFileSync(
        process.execPath,
        [
          "--input-type=module",
          "-e",
          // A file URL, not a path: on Windows the ESM loader reads `D:\…` as a URL whose scheme is `d:`.
          `import { browserTargets } from ${JSON.stringify(pathToFileURL(path.join(SCRIPTS, "lib/browser-targets.ts")).href)};` +
            `console.log(JSON.stringify(browserTargets(${JSON.stringify(root)})));`
        ],
        { encoding: "utf8", stdio: "pipe" }
      );
      return { status: 0, output: stdout.trim() };
    } catch (error) {
      const failed = error as { status?: number; stderr?: string };
      return { status: failed.status ?? 1, output: failed.stderr ?? "" };
    }
  }

  it("turns browserslist into rolldown's target", () => {
    expect(targets(["chrome >= 98", "safari >= 15.4"])).toEqual({
      status: 0,
      output: '["chrome98","safari15.4"]'
    });
  });

  it.each([["> 0.5%"], ["last 2 versions"], ["not dead"]])(
    "refuses %j rather than dropping it from the target",
    query => {
      const result = targets(["chrome >= 98", query]);
      expect(result.status).toBe(1);
      expect(result.output).toContain('is not "<browser> >= <version>"');
    }
  );

  it("refuses a missing browserslist", () => {
    expect(targets(undefined).status).toBe(1);
  });
});
