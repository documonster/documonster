/**
 * `pnpm api:changes` has to see every change to a public type.
 *
 * The generated list is presented as exhaustive, so a change it cannot see is a change a consumer meets unwarned.
 * Each case below is one the comparison used to miss: an index signature, an interface's own call signature, a
 * well-known symbol member, a narrowed generic constraint, and two distinct types whose names differ only by a
 * suffix the comparison used to strip.
 *
 * The comparison runs as a subprocess, as the other gate tests run theirs: `scripts/` is outside the typed project.
 */

import { execFileSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { afterAll, beforeAll, describe, expect, it } from "vitest";

const SCRIPTS = path.resolve(import.meta.dirname, "../../../scripts/api-changes");

let workspace: string;
let caseId = 0;

beforeAll(() => {
  workspace = mkdtempSync(path.join(tmpdir(), "api-changes-"));
});

afterAll(() => {
  rmSync(workspace, { recursive: true, force: true });
});

function fixture(dts: string): string {
  const dir = path.join(workspace, String(caseId++));
  mkdirSync(dir);
  writeFileSync(
    path.join(dir, "package.json"),
    JSON.stringify({ exports: { ".": { types: "./index.d.ts" } } })
  );
  writeFileSync(path.join(dir, "index.d.ts"), dts);
  return dir;
}

/** The names `diffApi` reports as changed or removed between two declaration files. */
function changes(before: string, after: string): string[] {
  const driver = `
    import { dumpApi } from ${JSON.stringify(path.join(SCRIPTS, "dump.ts"))};
    import { diffApi } from ${JSON.stringify(path.join(SCRIPTS, "diff.ts"))};
    const [a, b] = process.argv.slice(1);
    const d = diffApi(dumpApi(a, "node", "x", []), dumpApi(b, "node", "x", []));
    console.log(JSON.stringify([...d.changed, ...d.removed].map(r => r[0].split(" :: ")[1])));
  `;
  const out = execFileSync(
    process.execPath,
    ["--input-type=module", "-e", driver, fixture(before), fixture(after)],
    { encoding: "utf8" }
  );
  return JSON.parse(out) as string[];
}

describe("api:changes comparison", () => {
  it("sees an index signature removed", () => {
    expect(
      changes(
        "export interface Options { known?: string; [key: string]: unknown; }",
        "export interface Options { known?: string; }"
      )
    ).toEqual(["Options.[key: string]"]);
  });

  it("sees an interface's own call signature change", () => {
    expect(
      changes(
        "export interface Fn { (value: string): string; }",
        "export interface Fn { (value: number): number; }"
      )
    ).toEqual(["Fn.(call)"]);
  });

  it("sees a well-known symbol member change", () => {
    expect(
      changes(
        "export declare class Reader { [Symbol.asyncIterator](): AsyncIterator<string>; }",
        "export declare class Reader { [Symbol.asyncIterator](): AsyncIterator<number>; }"
      )
    ).toEqual(["Reader#[Symbol.asyncIterator]"]);
  });

  it("sees a generic constraint narrowed", () => {
    expect(
      changes(
        "export type Box<T extends string | number> = { value: T };",
        "export type Box<T extends string> = { value: T };"
      )
    ).toEqual(["Box"]);
  });

  it("does not equate two types whose names differ only by a suffix", () => {
    expect(
      changes(
        "export interface Array {} export interface ArrayType {} export declare function f(x: ArrayType): void;",
        "export interface Array {} export interface ArrayType {} export declare function f(x: Array): void;"
      )
    ).toEqual(["f"]);
  });

  it("reads an imported alias as the type it names, and leaves other names alone", () => {
    const dir = path.join(workspace, "alias");
    mkdirSync(dir);
    writeFileSync(path.join(dir, "impl.d.ts"), "export interface Result { col: number; }");
    writeFileSync(
      path.join(dir, "package.json"),
      JSON.stringify({ exports: { ".": { types: "./index.d.ts" } } })
    );
    writeFileSync(
      path.join(dir, "index.d.ts"),
      'import type { Result as ResultImpl } from "./impl";\nexport type Out = ResultImpl;\nexport declare function f(col: number): void;'
    );
    const driver = `
      import { dumpApi } from ${JSON.stringify(path.join(SCRIPTS, "dump.ts"))};
      const r = dumpApi(process.argv[1], "node", "x", []);
      console.log(JSON.stringify([r["x :: Out"].sig, r["x :: f"].sig]));
    `;
    const out = execFileSync(process.execPath, ["--input-type=module", "-e", driver, dir], {
      encoding: "utf8"
    });
    expect(JSON.parse(out)).toEqual(["type Out = Result;", "(col: number): void"]);
  });

  it("does not report a union printed in another order", () => {
    expect(
      changes(
        'export interface E {} export interface T {} export declare function f(x: E | T, y: "b" | "a"): (cb: (v: E | T) => void) => T | E;',
        'export interface E {} export interface T {} export declare function f(x: T | E, y: "a" | "b"): (cb: (v: T | E) => void) => E | T;'
      )
    ).toEqual([]);
  });

  it("still reports a union that gained a member", () => {
    expect(
      changes(
        "export interface E {} export interface T {} export declare function f(x: E | T): void;",
        "export interface E {} export interface T {} export declare function f(x: E | T | string): void;"
      )
    ).toEqual(["f"]);
  });

  it("does not report a parameter that was only renamed", () => {
    expect(
      changes(
        "export declare function f(opts: string): void;",
        "export declare function f(options: string): void;"
      )
    ).toEqual([]);
  });

  it("does not report a type qualified by a namespace import", () => {
    const before = 'import type { Ctx } from "./types"; export declare function f(c: Ctx): void;';
    const after = 'import type * as T from "./types"; export declare function f(c: T.Ctx): void;';
    const run = (dts: string): string => {
      const d = fixture(dts);
      writeFileSync(path.join(d, "types.d.ts"), "export interface Ctx { a: number; }");
      return d;
    };
    const driver = `
      import { dumpApi } from ${JSON.stringify(path.join(SCRIPTS, "dump.ts"))};
      import { diffApi } from ${JSON.stringify(path.join(SCRIPTS, "diff.ts"))};
      const [a, b] = process.argv.slice(1);
      const d = diffApi(dumpApi(a, "node", "x", []), dumpApi(b, "node", "x", []));
      console.log(JSON.stringify(d.changed.map(r => r[0])));
    `;
    const out = execFileSync(
      process.execPath,
      ["--input-type=module", "-e", driver, run(before), run(after)],
      { encoding: "utf8" }
    );
    expect(JSON.parse(out)).toEqual([]);
  });

  it("reports nothing for an unchanged surface", () => {
    const dts =
      "export interface O { [k: string]: number; (x: string): void; } export type B<T extends string> = { v: T };";
    expect(changes(dts, dts)).toEqual([]);
  });
});
