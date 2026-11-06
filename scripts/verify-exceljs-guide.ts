/**
 * MIGRATION_EXCELJS_TO_DOCUMONSTER.md has to account for every name exceljs published.
 *
 * The same check as `verify-excelts-guide.ts`, against `exceljs@4.4.0` — the last release, and the version the guide
 * maps from. exceljs publishes one declaration file whose objects are interfaces (`Worksheet`, `Cell`, …), so the
 * rules are:
 *
 * - an export is accounted for when `documonster/excel` publishes the same name, not deprecated, as a type or a value
 *   of the same kind (`Font`, `Style`, `PaperSize`) — or when the guide mentions it;
 * - a member is accounted for when its owner is still the same data shape in `documonster/excel` and has the member
 *   (`Font.bold`, `PageSetup.fitToPage`) — or when the guide mentions it. An owner that became a namespace of
 *   functions (`Worksheet`, `Cell`) has every member in the guide: `Worksheet.addRow` existing as a deprecated function
 *   says nothing about how `ws.addRow` is ported;
 * - every replacement the guide's tables give exists in the current build and is not deprecated.
 *
 * It needs `dist/types` (`pnpm build:esm`) and network access, so it is a command to run when the guide or the API
 * changes rather than part of `pnpm check`.
 *
 * Usage: node scripts/verify-exceljs-guide.ts
 */
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import ts from "typescript-api";

import { dumpApi } from "./api-changes/dump.ts";
import type { ApiEntry, ApiRecord } from "./api-changes/dump.ts";
import { checkReplacements, describesMember, mentionedNames } from "./api-changes/guide-check.ts";
import { fetchPackage } from "./api-changes/package.ts";

const ROOT = path.resolve(import.meta.dirname, "..");
const GUIDE = path.join(ROOT, "MIGRATION_EXCELJS_TO_DOCUMONSTER.md");
const EXCELJS = "exceljs@4.4.0";
const TYPE_ROOTS = [path.join(ROOT, "node_modules/@types")];

/** Kinds that describe data — an interface or a type a caller builds — rather than an object with methods. */
const DATA_KINDS = new Set(["interface", "type"]);

/**
 * Data types `documonster/excel` publishes under another name, as §9–§13 of the guide say: their members are compared
 * with the renamed type's.
 */
const TYPE_RENAMES: Record<string, string> = {
  DataValidation: "DataValidationRule",
  Address: "DecodedAddress",
  Cvfo: "Cfvo",
  Media: "WorkbookMedia",
  ImageRange: "AddImageRange",
  "stream.xlsx.ArchiverZipOptions": "Stream.WorkbookZipOptions"
};

/**
 * Each class and interface of the declaration file at `file`, by its qualified name (`stream.xlsx.WorkbookReader`),
 * with the ones it extends — so a member it inherits counts as described where its base's is.
 */
function bases(file: string): Map<string, string[]> {
  const sf = ts.createSourceFile(file, fs.readFileSync(file, "utf8"), ts.ScriptTarget.Latest, true);
  const declared = new Map<string, string[]>();
  const visit = (node: ts.Node, scope: string[]): void => {
    if (ts.isModuleDeclaration(node) && ts.isIdentifier(node.name)) {
      const inner = [...scope, node.name.text];
      node.body?.forEachChild(c => visit(c, inner));
      return;
    }
    if ((ts.isClassDeclaration(node) || ts.isInterfaceDeclaration(node)) && node.name) {
      const heritage = (node.heritageClauses ?? []).flatMap(h =>
        h.types.map(t => {
          // `Required<TableProperties>` extends what its argument has.
          const expr = t.expression.getText(sf);
          const arg = t.typeArguments?.[0];
          return expr === "Required" || expr === "Partial" ? (arg?.getText(sf) ?? expr) : expr;
        })
      );
      declared.set([...scope, node.name.text].join("."), heritage);
    }
    node.forEachChild(c => visit(c, scope));
  };
  visit(sf, []);
  // A base named within a namespace is that namespace's, when it declares one by the name.
  const out = new Map<string, string[]>();
  for (const [name, heritage] of declared) {
    const scope = name.split(".").slice(0, -1);
    out.set(
      name,
      heritage.map(h => {
        for (let i = scope.length; i >= 0; i--) {
          const candidate = [...scope.slice(0, i), h].join(".");
          if (declared.has(candidate)) {
            return candidate;
          }
        }
        return h;
      })
    );
  }
  return out;
}

/** `owner` and everything it extends, nearest first. */
function lineage(owner: string, of: Map<string, string[]>): string[] {
  const out: string[] = [];
  const walk = (name: string): void => {
    if (!out.includes(name)) {
      out.push(name);
      (of.get(name) ?? []).forEach(walk);
    }
  };
  walk(owner);
  return out;
}

/** The names `require("exceljs")` exports: its entry object's keys and the enums it assigns into it. */
function runtimeExports(pkg: string): string[] {
  const keysOf = (file: string): string[] => {
    const text = fs.readFileSync(path.join(pkg, file), "utf8");
    return [...text.matchAll(/^ {2}([A-Za-z_$][\w$]*):/gm)].map(m => m[1]!);
  };
  const entry = keysOf("lib/exceljs.nodejs.js");
  if (
    entry.length === 0 ||
    !/require\('\.\/doc\/enums'\)/.test(
      fs.readFileSync(path.join(pkg, "lib/exceljs.nodejs.js"), "utf8")
    )
  ) {
    throw new Error(
      "verify:exceljs-guide — exceljs's entry file no longer has the shape this reads"
    );
  }
  return [...entry, ...keysOf("lib/doc/enums.js")];
}

function main(): void {
  if (!fs.existsSync(path.join(ROOT, "dist/types"))) {
    console.error("✗ verify:exceljs-guide — dist/types is missing. Run pnpm build:esm.");
    process.exit(2);
  }
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-exceljs-guide-"));
  try {
    console.log(`verify:exceljs-guide — fetching ${EXCELJS}`);
    const pkg = fetchPackage(EXCELJS, work);
    // exceljs has no `exports` map; give its one declaration file one, so it is read as an entry.
    const dir = path.join(work, "entry");
    fs.mkdirSync(dir);
    fs.copyFileSync(path.join(pkg, "index.d.ts"), path.join(dir, "index.d.ts"));
    fs.writeFileSync(
      path.join(dir, "package.json"),
      JSON.stringify({ exports: { ".": { types: "./index.d.ts" } } })
    );
    const old = dumpApi(dir, "node", "x", TYPE_ROOTS);
    const heritage = bases(path.join(dir, "index.d.ts"));
    const now: ApiRecord = {
      ...dumpApi(ROOT, "node", "documonster", TYPE_ROOTS),
      ...dumpApi(ROOT, "browser", "documonster", TYPE_ROOTS)
    };
    const excel = (n: string): ApiEntry | undefined =>
      now[`documonster/excel :: ${n}`] ??
      now[`documonster/excel :: ${n} (type)`] ??
      now[`documonster/excel :: ${n}?`];

    const text = fs.readFileSync(GUIDE, "utf8");
    const mentioned = mentionedNames(text);
    const problems: string[] = [];

    for (const [k, v] of Object.entries(old)) {
      const name = k.split(" :: ")[1]!.replace(" (type)", "");
      // `Owner.member`, `Owner#member` or `Owner.[Symbol.iterator]` — a bracketed member keeps its dot.
      const split = /^(.+?)[.#](\[[^\]]+\]|[^.#[]+)$/.exec(name);
      const last = (split ? split[2]! : name).replace(/\?$/, "");
      if (!split || v.kind === "namespace") {
        // An export (or a namespace member such as `stream.xlsx.WorkbookWriter`, which is checked by its own name).
        const same = excel(name);
        if ((same && !same.deprecated && same.kind === v.kind) || mentioned.has(last)) {
          continue;
        }
        problems.push(
          `exceljs: \`${name}\` (${v.kind}) is not in documonster/excel and not in the guide`
        );
        continue;
      }
      // A member.
      const owner = split[1]!;
      // `stream.xlsx.WorkbookReader` is `Stream.WorkbookReader`, as §16 says.
      const renamed = TYPE_RENAMES[owner] ?? owner.replace(/^stream\.xlsx\./, "Stream.");
      const ownerNow = excel(renamed);
      const ownerThen = old[`x :: ${owner}`] ?? old[`x :: ${owner} (type)`];
      if (
        ownerThen &&
        !DATA_KINDS.has(ownerThen.kind) &&
        ownerThen.kind !== "class" &&
        ownerThen.kind !== "enum"
      ) {
        continue; // a member of a namespace — checked as an export above
      }
      if (ownerThen?.kind === "type") {
        continue; // a type alias's members are its target's (`HeaderArray` is a `string[]`); the alias is checked
      }
      if (ownerNow && DATA_KINDS.has(ownerNow.kind) && !ownerNow.deprecated) {
        const member = excel(`${renamed}.${last}`);
        if (member && !member.deprecated) {
          continue;
        }
      }
      // A class that is still a class keeps what it still has.
      if (ownerThen?.kind === "class" && ownerNow?.kind === "class" && !ownerNow.deprecated) {
        const member = excel(`${renamed}#${last}`);
        if (member && !member.deprecated) {
          continue;
        }
      }
      // An enum member of a constant object that kept the member (`ValueType.Null`).
      if (
        ownerNow?.kind === "const" &&
        !ownerNow.deprecated &&
        new RegExp(`(?:readonly )?${last.replace(/\$/g, "\\$")}:`).test(ownerNow.sig)
      ) {
        continue;
      }
      // A type the guide gives no counterpart (`CellMatrix`, `RelationshipType`) is described as a whole, members and
      // all. An owner that is still published under its name — as a namespace of functions, say — is not.
      const ownerTop = owner.split(".").pop()!;
      if (!ownerNow && mentioned.has(ownerTop)) {
        continue;
      }
      const named = lineage(owner, heritage).some(o => describesMember(text, o, last));
      if (!named) {
        problems.push(`exceljs: \`${name}\` is not described`);
      }
    }

    // What `require("exceljs")` gives is more than its declarations say (`ModelContainer` has none), so the entry
    // file and the enums it spreads in are read too: every name they export is checked as an export is.
    const declared = new Set(
      Object.keys(old).map(k => k.split(" :: ")[1]!.split(/[.#]/)[0]!.replace(" (type)", ""))
    );
    for (const name of runtimeExports(pkg)) {
      if (!declared.has(name) && !excel(name) && !mentioned.has(name)) {
        problems.push(
          `exceljs: \`${name}\` (exported at run time, undeclared) is not in documonster/excel and not in the guide`
        );
      }
    }

    problems.push(...checkReplacements(text, now));

    if (problems.length) {
      console.error(`✗ verify:exceljs-guide — ${problems.length} problem(s):`);
      for (const p of [...new Set(problems)]) {
        console.error(`  ${p}`);
      }
      process.exit(1);
    }
    console.log(
      `✓ verify:exceljs-guide — every name ${EXCELJS} publishes is accounted for, and every replacement exists.`
    );
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

main();
