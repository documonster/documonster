/**
 * Codemod: convert an object-literal namespace (`export const Cell = { ... }`)
 * into flat named exports, and rewrite every `Cell.member(` call site across the
 * repo to the flat name.
 *
 * Member → flat name mapping per namespace is defined in NS_MAP below. Members
 * not listed get a default `${prefix}${PascalCase(member)}` name.
 *
 * Usage: pnpm exec tsx scripts/codemod-namespace-to-flat.ts <Namespace> [--write]
 *
 * This only handles the *definition* file's object literal and the call-site
 * rewrites `<NS>.<member>` → flat. It does NOT touch property accesses on
 * handles (that is a separate worksheet/public-API codemod).
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

interface NsConfig {
  /** Path to the definition file containing `export const <NS> = {...}`. */
  defFile: string;
  /** Map of member name → flat export name. Unlisted members use the default. */
  map: Record<string, string>;
  /** Default prefix for unlisted members, e.g. "cell". */
  prefix: string;
}

const ROOT = path.resolve(import.meta.dirname, "..");

const NS_CONFIG: Record<string, NsConfig> = {
  Cell: {
    defFile: "src/modules/excel/cell.ts",
    prefix: "cell",
    map: {
      Types: "CellTypes",
      $col$row: "cellAbsoluteAddress"
    }
  },
  Row: {
    defFile: "src/modules/excel/row.ts",
    prefix: "row",
    map: {}
  },
  Column: {
    defFile: "src/modules/excel/column.ts",
    prefix: "column",
    map: {}
  },
  Range: {
    defFile: "src/modules/excel/range.ts",
    prefix: "range",
    map: {
      $t$l: "rangeAbsoluteTopLeft",
      $b$r: "rangeAbsoluteBottomRight",
      $range: "rangeAbsolute",
      $shortRange: "rangeAbsoluteShort"
    }
  }
};

function pascal(s: string): string {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/** Compute flat name for a member. */
function flatName(cfg: NsConfig, member: string): string {
  if (Object.hasOwn(cfg.map, member)) {
    return cfg.map[member];
  }
  // strip leading underscores, remember them to re-prefix (keep @internal marker)
  const underscores = member.match(/^_*/)?.[0] ?? "";
  const bare = member.slice(underscores.length);
  return `${underscores}${cfg.prefix}${pascal(bare)}`;
}

function listSourceFiles(): string[] {
  const out: string[] = [];
  const walk = (dir: string) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        if (entry.name === "node_modules" || entry.name === "dist") continue;
        walk(full);
      } else if (/\.ts$/.test(entry.name) && !entry.name.endsWith(".d.ts")) {
        out.push(full);
      }
    }
  };
  walk(path.join(ROOT, "src"));
  return out;
}

function main() {
  const nsName = process.argv[2];
  const write = process.argv.includes("--write");
  const cfg = NS_CONFIG[nsName];
  if (!cfg) {
    console.error(`Unknown namespace: ${nsName}. Known: ${Object.keys(NS_CONFIG).join(", ")}`);
    process.exit(1);
  }
  console.log(`[codemod] namespace=${nsName} write=${write}`);

  // 1. Discover member names by parsing the def file's object literal.
  const defPath = path.join(ROOT, cfg.defFile);
  const defText = fs.readFileSync(defPath, "utf8");
  const sf = ts.createSourceFile(defPath, defText, ts.ScriptTarget.Latest, true);

  let objLiteral: ts.ObjectLiteralExpression | undefined;
  const visit = (node: ts.Node) => {
    if (ts.isVariableStatement(node)) {
      for (const decl of node.declarationList.declarations) {
        if (
          ts.isIdentifier(decl.name) &&
          decl.name.text === nsName &&
          decl.initializer &&
          ts.isObjectLiteralExpression(decl.initializer)
        ) {
          objLiteral = decl.initializer;
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sf);
  if (!objLiteral) {
    console.error(`Could not find \`export const ${nsName} = {...}\` in ${cfg.defFile}`);
    process.exit(1);
  }

  const members: string[] = [];
  for (const prop of objLiteral.properties) {
    if (prop.name && (ts.isIdentifier(prop.name) || ts.isStringLiteral(prop.name))) {
      members.push(prop.name.text);
    }
  }
  console.log(`[codemod] found ${members.length} members: ${members.join(", ")}`);

  const nameOf = new Map<string, string>();
  for (const m of members) {
    nameOf.set(m, flatName(cfg, m));
  }

  // Report any name collisions.
  const seen = new Map<string, string>();
  for (const [m, f] of nameOf) {
    if (seen.has(f)) {
      console.error(`COLLISION: ${m} and ${seen.get(f)} both map to ${f}`);
      process.exit(1);
    }
    seen.set(f, m);
  }

  console.log("[codemod] name map:");
  for (const [m, f] of nameOf) {
    console.log(`    ${nsName}.${m} -> ${f}`);
  }

  if (!write) {
    console.log("[codemod] dry run (pass --write to apply)");
    return;
  }

  // 2a. Transform the def file: replace `export const NS = {...}` with flat
  // declarations, one per member. Method shorthand -> `export function`,
  // arrow/value property -> `export const`.
  {
    const declParts: string[] = [];
    for (const prop of objLiteral.properties) {
      const name =
        prop.name && (ts.isIdentifier(prop.name) || ts.isStringLiteral(prop.name))
          ? prop.name.text
          : undefined;
      if (!name) {
        console.error(
          `[codemod] unsupported property kind in object literal: ${ts.SyntaxKind[prop.kind]}`
        );
        process.exit(1);
      }
      const flat = nameOf.get(name)!;
      if (ts.isMethodDeclaration(prop)) {
        // `foo(args): Ret { body }` -> `export function flat(args): Ret { body }`
        const sigStart = prop.getStart(sf);
        const nameNode = prop.name;
        const afterName = nameNode.getEnd();
        const body = prop.getText(sf).slice(afterName - sigStart);
        const asyncKw = prop.modifiers?.some(m => m.kind === ts.SyntaxKind.AsyncKeyword)
          ? "async "
          : "";
        declParts.push(`export ${asyncKw}function ${flat}${body}`);
      } else if (ts.isPropertyAssignment(prop)) {
        // `bar: <expr>` -> `export const flat = <expr>;`
        const init = prop.initializer.getText(sf);
        declParts.push(`export const ${flat} = ${init};`);
      } else if (ts.isShorthandPropertyAssignment(prop)) {
        // `Types` (shorthand for `Types: Types`) -> `export const flat = Types;`
        declParts.push(`export const ${flat} = ${name};`);
      } else {
        console.error(`[codemod] unsupported property: ${name} (${ts.SyntaxKind[prop.kind]})`);
        process.exit(1);
      }
    }

    // Find the full `export const NS = {...};` statement span to replace.
    let stmt: ts.Node = objLiteral;
    while (stmt.parent && !ts.isVariableStatement(stmt)) {
      stmt = stmt.parent;
    }
    const stmtStart = stmt.getStart(sf);
    const stmtEnd = stmt.getEnd();
    const newDefText =
      defText.slice(0, stmtStart) + declParts.join("\n\n") + defText.slice(stmtEnd);
    fs.writeFileSync(defPath, newDefText);
    console.log(
      `[codemod] rewrote def file ${cfg.defFile} into ${declParts.length} flat declarations`
    );
  }

  // 2b. Rewrite all call/usage sites repo-wide: `<NS>.<member>` -> flat name.

  // We do a token-level replacement using the TS scanner to avoid string/comment
  // false-positives, restricted to PropertyAccessExpression on the NS identifier.
  let filesChanged = 0;
  for (const file of listSourceFiles()) {
    const text = fs.readFileSync(file, "utf8");
    if (!text.includes(`${nsName}.`)) continue;
    const src = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
    const edits: { start: number; end: number; repl: string }[] = [];

    const v = (node: ts.Node) => {
      if (
        ts.isPropertyAccessExpression(node) &&
        ts.isIdentifier(node.expression) &&
        node.expression.text === nsName &&
        nameOf.has(node.name.text)
      ) {
        // Replace `NS.member` with flat name (covers the whole prop-access span).
        edits.push({
          start: node.getStart(src),
          end: node.name.getEnd(),
          repl: nameOf.get(node.name.text)!
        });
      }
      ts.forEachChild(node, v);
    };
    v(src);

    if (edits.length === 0) continue;
    edits.sort((a, b) => b.start - a.start);
    let out = text;
    for (const e of edits) {
      out = out.slice(0, e.start) + e.repl + out.slice(e.end);
    }
    fs.writeFileSync(file, out);
    filesChanged++;
    console.log(`    rewrote ${edits.length} sites in ${path.relative(ROOT, file)}`);
  }
  console.log(`[codemod] rewrote call sites in ${filesChanged} files`);
  console.log(
    `[codemod] NOTE: the object literal in ${cfg.defFile} must still be manually` +
      ` split into flat \`export function\`/\`export const\` declarations.`
  );
}

main();
