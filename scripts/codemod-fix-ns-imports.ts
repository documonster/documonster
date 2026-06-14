/**
 * Follow-up to codemod-namespace-to-flat: fix imports.
 *
 * After call sites are rewritten from `NS.member` to flat `nsMember` names, the
 * `import { NS } from "<defModule>"` specifiers are stale. This script, for each
 * file, finds which flat names (from the given namespace's export set) are used
 * and rewrites the import to bring those in by name, dropping the `NS` import.
 *
 * Usage: node scripts/codemod-fix-ns-imports.ts <Namespace> [--write]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

interface NsConfig {
  defFile: string;
  /** module specifiers that import this namespace */
  moduleSpecifiers: string[];
}

const NS_CONFIG: Record<string, NsConfig> = {
  Cell: {
    defFile: "src/modules/excel/cell.ts",
    moduleSpecifiers: ["@excel/cell", "../cell", "./cell"]
  },
  Row: {
    defFile: "src/modules/excel/row.ts",
    moduleSpecifiers: ["@excel/row", "../row", "./row"]
  },
  Column: {
    defFile: "src/modules/excel/column.ts",
    moduleSpecifiers: ["@excel/column", "../column", "./column"]
  },
  Range: {
    defFile: "src/modules/excel/range.ts",
    moduleSpecifiers: ["@excel/range", "../range", "./range"]
  }
};

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

/** Collect the set of exported names from the def file. */
function exportedNames(defPath: string): Set<string> {
  const text = fs.readFileSync(defPath, "utf8");
  const sf = ts.createSourceFile(defPath, text, ts.ScriptTarget.Latest, true);
  const names = new Set<string>();
  const visit = (node: ts.Node) => {
    const isExported = ts.canHaveModifiers(node)
      ? ts.getModifiers(node)?.some(m => m.kind === ts.SyntaxKind.ExportKeyword)
      : false;
    if (isExported) {
      if (ts.isFunctionDeclaration(node) && node.name) {
        names.add(node.name.text);
      } else if (ts.isVariableStatement(node)) {
        for (const d of node.declarationList.declarations) {
          if (ts.isIdentifier(d.name)) names.add(d.name.text);
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sf);
  return names;
}

function main() {
  const nsName = process.argv[2];
  const write = process.argv.includes("--write");
  const cfg = NS_CONFIG[nsName];
  if (!cfg) {
    console.error(`Unknown namespace: ${nsName}`);
    process.exit(1);
  }
  const exported = exportedNames(path.join(ROOT, cfg.defFile));
  console.log(`[fix-imports] ${nsName}: ${exported.size} exported names`);

  let changed = 0;
  for (const file of listSourceFiles()) {
    const text = fs.readFileSync(file, "utf8");
    const src = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);

    // Find the import declaration that imports `NS` from one of the specifiers.
    let targetImport: ts.ImportDeclaration | undefined;
    let importsNs = false;
    for (const stmt of src.statements) {
      if (
        ts.isImportDeclaration(stmt) &&
        ts.isStringLiteral(stmt.moduleSpecifier) &&
        cfg.moduleSpecifiers.includes(stmt.moduleSpecifier.text) &&
        stmt.importClause?.namedBindings &&
        ts.isNamedImports(stmt.importClause.namedBindings)
      ) {
        for (const el of stmt.importClause.namedBindings.elements) {
          if (el.name.text === nsName && !el.isTypeOnly) {
            targetImport = stmt;
            importsNs = true;
          }
        }
      }
    }
    if (!targetImport || !importsNs) continue;

    // Which flat names from this namespace are actually used in the file?
    const used = new Set<string>();
    const v = (node: ts.Node) => {
      if (ts.isIdentifier(node) && exported.has(node.text)) {
        // exclude the import specifier identifiers themselves
        if (!(node.parent && ts.isImportSpecifier(node.parent))) {
          used.add(node.text);
        }
      }
      ts.forEachChild(node, v);
    };
    v(src);

    const moduleSpec = (targetImport.moduleSpecifier as ts.StringLiteral).text;
    const clause = targetImport.importClause!;
    const named = clause.namedBindings as ts.NamedImports;

    // Build the surviving specifiers: keep all non-NS named imports, add flat names.
    const survivors: string[] = [];
    for (const el of named.elements) {
      if (el.name.text === nsName) continue;
      const typePrefix = el.isTypeOnly ? "type " : "";
      const alias = el.propertyName ? `${el.propertyName.text} as ` : "";
      survivors.push(`${typePrefix}${alias}${el.name.text}`);
    }
    for (const u of [...used].sort()) {
      survivors.push(u);
    }

    const defaultPart = clause.name ? `${clause.name.text}, ` : "";
    let newImport: string;
    if (survivors.length === 0) {
      // Whole import becomes empty — drop the statement (side-effect import not intended here).
      newImport = "";
    } else {
      const typeOnlyClause = clause.isTypeOnly ? "type " : "";
      newImport = `import ${typeOnlyClause}${defaultPart}{ ${survivors.join(", ")} } from "${moduleSpec}";`;
    }

    const start = targetImport.getStart(src);
    const end = targetImport.getEnd();
    let out = text.slice(0, start) + newImport + text.slice(end);
    // Clean up a possible leftover blank line if we removed the import entirely.
    if (newImport === "") {
      out = out.slice(0, start) + out.slice(start).replace(/^\n/, "");
    }

    if (write) {
      fs.writeFileSync(file, out);
    }
    changed++;
    console.log(`    ${path.relative(ROOT, file)}: +${used.size} flat names`);
  }
  console.log(`[fix-imports] ${write ? "rewrote" : "would rewrite"} ${changed} files`);
}

main();
