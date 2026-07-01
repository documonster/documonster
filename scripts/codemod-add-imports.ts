/**
 * Add missing named imports for flat functions exported by a given module.
 *
 * After call-site rewrites turn `ws.getColumn(c)` into `getColumn(ws, c)`, the
 * consumer file references `getColumn` but doesn't import it. This script, for
 * each file, finds identifiers that (a) are exported by the target module and
 * (b) are used but not declared/imported locally, and adds them to (or creates)
 * an import from the target module.
 *
 * Usage: node scripts/codemod-add-imports.ts <defFileRel> <moduleSpecifier> [--write]
 *   e.g. node scripts/codemod-add-imports.ts src/modules/excel/worksheet.ts @excel/worksheet --write
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

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

function exportedFns(defPath: string): Set<string> {
  const text = fs.readFileSync(defPath, "utf8");
  const sf = ts.createSourceFile(defPath, text, ts.ScriptTarget.Latest, true);
  const names = new Set<string>();
  const visit = (n: ts.Node) => {
    if (
      ts.isFunctionDeclaration(n) &&
      n.name &&
      ts.getModifiers(n)?.some(m => m.kind === ts.SyntaxKind.ExportKeyword)
    ) {
      names.add(n.name.text);
    }
    ts.forEachChild(n, visit);
  };
  visit(sf);
  return names;
}

function main() {
  const defRel = process.argv[2];
  const moduleSpec = process.argv[3];
  const write = process.argv.includes("--write");
  if (!defRel || !moduleSpec) {
    console.error("usage: codemod-add-imports.ts <defFileRel> <moduleSpecifier> [--write]");
    process.exit(1);
  }
  const defPath = path.join(ROOT, defRel);
  const exported = exportedFns(defPath);
  console.log(`[add-imports] ${moduleSpec}: ${exported.size} exported fns`);

  let changed = 0;
  for (const file of listSourceFiles()) {
    if (path.resolve(file) === path.resolve(defPath)) continue;
    const text = fs.readFileSync(file, "utf8");
    const src = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);

    // Collect: locally declared names, already-imported names, used identifiers.
    const declared = new Set<string>();
    const importedFrom = new Map<string, ts.ImportDeclaration>(); // module -> import
    const alreadyImported = new Set<string>();
    let existingTargetImport: ts.ImportDeclaration | undefined;

    for (const stmt of src.statements) {
      if (ts.isImportDeclaration(stmt) && ts.isStringLiteral(stmt.moduleSpecifier)) {
        const spec = stmt.moduleSpecifier.text;
        if (spec === moduleSpec) existingTargetImport = stmt;
        const nb = stmt.importClause?.namedBindings;
        if (nb && ts.isNamedImports(nb)) {
          for (const el of nb.elements) alreadyImported.add(el.name.text);
        }
        if (stmt.importClause?.name) alreadyImported.add(stmt.importClause.name.text);
      }
      if (ts.isFunctionDeclaration(stmt) && stmt.name) declared.add(stmt.name.text);
      if (ts.isVariableStatement(stmt)) {
        for (const d of stmt.declarationList.declarations) {
          if (ts.isIdentifier(d.name)) declared.add(d.name.text);
        }
      }
    }

    // Find used exported names that are neither declared nor imported.
    const needed = new Set<string>();
    const visit = (n: ts.Node) => {
      if (
        ts.isIdentifier(n) &&
        exported.has(n.text) &&
        !declared.has(n.text) &&
        !alreadyImported.has(n.text)
      ) {
        // exclude property names (obj.foo) and import specifiers
        const p = n.parent;
        const isProp = ts.isPropertyAccessExpression(p) && p.name === n;
        const isImportSpec = ts.isImportSpecifier(p);
        const isPropAssign =
          (ts.isPropertyAssignment(p) || ts.isPropertySignature(p)) && p.name === n;
        if (!isProp && !isImportSpec && !isPropAssign) {
          needed.add(n.text);
        }
      }
      ts.forEachChild(n, visit);
    };
    visit(src);

    if (needed.size === 0) continue;
    const sorted = [...needed].sort();

    let out: string;
    if (existingTargetImport) {
      const nb = existingTargetImport.importClause?.namedBindings;
      if (nb && ts.isNamedImports(nb)) {
        // splice new names into existing braces
        const insertPos =
          nb.elements.length > 0
            ? nb.elements[nb.elements.length - 1].getEnd()
            : nb.getStart(src) + 1;
        out = text.slice(0, insertPos) + ", " + sorted.join(", ") + text.slice(insertPos);
      } else {
        // import has default or namespace only — append a separate import
        const end = existingTargetImport.getEnd();
        out =
          text.slice(0, end) +
          `\nimport { ${sorted.join(", ")} } from "${moduleSpec}";` +
          text.slice(end);
      }
    } else {
      // add a new import after the last import statement (or at top)
      const imports = src.statements.filter(ts.isImportDeclaration);
      const pos = imports.length > 0 ? imports[imports.length - 1].getEnd() : 0;
      const newImp = `\nimport { ${sorted.join(", ")} } from "${moduleSpec}";`;
      out =
        pos === 0
          ? `import { ${sorted.join(", ")} } from "${moduleSpec}";\n` + text
          : text.slice(0, pos) + newImp + text.slice(pos);
    }

    if (write) fs.writeFileSync(file, out);
    changed++;
    console.log(
      `    ${path.relative(ROOT, file)}: +${sorted.length} (${sorted.slice(0, 6).join(", ")}${sorted.length > 6 ? "…" : ""})`
    );
  }
  console.log(`[add-imports] ${write ? "rewrote" : "would rewrite"} ${changed} files`);
}

main();
