/**
 * Follow-up to codemod-tests-to-namespaces: fix imports.
 *
 * For each test/example file that now references the Excel dot-namespaces
 * (Workbook / Worksheet / Cell / Row / Column / Chart / Table / Image / Pivot /
 * Sparkline / Form / Range), ensure a single `import { … } from "@excel/index"`
 * brings in exactly the namespaces used, and remove any stale `Workbook` import
 * from the old root index.
 *
 * Run: node scripts/codemod-fix-ns-test-imports.ts [--write] [--glob substr]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");
const NAMESPACES = [
  "Workbook",
  "Worksheet",
  "Cell",
  "Row",
  "Column",
  "Range",
  "Chart",
  "Table",
  "Image",
  "Pivot",
  "Sparkline",
  "Form"
];

function listFiles(globSub?: string): string[] {
  const out: string[] = [];
  const walk = (dir: string) => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        if (e.name === "node_modules" || e.name === "dist") continue;
        walk(full);
      } else if (
        /\.(test|spec)\.ts$/.test(e.name) ||
        full.includes("/__tests__/") ||
        full.includes("/examples/")
      ) {
        if (!globSub || full.includes(globSub)) out.push(full);
      }
    }
  };
  walk(path.join(ROOT, "src"));
  return out;
}

function main() {
  const write = process.argv.includes("--write");
  const gi = process.argv.indexOf("--glob");
  const globSub = gi >= 0 ? process.argv[gi + 1] : undefined;
  let changed = 0;

  for (const file of listFiles(globSub)) {
    let text = fs.readFileSync(file, "utf8");
    const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);

    // Which namespaces are USED as a value (`Ns.member`) in the file?
    const used = new Set<string>();
    const declared = new Set<string>();
    for (const stmt of sf.statements) {
      // collect local declarations that would shadow a namespace name
      if (ts.isVariableStatement(stmt)) {
        for (const d of stmt.declarationList.declarations) {
          if (ts.isIdentifier(d.name)) declared.add(d.name.text);
        }
      }
      if ((ts.isClassDeclaration(stmt) || ts.isFunctionDeclaration(stmt)) && stmt.name) {
        declared.add(stmt.name.text);
      }
    }
    const visit = (n: ts.Node) => {
      if (
        ts.isPropertyAccessExpression(n) &&
        ts.isIdentifier(n.expression) &&
        NAMESPACES.includes(n.expression.text)
      ) {
        used.add(n.expression.text);
      }
      ts.forEachChild(n, visit);
    };
    visit(sf);

    const needed = [...used].filter(ns => !declared.has(ns)).sort();
    if (needed.length === 0) continue;

    // Remove a stale `import { Workbook } from "...index"` (root or relative).
    // We collect existing imports of namespace names to avoid double-import.
    const importEdits: { start: number; end: number; repl: string }[] = [];
    let anchorEnd = 0;
    for (const stmt of sf.statements) {
      if (!ts.isImportDeclaration(stmt) || !ts.isStringLiteral(stmt.moduleSpecifier)) continue;
      anchorEnd = Math.max(anchorEnd, stmt.getEnd());
      const nb = stmt.importClause?.namedBindings;
      if (nb && ts.isNamedImports(nb)) {
        const spec = stmt.moduleSpecifier.text;
        // Only strip a namespace-named specifier if we are re-adding it from
        // @excel/index (i.e. it's actually used as `Ns.member`). Names like a
        // class `Image` used as `new Image()` are NOT in `needed` → keep them.
        const nsEls = nb.elements.filter(
          el => needed.includes(el.name.text) && spec !== "@excel/index"
        );
        if (nsEls.length > 0) {
          // drop these namespace specifiers from this import (we re-add from @excel/index)
          const survivors = nb.elements.filter(
            el => !(needed.includes(el.name.text) && spec !== "@excel/index")
          );
          const typeOnly = stmt.importClause!.isTypeOnly ? "type " : "";
          const def = stmt.importClause!.name ? `${stmt.importClause!.name.text}, ` : "";
          let repl: string;
          if (survivors.length === 0 && !stmt.importClause!.name) {
            repl = "";
          } else {
            repl = `import ${typeOnly}${def}{ ${survivors
              .map(
                e =>
                  (e.isTypeOnly ? "type " : "") +
                  (e.propertyName ? e.propertyName.text + " as " : "") +
                  e.name.text
              )
              .join(", ")} } from "${spec}";`;
          }
          importEdits.push({ start: stmt.getStart(sf), end: stmt.getEnd(), repl });
        }
      }
    }

    // Apply removals (reverse order)
    importEdits.sort((a, b) => b.start - a.start);
    for (const e of importEdits) {
      let end = e.end;
      if (e.repl === "") while (text[end] === "\n") end++;
      text = text.slice(0, e.start) + e.repl + text.slice(end);
    }

    // Insert/merge the consolidated namespace import.
    const reSf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
    let excelIdx: ts.ImportDeclaration | undefined;
    let firstImport: ts.ImportDeclaration | undefined;
    for (const stmt of reSf.statements) {
      if (ts.isImportDeclaration(stmt) && ts.isStringLiteral(stmt.moduleSpecifier)) {
        firstImport = firstImport ?? stmt;
        if (stmt.moduleSpecifier.text === "@excel/index") excelIdx = excelIdx ?? stmt;
      }
    }
    if (
      excelIdx &&
      excelIdx.importClause?.namedBindings &&
      ts.isNamedImports(excelIdx.importClause.namedBindings)
    ) {
      // merge missing namespaces into the existing @excel/index import
      const nb = excelIdx.importClause.namedBindings;
      const have = new Set(nb.elements.map(e => e.name.text));
      const add = needed.filter(n => !have.has(n));
      if (add.length > 0) {
        const at =
          nb.elements.length > 0
            ? nb.elements[nb.elements.length - 1].getEnd()
            : nb.getStart(reSf) + 1;
        text = text.slice(0, at) + ", " + add.join(", ") + text.slice(at);
      }
    } else {
      const insertPos = firstImport ? firstImport.getEnd() : 0;
      const imp = `\nimport { ${needed.join(", ")} } from "@excel/index";`;
      text =
        insertPos === 0
          ? `import { ${needed.join(", ")} } from "@excel/index";\n` + text
          : text.slice(0, insertPos) + imp + text.slice(insertPos);
    }

    if (write) fs.writeFileSync(file, text);
    changed++;
    console.log(`  ${path.relative(ROOT, file)}: +{ ${needed.join(", ")} }`);
  }
  console.log(`[fix-ns-imports] ${write ? "rewrote" : "would rewrite"} ${changed} files`);
}

main();
