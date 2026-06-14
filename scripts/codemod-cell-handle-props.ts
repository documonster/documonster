/**
 * Codemod: rewrite handle-level cell property access on `CellData` values to
 * the handle-level flat helpers. Type-checker driven so mocks / unrelated
 * objects are left untouched.
 *
 *   <cellExpr>.value        -> cellGetValue(<cellExpr>)
 *   <cellExpr>.value = v     -> cellSetValue(<cellExpr>, v)
 *   <cellExpr>.text          -> cellText(<cellExpr>)
 *   <cellExpr>.type          -> cellType(<cellExpr>)
 *   <cellExpr>.numFmt        -> cellNumFmt / cellSetNumFmt
 *   <cellExpr>.font/fill/... -> cellFont/cellSetFont/...
 *   <cellExpr>.result        -> cellResult / cellSetResult
 *   <cellExpr>.formula       -> cellFormula
 *   <cellExpr>.master        -> cellMaster
 *   <cellExpr>.note          -> cellNote / cellSetNote
 *   <cellExpr>.address       -> (cellExpr).address  (data field — leave)
 *
 * Only fires when the checker says `<cellExpr>` is `CellData`.
 * Run: node scripts/codemod-cell-handle-props.ts [--write] [--glob substr]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

// prop -> [getterFn, setterFn|null]
const PROP_MAP: Record<string, [string, string | null]> = {
  value: ["cellGetValue", "cellSetValue"],
  text: ["cellText", null],
  type: ["cellType", null],
  effectiveType: ["cellEffectiveType", null],
  numFmt: ["cellNumFmt", "cellSetNumFmt"],
  font: ["cellFont", "cellSetFont"],
  fill: ["cellFill", "cellSetFill"],
  border: ["cellBorder", "cellSetBorder"],
  alignment: ["cellAlignment", "cellSetAlignment"],
  protection: ["cellProtection", "cellSetProtection"],
  result: ["cellResult", "cellSetResult"],
  formula: ["cellFormula", null],
  master: ["cellMaster", null],
  note: ["cellNote", "cellSetNote"],
  hyperlink: ["cellHyperlink", null],
  isMerged: ["cellIsMerged", null],
  isHyperlink: ["cellIsHyperlink", null],
  html: ["cellHtml", null],
  formulaType: ["cellFormulaType", null],
  displayText: ["cellDisplayText", null],
  comment: ["cellComment", "cellSetComment"],
  dataValidation: ["cellDataValidation", "cellSetDataValidation"],
  fullAddress: ["cellFullAddress", null],
  names: ["cellNames", "cellSetNames"],
  name: ["cellName", "cellSetName"]
};
// data fields on CellData — never rewrite
const DATA_FIELDS = new Set([
  "address",
  "row",
  "column",
  "style",
  "_value",
  "_mergeCount",
  "_comment"
]);

function loadProgram(): ts.Program {
  const cfgPath = ts.findConfigFile(ROOT, ts.sys.fileExists, "tsconfig.json")!;
  const { config } = ts.readConfigFile(cfgPath, ts.sys.readFile);
  const parsed = ts.parseJsonConfigFileContent(config, ts.sys, path.dirname(cfgPath));
  return ts.createProgram(parsed.fileNames, { ...parsed.options, noEmit: true });
}

function main() {
  const write = process.argv.includes("--write");
  const gi = process.argv.indexOf("--glob");
  const globSub = gi >= 0 ? process.argv[gi + 1] : undefined;

  const program = loadProgram();
  const checker = program.getTypeChecker();

  const isCellData = (expr: ts.Expression): boolean => {
    try {
      const t = checker.getTypeAtLocation(expr);
      const types = t.isUnion() ? t.types : [t];
      return types.some(x => {
        const sym = x.aliasSymbol ?? x.getSymbol();
        return sym?.getName() === "CellData";
      });
    } catch {
      return false;
    }
  };

  let filesChanged = 0;
  const usedHelpers = new Map<string, Set<string>>();

  for (const sf of program.getSourceFiles()) {
    if (sf.isDeclarationFile) continue;
    const file = sf.fileName;
    if (!file.startsWith(path.join(ROOT, "src"))) continue;
    if (!/__tests__/.test(file) && !/\.(test|spec)\.ts$/.test(file) && !file.includes("/examples/"))
      continue;
    if (globSub && !file.includes(globSub)) continue;

    const text = sf.getFullText();
    const edits: { start: number; end: number; repl: string }[] = [];
    const helpers = new Set<string>();

    const visit = (n: ts.Node) => {
      if (
        ts.isPropertyAccessExpression(n) &&
        Object.hasOwn(PROP_MAP, n.name.text) &&
        !DATA_FIELDS.has(n.name.text)
      ) {
        const recv = n.expression;
        // skip if receiver is itself a prop access we don't want (e.g. x.value.foo handled by recursion)
        if (isCellData(recv)) {
          const prop = n.name.text;
          const [getter, setter] = PROP_MAP[prop];
          const recvText = recv.getText(sf);
          const parent = n.parent;
          const isAssign =
            ts.isBinaryExpression(parent) &&
            parent.left === n &&
            parent.operatorToken.kind === ts.SyntaxKind.EqualsToken;
          if (isAssign && setter) {
            const rhs = (parent as ts.BinaryExpression).right.getText(sf);
            edits.push({
              start: parent.getStart(sf),
              end: parent.getEnd(),
              repl: `${setter}(${recvText}, ${rhs})`
            });
            helpers.add(setter);
            ts.forEachChild((parent as ts.BinaryExpression).right, visit);
            return;
          }
          if (!isAssign) {
            edits.push({ start: n.getStart(sf), end: n.getEnd(), repl: `${getter}(${recvText})` });
            helpers.add(getter);
            ts.forEachChild(recv, visit);
            return;
          }
        }
      }
      ts.forEachChild(n, visit);
    };
    visit(sf);

    if (edits.length === 0) continue;
    edits.sort((a, b) => a.start - b.start || a.end - b.end);
    const kept: typeof edits = [];
    let lastEnd = -1;
    for (const e of edits) {
      if (e.start < lastEnd) continue;
      kept.push(e);
      lastEnd = e.end;
    }
    kept.sort((a, b) => b.start - a.start);
    let out = text;
    for (const e of kept) out = out.slice(0, e.start) + e.repl + out.slice(e.end);

    // Ensure the used cell helpers are imported from @excel/cell.
    if (helpers.size > 0) {
      const reSf = ts.createSourceFile(file, out, ts.ScriptTarget.Latest, true);
      const alreadyImported = new Set<string>();
      let cellImport: ts.ImportDeclaration | undefined;
      let firstImportEnd = 0;
      for (const stmt of reSf.statements) {
        if (!ts.isImportDeclaration(stmt) || !ts.isStringLiteral(stmt.moduleSpecifier)) continue;
        firstImportEnd = firstImportEnd || stmt.getEnd();
        const nb = stmt.importClause?.namedBindings;
        if (nb && ts.isNamedImports(nb)) {
          for (const el of nb.elements) alreadyImported.add(el.name.text);
        }
        if (stmt.moduleSpecifier.text === "@excel/cell") cellImport = stmt;
      }
      const missing = [...helpers].filter(h => !alreadyImported.has(h)).sort();
      if (missing.length > 0) {
        if (
          cellImport &&
          cellImport.importClause?.namedBindings &&
          ts.isNamedImports(cellImport.importClause.namedBindings)
        ) {
          const nb = cellImport.importClause.namedBindings;
          const insertAt =
            nb.elements.length > 0
              ? nb.elements[nb.elements.length - 1].getEnd()
              : nb.getStart(reSf) + 1;
          out = out.slice(0, insertAt) + ", " + missing.join(", ") + out.slice(insertAt);
        } else {
          const imp = `\nimport { ${missing.join(", ")} } from "@excel/cell";`;
          out = firstImportEnd
            ? out.slice(0, firstImportEnd) + imp + out.slice(firstImportEnd)
            : `import { ${missing.join(", ")} } from "@excel/cell";\n` + out;
        }
      }
    }

    if (write) fs.writeFileSync(file, out);
    filesChanged++;
    usedHelpers.set(file, helpers);
    console.log(`  ${path.relative(ROOT, file)}: ${kept.length} sites`);
  }
  console.log(`[cell-handle-props] ${write ? "rewrote" : "would rewrite"} ${filesChanged} files`);
}

main();
