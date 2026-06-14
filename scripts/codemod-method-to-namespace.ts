/**
 * Codemod: rewrite remaining instance-method calls on Workbook/Worksheet
 * handles (`wb.addWorksheet(...)`, `ws.addTable(...)`, `ws.toJSON(...)`) to the
 * dot-namespace surface. Type-checker driven (receiver must be
 * WorkbookData/Worksheet/WorksheetData).
 *
 * Run: node scripts/codemod-method-to-namespace.ts [--write] [--glob substr]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

// member -> [Namespace, member]
const WORKBOOK: Record<string, [string, string]> = {
  addWorksheet: ["Workbook", "addWorksheet"],
  getWorksheet: ["Workbook", "getWorksheet"],
  removeWorksheet: ["Workbook", "removeWorksheet"],
  addChartsheet: ["Workbook", "addChartsheet"],
  getChartsheet: ["Workbook", "getChartsheet"],
  removeChartsheet: ["Workbook", "removeChartsheet"],
  eachSheet: ["Workbook", "eachSheet"],
  importSheet: ["Workbook", "importSheet"],
  addExternalLink: ["Workbook", "addExternalLink"],
  getExternalLink: ["Workbook", "getExternalLink"],
  registerPerson: ["Workbook", "registerPerson"],
  registerFunction: ["Workbook", "registerFunction"],
  unregisterFunction: ["Workbook", "unregisterFunction"]
};
const WORKSHEET: Record<string, [string, string]> = {
  mergeCells: ["Worksheet", "merge"],
  unMergeCells: ["Worksheet", "unmerge"],
  addRow: ["Worksheet", "addRow"],
  addRows: ["Worksheet", "addRows"],
  getRow: ["Worksheet", "getRow"],
  getRows: ["Worksheet", "getRows"],
  findRow: ["Worksheet", "findRow"],
  eachRow: ["Worksheet", "eachRow"],
  insertRow: ["Worksheet", "insertRow"],
  insertRows: ["Worksheet", "insertRows"],
  duplicateRow: ["Worksheet", "duplicateRow"],
  spliceRows: ["Worksheet", "spliceRows"],
  spliceColumns: ["Worksheet", "spliceColumns"],
  fillFormula: ["Worksheet", "fillFormula"],
  autoFitColumn: ["Worksheet", "autoFitColumn"],
  autoFitColumns: ["Worksheet", "autoFitColumns"],
  autoFitRow: ["Worksheet", "autoFitRow"],
  autoFitRows: ["Worksheet", "autoFitRows"],
  getSheetValues: ["Worksheet", "getValues"],
  toJSON: ["Worksheet", "toJson"],
  addJSON: ["Worksheet", "addJson"],
  toAOA: ["Worksheet", "toAoa"],
  addAOA: ["Worksheet", "addAoa"],
  addTable: ["Table", "add"],
  getTable: ["Table", "get"],
  getTables: ["Table", "list"],
  removeTable: ["Table", "remove"],
  addImage: ["Image", "place"],
  getImages: ["Image", "list"],
  addBackgroundImage: ["Image", "setBackground"],
  addPivotTable: ["Pivot", "add"],
  addSparklineGroup: ["Sparkline", "add"],
  getSparklineGroups: ["Sparkline", "list"],
  removeSparklineGroup: ["Sparkline", "remove"],
  addFormCheckbox: ["Form", "addCheckbox"],
  getFormCheckboxes: ["Form", "listCheckboxes"],
  addConditionalFormatting: ["Worksheet", "addConditionalFormatting"],
  removeConditionalFormatting: ["Worksheet", "removeConditionalFormatting"],
  addChart: ["Chart", "add"],
  addColumnChart: ["Chart", "addColumn"],
  addBarChart: ["Chart", "addBar"],
  addLineChart: ["Chart", "addLine"],
  addAreaChart: ["Chart", "addArea"],
  addPieChart: ["Chart", "addPie"],
  addScatterChart: ["Chart", "addScatter"],
  getCharts: ["Chart", "get"],
  removeChart: ["Chart", "remove"]
};

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

  const typeName = (expr: ts.Expression): string[] => {
    try {
      const t = checker.getTypeAtLocation(expr);
      const types = t.isUnion() ? t.types : [t];
      return types.map(x => (x.aliasSymbol ?? x.getSymbol())?.getName() ?? "");
    } catch {
      return [];
    }
  };
  // Heuristic: receiver identifier names that denote a worksheet/workbook handle
  // even when typed `any` (test helpers). Member name disambiguates the namespace.
  const WS_NAMES =
    /^(ws|ws\d|sheet|worksheet|s\d|wsResult|wsSquare|wsSingles|summary|revenue|expenses|data|cm|pivot|out)$/;
  const WB_NAMES = /^(wb|wb\d|workbook|loaded|clone|target|source|out)$/;

  let filesChanged = 0;
  for (const sf of program.getSourceFiles()) {
    if (sf.isDeclarationFile) continue;
    const file = sf.fileName;
    if (!file.startsWith(path.join(ROOT, "src"))) continue;
    if (!/__tests__/.test(file) && !/\.(test|spec)\.ts$/.test(file) && !file.includes("/examples/"))
      continue;
    if (globSub && !file.includes(globSub)) continue;

    const text = sf.getFullText();
    const edits: { start: number; end: number; repl: string }[] = [];
    const nsUsed = new Set<string>();

    const visit = (n: ts.Node) => {
      if (
        ts.isCallExpression(n) &&
        ts.isPropertyAccessExpression(n.expression) &&
        ts.isIdentifier(n.expression.name)
      ) {
        const member = n.expression.name.text;
        const recv = n.expression.expression;
        const tnames = typeName(recv);
        const recvIsAny = tnames.length === 0 || tnames.some(t => t === "" || t === "any");
        const recvId = ts.isIdentifier(recv) ? recv.text : "";
        let mapping: [string, string] | undefined;
        if (
          Object.hasOwn(WORKBOOK, member) &&
          (tnames.some(t => t === "WorkbookData" || t === "Workbook") ||
            (recvIsAny && WB_NAMES.test(recvId)))
        ) {
          mapping = WORKBOOK[member];
        } else if (
          Object.hasOwn(WORKSHEET, member) &&
          (tnames.some(t => t === "WorksheetData" || t === "Worksheet") ||
            (recvIsAny && WS_NAMES.test(recvId)))
        ) {
          mapping = WORKSHEET[member];
        }
        if (mapping) {
          const [ns, mem] = mapping;
          const recvText = recv.getText(sf);
          const open = n.getChildren(sf).find(c => c.kind === ts.SyntaxKind.OpenParenToken)!;
          const hasArgs = n.arguments.length > 0;
          // replace `recv.member(` with `Ns.mem(recv, `
          edits.push({
            start: n.expression.getStart(sf),
            end: open.getEnd(),
            repl: `${ns}.${mem}(${recvText}${hasArgs ? ", " : ""}`
          });
          nsUsed.add(ns);
          n.arguments.forEach(visit);
          return;
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

    // ensure namespace imports
    const reSf = ts.createSourceFile(file, out, ts.ScriptTarget.Latest, true);
    const importedNs = new Set<string>();
    let excelIdxImport: ts.ImportDeclaration | undefined;
    let firstImportEnd = 0;
    for (const stmt of reSf.statements) {
      if (!ts.isImportDeclaration(stmt) || !ts.isStringLiteral(stmt.moduleSpecifier)) continue;
      firstImportEnd = firstImportEnd || stmt.getEnd();
      if (stmt.moduleSpecifier.text === "@excel/index") {
        excelIdxImport = stmt;
        const nb = stmt.importClause?.namedBindings;
        if (nb && ts.isNamedImports(nb)) for (const el of nb.elements) importedNs.add(el.name.text);
      }
    }
    const missingNs = [...nsUsed].filter(x => !importedNs.has(x)).sort();
    if (missingNs.length > 0) {
      if (
        excelIdxImport &&
        excelIdxImport.importClause?.namedBindings &&
        ts.isNamedImports(excelIdxImport.importClause.namedBindings)
      ) {
        const nb = excelIdxImport.importClause.namedBindings;
        const at =
          nb.elements.length > 0
            ? nb.elements[nb.elements.length - 1].getEnd()
            : nb.getStart(reSf) + 1;
        out = out.slice(0, at) + ", " + missingNs.join(", ") + out.slice(at);
      } else {
        const imp = `\nimport { ${missingNs.join(", ")} } from "@excel/index";`;
        out = firstImportEnd
          ? out.slice(0, firstImportEnd) + imp + out.slice(firstImportEnd)
          : `import { ${missingNs.join(", ")} } from "@excel/index";\n` + out;
      }
    }

    if (write) fs.writeFileSync(file, out);
    filesChanged++;
    console.log(`  ${path.relative(ROOT, file)}: ${kept.length} sites (${[...nsUsed].join(",")})`);
  }
  console.log(`[method->ns] ${write ? "rewrote" : "would rewrite"} ${filesChanged} files`);
}

main();
