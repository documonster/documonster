/**
 * Codemod: rewrite member access on instances of a de-classed type to flat
 * function calls, using the TypeScript type checker to identify receivers.
 *
 *   ws.getCell(a, b)   -> getCell(ws, a, b)        (method)
 *   ws.name            -> getSheetName(ws)          (getter)
 *   ws.name = v        -> setSheetName(ws, v)       (setter)
 *   ws.field           -> ws.field                  (left alone — public data field)
 *
 * The set of method/getter/setter names + their flat targets is read from the
 * definition file's exported functions and a name map identical to
 * codemod-class-to-flat.ts.
 *
 * A receiver expression is rewritten only when the checker says its type is
 * (or includes) the target record type (e.g. `WorksheetData`).
 *
 * Usage: node scripts/codemod-class-callsites.ts <Class> [--write] [--glob <substr>]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

interface Config {
  recordType: string; // e.g. "WorksheetData"
  /** legacy class name alias also accepted as the record (e.g. "Worksheet") */
  aliasType: string;
  /** relative path to the def file whose exported fns are the method set */
  defFile: string;
  /** getter member -> flat fn */
  getterMap: Record<string, string>;
  /** setter member -> flat fn */
  setterMap: Record<string, string>;
  /** method members that are flat fns keeping their name */
  methodNames: Set<string>;
  /** members that remain plain data fields (do NOT rewrite) */
  dataFields: Set<string>;
}

// Built from worksheet.ts. Method names = exported functions taking ws first.
const WORKSHEET_GETTERS: Record<string, string> = {
  name: "getSheetName",
  workbook: "getSheetWorkbook",
  dimensions: "getSheetDimensions",
  columns: "getColumns",
  lastColumn: "getLastColumn",
  columnCount: "getColumnCount",
  actualColumnCount: "getActualColumnCount",
  lastRow: "getLastRow",
  rowCount: "getRowCount",
  actualRowCount: "getActualRowCount",
  hasMerges: "getHasMerges",
  mergedRegions: "getMergedRegions",
  sparklineGroups: "getSparklineGroups2",
  model: "getSheetModel"
};
const WORKSHEET_SETTERS: Record<string, string> = {
  name: "setSheetName",
  columns: "setColumns",
  sparklineGroups: "setSparklineGroups2",
  model: "setSheetModel"
};

const CONFIGS: Record<string, Config> = {
  Worksheet: {
    recordType: "WorksheetData",
    aliasType: "Worksheet",
    defFile: "src/modules/excel/worksheet.ts",
    getterMap: WORKSHEET_GETTERS,
    setterMap: WORKSHEET_SETTERS,
    // populated from def file below
    methodNames: new Set<string>(),
    // public data fields on WorksheetData — accessed directly, never rewritten
    dataFields: new Set<string>([
      "id",
      "orderNo",
      "state",
      "rowBreaks",
      "colBreaks",
      "properties",
      "pageSetup",
      "headerFooter",
      "dataValidations",
      "views",
      "autoFilter",
      "sheetProtection",
      "tables",
      "pivotTables",
      "conditionalFormattings",
      "formControls",
      "ignoredErrors",
      "threadedComments",
      // private underscore fields
      "_workbook",
      "_name",
      "_rows",
      "_columns",
      "_keys",
      "_merges",
      "_media",
      "_shapes",
      "_charts",
      "_sparklineGroups",
      "_headerRowCount",
      "_drawing",
      "_watermark"
    ])
  },
  Workbook: {
    recordType: "WorkbookData",
    aliasType: "Workbook",
    defFile: "src/modules/excel/workbook.browser.ts",
    getterMap: {
      defaultFont: "getDefaultFont",
      xlsx: "getXlsxIo",
      nextId: "getNextId",
      worksheets: "getWorksheets",
      chartsheets: "getChartsheets",
      definedNames: "getDefinedNames",
      persons: "getPersons",
      model: "getWorkbookModel"
    },
    setterMap: {
      defaultFont: "setDefaultFont",
      model: "setWorkbookModel"
    },
    methodNames: new Set<string>(),
    // public data fields on WorkbookData — accessed directly, never rewritten
    dataFields: new Set<string>([
      "category",
      "company",
      "created",
      "description",
      "keywords",
      "manager",
      "modified",
      "subject",
      "title",
      "creator",
      "lastModifiedBy",
      "lastPrinted",
      "language",
      "revision",
      "contentStatus",
      "properties",
      "calcProperties",
      "views",
      "media",
      "pivotTables",
      "protection",
      "externalLinks",
      "_worksheets",
      "_definedNames",
      "_themes",
      "_defaultFont",
      "_writerExternalLinkCache",
      "_tableNames",
      "_chartEntries",
      "_chartRels"
    ])
  }
};

function collectExportedFns(defPath: string): Set<string> {
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

function loadProgram(): ts.Program {
  const configPath = ts.findConfigFile(ROOT, ts.sys.fileExists, "tsconfig.json")!;
  const { config } = ts.readConfigFile(configPath, ts.sys.readFile);
  const parsed = ts.parseJsonConfigFileContent(config, ts.sys, path.dirname(configPath));
  return ts.createProgram(parsed.fileNames, { ...parsed.options, noEmit: true });
}

function main() {
  const className = process.argv[2];
  const write = process.argv.includes("--write");
  const globIdx = process.argv.indexOf("--glob");
  const globSub = globIdx >= 0 ? process.argv[globIdx + 1] : undefined;
  const cfg = CONFIGS[className];
  if (!cfg) {
    console.error(`Unknown class ${className}`);
    process.exit(1);
  }

  // Method names = exported fns minus the getters/setters/create.
  const defPath = path.join(ROOT, cfg.defFile);
  const exportedFns = collectExportedFns(defPath);
  const nonMethods = new Set([
    ...Object.values(cfg.getterMap),
    ...Object.values(cfg.setterMap),
    `create${className}`
  ]);
  for (const fn of exportedFns) {
    if (!nonMethods.has(fn)) cfg.methodNames.add(fn);
  }
  console.log(`[callsites] ${className}: ${cfg.methodNames.size} method names`);

  const program = loadProgram();
  const checker = program.getTypeChecker();

  const isRecordReceiver = (expr: ts.Expression): boolean => {
    const type = checker.getTypeAtLocation(expr);
    const types = type.isUnion() ? type.types : [type];
    for (const t of types) {
      const sym = t.aliasSymbol ?? t.getSymbol();
      const name = sym?.getName();
      if (name === cfg.recordType || name === cfg.aliasType) return true;
    }
    return false;
  };

  let filesChanged = 0;
  for (const sf of program.getSourceFiles()) {
    if (sf.isDeclarationFile) continue;
    const file = sf.fileName;
    if (!file.startsWith(path.join(ROOT, "src"))) continue;
    if (globSub && !file.includes(globSub)) continue;
    // Skip the def file itself.
    if (path.resolve(file) === path.resolve(defPath)) continue;

    const text = sf.getFullText();
    const edits: { start: number; end: number; repl: string }[] = [];

    const visit = (n: ts.Node) => {
      // setter assignment: recv.setter = value
      if (
        ts.isBinaryExpression(n) &&
        n.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
        ts.isPropertyAccessExpression(n.left) &&
        cfg.setterMap[n.left.name.text] &&
        isRecordReceiver(n.left.expression)
      ) {
        const fn = cfg.setterMap[n.left.name.text];
        const recv = n.left.expression.getText(sf);
        edits.push({
          start: n.left.getStart(sf),
          end: n.operatorToken.getEnd(),
          repl: `${fn}(${recv},`
        });
        edits.push({ start: n.right.getEnd(), end: n.right.getEnd(), repl: ")" });
        ts.forEachChild(n.right, visit);
        return;
      }
      if (ts.isPropertyAccessExpression(n)) {
        const member = n.name.text;
        const recvExpr = n.expression;
        // method call: recv.method(args)
        if (
          cfg.methodNames.has(member) &&
          ts.isCallExpression(n.parent) &&
          n.parent.expression === n &&
          isRecordReceiver(recvExpr)
        ) {
          const call = n.parent;
          const recv = recvExpr.getText(sf);
          const hasArgs = call.arguments.length > 0;
          const openParen = call
            .getChildren(sf)
            .find(c => c.kind === ts.SyntaxKind.OpenParenToken)!;
          edits.push({
            start: n.getStart(sf),
            end: openParen.getEnd(),
            repl: `${member}(${recv}${hasArgs ? ", " : ""}`
          });
          for (const a of call.arguments) visit(a);
          return;
        }
        // getter read: recv.getter  (not part of a setter assignment, handled above)
        if (cfg.getterMap[member] && isRecordReceiver(recvExpr)) {
          // skip if this is the LHS of an assignment (setter) — handled above
          const isAssignTarget =
            ts.isBinaryExpression(n.parent) &&
            n.parent.left === n &&
            n.parent.operatorToken.kind === ts.SyntaxKind.EqualsToken;
          if (!isAssignTarget) {
            const fn = cfg.getterMap[member];
            const recv = recvExpr.getText(sf);
            edits.push({ start: n.getStart(sf), end: n.getEnd(), repl: `${fn}(${recv})` });
            ts.forEachChild(recvExpr, visit);
            return;
          }
        }
      }
      ts.forEachChild(n, visit);
    };
    visit(sf);

    if (edits.length === 0) continue;
    // dedup/overlap guard
    edits.sort((a, b) => a.start - b.start || a.end - b.end);
    const kept: typeof edits = [];
    let lastEnd = -1;
    for (const e of edits) {
      if (e.start === e.end) {
        kept.push(e);
        continue;
      }
      if (e.start < lastEnd) continue;
      kept.push(e);
      lastEnd = e.end;
    }
    kept.sort((a, b) => b.start - a.start);
    let out = text;
    for (const e of kept) {
      out = out.slice(0, e.start) + e.repl + out.slice(e.end);
    }
    if (write) fs.writeFileSync(file, out);
    filesChanged++;
    console.log(`    ${path.relative(ROOT, file)}: ${kept.length} sites`);
  }
  console.log(`[callsites] ${write ? "rewrote" : "would rewrite"} ${filesChanged} files`);
}

main();
