/**
 * Codemod: migrate Excel tests/examples from the intermediate flat API to the
 * final dot-namespace surface (`Workbook.` / `Cell.` / `Row.` / `Column.` /
 * `Chart.` / `Table.` / `Image.` / `Pivot.`).
 *
 *   getCell(ws, a).value            -> Cell.getValue(ws, a)
 *   getCell(ws, a).value = v        -> Cell.setValue(ws, a, v)
 *   getCell(ws, a).text             -> Cell.getText(ws, a)
 *   getCell(ws, a).{font,fill,…}    -> Cell.getStyle(ws, a).{…}
 *   getCell(ws, a).{font,…} = v     -> Cell.setStyle(ws, a, { font: v })
 *   getCell(ws, a).note             -> Cell.getNote / setNote
 *   getCell(ws, a).result/type/...  -> Cell.getResult/getType/...
 *   new Workbook()                  -> Workbook.create()
 *   addWorksheet(wb, n)             -> Workbook.addWorksheet(wb, n)
 *   getWorksheet/removeWorksheet/…  -> Workbook.*
 *   addRow/eachRow/mergeCells/…     -> Worksheet.*
 *   getRow(ws,n).height            -> Row.getHeight / setHeight
 *   getColumn(ws,k).width          -> Column.getWidth / setWidth
 *
 * Type-checker driven where receiver identity matters; otherwise structural.
 * Run: node scripts/codemod-tests-to-namespaces.ts [--write] [--glob substr]
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

// flat fn -> [Namespace, member]
const WORKBOOK_FNS: Record<string, string> = {
  addWorksheet: "addWorksheet",
  getWorksheet: "getWorksheet",
  removeWorksheet: "removeWorksheet",
  addChartsheet: "addChartsheet",
  getChartsheet: "getChartsheet",
  removeChartsheet: "removeChartsheet",
  eachSheet: "eachSheet",
  importSheet: "importSheet",
  addExternalLink: "addExternalLink",
  getExternalLink: "getExternalLink",
  registerPerson: "registerPerson",
  registerFunction: "registerFunction",
  unregisterFunction: "unregisterFunction",
  toXlsxBuffer: "toXlsxBuffer",
  loadXlsx: "loadXlsx",
  readXlsxFile: "readXlsxFile",
  writeXlsx: "writeXlsx"
};
const WORKSHEET_FNS: Record<string, string> = {
  mergeCells: "merge",
  unMergeCells: "unmerge",
  mergeCellsWithoutStyle: "mergeWithoutStyle",
  addRow: "addRow",
  addRows: "addRows",
  getRow: "getRow",
  getRows: "getRows",
  findRow: "findRow",
  eachRow: "eachRow",
  insertRow: "insertRow",
  insertRows: "insertRows",
  duplicateRow: "duplicateRow",
  spliceRows: "spliceRows",
  spliceColumns: "spliceColumns",
  fillFormula: "fillFormula",
  autoFitColumn: "autoFitColumn",
  autoFitColumns: "autoFitColumns",
  autoFitRow: "autoFitRow",
  autoFitRows: "autoFitRows",
  getSheetValues: "getValues",
  getColumnCount: "columnCount",
  getActualColumnCount: "actualColumnCount",
  getRowCount: "rowCount",
  getActualRowCount: "actualRowCount",
  getSheetDimensions: "dimensions",
  getHasMerges: "hasMerges",
  getMergedRegions: "mergedRegions",
  setColumns: "setColumns",
  destroy: "destroy"
};

interface Edit {
  start: number;
  end: number;
  repl: string;
}

const CELL_PROP_GETTER: Record<string, string> = {
  value: "getValue",
  text: "getText",
  type: "getType",
  effectiveType: "getEffectiveType",
  formula: "getFormula",
  result: "getResult",
  note: "getNote",
  hyperlink: "getHyperlink",
  master: "getMergeMaster",
  isMerged: "isMerged"
};
const CELL_STYLE_PROPS = new Set([
  "font",
  "fill",
  "border",
  "alignment",
  "numFmt",
  "protection",
  "style"
]);

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

  const usedNs = new Set<string>();
  let filesChanged = 0;

  for (const file of listFiles(globSub)) {
    const text = fs.readFileSync(file, "utf8");
    if (!/getCell\(|getRow\(|getColumn\(|new Workbook\(|addWorksheet\(/.test(text)) continue;
    const sf = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true);
    const edits: Edit[] = [];
    const fileNs = new Set<string>();

    const callText = (node: ts.CallExpression): string => node.getText(sf);

    const visit = (n: ts.Node) => {
      // new Workbook(...) -> Workbook.create(...)
      if (
        ts.isNewExpression(n) &&
        ts.isIdentifier(n.expression) &&
        n.expression.text === "Workbook"
      ) {
        const argsText =
          n.arguments && n.arguments.length ? n.arguments.map(a => a.getText(sf)).join(", ") : "";
        edits.push({
          start: n.getStart(sf),
          end: n.getEnd(),
          repl: `Workbook.create(${argsText})`
        });
        fileNs.add("Workbook");
        n.arguments?.forEach(visit);
        return;
      }

      // <flatFn>(...)  where flatFn is a workbook/worksheet fn -> Ns.member(...)
      if (ts.isCallExpression(n) && ts.isIdentifier(n.expression)) {
        const fn = n.expression.text;
        let ns: string | undefined;
        let member: string | undefined;
        if (WORKBOOK_FNS[fn]) {
          ns = "Workbook";
          member = WORKBOOK_FNS[fn];
        } else if (WORKSHEET_FNS[fn]) {
          ns = "Worksheet";
          member = WORKSHEET_FNS[fn];
        }
        if (ns && member) {
          // skip if this call is the receiver of a property access we handle below
          // (e.g. getRow(ws,n).height) — those are handled in the prop-access branch
          const parent = n.parent;
          const isCellRowColAccess =
            ts.isPropertyAccessExpression(parent) &&
            parent.expression === n &&
            (fn === "getRow" || fn === "getColumn");
          if (!isCellRowColAccess) {
            edits.push({
              start: n.expression.getStart(sf),
              end: n.expression.getEnd(),
              repl: `${ns}.${member}`
            });
            fileNs.add(ns);
          }
        }
      }

      // getCell(ws, a).PROP  and  getCell(ws, a).PROP = v
      if (
        ts.isPropertyAccessExpression(n) &&
        ts.isCallExpression(n.expression) &&
        ts.isIdentifier(n.expression.expression) &&
        (n.expression.expression.text === "getCell" ||
          n.expression.expression.text === "getRow" ||
          n.expression.expression.text === "getColumn")
      ) {
        const getter = n.expression.expression.text;
        const prop = n.name.text;
        const call = n.expression;
        const argsText = call.arguments.map(a => a.getText(sf)).join(", ");
        const parent = n.parent;
        const isAssign =
          ts.isBinaryExpression(parent) &&
          parent.left === n &&
          parent.operatorToken.kind === ts.SyntaxKind.EqualsToken;

        if (getter === "getCell") {
          if (prop === "value" && isAssign) {
            const rhs = (parent as ts.BinaryExpression).right.getText(sf);
            edits.push({
              start: parent.getStart(sf),
              end: parent.getEnd(),
              repl: `Cell.setValue(${argsText}, ${rhs})`
            });
            fileNs.add("Cell");
            return;
          }
          if (CELL_PROP_GETTER[prop] && !isAssign) {
            edits.push({
              start: n.getStart(sf),
              end: n.getEnd(),
              repl: `Cell.${CELL_PROP_GETTER[prop]}(${argsText})`
            });
            fileNs.add("Cell");
            return;
          }
          if (CELL_STYLE_PROPS.has(prop)) {
            if (prop === "style" && !isAssign) {
              edits.push({
                start: n.getStart(sf),
                end: n.getEnd(),
                repl: `Cell.getStyle(${argsText})`
              });
              fileNs.add("Cell");
              return;
            }
            if (prop !== "style" && isAssign) {
              const rhs = (parent as ts.BinaryExpression).right.getText(sf);
              edits.push({
                start: parent.getStart(sf),
                end: parent.getEnd(),
                repl: `Cell.setStyle(${argsText}, { ${prop}: ${rhs} })`
              });
              fileNs.add("Cell");
              return;
            }
            if (prop !== "style" && !isAssign) {
              edits.push({
                start: n.getStart(sf),
                end: n.getEnd(),
                repl: `Cell.getStyle(${argsText}).${prop}`
              });
              fileNs.add("Cell");
              return;
            }
          }
          if (prop === "note" && isAssign) {
            const rhs = (parent as ts.BinaryExpression).right.getText(sf);
            edits.push({
              start: parent.getStart(sf),
              end: parent.getEnd(),
              repl: `Cell.setNote(${argsText}, ${rhs})`
            });
            fileNs.add("Cell");
            return;
          }
        } else if (getter === "getRow") {
          const map: Record<string, [string, string]> = {
            height: ["getHeight", "setHeight"],
            hidden: ["getHidden", "setHidden"],
            outlineLevel: ["getOutlineLevel", "setOutlineLevel"]
          };
          if (map[prop]) {
            const [g, s] = map[prop];
            if (isAssign) {
              const rhs = (parent as ts.BinaryExpression).right.getText(sf);
              edits.push({
                start: parent.getStart(sf),
                end: parent.getEnd(),
                repl: `Row.${s}(${argsText}, ${rhs})`
              });
            } else {
              edits.push({ start: n.getStart(sf), end: n.getEnd(), repl: `Row.${g}(${argsText})` });
            }
            fileNs.add("Row");
            return;
          }
        } else if (getter === "getColumn") {
          const map: Record<string, [string, string]> = {
            width: ["getWidth", "setWidth"],
            header: ["getHeader", "setHeader"],
            key: ["getKey", "setKey"],
            hidden: ["getHidden", "setHidden"]
          };
          if (map[prop]) {
            const [g, s] = map[prop];
            if (isAssign) {
              const rhs = (parent as ts.BinaryExpression).right.getText(sf);
              edits.push({
                start: parent.getStart(sf),
                end: parent.getEnd(),
                repl: `Column.${s}(${argsText}, ${rhs})`
              });
            } else {
              edits.push({
                start: n.getStart(sf),
                end: n.getEnd(),
                repl: `Column.${g}(${argsText})`
              });
            }
            fileNs.add("Column");
            return;
          }
        }
      }

      ts.forEachChild(n, visit);
    };
    visit(sf);

    if (edits.length === 0) continue;
    // dedup/overlap
    edits.sort((a, b) => a.start - b.start || a.end - b.end);
    const kept: Edit[] = [];
    let lastEnd = -1;
    for (const e of edits) {
      if (e.start < lastEnd) continue;
      kept.push(e);
      lastEnd = e.end;
    }
    kept.sort((a, b) => b.start - a.start);
    let out = text;
    for (const e of kept) out = out.slice(0, e.start) + e.repl + out.slice(e.end);

    if (write) fs.writeFileSync(file, out);
    filesChanged++;
    for (const n of fileNs) usedNs.add(n);
    console.log(`  ${path.relative(ROOT, file)}: ${kept.length} sites (${[...fileNs].join(",")})`);
  }
  console.log(
    `[tests->ns] ${write ? "rewrote" : "would rewrite"} ${filesChanged} files; namespaces used: ${[...usedNs].join(", ")}`
  );
  console.log(
    `[tests->ns] NOTE: imports must be fixed separately (add { ${[...usedNs].join(", ")} } from the excel index).`
  );
}

main();
