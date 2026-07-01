/**
 * Codemod: convert a single class into a plain-data record interface + flat
 * named exports. Tailored for the excel Worksheet/Workbook de-classing.
 *
 * Transforms within the DEFINITION file:
 *   - `class C { fieldDecls; constructor(o){...}; method(args){...}; get x(){}; set x(v){} }`
 *     into:
 *       interface CData { ...fields... }
 *       export function create<C>(o): CData { ... }   // from constructor
 *       export function <method>(self: CData, ...args) { ... }   // public methods keep name
 *       export function _<priv>(self: CData, ...args) { ... }    // private methods get _ prefix kept
 *       export function get<Getter>(self): T { ... }  // getters
 *       export function set<Setter>(self, v): void {} // setters
 *   - inside bodies: `this.field` -> `self.field`; `this.method(a)` -> `method(self, a)`;
 *     `this.getter` -> `getGetter(self)`; `this.setter = v` -> `setSetter(self, v)`.
 *
 * The method/getter NAME MAP is supplied per class (see CLASS_CONFIG). Public
 * methods keep their identifier; getters/setters map to get/set<Name>.
 *
 * Call-site rewriting across the repo is handled by a SEPARATE codemod
 * (codemod-class-callsites.ts) because it needs type info to know the receiver
 * is an instance of the class.
 *
 * Usage: node scripts/codemod-class-to-flat.ts <Class> [--write]
 *
 * This script only rewrites the DEFINITION file and prints the name map.
 */
import * as fs from "node:fs";
import * as path from "node:path";

import ts from "typescript";

const ROOT = path.resolve(import.meta.dirname, "..");

interface ClassConfig {
  defFile: string;
  /** name of the `create*` function generated from the constructor */
  createName: string;
  /** the self parameter name + type used in generated functions */
  selfParam: string;
  selfType: string;
  /** explicit name overrides for getters/setters/methods (member -> flatName) */
  getterMap: Record<string, string>;
  setterMap: Record<string, string>;
  methodMap: Record<string, string>;
}

const CLASS_CONFIG: Record<string, ClassConfig> = {
  Worksheet: {
    defFile: "src/modules/excel/worksheet.ts",
    createName: "createWorksheet",
    selfParam: "ws",
    selfType: "WorksheetData",
    getterMap: {
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
    },
    setterMap: {
      name: "setSheetName",
      columns: "setColumns",
      sparklineGroups: "setSparklineGroups2",
      model: "setSheetModel"
    },
    methodMap: {}
  },
  Workbook: {
    defFile: "src/modules/excel/workbook.browser.ts",
    createName: "createWorkbook",
    selfParam: "wb",
    selfType: "WorkbookData",
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
    methodMap: {}
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

function main() {
  const className = process.argv[2];
  const write = process.argv.includes("--write");
  const cfg = CLASS_CONFIG[className];
  if (!cfg) {
    console.error(`Unknown class: ${className}`);
    process.exit(1);
  }

  const defPath = path.join(ROOT, cfg.defFile);
  const text = fs.readFileSync(defPath, "utf8");
  const sf = ts.createSourceFile(defPath, text, ts.ScriptTarget.Latest, true);

  let classNode: ts.ClassDeclaration | undefined;
  const find = (n: ts.Node) => {
    if (ts.isClassDeclaration(n) && n.name?.text === className) {
      classNode = n;
    }
    ts.forEachChild(n, find);
  };
  find(sf);
  if (!classNode) {
    console.error(`class ${className} not found in ${cfg.defFile}`);
    process.exit(1);
  }

  // Collect member categories. Static members are skipped (handled manually).
  const fields: ts.PropertyDeclaration[] = [];
  const methods: ts.MethodDeclaration[] = [];
  const getters: ts.GetAccessorDeclaration[] = [];
  const setters: ts.SetAccessorDeclaration[] = [];
  let ctor: ts.ConstructorDeclaration | undefined;
  const isStatic = (m: ts.ClassElement): boolean =>
    ts.canHaveModifiers(m) &&
    (ts.getModifiers(m)?.some(x => x.kind === ts.SyntaxKind.StaticKeyword) ?? false);
  for (const m of classNode.members) {
    if (isStatic(m)) continue;
    if (ts.isPropertyDeclaration(m)) fields.push(m);
    else if (ts.isMethodDeclaration(m)) methods.push(m);
    else if (ts.isGetAccessor(m)) getters.push(m);
    else if (ts.isSetAccessor(m)) setters.push(m);
    else if (ts.isConstructorDeclaration(m)) ctor = m;
  }

  console.log(
    `[class->flat] ${className}: ${fields.length} fields, ${methods.length} methods, ${getters.length} getters, ${setters.length} setters`
  );

  const methodName = (m: ts.MethodDeclaration): string => {
    const n = (m.name as ts.Identifier).text;
    return cfg.methodMap[n] ?? n;
  };
  const getterName = (g: ts.GetAccessorDeclaration): string => {
    const n = (g.name as ts.Identifier).text;
    return cfg.getterMap[n] ?? `get${n.charAt(0).toUpperCase()}${n.slice(1)}`;
  };
  const setterName = (s: ts.SetAccessorDeclaration): string => {
    const n = (s.name as ts.Identifier).text;
    return cfg.setterMap[n] ?? `set${n.charAt(0).toUpperCase()}${n.slice(1)}`;
  };

  // Build the member-name -> kind maps for `this.X` rewriting.
  const fieldNames = new Set(fields.map(f => (f.name as ts.Identifier).text));
  const methodNames = new Map<string, string>();
  for (const m of methods) methodNames.set((m.name as ts.Identifier).text, methodName(m));
  const getterNames = new Map<string, string>();
  for (const g of getters) getterNames.set((g.name as ts.Identifier).text, getterName(g));
  const setterNames = new Map<string, string>();
  for (const s of setters) setterNames.set((s.name as ts.Identifier).text, setterName(s));

  console.log("[class->flat] getter map:");
  for (const [k, v] of getterNames) console.log(`    get ${k} -> ${v}`);
  console.log("[class->flat] setter map:");
  for (const [k, v] of setterNames) console.log(`    set ${k} -> ${v}`);

  if (!write) {
    console.log("[class->flat] dry run");
    return;
  }

  // ---- Emit ----
  // Single-pass edit collector: walks the whole class once, producing
  // file-absolute, NON-OVERLAPPING edits. `this.X` rewrites only touch the
  // `this.X` token span (and, for setters, the trailing `=`), never the RHS or
  // argument sub-expressions — those are visited normally and get their own
  // disjoint edits.
  type Edit = { start: number; end: number; repl: string };
  const allEdits: Edit[] = [];

  function collect(n: ts.Node, src: ts.SourceFile) {
    // `this.setter = value` -> `setSetter(self, value )`
    if (
      ts.isBinaryExpression(n) &&
      n.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
      ts.isPropertyAccessExpression(n.left) &&
      n.left.expression.kind === ts.SyntaxKind.ThisKeyword &&
      setterNames.has(n.left.name.text) &&
      !methodNames.has(n.left.name.text)
    ) {
      const fn = setterNames.get(n.left.name.text)!;
      // replace `this.setter =` (LHS through operator) with `setSetter(self,`
      allEdits.push({
        start: n.left.getStart(src),
        end: n.operatorToken.getEnd(),
        repl: `${fn}(${cfg.selfParam},`
      });
      // append `)` right after the RHS expression
      allEdits.push({ start: n.right.getEnd(), end: n.right.getEnd(), repl: ")" });
      // descend into RHS only
      collect(n.right, src);
      return;
    }

    if (ts.isPropertyAccessExpression(n) && n.expression.kind === ts.SyntaxKind.ThisKeyword) {
      const prop = n.name.text;
      if (methodNames.has(prop)) {
        const parent = n.parent;
        if (ts.isCallExpression(parent) && parent.expression === n) {
          const fn = methodNames.get(prop)!;
          const hasArgs = parent.arguments.length > 0;
          const openParen = parent
            .getChildren(src)
            .find(c => c.kind === ts.SyntaxKind.OpenParenToken)!;
          // replace `this.method(` -> `method(self, `  (only this token span)
          allEdits.push({
            start: n.getStart(src),
            end: openParen.getEnd(),
            repl: `${fn}(${cfg.selfParam}${hasArgs ? ", " : ""}`
          });
          // descend into arguments (their own edits, disjoint)
          for (const a of parent.arguments) collect(a, src);
          return;
        }
        // method reference (not directly called) -> bound arrow
        const fn = methodNames.get(prop)!;
        allEdits.push({
          start: n.getStart(src),
          end: n.getEnd(),
          repl: `((...a: never[]) => (${fn} as (...x: never[]) => unknown)(${cfg.selfParam}, ...a))`
        });
        return;
      }
      if (getterNames.has(prop)) {
        const fn = getterNames.get(prop)!;
        allEdits.push({ start: n.getStart(src), end: n.getEnd(), repl: `${fn}(${cfg.selfParam})` });
        return;
      }
      if (fieldNames.has(prop)) {
        // `this.field` -> `self.field` (replace only the `this` keyword + dot? no — replace `this` with self)
        allEdits.push({
          start: n.expression.getStart(src),
          end: n.expression.getEnd(),
          repl: cfg.selfParam
        });
        // descend in case field access has computed parts (rare); but name is fine
        return;
      }
      // unknown member on this (e.g. inherited) — leave `this` -> self
      allEdits.push({
        start: n.expression.getStart(src),
        end: n.expression.getEnd(),
        repl: cfg.selfParam
      });
      return;
    }

    if (n.kind === ts.SyntaxKind.ThisKeyword) {
      allEdits.push({ start: n.getStart(src), end: n.getEnd(), repl: cfg.selfParam });
      return;
    }
    ts.forEachChild(n, c => collect(c, src));
  }

  // Collect edits for every member body up-front (single global pass).
  for (const m of [...methods, ...getters, ...setters]) {
    if (m.body) collect(m.body, sf);
  }
  if (ctor?.body) collect(ctor.body, sf);

  // Apply global edits to produce a rewritten copy of the whole file, then
  // extract member text from THAT. Build an offset-shifted full text.
  function applyEditsToRange(nodeStart: number, nodeEnd: number): string {
    const within = allEdits
      .filter(e => e.start >= nodeStart && e.end <= nodeEnd)
      .sort((a, b) => a.start - b.start || a.end - b.end);
    // Drop edits that overlap a previously kept edit (defensive against
    // nested-node double collection).
    const kept: Edit[] = [];
    let lastEnd = -1;
    for (const e of within) {
      if (e.start === e.end) {
        // zero-width insertion — always keep, doesn't overlap
        kept.push(e);
        continue;
      }
      if (e.start < lastEnd) continue; // overlaps previous — skip
      kept.push(e);
      lastEnd = e.end;
    }
    kept.sort((a, b) => b.start - a.start);
    let out = text.slice(nodeStart, nodeEnd);
    for (const e of kept) {
      out = out.slice(0, e.start - nodeStart) + e.repl + out.slice(e.end - nodeStart);
    }
    return out;
  }

  const parts: string[] = [];

  // 1. interface WorksheetData { fields }
  const fieldLines: string[] = [];
  for (const f of fields) {
    const name = (f.name as ts.Identifier).text;
    const opt = f.questionToken ? "?" : "";
    const typeText = f.type ? f.type.getText(sf) : "unknown";
    fieldLines.push(`  ${name}${opt}: ${typeText};`);
  }
  parts.push(`export interface ${cfg.selfType} {\n${fieldLines.join("\n")}\n}`);

  // 2. createWorksheet from constructor
  if (ctor) {
    const params = ctor.parameters.map(p => p.getText(sf)).join(", ");
    const bodyInner = ctor.body
      ? applyEditsToRange(ctor.body.getStart(sf), ctor.body.getEnd())
      : "{}";
    const innerText = bodyInner.replace(/^\s*\{/, "").replace(/\}\s*$/, "");
    // Field initializers (e.g. `readonly _x = new Set()`) run before the ctor body.
    const initLines: string[] = [];
    for (const f of fields) {
      if (f.initializer) {
        const name = (f.name as ts.Identifier).text;
        initLines.push(`  ${cfg.selfParam}.${name} = ${f.initializer.getText(sf)};`);
      }
    }
    parts.push(
      `export function ${cfg.createName}(${params}): ${cfg.selfType} {\n` +
        `  const ${cfg.selfParam} = {} as ${cfg.selfType};\n` +
        (initLines.length ? initLines.join("\n") + "\n" : "") +
        innerText +
        `\n  return ${cfg.selfParam};\n}`
    );
  }

  const selfP = `${cfg.selfParam}: ${cfg.selfType}`;
  // Replace a bare `this` return/param type with the self type name.
  const fixThisType = (s: string): string => s.replace(/\bthis\b/g, cfg.selfType);

  // 3. methods (handle overloads: signature-only decls share a name with one impl)
  // Group consecutive methods by name to emit overload signatures + impl.
  const methodsByName = new Map<string, ts.MethodDeclaration[]>();
  const methodOrder: string[] = [];
  for (const m of methods) {
    const n = (m.name as ts.Identifier).text;
    if (!methodsByName.has(n)) {
      methodsByName.set(n, []);
      methodOrder.push(n);
    }
    methodsByName.get(n)!.push(m);
  }
  for (const n of methodOrder) {
    const group = methodsByName.get(n)!;
    const flat = cfg.methodMap[n] ?? n;
    const impls = group.filter(m => m.body);
    const sigs = group.filter(m => !m.body);
    // emit overload signatures first (only when there are real overloads)
    if (impls.length === 1 && sigs.length > 0) {
      for (const s of sigs) {
        const params = s.parameters.map(p => fixThisType(p.getText(sf))).join(", ");
        const ret = s.type ? `: ${fixThisType(s.type.getText(sf))}` : "";
        const tp = s.typeParameters
          ? `<${s.typeParameters.map(t => t.getText(sf)).join(", ")}>`
          : "";
        const allP = params ? `${selfP}, ${params}` : selfP;
        parts.push(`export function ${flat}${tp}(${allP})${ret};`);
      }
    }
    for (const m of impls) {
      const params = m.parameters.map(p => fixThisType(p.getText(sf))).join(", ");
      const ret = m.type ? `: ${fixThisType(m.type.getText(sf))}` : "";
      const isAsync = m.modifiers?.some(x => x.kind === ts.SyntaxKind.AsyncKeyword) ? "async " : "";
      const tp = m.typeParameters ? `<${m.typeParameters.map(t => t.getText(sf)).join(", ")}>` : "";
      const body = applyEditsToRange(m.body!.getStart(sf), m.body!.getEnd());
      const allP = params ? `${selfP}, ${params}` : selfP;
      parts.push(`export ${isAsync}function ${flat}${tp}(${allP})${ret} ${body}`);
    }
  }

  // 4. getters
  for (const g of getters) {
    const name = getterName(g);
    const ret = g.type ? `: ${fixThisType(g.type.getText(sf))}` : "";
    const body = g.body ? applyEditsToRange(g.body.getStart(sf), g.body.getEnd()) : "{}";
    parts.push(`export function ${name}(${selfP})${ret} ${body}`);
  }

  // 5. setters
  for (const s of setters) {
    const name = setterName(s);
    const param = fixThisType(s.parameters[0].getText(sf));
    const body = s.body ? applyEditsToRange(s.body.getStart(sf), s.body.getEnd()) : "{}";
    parts.push(`export function ${name}(${selfP}, ${param}): void ${body}`);
  }

  const start = classNode.getStart(sf);
  const end = classNode.getEnd();
  const newText = text.slice(0, start) + parts.join("\n\n") + text.slice(end);
  fs.writeFileSync(defPath, newText);
  console.log(`[class->flat] emitted ${parts.length} declarations to ${cfg.defFile}`);
}

main();
