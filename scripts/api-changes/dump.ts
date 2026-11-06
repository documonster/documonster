/**
 * Read the public API of a package — every name every `exports` subpath publishes — into a flat record, so two
 * releases can be compared name by name. Used by `scripts/api-changes.ts`.
 *
 * Each record key is `<entry> :: <name>`, where `<name>` is the path to the member: `Cell.setValue` for a namespace
 * member, `Pdf.Editor#toBuffer` for an instance member, `Pdf.Editor.read` for a static, `WorkbookReadOptions.password?`
 * for a field. Namespaces are walked recursively; a function keeps every overload; a class keeps its constructor,
 * its public instance members and its public statics; an interface or an object type alias keeps every field.
 *
 * The TypeScript checker is needed for this, and the repository's own compiler (TypeScript 7) has no JavaScript API,
 * so this reads declarations with `typescript-api` (TypeScript 5, a devDependency kept for this script alone).
 */
import fs from "node:fs";
import path from "node:path";

import ts from "typescript-api";

export interface ApiEntry {
  kind: string;
  /** The printed signature: one line per overload for a callable. */
  sig: string;
  /**
   * For a callable: its overloads without parameter names, which are not part of a signature's type — `(opts: X)` and
   * `(options: X)` accept the same calls.
   */
  shape?: string;
  /** For a type alias: the type it resolves to, printed with its union members sorted. */
  resolved?: string;
  /** For a type alias: its type parameters. */
  tparams?: string;
  /** The `@deprecated` text when the whole name is deprecated. */
  deprecated: string | false;
  /** The `@deprecated` text when only some of its overloads are. */
  partial?: string;
}

export type ApiRecord = Record<string, ApiEntry>;

const FORMAT =
  ts.TypeFormatFlags.NoTruncation | ts.TypeFormatFlags.UseAliasDefinedOutsideCurrentScope;

function norm(s: string): string {
  return s
    .replace(/import\("[^"]*"\)\./g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/** TSDoc text with `{@link target | label}` shown as its label (or its target), in backticks. */
function partsText(parts: readonly ts.SymbolDisplayPart[] | undefined): string {
  let out = "";
  let link: { name: string; text: string } | null = null;
  for (const p of parts ?? []) {
    if (p.kind === "link") {
      if (link === null) {
        link = { name: "", text: "" };
      } else {
        const shown = link.text.trim() || link.name.trim();
        out += shown ? "`" + shown + "`" : "";
        link = null;
      }
      continue;
    }
    if (link !== null) {
      if (p.kind === "linkName") {
        link.name += p.text;
      } else {
        link.text += p.text;
      }
      continue;
    }
    out += p.text;
  }
  return out.replace(/\s+/g, " ").trim();
}

function deprecationOf(tags: readonly ts.JSDocTagInfo[]): string | false {
  const tag = tags.find(t => t.name === "deprecated");
  return tag ? partsText(tag.text) || "deprecated" : false;
}

interface EntryFile {
  sub: string;
  cond: "node" | "browser";
  file: string;
}

function entryFiles(root: string): EntryFile[] {
  const pkg = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));
  const out: EntryFile[] = [];
  for (const [sub, spec] of Object.entries<any>(pkg.exports)) {
    if (sub === "./package.json" || typeof spec === "string") {
      continue;
    }
    const node = spec.default?.types ?? spec.types ?? spec.import?.types;
    if (!node) {
      continue;
    }
    const browser = spec.browser?.types ?? node;
    out.push({ sub, cond: "node", file: path.join(root, node) });
    if (browser !== node) {
      out.push({ sub, cond: "browser", file: path.join(root, browser) });
    }
  }
  return out;
}

/**
 * The public API of the package at `root`, under one condition. `label` names the package in the record keys
 * (`documonster/excel`, `@documonster/mcp`); browser entries carry a ` [browser]` suffix.
 */
export function dumpApi(
  root: string,
  cond: "node" | "browser",
  label: string,
  typeRoots: string[],
  names: "published" | "declared" = "published"
): ApiRecord {
  const files = entryFiles(root).filter(e => e.cond === cond);
  const out: ApiRecord = {};
  if (files.length === 0) {
    return out;
  }
  const program = ts.createProgram(
    files.map(e => e.file),
    {
      strict: true,
      target: ts.ScriptTarget.ESNext,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
      customConditions: cond === "browser" ? ["browser"] : [],
      skipLibCheck: true,
      lib: ["lib.esnext.d.ts", "lib.dom.d.ts"],
      types: ["node"],
      typeRoots
    }
  );
  const checker = program.getTypeChecker();
  const symDep0 = (sym: ts.Symbol): string | false => deprecationOf(sym.getJsDocTags(checker));
  // The name each type is published under. A type's declared name and its published name can differ (`FlowNode`
  // published as `MermaidFlowNode`), and a declaration file can name it under a third, the alias it was imported as
  // (`FlowNodeImpl`). Both releases are printed with the published name, so a rename that kept the type reads as the
  // same type and only a real change differs. A declared name published under two different names is ambiguous and
  // kept as it is.
  const TYPE_LIKE =
    ts.SymbolFlags.Interface |
    ts.SymbolFlags.TypeAlias |
    ts.SymbolFlags.Class |
    ts.SymbolFlags.Enum;
  const resolve = (sym: ts.Symbol): ts.Symbol =>
    sym.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym) : sym;
  const ownFiles = program
    .getSourceFiles()
    .filter(f => f.fileName.startsWith(root) && !f.fileName.includes("/node_modules/"));
  // Type-like symbols by every name text could show them under: the name they were declared with, and any name a
  // declaration file imported them as. A name that stands for two different types is ambiguous in printed text.
  const byName = new Map<string, Set<ts.Symbol>>();
  const note = (name: string, sym: ts.Symbol): void => {
    (byName.get(name) ?? byName.set(name, new Set()).get(name)!).add(sym);
  };
  for (const f of ownFiles) {
    f.forEachChild(function visit(n: ts.Node): void {
      if (
        (ts.isInterfaceDeclaration(n) ||
          ts.isTypeAliasDeclaration(n) ||
          ts.isClassDeclaration(n) ||
          ts.isEnumDeclaration(n)) &&
        n.name
      ) {
        const sym = checker.getSymbolAtLocation(n.name);
        if (sym) {
          note(n.name.text, sym);
        }
      } else if (ts.isImportSpecifier(n) || ts.isExportSpecifier(n)) {
        const sym = checker.getSymbolAtLocation(n.name);
        const target = sym && resolve(sym);
        if (target && target.flags & TYPE_LIKE) {
          note(n.name.text, target);
        }
      }
      n.forEachChild(visit);
    });
  }
  // The name each type is published under, from the entries' own exports (a deprecated forward is not a publication).
  const published = new Map<ts.Symbol, string>();
  for (const e of files) {
    const sf = program.getSourceFile(e.file);
    if (!sf) {
      continue;
    }
    for (const x of checker.getExportsOfModule(checker.getSymbolAtLocation(sf)!)) {
      const target = resolve(x);
      if (target.flags & TYPE_LIKE && !symDep0(x) && !published.has(target)) {
        published.set(target, x.name);
      }
    }
  }
  // `"declared"` spells every type as it was declared instead, which is what tells a type newly published under a
  // prefixed name (`TransformEntryData` → `ArchiveTransformEntryData`) from a different type.
  const nameOf = (sym: ts.Symbol): string =>
    names === "published" ? (published.get(sym) ?? sym.name) : sym.name;
  const textNames = new Map<string, string>();
  for (const [name, syms] of byName) {
    if (syms.size === 1) {
      const to = nameOf([...syms][0]!);
      if (to !== name) {
        textNames.set(name, to);
      }
    }
  }
  // A declaration file that imports a module as a namespace (`import type * as CsvTypes`) gets types printed under
  // it (`CsvTypes.TransformContext`); the qualifier is only the file's own name for the module.
  const namespaceImports = new Set<string>();
  for (const f of ownFiles) {
    f.forEachChild(function visit(n: ts.Node): void {
      if (ts.isNamespaceImport(n)) {
        namespaceImports.add(n.name.text);
      }
      n.forEachChild(visit);
    });
  }
  const qualifierRe = namespaceImports.size
    ? new RegExp("(?<![\\w$.])(?:" + [...namespaceImports].join("|") + ")\\.(?=[A-Za-z_$])", "g")
    : undefined;
  const textRe = textNames.size
    ? new RegExp("(?<![\\w$.])(" + [...textNames.keys()].join("|") + ")(?![\\w$]|\\??:)", "g")
    : undefined;
  /** Compiler-printed type text, with each type under its published name. */
  const print = (text: string): string => {
    const t = qualifierRe ? norm(text).replace(qualifierRe, "") : norm(text);
    return textRe ? t.replace(textRe, m => textNames.get(m)!) : t;
  };
  // Declaration text (a type alias's right-hand side, a heritage clause, a constraint) names types as the file
  // imported them, possibly qualified by a namespace import (`CsvTypes.ColumnConfig`). Resolve each reference and
  // print the type by its published name.
  const declText = (node: ts.Node): string => {
    const sf = node.getSourceFile();
    const edits: [number, number, string][] = [];
    const visit = (n: ts.Node): void => {
      const ref = ts.isTypeReferenceNode(n)
        ? n.typeName
        : ts.isExpressionWithTypeArguments(n)
          ? n.expression
          : undefined;
      if (ref) {
        const id = ts.isQualifiedName(ref)
          ? ref.right
          : ts.isPropertyAccessExpression(ref)
            ? ref.name
            : ref;
        const sym = checker.getSymbolAtLocation(id);
        const target = sym && resolve(sym);
        if (target && target.flags & TYPE_LIKE) {
          const name = nameOf(target);
          if (name !== ref.getText(sf)) {
            edits.push([ref.getStart(sf), ref.getEnd(), name]);
          }
        }
      }
      n.forEachChild(visit);
    };
    visit(node);
    let text = node.getText(sf);
    const base = node.getStart(sf);
    for (const [from, to, name] of edits.sort((x, y) => y[0] - x[0])) {
      text = text.slice(0, from - base) + name + text.slice(to - base);
    }
    return norm(text);
  };
  const hidden = (d: ts.Declaration | undefined): boolean =>
    !!d &&
    (ts.getCombinedModifierFlags(d) & (ts.ModifierFlags.Private | ts.ModifierFlags.Protected)) !==
      0;
  const symDep = (sym: ts.Symbol): string | false => deprecationOf(sym.getJsDocTags(checker));

  function shapeOf(s: ts.Signature, ctor: boolean): string {
    const tps = (s.getTypeParameters() ?? []).map(tp => {
      const c = tp.getConstraint();
      return (
        checker.typeToString(tp) +
        (c ? ` extends ${checker.typeToString(c, undefined, FORMAT)}` : "")
      );
    });
    const params = s.getParameters().map(p => {
      const d = p.valueDeclaration as ts.ParameterDeclaration | undefined;
      const t = d ? checker.getTypeOfSymbolAtLocation(p, d) : checker.getTypeOfSymbol(p);
      const rest = d?.dotDotDotToken ? "..." : "";
      const opt = d && checker.isOptionalParameter(d) ? "?" : "";
      return `${rest}${opt}${checker.typeToString(t, undefined, FORMAT)}`;
    });
    return print(
      `${ctor ? "new " : ""}${tps.length ? `<${tps.join(", ")}>` : ""}(${params.join(", ")}) => ${checker.typeToString(s.getReturnType(), undefined, FORMAT)}`
    );
  }
  function signatures(
    type: ts.Type,
    ctor: boolean
  ): { text: string; shape: string; dep: string | false }[] {
    const sigs = ctor ? type.getConstructSignatures() : type.getCallSignatures();
    return sigs.map(s => ({
      shape: shapeOf(s, ctor),
      text:
        ctor && s.declaration && hidden(s.declaration as ts.Declaration)
          ? "private constructor — not constructible"
          : (ctor ? "new " : "") +
            print(
              checker.signatureToString(
                s,
                undefined,
                FORMAT | (ctor ? ts.TypeFormatFlags.WriteCallStyleSignature : 0)
              )
            ),
      dep: deprecationOf(s.getJsDocTags())
    }));
  }
  // A name whose every overload is deprecated is deprecated; one with some is a "deprecated call form".
  function depOf(
    list: { dep: string | false }[],
    fallback: string | false
  ): { deprecated: string | false; partial?: string } {
    const d = list.filter(s => s.dep);
    if (list.length === 0 || d.length === 0) {
      return { deprecated: fallback };
    }
    if (d.length === list.length) {
      return { deprecated: d[0]!.dep };
    }
    return { deprecated: false, partial: d[0]!.dep as string };
  }
  function typeParams(decl: ts.Declaration | undefined): string {
    const tps = decl
      ? ts.getEffectiveTypeParameterDeclarations(decl as ts.DeclarationWithTypeParameters)
      : [];
    if (!tps.length) {
      return "";
    }
    return "<" + tps.map(tp => declText(tp)).join(", ") + ">";
  }
  function put(entry: string, key: string, rec: ApiEntry): void {
    out[`${entry} :: ${key}`] = rec;
  }
  function members(entry: string, prefix: string, type: ts.Type, sep: string, at: ts.Node): void {
    // What a type is beyond its named properties: its index signatures and, for an interface or an object type, its
    // own call and construct signatures. A change to any of these changes what code compiles against it.
    for (const info of checker.getIndexInfosOfType(type)) {
      put(entry, `${prefix}${sep}[key: ${checker.typeToString(info.keyType)}]`, {
        kind: "index",
        sig:
          (info.isReadonly ? "readonly " : "") + print(checker.typeToString(info.type, at, FORMAT)),
        deprecated: false
      });
    }
    if (sep === ".") {
      for (const [kind, ctor] of [
        ["call", false],
        ["construct", true]
      ] as const) {
        const list = signatures(type, ctor);
        if (list.length) {
          put(entry, `${prefix}${sep}(${kind})`, {
            kind,
            sig: list.map(s => s.text).join("\n"),
            shape: list.map(s => s.shape).join("\n"),
            ...depOf(list, false)
          });
        }
      }
    }
    for (const p of checker.getPropertiesOfType(type)) {
      if (p.name.startsWith("#")) {
        continue;
      }
      // A well-known symbol member (`[Symbol.asyncIterator]`) is public API like any other.
      const memberName = p.name.startsWith("__@") ? checker.symbolToString(p) : p.name;
      if (memberName.includes("__@")) {
        continue;
      }
      const d = p.valueDeclaration ?? p.declarations?.[0];
      if (hidden(d)) {
        continue;
      }
      const t = checker.getTypeOfSymbolAtLocation(p, at);
      const opt = p.flags & ts.SymbolFlags.Optional ? "?" : "";
      if (p.flags & ts.SymbolFlags.Method && t.getCallSignatures().length) {
        const list = signatures(t, false);
        const dd = depOf(list, symDep(p));
        put(entry, prefix + sep + memberName + opt, {
          kind: "method",
          sig: list.map(s => s.text).join("\n"),
          shape: list.map(s => s.shape).join("\n"),
          ...dd
        });
      } else {
        const ro =
          d && ts.getCombinedModifierFlags(d) & ts.ModifierFlags.Readonly ? "readonly " : "";
        put(entry, prefix + sep + memberName + opt, {
          kind: "prop",
          sig: ro + print(checker.typeToString(t, at, FORMAT)),
          deprecated: symDep(p)
        });
      }
    }
  }
  function classLike(
    entry: string,
    name: string,
    ctorType: ts.Type,
    at: ts.Node,
    dep: string | false
  ): void {
    const ctors = signatures(ctorType, true);
    put(entry, name, {
      kind: "class",
      sig: ctors.map(s => s.text).join("\n"),
      shape: ctors.map(s => s.shape).join("\n"),
      ...depOf(ctors, dep)
    });
    const first = ctorType.getConstructSignatures()[0];
    if (first) {
      members(entry, name, checker.getReturnTypeOfSignature(first), "#", at);
    }
    for (const p of checker.getPropertiesOfType(ctorType)) {
      if (
        p.name === "prototype" ||
        p.name.startsWith("__@") ||
        hidden(p.valueDeclaration ?? p.declarations?.[0])
      ) {
        continue;
      }
      const t = checker.getTypeOfSymbolAtLocation(p, at);
      const list = signatures(t, false);
      put(entry, `${name}.${p.name}`, {
        kind: "static",
        sig: list.length
          ? list.map(s => s.text).join("\n")
          : print(checker.typeToString(t, at, FORMAT)),
        ...(list.length ? { shape: list.map(s => s.shape).join("\n") } : {}),
        ...depOf(list, symDep(p))
      });
    }
  }
  function visit(entry: string, name: string, sym0: ts.Symbol): void {
    const sym = sym0.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(sym0) : sym0;
    const aliasDep = sym0 !== sym ? symDep(sym0) : false;
    const decl = sym.valueDeclaration ?? sym.declarations?.[0];
    if (!decl) {
      put(entry, name, { kind: "unknown", sig: "", deprecated: aliasDep });
      return;
    }
    const f = sym.flags;
    if (
      f & (ts.SymbolFlags.ValueModule | ts.SymbolFlags.NamespaceModule) &&
      !(f & (ts.SymbolFlags.Class | ts.SymbolFlags.Function | ts.SymbolFlags.Enum))
    ) {
      put(entry, name, {
        kind: "namespace",
        sig: "",
        deprecated: aliasDep || symDep(sym)
      });
      for (const m of checker.getExportsOfModule(sym)) {
        visit(entry, `${name}.${m.name}`, m);
      }
      return;
    }
    const baseDep = aliasDep || symDep(sym);
    if (f & ts.SymbolFlags.Class) {
      classLike(entry, name, checker.getTypeOfSymbolAtLocation(sym, decl), decl, baseDep);
      return;
    }
    if (f & ts.SymbolFlags.Function) {
      const list = signatures(checker.getTypeOfSymbolAtLocation(sym, decl), false);
      put(entry, name, {
        kind: "function",
        sig: list.map(s => s.text).join("\n"),
        shape: list.map(s => s.shape).join("\n"),
        ...depOf(list, aliasDep || (list.length === 1 ? symDep(sym) : false))
      });
      return;
    }
    if (f & ts.SymbolFlags.Enum) {
      put(entry, name, { kind: "enum", sig: "", deprecated: baseDep });
      for (const m of checker.getExportsOfModule(sym)) {
        put(entry, `${name}.${m.name}`, {
          kind: "enum-member",
          sig: JSON.stringify(checker.getConstantValue(m.valueDeclaration as ts.EnumMember)),
          deprecated: symDep(m)
        });
      }
      return;
    }
    let typeKey = name;
    if (f & ts.SymbolFlags.Variable) {
      const t = checker.getTypeOfSymbolAtLocation(sym, decl);
      if (t.getConstructSignatures().length) {
        classLike(entry, name, t, decl, baseDep);
      } else if (t.getCallSignatures().length) {
        const list = signatures(t, false);
        const fallback =
          aliasDep || (list.length === 1 || list.every(s => !s.dep) ? symDep(sym) : false);
        put(entry, name, {
          kind: "function",
          sig: list.map(s => s.text).join("\n"),
          shape: list.map(s => s.shape).join("\n"),
          ...depOf(list, fallback)
        });
      } else {
        put(entry, name, {
          kind: "const",
          sig: print(checker.typeToString(t, decl, FORMAT)),
          deprecated: baseDep
        });
      }
      typeKey = name + " (type)";
    }
    if (f & ts.SymbolFlags.Interface) {
      const idecl = sym.declarations!.find(ts.isInterfaceDeclaration) ?? decl;
      const heritage = (sym.declarations ?? []).flatMap(d =>
        ((d as ts.InterfaceDeclaration).heritageClauses ?? []).map(h => declText(h))
      );
      put(entry, typeKey, {
        kind: "interface",
        sig: (typeParams(idecl) + " " + heritage.join(" ")).trim(),
        deprecated: baseDep
      });
      if (typeKey === name) {
        members(entry, name, checker.getDeclaredTypeOfSymbol(sym), ".", idecl);
      }
      return;
    }
    if (f & ts.SymbolFlags.TypeAlias) {
      const d = sym.declarations!.find(ts.isTypeAliasDeclaration)!;
      const t = checker.getDeclaredTypeOfSymbol(sym);
      const printType = (x: ts.Type): string =>
        print(checker.typeToString(x, undefined, FORMAT | ts.TypeFormatFlags.InTypeAlias));
      const resolved = t.isUnion() ? t.types.map(printType).sort().join(" | ") : printType(t);
      put(entry, typeKey, {
        kind: "type",
        sig: declText(d)
          .replace(/^export\s+/, "")
          .replace(/^declare\s+/, ""),
        resolved,
        tparams: typeParams(d),
        deprecated: baseDep
      });
      if (
        typeKey === name &&
        t.flags & ts.TypeFlags.Object &&
        t.getConstructSignatures().length === 0
      ) {
        const target = ts.isTypeReferenceNode(d.type)
          ? checker.getSymbolAtLocation(d.type.typeName)
          : undefined;
        const tsym =
          target && target.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(target) : target;
        members(entry, name, t, tsym && tsym.flags & ts.SymbolFlags.Class ? "#" : ".", d);
      }
      return;
    }
    if (!(f & ts.SymbolFlags.Variable)) {
      put(entry, name, { kind: "other", sig: "", deprecated: baseDep });
    }
  }
  for (const e of files) {
    const sf = program.getSourceFile(e.file);
    if (!sf) {
      throw new Error(`api-changes: ${e.file} is missing — build the package first`);
    }
    const moduleSymbol = checker.getSymbolAtLocation(sf)!;
    const entry = `${label}${e.sub === "." ? "" : e.sub.slice(1)}${cond === "browser" ? " [browser]" : ""}`;
    for (const x of checker.getExportsOfModule(moduleSymbol)) {
      visit(entry, x.name, x);
    }
  }
  return out;
}
