/**
 * The worker scripts the package starts, and how to read their text from a built tree.
 *
 * A worker's source reaches a browser as a *string* — handed to `new Blob([…])`, or exported by a module that
 * generates it — so no parse of the files that ship sees inside it, and a syntax the floor cannot parse there fails
 * only once a worker starts. Two checks read these scripts: `verify-browser-syntax.ts` refuses the syntax it knows
 * of, and `verify-browser-floor.ts` hands them to the floor's own engines. Both read them here, so they cannot
 * disagree about which workers there are or what each one runs.
 */
import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

import ts from "typescript-api";

/**
 * Every built file that starts a worker, relative to the built tree (`dist`), and where the worker's source is:
 * `"inline"` when the file spells it out in a `new Blob([…])`, or the built module that exports it — a string, or a
 * function returning one.
 */
export const WORKER_SPAWNERS: readonly {
  readonly name: string;
  readonly file: string;
  readonly script: "inline" | { readonly module: string; readonly export: string };
}[] = [
  {
    name: "inflate worker",
    file: "esm/modules/archive/unzip/stream.browser.js",
    script: "inline"
  },
  {
    name: "csv worker",
    file: "esm/modules/csv/worker/pool.browser.js",
    script: {
      module: "esm/modules/csv/worker/worker-script.generated.js",
      export: "CSV_WORKER_SCRIPT"
    }
  },
  {
    name: "compression worker",
    file: "esm/modules/archive/compression/worker-pool/pool.browser.js",
    script: {
      module: "esm/modules/archive/compression/worker-pool/worker-script.js",
      export: "generateWorkerScript"
    }
  }
];

/** Every `const` declared exactly once in `file`, by name, to its initializer. */
export function singleConsts(file: ts.SourceFile): Map<string, ts.Expression | undefined> {
  const consts = new Map<string, ts.Expression | undefined>();
  const visit = (node: ts.Node): void => {
    if (ts.isVariableDeclaration(node) && ts.isIdentifier(node.name)) {
      const list = node.parent;
      const isConst =
        ts.isVariableDeclarationList(list) &&
        (list.flags & ts.NodeFlags.BlockScoped) === ts.NodeFlags.Const;
      const name = node.name.text;
      consts.set(name, consts.has(name) || !isConst ? undefined : node.initializer);
    } else if (
      (ts.isParameter(node) || ts.isBindingElement(node) || ts.isFunctionDeclaration(node)) &&
      node.name &&
      ts.isIdentifier(node.name)
    ) {
      consts.set(node.name.text, undefined);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return consts;
}

/** The text of `node` as far as `consts` lets it be known, with every computed part as `fill`. */
export function staticText(
  node: ts.Expression,
  fill: string,
  consts: Map<string, ts.Expression | undefined>,
  seen = new Set<string>()
): string | undefined {
  const at = (child: ts.Expression): string | undefined => staticText(child, fill, consts, seen);
  if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) {
    return node.text;
  }
  if (
    ts.isParenthesizedExpression(node) ||
    ts.isAsExpression(node) ||
    ts.isSatisfiesExpression(node)
  ) {
    return at(node.expression);
  }
  if (ts.isTemplateExpression(node)) {
    return (
      node.head.text +
      node.templateSpans.map(span => (at(span.expression) ?? fill) + span.literal.text).join("")
    );
  }
  if (
    ts.isTaggedTemplateExpression(node) &&
    ts.isPropertyAccessExpression(node.tag) &&
    ts.isIdentifier(node.tag.expression) &&
    node.tag.expression.text === "String" &&
    node.tag.name.text === "raw"
  ) {
    const template = node.template;
    return ts.isNoSubstitutionTemplateLiteral(template)
      ? (template.rawText ?? template.text)
      : (template.head.rawText ?? template.head.text) +
          template.templateSpans
            .map(
              span => (at(span.expression) ?? fill) + (span.literal.rawText ?? span.literal.text)
            )
            .join("");
  }
  if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.PlusToken) {
    const left = at(node.left);
    const right = at(node.right);
    return left === undefined && right === undefined ? undefined : (left ?? fill) + (right ?? fill);
  }
  if (ts.isIdentifier(node) && !seen.has(node.text)) {
    const initializer = consts.get(node.text);
    return initializer
      ? staticText(initializer, fill, consts, new Set(seen).add(node.text))
      : undefined;
  }
  return undefined;
}

/** The `type` a `new Blob(parts, { type })` call gives, if the file spells it out. */
function blobType(
  node: ts.NewExpression,
  consts: Map<string, ts.Expression | undefined>
): string | undefined {
  const options = node.arguments?.[1];
  if (!options || !ts.isObjectLiteralExpression(options)) {
    return undefined;
  }
  for (const property of options.properties) {
    if (
      ts.isPropertyAssignment(property) &&
      ts.isIdentifier(property.name) &&
      property.name.text === "type"
    ) {
      return staticText(property.initializer, "", consts);
    }
  }
  return undefined;
}

/**
 * The script `node` puts in a Blob, when it is a `new Blob([…], { type: "…javascript" })` whose text the file spells
 * out. `undefined` for any other node, and for a script assembled at run time, which cannot be read here.
 */
export function blobScript(
  node: ts.Node,
  consts: Map<string, ts.Expression | undefined>
): string | undefined {
  if (
    !ts.isNewExpression(node) ||
    !ts.isIdentifier(node.expression) ||
    node.expression.text !== "Blob" ||
    !/javascript/i.test(blobType(node, consts) ?? "")
  ) {
    return undefined;
  }
  const parts = node.arguments?.[0];
  if (!parts || !ts.isArrayLiteralExpression(parts) || parts.elements.length === 0) {
    return undefined;
  }
  const text = parts.elements.map(part => staticText(part, "", consts));
  return text.every(part => part !== undefined) ? text.join("") : undefined;
}

/** Every Blob script `source` spells out, in order. */
function blobScripts(source: string): string[] {
  const file = ts.createSourceFile("x.js", source, ts.ScriptTarget.Latest, true, ts.ScriptKind.JS);
  const consts = singleConsts(file);
  const found: string[] = [];
  const visit = (node: ts.Node): void => {
    const script = blobScript(node, consts);
    if (script !== undefined) {
      found.push(script);
    }
    ts.forEachChild(node, visit);
  };
  visit(file);
  return found;
}

/** A worker's script, as read from a built tree. `inline` when its spawner spells it out in a Blob. */
export interface WorkerScript {
  readonly name: string;
  /** The built file the text was read from, relative to the tree: the spawner, or the module that exports it. */
  readonly source: string;
  readonly inline: boolean;
  readonly text: string;
}

/**
 * The text of every worker in {@link WORKER_SPAWNERS}, read from the built tree at `dist`, and what stopped any of
 * them being read — a file that is gone, an inline script no longer spelled out, an export that is not a script.
 * A worker that cannot be read is reported rather than left out, so neither check can pass without it.
 */
export async function readWorkerScripts(
  dist: string
): Promise<{ scripts: WorkerScript[]; problems: string[] }> {
  const scripts: WorkerScript[] = [];
  const problems: string[] = [];
  for (const { name, file, script } of WORKER_SPAWNERS) {
    const spawner = path.join(dist, file);
    if (!fs.existsSync(spawner)) {
      problems.push(`${file}: listed in WORKER_SPAWNERS, but does not exist`);
      continue;
    }
    if (script === "inline") {
      const found = blobScripts(fs.readFileSync(spawner, "utf8"));
      if (found.length !== 1) {
        problems.push(
          `${file}: WORKER_SPAWNERS says it spells out its worker script in a Blob, but ${found.length} were read`
        );
        continue;
      }
      scripts.push({ name, source: file, inline: true, text: found[0]! });
      continue;
    }
    const module = path.join(dist, script.module);
    if (!fs.existsSync(module)) {
      problems.push(`${script.module}: listed as holding a worker script, but does not exist`);
      continue;
    }
    const exported = ((await import(pathToFileURL(module).href)) as Record<string, unknown>)[
      script.export
    ];
    const text = typeof exported === "function" ? (exported as () => unknown)() : exported;
    if (typeof text !== "string" || text.length === 0) {
      problems.push(`${script.module}: ${script.export} did not yield a worker script`);
      continue;
    }
    scripts.push({ name, source: script.module, inline: false, text });
  }
  return { scripts, problems };
}
