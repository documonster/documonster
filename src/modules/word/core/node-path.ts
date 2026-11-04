/**
 * The writer's choke point for models it cannot write: an unknown node kind,
 * a list that is not an array, a value with no XML form. Every model — built
 * with `Build.*`, assembled by hand or read — goes through the writer, so this
 * is the one place such a node is refused, with a {@link DocxArgumentError}
 * naming where it is (`body[3].children[0].content[1].type`) rather than being
 * dropped or failing as another module's error.
 *
 * The loops that know an index (body blocks, paragraph children, run content,
 * table rows and cells) prefix the field of an error passing through them;
 * the checks themselves cost one comparison on the way, and a `try` that does
 * not throw costs nothing.
 */

import { describeValue } from "@word/core/argument-checks";
import { isRangeMarker } from "@word/core/text-utils";
import { DocxArgumentError } from "@word/errors";
import type { DocxDocument } from "@word/types";
import { XmlError } from "@xml/errors";

/** Refuse a node the writer cannot write, at `field` (relative to its parent). */
export function unwritableNode(field: string, expected: string, value: unknown): never {
  throw new DocxArgumentError(field, `${field} must be ${expected}, got ${describeValue(value)}`);
}

/** An array, or {@link unwritableNode}. */
export function writableList<T>(value: readonly T[] | undefined, field: string): readonly T[] {
  if (!Array.isArray(value)) {
    unwritableNode(field, "an array", value);
  }
  return value;
}

/**
 * `render()`, with `prefix` put before the field of a model check failing in
 * it ({@link atNode}). Each renderer names fields relative to the object it was
 * given; the caller that handed it `holder[key]` wraps the call in
 * `within(key, …)`, so the field a caller sees is the exact path from the
 * document (`body[0].properties.alignment`) to the value.
 */
export function within<T>(prefix: string, render: () => T): T {
  try {
    return render();
  } catch (error) {
    throw atNode(prefix, error);
  }
}

/** What the innermost check said, and the error it threw, for each relocated error. */
const ORIGIN = new WeakMap<Error, { readonly message: string; readonly cause: unknown }>();

/**
 * `error`, thrown while writing the node at `prefix`, with that prefix on its
 * field: a {@link DocxArgumentError} (of whatever subclass) keeps its
 * class; an `XmlError` — a value with no XML form — becomes one. Anything
 * else is returned as it is.
 */
export function atNode(prefix: string, error: unknown): unknown {
  if (error instanceof DocxArgumentError) {
    const origin = ORIGIN.get(error) ?? { message: error.message, cause: error };
    const field =
      error.field === ""
        ? prefix
        : error.field.startsWith("[")
          ? `${prefix}${error.field}`
          : `${prefix}.${error.field}`;
    const Class = error.constructor as new (
      field: string,
      message: string,
      options?: { cause?: unknown }
    ) => DocxArgumentError;
    const located = new Class(field, `${origin.message} (at ${field})`, { cause: origin.cause });
    ORIGIN.set(located, origin);
    return located;
  }
  if (error instanceof XmlError) {
    const origin = { message: error.message, cause: error };
    const located = new DocxArgumentError(prefix, `${origin.message} (at ${prefix})`, {
      cause: error
    });
    ORIGIN.set(located, origin);
    return located;
  }
  return error;
}

const BLOCK_KINDS: ReadonlySet<string> = new Set([
  "paragraph",
  "table",
  "floatingImage",
  "tableOfContents",
  "math",
  "textBox",
  "checkBox",
  "drawingShape",
  "opaqueDrawing",
  "chart",
  "chartEx",
  "altChunk",
  "sdt"
]);

const PARAGRAPH_CHILD_KINDS: ReadonlySet<string> = new Set([
  "hyperlink",
  "commentReference",
  "opaqueParagraphChild",
  "insertedRun",
  "deletedRun",
  "movedFromRun",
  "movedToRun"
]);

const RUN_CONTENT_KINDS: ReadonlySet<string> = new Set([
  "text",
  "break",
  "tab",
  "ptab",
  "ruby",
  "symbol",
  "footnoteRef",
  "endnoteRef",
  "field",
  "image",
  "carriageReturn",
  "noBreakHyphen",
  "softHyphen",
  "lastRenderedPageBreak",
  "annotationReference",
  "dateField",
  "opaqueRun"
]);

type Node = { readonly type?: unknown } & Record<string, unknown>;

/** `value` is an object, or {@link unwritableNode} at `field`. */
function node(value: unknown, field: string, expected: string): Node {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    unwritableNode(field, expected, value);
  }
  return value as Node;
}

// The checks below are plain loops with a `try` per item rather than a
// callback per item: a `try` that does not throw is free, a closure is not,
// and this runs once per paragraph, run and cell of every document written.

function checkRun(run: Node): void {
  const content = writableList(run.content as readonly unknown[], "content");
  for (let i = 0; i < content.length; i++) {
    const item = content[i];
    const type = (item as Node | null)?.type;
    if (typeof type === "string" && RUN_CONTENT_KINDS.has(type)) {
      continue;
    }
    if (typeof item !== "object" || item === null) {
      unwritableNode(`content[${i}]`, "a run content node", item);
    }
    unwritableNode(`content[${i}].type`, "a run content kind", type);
  }
}

function checkParagraphChild(item: unknown): void {
  const child = node(item, "", "a paragraph child");
  const type = child.type;
  if (type === undefined) {
    checkRun(child);
    return;
  }
  if (typeof type !== "string") {
    unwritableNode("type", "a paragraph child kind", type);
  }
  if (PARAGRAPH_CHILD_KINDS.has(type)) {
    if (type === "hyperlink") {
      writableList(child.children as readonly unknown[], "children");
    } else if (type !== "commentReference" && type !== "opaqueParagraphChild") {
      checkRun(node(child.run, "run", "a run"));
    }
    return;
  }
  if (!isRangeMarker(child)) {
    unwritableNode("type", "a paragraph child kind", type);
  }
}

function checkParagraph(p: Node): void {
  const children = writableList(p.children as readonly unknown[], "children");
  for (let i = 0; i < children.length; i++) {
    try {
      checkParagraphChild(children[i]);
    } catch (error) {
      throw atNode(`children[${i}]`, error);
    }
  }
}

function checkCellBlock(item: unknown): void {
  const block = node(item, "", "a paragraph or a table");
  if (block.type === "paragraph") {
    checkParagraph(block);
  } else if (block.type === "table") {
    checkTable(block);
  } else {
    unwritableNode("type", "a paragraph or a table", block.type);
  }
}

function checkTable(t: Node): void {
  const rows = writableList(t.rows as readonly unknown[], "rows");
  for (let r = 0; r < rows.length; r++) {
    try {
      const cells = writableList(
        node(rows[r], "", "a table row").cells as readonly unknown[],
        "cells"
      );
      for (let c = 0; c < cells.length; c++) {
        try {
          const content = writableList(
            node(cells[c], "", "a table cell").content as readonly unknown[],
            "content"
          );
          for (let k = 0; k < content.length; k++) {
            try {
              checkCellBlock(content[k]);
            } catch (error) {
              throw atNode(`content[${k}]`, error);
            }
          }
        } catch (error) {
          throw atNode(`cells[${c}]`, error);
        }
      }
    } catch (error) {
      throw atNode(`rows[${r}]`, error);
    }
  }
}

function checkParagraphList(list: unknown, field: string): void {
  if (list === undefined) {
    return;
  }
  const items = writableList(list as readonly unknown[], field);
  for (let i = 0; i < items.length; i++) {
    try {
      const p = node(items[i], "", "a paragraph");
      if (p.type !== "paragraph") {
        unwritableNode("type", "a paragraph", p.type);
      }
      checkParagraph(p);
    } catch (error) {
      throw atNode(`${field}[${i}]`, error);
    }
  }
}

function checkBlockList(list: unknown, field: string): void {
  const items = writableList(list as readonly unknown[], field);
  for (let i = 0; i < items.length; i++) {
    try {
      checkWritableBlock(items[i]);
    } catch (error) {
      throw atNode(`${field}[${i}]`, error);
    }
  }
}

/** One block of the body, a cell, a content control or a story. */
export function checkWritableBlock(value: unknown): void {
  const block = node(value, "", "a body block");
  switch (block.type) {
    case "paragraph":
      checkParagraph(block);
      return;
    case "table":
      checkTable(block);
      return;
    case "textBox":
      checkParagraphList(block.content, "content");
      return;
    case "drawingShape":
      checkParagraphList(block.textContent, "textContent");
      return;
    case "tableOfContents":
      checkParagraphList(block.cachedParagraphs, "cachedParagraphs");
      return;
    case "sdt": {
      const content = writableList(block.content as readonly unknown[], "content");
      for (let i = 0; i < content.length; i++) {
        try {
          const inner = node(content[i], "", "a block or a run");
          if (inner.type === undefined) {
            checkRun(inner);
          } else {
            checkWritableBlock(inner);
          }
        } catch (error) {
          throw atNode(`content[${i}]`, error);
        }
      }
      return;
    }
    default:
      if (typeof block.type !== "string" || !BLOCK_KINDS.has(block.type)) {
        unwritableNode("type", "a body block kind", block.type);
      }
  }
}

/**
 * Every node of `doc` the writer walks, checked for its kind and its lists
 * before anything is written — the pre-passes that collect images, links and
 * styles walk the same tree, and must not be the ones to trip over it.
 */
export function checkWritableDocument(doc: DocxDocument): void {
  checkBlockList(doc.body, "body");
  for (const kind of ["headers", "footers"] as const) {
    for (const [key, part] of doc[kind] ?? []) {
      try {
        const content = node(node(part, "", "a header or footer").content, "content", "an object");
        const children = writableList(content.children as readonly unknown[], "children");
        for (let i = 0; i < children.length; i++) {
          try {
            checkCellBlock(children[i]);
          } catch (error) {
            throw atNode(`children[${i}]`, error);
          }
        }
      } catch (error) {
        throw atNode(`${kind}[${JSON.stringify(key)}]`, error);
      }
    }
  }
  for (const kind of ["footnotes", "endnotes", "comments"] as const) {
    const list = doc[kind];
    if (list === undefined) {
      continue;
    }
    const items = writableList(list as readonly unknown[], kind);
    for (let i = 0; i < items.length; i++) {
      try {
        checkParagraphList(node(items[i], "", "an object").content, "content");
      } catch (error) {
        throw atNode(`${kind}[${i}]`, error);
      }
    }
  }
}
