/**
 * What `PdfEditor.splitPages()` must produce, stated without it: each part is the document
 * as it would be written now — with the editor's edits — with every other page removed,
 * written by `toBuffer()`. `splitPages` writes its parts
 * on a dedicated path, so these helpers compare the two object for object — the object graph
 * reachable from the trailer, with references renumbered by first visit — and by object count,
 * so a part may differ from the reference only in how its objects are numbered and ordered.
 */
import { initDecryption, isEncrypted } from "@pdf/reader/pdf-decrypt";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfObject, PdfRef } from "@pdf/reader/pdf-parser";
import { expect } from "vitest";

import { Pdf, PdfError } from "../index";
import { expectPruningMatchesOracle } from "./field-pruning.oracle";

/** How a write ended: the parts, or the class and message of what it threw. */
export type SplitOutcome = { parts: Uint8Array[] } | { error: string };

async function settle(write: () => Promise<Uint8Array[]>): Promise<SplitOutcome> {
  try {
    return { parts: await write() };
  } catch (error) {
    expect(error).toBeInstanceOf(PdfError);
    return { error: `${(error as Error).constructor.name}: ${(error as Error).message}` };
  }
}

/** An editor, as `Pdf.Editor.read` returns it. */
export type Editor = ReturnType<typeof Pdf.Editor.read>;

/**
 * The reference: one editor per page, `edit` applied, every other page removed, written with
 * `toBuffer()`. `edit` may not add pages (a part that is an added page is not expressible
 * here), and pages it removes are not parts.
 */
export function referenceSplit(
  bytes: Uint8Array,
  password: string,
  indices?: readonly number[],
  edit: (editor: Editor) => void = () => {}
): Promise<SplitOutcome> {
  return settle(async () => {
    const probe = Pdf.Editor.read(bytes, { password });
    edit(probe);
    const count = probe.pageCount;
    const parts: Uint8Array[] = [];
    for (const idx of indices ?? Array.from({ length: count }, (_, i) => i)) {
      if (idx < 0 || idx >= count || probe.isPageRemoved(idx)) {
        continue;
      }
      const editor = Pdf.Editor.read(bytes, { password });
      edit(editor);
      for (let i = 0; i < count; i++) {
        if (i !== idx) {
          editor.removePage(i);
        }
      }
      parts.push(await editor.toBuffer());
    }
    return parts;
  });
}

/**
 * A written file as text that is independent of object numbering and placement: every object
 * reachable from the trailer, numbered by first visit, streams by their decoded bytes.
 */
export function canonicalPdf(bytes: Uint8Array, password: string): string {
  const doc = new PdfDocument(bytes);
  if (isEncrypted(doc)) {
    expect(initDecryption(doc, password)).toBe(true);
  }
  const ids = new Map<string, number>();
  const queue: PdfRef[] = [];
  const hex = (data: Uint8Array): string =>
    Array.from(data, b => b.toString(16).padStart(2, "0")).join("");
  const value = (val: PdfObject | undefined): string => {
    if (val === undefined || val === null) {
      return "null";
    }
    if (isPdfRef(val)) {
      const key = `${val.objNum}:${val.gen}`;
      let id = ids.get(key);
      if (id === undefined) {
        id = ids.size;
        ids.set(key, id);
        queue.push(val);
      }
      return `@${id}`;
    }
    if (isPdfArray(val)) {
      return `[${val.map(value).join(" ")}]`;
    }
    if (val instanceof Map) {
      return `<<${[...val].map(([k, v]) => `/${k} ${value(v)}`).join(" ")}>>`;
    }
    if (val instanceof Uint8Array) {
      return `<${hex(val)}>`;
    }
    return JSON.stringify(val);
  };

  const lines: string[] = [];
  for (const key of ["Root", "Info", "Encrypt"]) {
    lines.push(`trailer /${key} ${value(doc.trailer.get(key))}`);
  }
  for (let i = 0; i < queue.length; i++) {
    const ref = queue[i];
    const stream = doc.derefStreamWithObjNum(ref);
    if (stream) {
      const dict = new Map(stream.stream.dict);
      for (const key of ["Length", "Filter", "DecodeParms"]) {
        dict.delete(key);
      }
      const data = doc.getStreamData(stream.stream, stream.objNum, stream.gen);
      lines.push(`${i} stream ${value(dict)} <${hex(data)}>`);
    } else {
      lines.push(`${i} ${value(doc.deref(ref))}`);
    }
  }
  const objects = new TextDecoder("latin1").decode(bytes).match(/(?:^|[\r\n])\d+ \d+ obj\b/g);
  lines.push(`objects written: ${objects?.length ?? 0}`);
  return lines.join("\n");
}

/** The original pages `edit` removes; none when the document cannot be read. */
function removedBy(
  bytes: Uint8Array,
  password: string,
  edit: ((editor: Editor) => void) | undefined
): Set<number> {
  try {
    const editor = Pdf.Editor.read(bytes, { password });
    edit?.(editor);
    return new Set(
      Array.from({ length: editor.pageCount }, (_, i) => i).filter(i => editor.isPageRemoved(i))
    );
  } catch {
    return new Set();
  }
}

/** Assert that `splitPages()` writes what the reference writes, part by part. */
export async function expectSplitMatchesReference(
  bytes: Uint8Array,
  options: { password?: string; indices?: number[]; edit?: (editor: Editor) => void } = {}
): Promise<SplitOutcome> {
  const password = options.password ?? "";
  // Both sides below run the pruning rules of `page-index.ts`; the oracle checks the rules.
  expectPruningMatchesOracle(bytes, password, [removedBy(bytes, password, options.edit)]);
  const expected = await referenceSplit(bytes, password, options.indices, options.edit);
  let actual: SplitOutcome;
  try {
    const editor = Pdf.Editor.read(bytes, { password });
    options.edit?.(editor);
    actual = await settle(() => editor.splitPages(options.indices));
  } catch (error) {
    // Reading failed, so the reference failed the same way before writing anything.
    expect(error).toBeInstanceOf(PdfError);
    actual = { error: `${(error as Error).constructor.name}: ${(error as Error).message}` };
  }
  if ("error" in expected || "error" in actual) {
    // The message names the call that refused; everything else must match.
    const named = (outcome: SplitOutcome) =>
      "error" in outcome
        ? { error: outcome.error.replaceAll("toBuffer()", "splitPages()") }
        : outcome;
    expect(actual).toEqual(named(expected));
    return actual;
  }
  expect(actual.parts.length).toBe(expected.parts.length);
  for (let i = 0; i < expected.parts.length; i++) {
    expect(canonicalPdf(actual.parts[i], password), `part ${i}`).toBe(
      canonicalPdf(expected.parts[i], password)
    );
  }
  return actual;
}
