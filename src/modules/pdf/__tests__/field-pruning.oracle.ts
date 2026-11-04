/**
 * An oracle for `page-index.ts`: what a full rewrite keeping some original pages leaves out,
 * worked out the plain way — one recursive walk of /Fields per selection of pages, nothing
 * shared between selections, no labels, no fixpoint. It is the walk the editor ran before the
 * rules moved into `page-index.ts`, so it is the established behaviour; keeping it here, apart
 * from the code it checks, is what lets a bug in the shared rules show up even though
 * `splitPages()` and its `toBuffer()` reference now both run them.
 *
 * Deliberately naive: recursion (the inputs are small), a fresh walk per selection.
 */
import { buildPageIndex, pruneRemovedPages, pruneToPage } from "@pdf/builder/page-index";
import type { Pruning } from "@pdf/builder/page-index";
import { initDecryption, isEncrypted } from "@pdf/reader/pdf-decrypt";
import { PdfDocument } from "@pdf/reader/pdf-document";
import { isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfObject } from "@pdf/reader/pdf-parser";
import { expect } from "vitest";

/** What the oracle says a selection leaves out. */
export interface NaivePruning {
  readonly dropped: ReadonlySet<string>;
  readonly kids: ReadonlyMap<string, readonly PdfObject[]>;
  readonly fields: readonly PdfObject[] | undefined;
}

const keyOf = (ref: { objNum: number; gen: number }): string => `${ref.objNum}:${ref.gen}`;

/** The rules, applied to `doc` keeping the original pages `kept` says it keeps. */
export function naivePruning(doc: PdfDocument, kept: (page: number) => boolean): NaivePruning {
  const dropped = new Set<string>();
  const kids = new Map<string, PdfObject[]>();
  const pages = doc.getPagesWithObjInfo();
  pages.forEach(({ objNum, gen }, i) => {
    if (!kept(i) && objNum > 0) {
      dropped.add(keyOf({ objNum, gen }));
    }
  });

  const onKept = new Set<string>();
  const onRemoved = new Set<string>();
  pages.forEach(({ dict }, i) => {
    const annots = doc.deref(dict.get("Annots") ?? null);
    for (const ref of isPdfArray(annots) ? annots : []) {
      if (isPdfRef(ref)) {
        (kept(i) ? onKept : onRemoved).add(keyOf(ref));
      }
    }
  });
  const removed = [...onRemoved].filter(key => !onKept.has(key));
  if (removed.length === 0) {
    return { dropped, kids, fields: undefined };
  }
  for (const key of removed) {
    dropped.add(key);
  }

  let fields: PdfObject | null;
  try {
    const acroForm = doc.derefDict(doc.getCatalog().get("AcroForm") ?? null);
    fields = acroForm ? doc.deref(acroForm.get("Fields") ?? null) : null;
  } catch {
    return { dropped, kids, fields: undefined };
  }
  if (!isPdfArray(fields)) {
    return { dropped, kids, fields: undefined };
  }

  // A node's outcome is decided the first time it is walked; one met again while still being
  // walked (a /Kids cycle) counts as kept.
  const outcome = new Map<string, boolean>();
  const walking = new Set<string>();
  const survives = (item: PdfObject): boolean => {
    if (!isPdfRef(item)) {
      return true;
    }
    const key = keyOf(item);
    if (removed.includes(key)) {
      return false;
    }
    if (outcome.has(key)) {
      return outcome.get(key)!;
    }
    if (walking.has(key)) {
      return true;
    }
    const node = doc.derefDict(item);
    const list = node ? doc.deref(node.get("Kids") ?? null) : null;
    if (!isPdfArray(list) || list.length === 0) {
      return true;
    }
    walking.add(key);
    const survivors = list.filter(kid => {
      const kept = survives(kid);
      if (!kept && isPdfRef(kid)) {
        dropped.add(keyOf(kid));
      }
      return kept;
    });
    walking.delete(key);
    outcome.set(key, survivors.length > 0);
    if (survivors.length === 0) {
      dropped.add(key);
    } else if (survivors.length < list.length) {
      kids.set(key, survivors);
    }
    return survivors.length > 0;
  };
  const pruned = fields.filter(field => {
    const kept = survives(field);
    if (!kept && isPdfRef(field)) {
      dropped.add(keyOf(field));
    }
    return kept;
  });
  return { dropped, kids, fields: pruned };
}

/** Every object the rules could have something to say about, by `"objNum:gen"`. */
function candidateKeys(doc: PdfDocument): string[] {
  const keys = new Set<string>();
  const visit = (item: PdfObject): void => {
    if (!isPdfRef(item) || keys.has(keyOf(item))) {
      return;
    }
    keys.add(keyOf(item));
    const node = doc.derefDict(item);
    const list = node ? doc.deref(node.get("Kids") ?? null) : null;
    for (const kid of isPdfArray(list) ? list : []) {
      visit(kid);
    }
  };
  for (const { dict, objNum, gen } of doc.getPagesWithObjInfo()) {
    if (objNum > 0) {
      keys.add(keyOf({ objNum, gen }));
    }
    const annots = doc.deref(dict.get("Annots") ?? null);
    for (const ref of isPdfArray(annots) ? annots : []) {
      visit(ref);
    }
  }
  try {
    const acroForm = doc.derefDict(doc.getCatalog().get("AcroForm") ?? null);
    const fields = acroForm ? doc.deref(acroForm.get("Fields") ?? null) : null;
    for (const field of isPdfArray(fields) ? fields : []) {
      visit(field);
    }
  } catch {
    // No catalog: nothing in a form to visit.
  }
  return [...keys];
}

const show = (list: readonly PdfObject[] | undefined): string | undefined =>
  list?.map(item => (isPdfRef(item) ? `${keyOf(item)}R` : JSON.stringify(item))).join(" ");

function describePruning(pruning: Pruning | NaivePruning, keys: readonly string[]): string[] {
  const naive = "dropped" in pruning ? pruning : undefined;
  const fields = naive ? naive.fields : (pruning as Pruning).fields();
  const lines = [`fields: ${show(fields)}`];
  for (const key of keys) {
    const dropped = naive ? naive.dropped.has(key) : (pruning as Pruning).isDropped(key);
    const kids = naive ? naive.kids.get(key) : (pruning as Pruning).kidsOverride(key);
    if (dropped || kids) {
      lines.push(`${key}: ${dropped ? "dropped" : `kids ${show(kids)}`}`);
    }
  }
  return lines;
}

/** A parsed document, decrypted with `password`; `undefined` when it cannot be read. */
function open(bytes: Uint8Array, password: string): PdfDocument | undefined {
  try {
    const doc = new PdfDocument(bytes);
    if (isEncrypted(doc) && !initDecryption(doc, password)) {
      return undefined;
    }
    doc.getPagesWithObjInfo();
    return doc;
  } catch {
    return undefined;
  }
}

/**
 * Assert that `page-index.ts` answers as the oracle does for `bytes`: for every one-page
 * selection, as a part of `splitPages` and as a removal of every other page; for the
 * selection that keeps no original page; for removing each page alone; and for each
 * `removals` selection.
 */
export function expectPruningMatchesOracle(
  bytes: Uint8Array,
  password = "",
  removals: ReadonlyArray<ReadonlySet<number>> = []
): void {
  const doc = open(bytes, password);
  if (!doc) {
    return;
  }
  const pages = doc.getPagesWithObjInfo();
  const count = pages.length;
  const keys = candidateKeys(doc);
  const index = buildPageIndex(doc, pages);
  const all = Array.from({ length: count }, (_, i) => i);
  const check = (label: string, kept: (page: number) => boolean, shared: Pruning[]) => {
    const expected = describePruning(naivePruning(doc, kept), keys);
    for (const pruning of shared) {
      expect(describePruning(pruning, keys), label).toEqual(expected);
    }
  };
  for (const page of [...all, -1]) {
    if (count <= (page < 0 ? 0 : 1)) {
      continue; // Keeps every page: nothing is pruned, and the editor does not ask.
    }
    const others = new Set(all.filter(i => i !== page));
    check(`only page ${page}`, i => i === page, [
      pruneToPage(index, page),
      pruneRemovedPages(index, others)
    ]);
  }
  for (const removed of [...all.map(page => new Set([page])), ...removals]) {
    if (removed.size > 0 && count > 1) {
      check(`without pages ${[...removed].join(",")}`, i => !removed.has(i), [
        pruneRemovedPages(index, removed)
      ]);
    }
  }
}
