/**
 * Which objects of a document go with which page, and what a full rewrite that leaves pages
 * out drops with them. This is the one statement of those rules: `toBuffer()` after
 * `removePage()` and every part of `splitPages()` ask it.
 *
 * A full rewrite does not write a removed page's annotations, but the AcroForm still lists
 * them — directly in /Fields or as /Kids of a parent field — and cloning that list
 * resurrected each one as an orphan the form would still report. The rules:
 *
 * - a page object is dropped when any place the page tree lists it is removed;
 * - an annotation is dropped when every page listing it is removed;
 * - a field node with /Kids keeps the kids that survive, and is dropped only when none do — a
 *   field whose every widget was deleted is gone. A node that is also a kept annotation is
 *   such a node too; a dropped annotation is dropped whatever its kids;
 * - any other node — inline, without kids, or not a dictionary — is kept as it is: removing a
 *   page says nothing about it;
 * - a /Kids cycle of kept nodes keeps itself: survival is the greatest fixpoint of "some kid
 *   survives", which is what walking the graph and taking a node met again while it is still
 *   being walked as kept gives, in whatever order it is walked;
 * - the decisions about a node with kids apply only where the walk from /Fields reaches it
 *   through kept nodes; one reachable only through a dropped annotation is left as it is.
 *
 * None of this depends on the edits, only on the document and on which pages are kept, so
 * the graph is read once per document ({@link buildPageIndex}). Survival is then evaluated
 * over sets of "labels": a split labels each annotation with the pages listing it, which
 * answers every one-page selection at once; a removal labels it `0` when a kept page lists
 * it, which answers that selection. Both run the same evaluation, so there is one rule.
 */
import type { PdfDocument } from "@pdf/reader/pdf-document";
import { isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfDictValue, PdfObject, PdfRef } from "@pdf/reader/pdf-parser";

type PageInfo = { readonly dict: PdfDictValue; readonly objNum: number; readonly gen: number };

/** The AcroForm field graph as read: /Fields, and the /Kids of every node reached from it. */
interface FieldGraph {
  readonly roots: readonly PdfObject[];
  /**
   * The /Kids of a node when it is a non-empty array, read on first use: a removal never
   * looks beneath a dropped annotation, so it need not parse one.
   */
  kidsOf(key: string, ref: PdfRef): readonly PdfObject[] | undefined;
}

/** The page ownership of a document's objects, read once and shared by every write. */
export interface PageIndex {
  /** `"objNum:gen"` of each page object → the indices it appears at in the page tree. */
  readonly pagesByKey: ReadonlyMap<string, readonly number[]>;
  /** For each page, `"objNum:gen"` of each annotation it lists by reference, once each. */
  readonly annotsByPage: ReadonlyArray<readonly string[]>;
  /** `"objNum:gen"` of each annotation a page lists by reference → the last page listing it. */
  readonly lastListedOn: ReadonlyMap<string, number>;
  /** The field graph; `undefined` when the document has no /Fields array to prune. */
  readonly fields: FieldGraph | undefined;
}

/** What a full rewrite keeping some of the original pages leaves out. */
export interface Pruning {
  /** Whether an object (by `"objNum:gen"`) is left out, its references written as `null`. */
  isDropped(key: string): boolean;
  /** The /Kids written for a field node (by `"objNum:gen"`), when it loses some. */
  kidsOverride(key: string): PdfObject[] | undefined;
  /** The /Fields written; `undefined` when no annotation is dropped and the form is as read. */
  fields(): PdfObject[] | undefined;
}

const refKey = (ref: { objNum: number; gen: number }): string => `${ref.objNum}:${ref.gen}`;

/** Read the page ownership of `doc`'s objects. */
export function buildPageIndex(doc: PdfDocument, pagesInfo: readonly PageInfo[]): PageIndex {
  const pagesByKey = new Map<string, number[]>();
  /** Each annotation → the last page that listed it, to list it once per page. */
  const lastPage = new Map<string, number>();
  const annotsByPage = pagesInfo.map(({ dict, objNum, gen }, i) => {
    if (objNum > 0) {
      const key = refKey({ objNum, gen });
      const at = pagesByKey.get(key);
      if (at) {
        at.push(i);
      } else {
        pagesByKey.set(key, [i]);
      }
    }
    const annots = doc.deref(dict.get("Annots") ?? null);
    const keys: string[] = [];
    if (isPdfArray(annots)) {
      for (const ref of annots) {
        if (!isPdfRef(ref)) {
          continue;
        }
        const key = refKey(ref);
        if (lastPage.get(key) !== i) {
          lastPage.set(key, i);
          keys.push(key);
        }
      }
    }
    return keys;
  });
  return { pagesByKey, annotsByPage, lastListedOn: lastPage, fields: readFieldGraph(doc) };
}

function readFieldGraph(doc: PdfDocument): FieldGraph | undefined {
  let roots: PdfObject | null;
  try {
    const acroForm = doc.derefDict(doc.getCatalog().get("AcroForm") ?? null);
    roots = acroForm ? doc.deref(acroForm.get("Fields") ?? null) : null;
  } catch {
    return undefined;
  }
  if (!isPdfArray(roots)) {
    return undefined;
  }
  const read = new Map<string, readonly PdfObject[] | undefined>();
  return {
    roots,
    kidsOf(key, ref) {
      if (read.has(key)) {
        return read.get(key);
      }
      const node = doc.derefDict(ref);
      const list = node ? doc.deref(node.get("Kids") ?? null) : null;
      const kids = isPdfArray(list) && list.length > 0 ? list : undefined;
      read.set(key, kids);
      return kids;
    }
  };
}

/** The labels a node survives for: every label, or the ones in the set. */
type Labels = "all" | ReadonlySet<number>;

const NONE: ReadonlySet<number> = new Set();
const KEPT: ReadonlySet<number> = new Set([0]);

const has = (labels: Labels, label: number): boolean => labels === "all" || labels.has(label);

function meet(a: Labels, b: Labels): Labels {
  if (a === "all") {
    return b;
  }
  if (b === "all") {
    return a;
  }
  const [small, large] = a.size <= b.size ? [a, b] : [b, a];
  const out = new Set<number>();
  for (const label of small) {
    if (large.has(label)) {
      out.add(label);
    }
  }
  return out;
}

function sameLabels(a: Labels | undefined, b: Labels): boolean {
  if (a === "all" || b === "all" || a === undefined) {
    return a === b;
  }
  if (a.size !== b.size) {
    return false;
  }
  for (const label of a) {
    if (!b.has(label)) {
      return false;
    }
  }
  return true;
}

/** Which of a node's kids survive for each label, as indices into its /Kids, ascending. */
interface KidsByLabel {
  readonly always: readonly number[];
  readonly byLabel: ReadonlyMap<number, readonly number[]>;
}

/** The survival of the field nodes for one labelling of the annotations. */
interface Survival {
  /** The labels each node the walk from /Fields gets to survives for, by `"objNum:gen"`. */
  readonly survives: ReadonlyMap<string, Labels>;
  /**
   * The labels for which that walk reaches each of those nodes through kept nodes; `undefined`
   * when it reaches each of them for every label.
   */
  readonly reached: ReadonlyMap<string, Labels> | undefined;
  /** The /Kids of each of those nodes. */
  readonly kids: ReadonlyMap<string, readonly PdfObject[]>;
  /** Which kids of `kids` survive for each label (cached per list). */
  kidsByLabel(kids: readonly PdfObject[]): KidsByLabel;
}

/**
 * Evaluate the rules of the module comment for `labelsOf`, which gives an annotation's labels
 * (`undefined` for an object no page lists). An annotation with no labels is dropped whatever
 * is beneath it, so the walk does not go there.
 */
function evaluate(
  graph: FieldGraph,
  labelsOf: (key: string) => ReadonlySet<number> | undefined
): Survival {
  const survives = new Map<string, Labels>();
  const nodeKids = new Map<string, readonly PdfObject[]>();
  /** The /Kids the walk descends into from `ref`; `undefined` for a leaf or a dropped annotation. */
  const descend = (key: string, ref: PdfRef): readonly PdfObject[] | undefined =>
    labelsOf(key)?.size === 0 ? undefined : graph.kidsOf(key, ref);
  const valueOf = (kid: PdfObject): Labels => {
    if (!isPdfRef(kid)) {
      return "all";
    }
    const key = refKey(kid);
    const own = labelsOf(key);
    if (own?.size === 0) {
      return NONE;
    }
    return graph.kidsOf(key, kid) ? (survives.get(key) ?? "all") : (own ?? "all");
  };
  const evalNode = (key: string): Labels => {
    const acc = new Set<number>();
    for (const kid of nodeKids.get(key)!) {
      const value = valueOf(kid);
      if (value === "all") {
        return labelsOf(key) ?? "all";
      }
      for (const label of value) {
        acc.add(label);
      }
    }
    return meet(acc, labelsOf(key) ?? "all");
  };

  // Post-order from /Fields with an explicit stack: the depth of a field tree is the file's to
  // choose. A node met while it is being walked is taken to survive for now, which
  // over-estimates; when that happened, iterating from there settles on the greatest fixpoint.
  const order: string[] = [];
  const inProgress = new Set<string>();
  let cyclic = false;
  /** Whether an annotation has kids: only beneath one can the walk be stopped. */
  let gated = false;
  interface Frame {
    key: string | undefined;
    kids: readonly PdfObject[];
    next: number;
  }
  const stack: Frame[] = [{ key: undefined, kids: graph.roots, next: 0 }];
  while (stack.length > 0) {
    const top = stack[stack.length - 1];
    if (top.next < top.kids.length) {
      const kid = top.kids[top.next++];
      if (!isPdfRef(kid)) {
        continue;
      }
      const key = refKey(kid);
      const kids = descend(key, kid);
      if (!kids) {
        continue;
      }
      gated ||= labelsOf(key) !== undefined;
      if (survives.has(key)) {
        continue;
      }
      if (inProgress.has(key)) {
        cyclic = true;
        continue;
      }
      inProgress.add(key);
      nodeKids.set(key, kids);
      stack.push({ key, kids, next: 0 });
      continue;
    }
    stack.pop();
    if (top.key !== undefined) {
      inProgress.delete(top.key);
      survives.set(top.key, evalNode(top.key));
      order.push(top.key);
    }
  }
  while (cyclic) {
    cyclic = false;
    for (const key of order) {
      const value = evalNode(key);
      if (!sameLabels(survives.get(key), value)) {
        survives.set(key, value);
        cyclic = true;
      }
    }
  }

  // For which labels the walk gets to each node: through every node that is not dropped. With
  // no annotation that has kids, it gets to every node it evaluated, for every label.
  const reached = gated ? new Map<string, Labels>() : undefined;
  const queue: string[] = [];
  const reach = (kid: PdfObject, labels: Labels): void => {
    if (!reached || !isPdfRef(kid) || (labels !== "all" && labels.size === 0)) {
      return;
    }
    const key = refKey(kid);
    if (!descend(key, kid)) {
      return;
    }
    const current = reached.get(key);
    if (current === "all") {
      return;
    }
    let grown: Labels;
    if (labels === "all") {
      grown = "all";
    } else {
      const merged = new Set(current ?? NONE);
      for (const label of labels) {
        merged.add(label);
      }
      if (current && merged.size === current.size) {
        return;
      }
      grown = merged;
    }
    reached.set(key, grown);
    queue.push(key);
  };
  for (const root of graph.roots) {
    reach(root, "all");
  }
  while (queue.length > 0) {
    const key = queue.pop()!;
    const through = meet(reached!.get(key)!, labelsOf(key) ?? "all");
    for (const kid of nodeKids.get(key)!) {
      reach(kid, through);
    }
  }

  const cache = new Map<readonly PdfObject[], KidsByLabel>();
  return {
    survives,
    reached,
    kids: nodeKids,
    kidsByLabel(kids) {
      let found = cache.get(kids);
      if (!found) {
        const always: number[] = [];
        const byLabel = new Map<number, number[]>();
        kids.forEach((kid, i) => {
          const value = valueOf(kid);
          if (value === "all") {
            always.push(i);
            return;
          }
          for (const label of value) {
            const list = byLabel.get(label);
            if (list) {
              list.push(i);
            } else {
              byLabel.set(label, [i]);
            }
          }
        });
        found = { always, byLabel };
        cache.set(kids, found);
      }
      return found;
    }
  };
}

/** What a split works out once for all of its parts: the pages listing each annotation, and the survival of every one-page selection. */
const perPage = new WeakMap<
  PageIndex,
  { annotPages: ReadonlyMap<string, ReadonlySet<number>>; survival?: Survival }
>();

/** What a part of `splitPages` keeping only original page `page` leaves out; `-1` keeps none. */
export function pruneToPage(index: PageIndex, page: number): Pruning {
  let shared = perPage.get(index);
  if (!shared) {
    const annotPages = new Map<string, Set<number>>();
    index.annotsByPage.forEach((keys, i) => {
      for (const key of keys) {
        const pages = annotPages.get(key);
        if (pages) {
          pages.add(i);
        } else {
          annotPages.set(key, new Set([i]));
        }
      }
    });
    shared = { annotPages };
    perPage.set(index, shared);
  }
  const { annotPages } = shared;
  const removesAnnots = (index.annotsByPage[page]?.length ?? 0) < index.lastListedOn.size;
  const keepsAnnot = (key: string): boolean => annotPages.get(key)?.has(page) ?? true;
  return makePruning(
    index,
    at => at === page,
    keepsAnnot,
    removesAnnots,
    page,
    graph => {
      shared.survival ??= evaluate(graph, key => annotPages.get(key));
      return shared.survival;
    }
  );
}

/** What a full rewrite leaving out the original pages in `removed` leaves out. */
export function pruneRemovedPages(index: PageIndex, removed: ReadonlySet<number>): Pruning {
  const keeps = (at: number): boolean => !removed.has(at);
  const kept = new Set<string>();
  index.annotsByPage.forEach((keys, i) => {
    if (keeps(i)) {
      for (const key of keys) {
        kept.add(key);
      }
    }
  });
  const removesAnnots = kept.size < index.lastListedOn.size;
  // A dropped annotation is one some page lists and no kept page does.
  const keepsAnnot = (key: string): boolean => kept.has(key) || !index.lastListedOn.has(key);
  return makePruning(index, keeps, keepsAnnot, removesAnnots, 0, graph =>
    evaluate(graph, key => (kept.has(key) ? KEPT : index.lastListedOn.has(key) ? NONE : undefined))
  );
}

/**
 * The answers for one selection of pages: `keeps` says which original pages it keeps,
 * `keepsAnnot` which annotations (by `"objNum:gen"`; one no page lists is kept), and `label`
 * is the label `survival` gives what this selection keeps.
 */
function makePruning(
  index: PageIndex,
  keeps: (page: number) => boolean,
  keepsAnnot: (key: string) => boolean,
  removesAnnots: boolean,
  label: number,
  survival: (graph: FieldGraph) => Survival
): Pruning {
  const graph = removesAnnots ? index.fields : undefined;
  const evaluated = graph ? survival(graph) : undefined;
  const keptKids = (kids: readonly PdfObject[]): PdfObject[] => {
    const { always, byLabel } = evaluated!.kidsByLabel(kids);
    const other = byLabel.get(label) ?? [];
    const out: PdfObject[] = [];
    let i = 0;
    let j = 0;
    while (i < always.length || j < other.length) {
      const at =
        j >= other.length || (i < always.length && always[i] < other[j]) ? always[i++] : other[j++];
      out.push(kids[at]);
    }
    return out;
  };
  /** Whether the walk from /Fields reaches node `key`, and with which outcome. */
  const visit = (key: string): "kept" | "dropped" | undefined => {
    const reached = evaluated?.reached;
    if (!evaluated?.survives.has(key) || (reached && !has(reached.get(key) ?? NONE, label))) {
      return undefined;
    }
    return has(evaluated.survives.get(key)!, label) ? "kept" : "dropped";
  };
  return {
    isDropped(key) {
      if (index.pagesByKey.get(key)?.some(at => !keeps(at))) {
        return true;
      }
      return removesAnnots && (!keepsAnnot(key) || visit(key) === "dropped");
    },
    kidsOverride(key) {
      if (visit(key) !== "kept") {
        return undefined;
      }
      const kids = evaluated!.kids.get(key)!;
      const kept = keptKids(kids);
      return kept.length < kids.length ? kept : undefined;
    },
    fields() {
      return graph ? keptKids(graph.roots) : undefined;
    }
  };
}
