/**
 * The document-level entries of the catalog, as a full rewrite of `PdfEditor` writes them.
 *
 * A full rewrite builds a new catalog. It used to carry over only `/AcroForm`, `/Metadata`
 * and `/NeedsRendering`, so every rewrite — `toBuffer()`, `removePage`, each part of
 * `splitPages` — lost the bookmarks, the named destinations, the page labels, the optional
 * content configuration, the output intent, the language and every other catalog entry. Most
 * of them do not depend on which pages are written and are copied as they are. The ones that
 * name pages are rewritten for the pages that are:
 *
 * - **`/Outlines`** — an item whose destination is a page that is left out loses the
 *   destination; it is kept, without one, when an item beneath it survives (its children stay
 *   where they were in the hierarchy), and dropped otherwise. An item with no page destination
 *   (none, an unresolvable one, a URI or other action) is kept by a rewrite and dropped from a
 *   part of a split, which keeps only the items that lead to its page and their ancestors.
 *   `/Count`, `/First`, `/Last`, `/Prev` and `/Next` are recomputed, and an item's open or
 *   closed state is kept.
 * - **`/Dests`** and **`/Names /Dests`** — named destinations on pages that are left out, or
 *   on no page at all, are removed. The name tree is written as a single sorted leaf.
 * - **`/PageLabels`** — every page keeps the label it had; a page added to the document
 *   continues the numbering of the page before it.
 * - **`/OpenAction`** — dropped when it goes to a page that is left out.
 * - **`/ViewerPreferences /PrintPageRange`** — renumbered to the pages written.
 * - **`/Threads`** — a bead on a page that is left out is removed and the beads that remain are
 *   relinked into the thread's ring (`/N`, `/V`; `/F` and the first bead's `/T` follow the new
 *   first bead); a thread left with no bead is dropped (see {@link planThreads}).
 *
 * Some entries describe the pages as a whole and are not rewritten piecemeal: the structure
 * tree (`/StructTreeRoot`, with `/MarkInfo` and each page's `/StructParents`) and document
 * parts (`/DPartRoot`) are kept when every original page is written and dropped when one is
 * left out, since their elements would otherwise name content that is gone. `/Perms` and
 * `/Legal` are always dropped: they are about signatures, and a full rewrite invalidates every
 * signature — keeping a usage-rights signature makes a viewer report the document as altered.
 *
 * What leads where is worked out once per document ({@link buildCatalogIndex}), like the page
 * ownership `splitPages` uses: a part of a split then looks only at what concerns its own page,
 * so splitting a document with a bookmark per page does not walk the whole outline per part.
 */
import { pdfNameBytes, pdfNumber, pdfRef } from "@pdf/core/pdf-object";
import type { PdfDocument } from "@pdf/reader/pdf-document";
import { dictGetName, isPdfArray, isPdfRef } from "@pdf/reader/pdf-parser";
import type { PdfDictValue, PdfObject, PdfRef } from "@pdf/reader/pdf-parser";

/** What the editor's full rewrite provides to write the catalog's entries. */
export interface CatalogRewriteHost {
  readonly doc: PdfDocument;
  /**
   * The pages written, in order: the index of an original page, or `null` for a page added
   * with `addPage` or `copyPagesFrom`.
   */
  readonly outputPages: ReadonlyArray<number | null>;
  /** Whether this is a part of `splitPages` rather than the whole document. */
  readonly split: boolean;
  /** The article threads as this write leaves them (see {@link planThreads}). */
  readonly threads: ThreadPlan | null;
  /** Clone a value into the new file, as every other copied object is. */
  serialize(value: PdfObject): string;
  allocObject(): number;
  addObject(objNum: number, body: string): void;
  /** Make every later clone of a reference to `ref` point at `objNum`. */
  mapRef(ref: PdfRef, objNum: number): void;
  /**
   * Make every later clone of a reference to an original object (`"objNum:gen"`) the predicate
   * holds for, and that {@link mapRef} did not map, a `null`.
   */
  dropRefs(isDropped: (key: string) => boolean): void;
}

/** Catalog entries the editor writes itself. */
const WRITTEN_BY_EDITOR = new Set(["Type", "Pages", "AcroForm", "NeedsRendering", "Metadata"]);

/** Catalog entries a full rewrite never carries over (see the module comment). */
const ALWAYS_DROPPED = new Set(["Perms", "Legal"]);

/** Catalog entries kept only when every original page is written. */
const KEPT_WITH_EVERY_PAGE = new Set(["StructTreeRoot", "DPartRoot"]);

/** Outline item keys the rewrite recomputes. */
const OUTLINE_LINKS = new Set(["Parent", "Prev", "Next", "First", "Last", "Count"]);

/** Where a destination leads, in the document as it was read. */
type Target =
  /** Pages of this document: the original indices of the page it names. */
  | { readonly kind: "page"; readonly pages: readonly number[] }
  /** No destination at all. */
  | { readonly kind: "none" }
  /** An action that goes to no page here: a URI, a script, another file. */
  | { readonly kind: "action" }
  /** A destination that names no page: a missing name, a reference to something else. */
  | { readonly kind: "dead" };

const NONE: Target = { kind: "none" };
const ACTION: Target = { kind: "action" };
const DEAD: Target = { kind: "dead" };

interface IndexedOutlineItem {
  readonly ref: PdfRef | undefined;
  readonly dict: PdfDictValue;
  /** Index of the parent item in {@link OutlineIndex.items}; `-1` for a top-level item. */
  readonly parent: number;
  /** Indices of the children, in order. */
  readonly children: readonly number[];
  readonly open: boolean;
  readonly target: Target;
}

interface OutlineIndex {
  readonly root: PdfObject;
  /** Every item, each before its children, siblings in order. */
  readonly items: readonly IndexedOutlineItem[];
  /** The items whose destination is each original page. */
  readonly byPage: ReadonlyMap<number, readonly number[]>;
  /** Every outline object, by `"objNum:gen"`. */
  readonly keys: ReadonlySet<string>;
}

interface NamedEntry<K> {
  readonly key: K;
  readonly value: PdfObject;
  readonly target: Target;
}

interface NamedIndex<K> {
  /** Every entry, in the order written: the tree's sorted by key, the dictionary's as read. */
  readonly entries: ReadonlyArray<NamedEntry<K>>;
  /** Indices into {@link entries} of those leading to each original page. */
  readonly byPage: ReadonlyMap<number, readonly number[]>;
}

/** One action of a chain that is left (see `RewriteContext.actionChain`). */
interface ActionNode {
  readonly dict: PdfDictValue;
  /** The actions of its `/Next` that are left, in order. */
  readonly next: readonly ActionNode[];
  /** For a `GoTo`, the output index of its page. */
  readonly landing: number | undefined;
  /** Whether it or its `/Next` holds a `GoTo`, so it must be rewritten. */
  readonly changes: boolean;
}

interface LabelRange {
  readonly start: number;
  readonly dict: PdfDictValue;
  /** The value of the range's first page. */
  readonly first: number;
}

interface BeadEntry {
  readonly ref: PdfRef;
  /** The original pages its `/P` is at; empty when it names no page of the document. */
  readonly pages: readonly number[];
}

interface ThreadEntry {
  /** The entry of `/Threads`: a reference to the thread, or the thread itself. */
  readonly value: PdfObject;
  readonly dict: PdfDictValue;
  /** The beads, in thread order from `/F` along `/N`, each once. */
  readonly beads: readonly BeadEntry[];
  /** Positions in {@link beads} of the beads on each original page, ascending. */
  readonly byPage: ReadonlyMap<number, readonly number[]>;
}

interface ThreadIndex {
  readonly threads: readonly ThreadEntry[];
  /** Indices into {@link threads} of those with a bead on each original page, ascending. */
  readonly byPage: ReadonlyMap<number, readonly number[]>;
  /** Each bead (`"objNum:gen"`) → its thread and position. */
  readonly beads: ReadonlyMap<string, { readonly thread: number; readonly at: number }>;
  /** Each thread written by reference (`"objNum:gen"`) → its index. */
  readonly byKey: ReadonlyMap<string, number>;
}

/** What leads where in a document's catalog, worked out once for every write. */
export interface CatalogIndex {
  readonly pageCount: number;
  /** Original page object (`"objNum:gen"`) → the original indices it is at. */
  readonly pageIndices: ReadonlyMap<string, readonly number[]>;
  /** Name → value of `/Dests` and `/Names /Dests`, for resolving a named destination. */
  readonly lookupName: (name: string | Uint8Array) => PdfObject | undefined;
  outline?: OutlineIndex | null;
  destsTree?: NamedIndex<Uint8Array> | null;
  destsDict?: NamedIndex<string> | null;
  labels?: readonly LabelRange[] | null;
  threads?: ThreadIndex | null;
}

/** Index a document's catalog. The parts that are expensive to read are filled on first use. */
export function buildCatalogIndex(
  doc: PdfDocument,
  catalog: PdfDictValue,
  pagesInfo: ReadonlyArray<{ readonly objNum: number; readonly gen: number }>
): CatalogIndex {
  const pageIndices = new Map<string, number[]>();
  pagesInfo.forEach(({ objNum, gen }, i) => {
    if (objNum > 0) {
      const key = `${objNum}:${gen}`;
      const list = pageIndices.get(key);
      if (list) {
        list.push(i);
      } else {
        pageIndices.set(key, [i]);
      }
    }
  });
  let lookup: ((name: string | Uint8Array) => PdfObject | undefined) | undefined;
  return {
    pageCount: pagesInfo.length,
    pageIndices,
    lookupName: name => {
      lookup ??= namedLookup(doc, catalog);
      return lookup(name);
    }
  };
}

/**
 * The catalog entries of a full rewrite other than {@link WRITTEN_BY_EDITOR}, in the
 * original's order, as `[key, serialized value]`.
 */
export function rewriteCatalogEntries(
  host: CatalogRewriteHost,
  catalog: PdfDictValue,
  index: CatalogIndex
): Array<[string, string]> {
  const ctx = new RewriteContext(host, index);
  const out: Array<[string, string]> = [];
  for (const [key, value] of catalog) {
    if (WRITTEN_BY_EDITOR.has(key) || ALWAYS_DROPPED.has(key)) {
      continue;
    }
    if (KEPT_WITH_EVERY_PAGE.has(key) && !ctx.keepsStructure) {
      continue;
    }
    if (key === "MarkInfo" && !ctx.keepsStructure && catalog.has("StructTreeRoot")) {
      continue;
    }
    let serialized: string | undefined;
    try {
      serialized = ctx.rewrite(key, value);
    } catch {
      // An entry that cannot be read is left out, like any unreadable object in a rewrite.
      serialized = undefined;
    }
    if (serialized !== undefined && serialized !== "" && serialized !== "null") {
      out.push([key, serialized]);
    }
  }
  return out;
}

/**
 * Whether a full rewrite keeps the structure tree: every original page is written. A page's
 * `/StructParents` is carried over exactly when this holds.
 */
export function keepsStructureTree(
  pageCount: number,
  outputPages: ReadonlyArray<number | null>
): boolean {
  if (outputPages.length < pageCount) {
    return false;
  }
  const written = new Set<number>();
  for (const page of outputPages) {
    if (page !== null) {
      written.add(page);
    }
  }
  return written.size === pageCount;
}

/** Whether `outputPages` writes every original page, each at the index it had. */
function writesPagesInPlace(pageCount: number, outputPages: ReadonlyArray<number | null>): boolean {
  if (outputPages.length < pageCount) {
    return false;
  }
  for (let i = 0; i < pageCount; i++) {
    if (outputPages[i] !== i) {
      return false;
    }
  }
  return true;
}

class RewriteContext {
  readonly keepsStructure: boolean;
  /** Original page index → its index in the output. */
  private readonly newIndex = new Map<number, number>();
  /** Whether every original page is written, at the index it had. */
  private readonly pagesUnchanged: boolean;
  /** The one original page a part of a split writes, if it is an original page. */
  private readonly splitPage: number | undefined;

  constructor(
    private readonly host: CatalogRewriteHost,
    private readonly index: CatalogIndex
  ) {
    host.outputPages.forEach((page, i) => {
      if (page !== null && !this.newIndex.has(page)) {
        this.newIndex.set(page, i);
      }
    });
    this.keepsStructure = keepsStructureTree(index.pageCount, host.outputPages);
    this.pagesUnchanged = writesPagesInPlace(index.pageCount, host.outputPages);
    this.splitPage = host.split ? (host.outputPages[0] ?? undefined) : undefined;
  }

  private get doc(): PdfDocument {
    return this.host.doc;
  }

  rewrite(key: string, value: PdfObject): string | undefined {
    switch (key) {
      case "Outlines":
        return this.outlines(value);
      case "Dests":
        return this.destsDict(value);
      case "Names":
        return this.names(value);
      case "PageLabels":
        return this.pageLabels(value);
      case "OpenAction":
        return this.openAction(value);
      case "ViewerPreferences":
        return this.viewerPreferences(value);
      case "Threads":
        return this.threads(value);
      default:
        return this.host.serialize(value);
    }
  }

  // ===========================================================================
  // Destinations
  // ===========================================================================

  /** Where a target lands in the output: the page's new index, or `undefined`. */
  private landing(target: Target): number | undefined {
    if (target.kind !== "page") {
      return undefined;
    }
    for (const page of target.pages) {
      const at = this.newIndex.get(page);
      if (at !== undefined) {
        return at;
      }
    }
    return undefined;
  }

  /** A destination whose target is written, serialized for the output. */
  private destination(dest: PdfObject, newIndex: number): string {
    const value = this.doc.deref(dest);
    if (isPdfArray(value) && typeof value[0] === "number") {
      const rest = value.slice(1).map(v => this.host.serialize(v));
      return `[${[pdfNumber(newIndex), ...rest].join(" ")}]`;
    }
    if (value instanceof Map) {
      // `<< /D … >>`, as a named destination's value may be.
      return this.dictWith(value, "D", entry => this.destination(entry, newIndex));
    }
    return this.host.serialize(isPdfArray(value) ? value : dest);
  }

  /**
   * The actions of `action` and its `/Next` that are left, in the order they run (§12.6.2): a
   * `GoTo` whose page is not written is replaced by its own `/Next`, so the rest of the chain
   * still runs. `path` guards against a `/Next` cycle.
   */
  private actionChain(action: PdfObject, path = new Set<PdfDictValue>()): ActionNode[] {
    const dict = this.doc.derefDict(action);
    if (!dict || path.has(dict)) {
      return [];
    }
    path.add(dict);
    const nextValue = dict.get("Next");
    const resolved = this.doc.deref(nextValue ?? null);
    const next = isPdfArray(resolved)
      ? resolved.flatMap(item => this.actionChain(item, path))
      : nextValue === undefined
        ? []
        : this.actionChain(nextValue, path);
    path.delete(dict);
    if (dictGetName(dict, "S") !== "GoTo") {
      return [{ dict, next, landing: undefined, changes: next.some(n => n.changes) }];
    }
    const landing = this.landing(destTarget(this.doc, this.index, dict.get("D")));
    return landing === undefined ? next : [{ dict, next, landing, changes: true }];
  }

  /**
   * An action entry as written: unchanged when its chain has no `GoTo`, otherwise rebuilt from
   * {@link actionChain} — its first action, with the rest as `/Next` — or `undefined` when no
   * action is left.
   */
  private rewriteAction(action: PdfObject): string | undefined {
    const chain = this.actionChain(action);
    if (chain.length === 1 && !chain[0].changes && chain[0].dict === this.doc.derefDict(action)) {
      return this.host.serialize(action);
    }
    return chain.length === 0 ? undefined : this.writeActions(chain);
  }

  /** `nodes` as one action: the first, followed by its own `/Next` and then the others. */
  private writeActions(nodes: readonly ActionNode[]): string {
    const [first, ...rest] = nodes;
    const next = [...first.next, ...rest];
    const parts = ["<<"];
    for (const [key, entry] of first.dict) {
      if (key === "Next") {
        continue;
      }
      const value =
        key === "D" && first.landing !== undefined
          ? this.destination(entry, first.landing)
          : this.host.serialize(entry);
      parts.push(`${pdfNameBytes(key)} ${value}`);
    }
    if (next.length === 1) {
      parts.push(`/Next ${this.writeActions(next)}`);
    } else if (next.length > 1) {
      parts.push(`/Next [${next.map(node => this.writeActions([node])).join(" ")}]`);
    }
    parts.push(">>");
    return parts.join(" ");
  }

  /** `dict` serialized, with the value of `key` serialized by `rewrite`. */
  private dictWith(dict: PdfDictValue, key: string, rewrite: (entry: PdfObject) => string): string {
    const parts = ["<<"];
    for (const [k, entry] of dict) {
      parts.push(`${pdfNameBytes(k)} ${k === key ? rewrite(entry) : this.host.serialize(entry)}`);
    }
    parts.push(">>");
    return parts.join(" ");
  }

  /** The indices of a named index's entries that land, in order. */
  private landingEntries<K>(named: NamedIndex<K>): number[] {
    if (this.host.split) {
      return this.splitPage === undefined ? [] : [...(named.byPage.get(this.splitPage) ?? [])];
    }
    const out: number[] = [];
    named.entries.forEach((entry, i) => {
      if (this.landing(entry.target) !== undefined) {
        out.push(i);
      }
    });
    return out;
  }

  // ===========================================================================
  // Named destinations
  // ===========================================================================

  private destsDict(value: PdfObject): string | undefined {
    const named = (this.index.destsDict ??= indexDestsDict(this.doc, this.index, value));
    if (!named) {
      return undefined;
    }
    const parts = this.landingEntries(named).map(i => {
      const { key, value: dest, target } = named.entries[i];
      return `${pdfNameBytes(key)} ${this.destination(dest, this.landing(target)!)}`;
    });
    if (parts.length === 0) {
      return undefined;
    }
    const objNum = this.host.allocObject();
    this.host.addObject(objNum, `<< ${parts.join(" ")} >>`);
    return pdfRef(objNum);
  }

  private names(value: PdfObject): string | undefined {
    const dict = this.doc.derefDict(value);
    if (!dict) {
      return undefined;
    }
    const parts: string[] = [];
    for (const [key, entry] of dict) {
      const serialized = key === "Dests" ? this.destsTree(entry) : this.host.serialize(entry);
      if (serialized !== undefined && serialized !== "null") {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }
    return parts.length === 0 ? undefined : `<< ${parts.join(" ")} >>`;
  }

  /** The `/Names /Dests` tree, as one sorted leaf of the destinations that land. */
  private destsTree(root: PdfObject): string | undefined {
    const named = (this.index.destsTree ??= indexDestsTree(this.doc, this.index, root));
    const kept = this.landingEntries(named);
    if (kept.length === 0) {
      return undefined;
    }
    const body = kept
      .map(i => {
        const { key, value: dest, target } = named.entries[i];
        return `${this.host.serialize(key)} ${this.destination(dest, this.landing(target)!)}`;
      })
      .join(" ");
    const objNum = this.host.allocObject();
    this.host.addObject(objNum, `<< /Names [${body}] >>`);
    return pdfRef(objNum);
  }

  // ===========================================================================
  // Outlines
  // ===========================================================================

  private outlines(value: PdfObject): string | undefined {
    const outline = (this.index.outline ??= indexOutline(this.doc, this.index, value));
    if (!outline) {
      return undefined;
    }
    const { items } = outline;
    // Every original outline object not written below is left out: a reference to one from
    // elsewhere is written as `null` rather than cloning the original outline back in.
    this.host.dropRefs(key => outline.keys.has(key));

    /** The surviving items, each before its children, and the children each keeps. */
    const children = new Map<number, number[]>();
    const survivors: number[] = [];
    if (this.host.split) {
      // The items on the part's page, and their ancestors.
      const kept = new Set<number>();
      for (const i of this.splitPage === undefined
        ? []
        : (outline.byPage.get(this.splitPage) ?? [])) {
        for (let at = i; at >= 0 && !kept.has(at); at = items[at].parent) {
          kept.add(at);
        }
      }
      survivors.push(...[...kept].sort((a, b) => a - b));
    } else {
      // Post-order: an item follows its parent in `items`, so walking backwards settles
      // every child first.
      const survives = new Array<boolean>(items.length).fill(false);
      for (let i = items.length - 1; i >= 0; i--) {
        const item = items[i];
        survives[i] =
          this.landing(item.target) !== undefined ||
          item.target.kind !== "page" ||
          // A GoTo whose page is gone, with more of its chain left to run.
          (!item.dict.has("Dest") &&
            item.dict.has("A") &&
            this.actionChain(item.dict.get("A")!).length > 0) ||
          item.children.some(child => survives[child]);
      }
      items.forEach((_, i) => survives[i] && survivors.push(i));
    }
    const top: number[] = [];
    for (const i of survivors) {
      const parent = items[i].parent;
      if (parent < 0) {
        top.push(i);
      } else {
        const list = children.get(parent);
        if (list) {
          list.push(i);
        } else {
          children.set(parent, [i]);
        }
      }
    }
    if (top.length === 0) {
      return undefined;
    }

    // Descendants shown when each item is open, settled children first.
    const visible = new Map<number, number>();
    for (let k = survivors.length - 1; k >= 0; k--) {
      const i = survivors[k];
      let count = 0;
      for (const child of children.get(i) ?? []) {
        count += 1 + (items[child].open ? visible.get(child)! : 0);
      }
      visible.set(i, count);
    }

    const objNums = new Map<number, number>();
    const rootNum = this.host.allocObject();
    if (isPdfRef(value)) {
      this.host.mapRef(value, rootNum);
    }
    for (const i of survivors) {
      const objNum = this.host.allocObject();
      objNums.set(i, objNum);
      const ref = items[i].ref;
      if (ref) {
        this.host.mapRef(ref, objNum);
      }
    }

    const writeLevel = (level: readonly number[], parentNum: number) => {
      level.forEach((i, position) => {
        const item = items[i];
        const kids = children.get(i) ?? [];
        const newIndex = this.landing(item.target);
        const parts = ["<<"];
        for (const [key, entry] of item.dict) {
          const serialized = OUTLINE_LINKS.has(key)
            ? undefined
            : this.itemEntry(item, key, entry, newIndex);
          if (serialized !== undefined) {
            parts.push(`${pdfNameBytes(key)} ${serialized}`);
          }
        }
        parts.push(`/Parent ${pdfRef(parentNum)}`);
        if (position > 0) {
          parts.push(`/Prev ${pdfRef(objNums.get(level[position - 1])!)}`);
        }
        if (position < level.length - 1) {
          parts.push(`/Next ${pdfRef(objNums.get(level[position + 1])!)}`);
        }
        if (kids.length > 0) {
          const count = visible.get(i)!;
          parts.push(`/First ${pdfRef(objNums.get(kids[0])!)}`);
          parts.push(`/Last ${pdfRef(objNums.get(kids[kids.length - 1])!)}`);
          parts.push(`/Count ${item.open ? count : -count}`);
        }
        parts.push(">>");
        this.host.addObject(objNums.get(i)!, parts.join(" "));
      });
    };
    writeLevel(top, rootNum);
    for (const [parent, level] of children) {
      writeLevel(level, objNums.get(parent)!);
    }

    let rootVisible = 0;
    for (const i of top) {
      rootVisible += 1 + (items[i].open ? visible.get(i)! : 0);
    }
    const first = pdfRef(objNums.get(top[0])!);
    const last = pdfRef(objNums.get(top[top.length - 1])!);
    // Table 152: the root's /Count is omitted when no item is open.
    const anyOpen = survivors.some(i => items[i].open && children.has(i));
    this.host.addObject(
      rootNum,
      `<< /Type /Outlines /First ${first} /Last ${last}${anyOpen ? ` /Count ${rootVisible}` : ""} >>`
    );
    return pdfRef(rootNum);
  }

  /**
   * One entry of a surviving outline item. Its destination — `/Dest`, or a `GoTo` `/A` — is
   * kept when it lands, rewritten for the output, and left out when it does not; any other
   * action is kept as it is.
   */
  private itemEntry(
    item: IndexedOutlineItem,
    key: string,
    entry: PdfObject,
    newIndex: number | undefined
  ): string | undefined {
    if (key === "SE") {
      return this.keepsStructure ? this.host.serialize(entry) : undefined;
    }
    if (key === "Dest") {
      return newIndex === undefined ? undefined : this.destination(entry, newIndex);
    }
    if (key !== "A") {
      return this.host.serialize(entry);
    }
    const action = this.doc.derefDict(entry);
    if (action && item.dict.has("Dest") && dictGetName(action, "S") === "GoTo") {
      // /Dest already says where the item goes (Table 153 allows only one of the two).
      return undefined;
    }
    // A URI, a script, another file goes where it went; a GoTo follows its page.
    return action ? this.rewriteAction(entry) : undefined;
  }

  // ===========================================================================
  // Page labels
  // ===========================================================================

  private pageLabels(value: PdfObject): string | undefined {
    if (this.pagesUnchanged && !this.host.split) {
      return this.host.serialize(value);
    }
    const ranges = (this.index.labels ??= indexPageLabels(this.doc, value));
    if (!ranges || ranges.length === 0) {
      return undefined;
    }
    /** The range governing original page `page`, and the page's number within it. */
    const labelOf = (page: number): { range: number; value: number } | undefined => {
      let lo = 0;
      let hi = ranges.length - 1;
      let found = -1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (ranges[mid].start <= page) {
          found = mid;
          lo = mid + 1;
        } else {
          hi = mid - 1;
        }
      }
      return found < 0
        ? undefined
        : { range: found, value: ranges[found].first + page - ranges[found].start };
    };

    const nums: string[] = [];
    let previous: { range: number; next: number } | undefined;
    this.host.outputPages.forEach((page, i) => {
      const label = page === null ? undefined : labelOf(page);
      if (!label) {
        // An added page, or one before every range: it continues the page before it.
        if (previous) {
          previous.next++;
        } else {
          nums.push(`${i} << /S /D /St ${i + 1} >>`);
          previous = { range: -1, next: i + 2 };
        }
        return;
      }
      if (previous && previous.range === label.range && previous.next === label.value) {
        previous.next++;
        return;
      }
      const parts = ["<<"];
      for (const [key, entry] of ranges[label.range].dict) {
        if (key !== "St") {
          parts.push(`${pdfNameBytes(key)} ${this.host.serialize(entry)}`);
        }
      }
      if (label.value !== 1) {
        parts.push(`/St ${label.value}`);
      }
      parts.push(">>");
      nums.push(`${i} ${parts.join(" ")}`);
      previous = { range: label.range, next: label.value + 1 };
    });
    return nums.length === 0 ? undefined : `<< /Nums [${nums.join(" ")}] >>`;
  }

  // ===========================================================================
  // Other entries that name pages
  // ===========================================================================

  private openAction(value: PdfObject): string | undefined {
    if (this.doc.deref(value) instanceof Map) {
      return this.rewriteAction(value);
    }
    const at = this.landing(destTarget(this.doc, this.index, value));
    return at === undefined ? undefined : this.destination(value, at);
  }

  private viewerPreferences(value: PdfObject): string | undefined {
    const dict = this.doc.derefDict(value);
    if (!dict) {
      return undefined;
    }
    const parts: string[] = [];
    for (const [key, entry] of dict) {
      const serialized =
        key === "PrintPageRange" && !this.pagesUnchanged
          ? this.printPageRange(entry)
          : this.host.serialize(entry);
      if (serialized !== undefined) {
        parts.push(`${pdfNameBytes(key)} ${serialized}`);
      }
    }
    return `<< ${parts.join(" ")} >>`;
  }

  /** `/PrintPageRange` (1-based pairs of first and last page), renumbered to the output. */
  private printPageRange(value: PdfObject): string | undefined {
    const pairs = this.doc.deref(value);
    if (!isPdfArray(pairs)) {
      return undefined;
    }
    const included = (page: number): boolean => {
      for (let i = 0; i + 1 < pairs.length; i += 2) {
        const first = this.doc.deref(pairs[i]);
        const last = this.doc.deref(pairs[i + 1]);
        if (
          typeof first === "number" &&
          typeof last === "number" &&
          first <= page &&
          page <= last
        ) {
          return true;
        }
      }
      return false;
    };
    // The output pages, 1-based, that were in a range; each original page is asked about once.
    const out: number[] = [];
    this.host.outputPages.forEach((page, i) => {
      if (page === null || this.newIndex.get(page) !== i || !included(page + 1)) {
        return;
      }
      if (out.length > 0 && out[out.length - 1] === i) {
        out[out.length - 1] = i + 1;
      } else {
        out.push(i + 1, i + 1);
      }
    });
    return out.length === 0 ? undefined : `[${out.join(" ")}]`;
  }

  /** `/Threads`, each thread with the beads on the pages written (see {@link planThreads}). */
  private threads(value: PdfObject): string | undefined {
    if (this.pagesUnchanged && !this.host.split) {
      return this.host.serialize(value);
    }
    // A `/Threads` that cannot be read cannot be followed to the pages that are left out.
    const plan = this.host.threads;
    if (!plan) {
      return undefined;
    }
    const written = plan.threadsWritten().map(i => {
      const { value: thread, dict } = plan.index.threads[i];
      if (isPdfRef(thread)) {
        // Cloned with the plan's `/F`.
        return this.host.serialize(thread);
      }
      const first = plan.survivors(i)[0];
      return this.dictWith(dict, "F", () => this.host.serialize(first.ref));
    });
    return written.length === 0 ? undefined : `[${written.join(" ")}]`;
  }
}

/** The article threads of a full rewrite that leaves original pages out, or moves them. */
export interface ThreadPlan {
  readonly index: ThreadIndex;
  /** Indices of the threads written: those with a bead left, in `/Threads` order. */
  threadsWritten(): readonly number[];
  /** The beads of thread `i` that are left, in thread order. */
  survivors(i: number): readonly BeadEntry[];
  /** Whether a bead or thread (by `"objNum:gen"`) is left out. */
  isDropped(key: string): boolean;
  /** The entries a bead or thread (by `"objNum:gen"`) is written with instead of its own. */
  overrides(key: string): ReadonlyMap<string, PdfObject> | undefined;
  /** Whether an object (by `"objNum:gen"`) is a bead of a thread in `/Threads`. */
  isBead(key: string): boolean;
}

/**
 * The article threads of a full rewrite writing `outputPages`; `null` when it writes every
 * original page where it was, which leaves every thread as it is.
 *
 * ISO 32000-1 §12.4.3: a thread's `/F` is its first bead, the beads form a ring through `/N`
 * (next; the last bead's is the first) and `/V` (previous; the first bead's is the last), each
 * bead's `/P` is its page, and the first bead carries `/T`, the thread. A bead is on a page
 * that is written or it is left out; the beads that remain are relinked in their order into a
 * ring of their own, so the thread still reads as it did over the pages that are left. A
 * thread with no bead left is left out; one that loses no bead is written as it is.
 */
export function planThreads(
  doc: PdfDocument,
  catalog: PdfDictValue,
  index: CatalogIndex,
  outputPages: ReadonlyArray<number | null>,
  split: boolean
): ThreadPlan | null {
  const value = catalog.get("Threads");
  if ((!split && writesPagesInPlace(index.pageCount, outputPages)) || value === undefined) {
    return null;
  }
  let threads: ThreadIndex | null;
  try {
    threads = index.threads ??= indexThreads(doc, index, value);
  } catch {
    threads = null;
  }
  if (!threads) {
    return null;
  }
  const ix = threads;
  const written = new Set<number>();
  for (const page of outputPages) {
    if (page !== null) {
      written.add(page);
    }
  }
  // A part of a split writes one page: what concerns it is looked up rather than searched for,
  // so a split does not walk every thread once per page.
  const only = split ? (outputPages[0] ?? null) : undefined;
  /** For each thread asked about: the positions of its beads that are left, in order. */
  const left = new Map<
    number,
    { readonly at: readonly number[]; readonly slot: Map<number, number> }
  >();
  const leftOf = (i: number) => {
    let found = left.get(i);
    if (!found) {
      const thread = ix.threads[i];
      const at =
        only !== undefined
          ? only === null
            ? []
            : (thread.byPage.get(only) ?? [])
          : thread.beads.flatMap((bead, k) =>
              bead.pages.some(page => written.has(page)) ? [k] : []
            );
      found = { at, slot: new Map(at.map((k, n) => [k, n])) };
      left.set(i, found);
    }
    return found;
  };
  const survivors = (i: number): readonly BeadEntry[] =>
    leftOf(i).at.map(k => ix.threads[i].beads[k]);
  const beadOf = (key: string) => ix.beads.get(key);
  return {
    index: ix,
    survivors,
    threadsWritten() {
      if (only !== undefined) {
        return only === null ? [] : (ix.byPage.get(only) ?? []);
      }
      return ix.threads.map((_, i) => i).filter(i => leftOf(i).at.length > 0);
    },
    isBead: key => ix.beads.has(key),
    isDropped(key) {
      const bead = beadOf(key);
      if (bead) {
        return !leftOf(bead.thread).slot.has(bead.at);
      }
      const thread = ix.byKey.get(key);
      return thread !== undefined && leftOf(thread).at.length === 0;
    },
    overrides(key) {
      const bead = beadOf(key);
      if (bead) {
        const { beads, value: threadRef } = ix.threads[bead.thread];
        const { at: kept, slot } = leftOf(bead.thread);
        const n = slot.get(bead.at);
        if (n === undefined || kept.length === beads.length) {
          return undefined;
        }
        const out = new Map<string, PdfObject>([
          ["N", beads[kept[(n + 1) % kept.length]].ref],
          ["V", beads[kept[(n - 1 + kept.length) % kept.length]].ref]
        ]);
        if (n === 0 && isPdfRef(threadRef)) {
          out.set("T", threadRef);
        }
        return out;
      }
      const thread = ix.byKey.get(key);
      if (thread === undefined) {
        return undefined;
      }
      const { at: kept } = leftOf(thread);
      return kept.length > 0 && kept[0] !== 0
        ? new Map<string, PdfObject>([["F", ix.threads[thread].beads[kept[0]].ref]])
        : undefined;
    }
  };
}

// =============================================================================
// Resolving destinations
// =============================================================================

/** Where a destination (an array, a name, a string, or a `<< /D … >>`) leads. */
function destTarget(
  doc: PdfDocument,
  index: CatalogIndex,
  dest: PdfObject | undefined,
  depth = 0
): Target {
  if (dest === undefined || dest === null) {
    return NONE;
  }
  const value = doc.deref(dest);
  if (isPdfArray(value)) {
    const page = value[0];
    let pages: readonly number[] | undefined;
    if (isPdfRef(page)) {
      pages = index.pageIndices.get(`${page.objNum}:${page.gen}`);
    } else if (
      typeof page === "number" &&
      Number.isInteger(page) &&
      page >= 0 &&
      page < index.pageCount
    ) {
      // A local destination should name its page by reference, but some writers put the
      // index there, as a remote one does.
      pages = [page];
    }
    return pages ? { kind: "page", pages } : DEAD;
  }
  if ((typeof value === "string" || value instanceof Uint8Array) && depth < 2) {
    const entry = index.lookupName(value);
    return entry === undefined ? DEAD : destTarget(doc, index, entry, depth + 1);
  }
  if (value instanceof Map && depth < 2) {
    return destTarget(doc, index, value.get("D"), depth + 1);
  }
  return DEAD;
}

/** Where an action leads; `GoTo` is the only kind that goes to a page of this document. */
function actionTarget(
  doc: PdfDocument,
  index: CatalogIndex,
  action: PdfObject | undefined
): Target {
  if (action === undefined || action === null) {
    return NONE;
  }
  const dict = doc.derefDict(action);
  if (!dict) {
    return DEAD;
  }
  return dictGetName(dict, "S") === "GoTo" ? destTarget(doc, index, dict.get("D")) : ACTION;
}

/** Group the indices of `targets` by the original pages they lead to. */
function byPage(targets: readonly Target[]): Map<number, number[]> {
  const out = new Map<number, number[]>();
  targets.forEach((target, i) => {
    if (target.kind !== "page") {
      return;
    }
    for (const page of new Set(target.pages)) {
      const list = out.get(page);
      if (list) {
        list.push(i);
      } else {
        out.set(page, [i]);
      }
    }
  });
  return out;
}

// =============================================================================
// Indexing
// =============================================================================

/**
 * The outline. It is the file's: an item reached twice (a `/Next` or `/First` cycle) is read
 * once, and the walk uses an explicit stack, so neither a cycle nor depth can stop it.
 */
function indexOutline(
  doc: PdfDocument,
  index: CatalogIndex,
  value: PdfObject
): OutlineIndex | null {
  const root = doc.derefDict(value);
  if (!root) {
    return null;
  }
  const keys = new Set<string>();
  if (isPdfRef(value)) {
    keys.add(`${value.objNum}:${value.gen}`);
  }
  const seenDicts = new Set<PdfDictValue>([root]);
  interface Pending {
    ref: PdfRef | undefined;
    dict: PdfDictValue;
    parent: number;
    children: number[];
  }
  /** The sibling chain starting at `first`. */
  const siblings = (first: PdfObject | undefined, parent: number): Pending[] => {
    const level: Pending[] = [];
    let current = first;
    while (current !== undefined && current !== null) {
      let ref: PdfRef | undefined;
      if (isPdfRef(current)) {
        const key = `${current.objNum}:${current.gen}`;
        if (keys.has(key)) {
          break;
        }
        keys.add(key);
        ref = current;
      }
      const dict = doc.derefDict(current);
      if (!dict || seenDicts.has(dict)) {
        break;
      }
      seenDicts.add(dict);
      level.push({ ref, dict, parent, children: [] });
      current = dict.get("Next");
    }
    return level;
  };

  const pending: Pending[] = [];
  const stack = siblings(root.get("First"), -1).reverse();
  while (stack.length > 0) {
    const item = stack.pop()!;
    const at = pending.length;
    pending.push(item);
    if (item.parent >= 0) {
      pending[item.parent].children.push(at);
    }
    const kids = siblings(item.dict.get("First"), at);
    for (let i = kids.length - 1; i >= 0; i--) {
      stack.push(kids[i]);
    }
  }
  const items = pending.map(({ ref, dict, parent, children }): IndexedOutlineItem => {
    const count = doc.deref(dict.get("Count") ?? null);
    return {
      ref,
      dict,
      parent,
      children,
      open: typeof count === "number" && count > 0,
      target: dict.has("Dest")
        ? destTarget(doc, index, dict.get("Dest"))
        : actionTarget(doc, index, dict.get("A"))
    };
  });
  return { root: value, items, byPage: byPage(items.map(item => item.target)), keys };
}

function indexDestsTree(
  doc: PdfDocument,
  index: CatalogIndex,
  root: PdfObject
): NamedIndex<Uint8Array> {
  const seen = new Set<string>();
  const entries: Array<NamedEntry<Uint8Array>> = [];
  for (const [key, value] of readNameTree(doc, root)) {
    const text = bytesToLatin1(key);
    if (!seen.has(text)) {
      seen.add(text);
      entries.push({ key, value, target: destTarget(doc, index, value) });
    }
  }
  entries.sort((a, b) => compareBytes(a.key, b.key));
  return { entries, byPage: byPage(entries.map(entry => entry.target)) };
}

function indexDestsDict(
  doc: PdfDocument,
  index: CatalogIndex,
  value: PdfObject
): NamedIndex<string> | null {
  const dict = doc.derefDict(value);
  if (!dict) {
    return null;
  }
  const entries = [...dict].map(([key, dest]): NamedEntry<string> => ({
    key,
    value: dest,
    target: destTarget(doc, index, dest)
  }));
  return { entries, byPage: byPage(entries.map(entry => entry.target)) };
}

function indexPageLabels(doc: PdfDocument, value: PdfObject): LabelRange[] {
  const ranges: LabelRange[] = [];
  for (const [start, entry] of readTree(doc, value, "Nums")) {
    const key = doc.deref(start);
    const dict = doc.derefDict(entry);
    if (typeof key !== "number" || !Number.isInteger(key) || key < 0 || !dict) {
      continue;
    }
    const st = doc.deref(dict.get("St") ?? null);
    ranges.push({
      start: key,
      dict,
      first: typeof st === "number" && Number.isInteger(st) && st >= 1 ? st : 1
    });
  }
  return ranges.sort((a, b) => a.start - b.start);
}

function indexThreads(doc: PdfDocument, index: CatalogIndex, value: PdfObject): ThreadIndex | null {
  const list = doc.deref(value);
  if (!isPdfArray(list)) {
    return null;
  }
  const threads: ThreadEntry[] = [];
  const beadAt = new Map<string, { thread: number; at: number }>();
  const byKey = new Map<string, number>();
  const threadsByPage = new Map<number, number[]>();
  for (const thread of list) {
    const dict = doc.derefDict(thread);
    if (!dict) {
      continue;
    }
    const i = threads.length;
    const beads: BeadEntry[] = [];
    const byPage = new Map<number, number[]>();
    // The ring from /F along /N, each bead once: a bead already in a thread ends this one.
    let bead: PdfObject | undefined = dict.get("F");
    while (isPdfRef(bead) && !beadAt.has(`${bead.objNum}:${bead.gen}`)) {
      const beadDict = doc.derefDict(bead);
      if (!beadDict) {
        break;
      }
      const page = beadDict.get("P");
      const pages = (isPdfRef(page) && index.pageIndices.get(`${page.objNum}:${page.gen}`)) || [];
      beadAt.set(`${bead.objNum}:${bead.gen}`, { thread: i, at: beads.length });
      for (const p of new Set(pages)) {
        const at = byPage.get(p);
        if (at) {
          at.push(beads.length);
        } else {
          byPage.set(p, [beads.length]);
        }
        const onPage = threadsByPage.get(p);
        if (!onPage) {
          threadsByPage.set(p, [i]);
        } else if (onPage[onPage.length - 1] !== i) {
          onPage.push(i);
        }
      }
      beads.push({ ref: bead, pages });
      bead = beadDict.get("N");
    }
    if (isPdfRef(thread)) {
      byKey.set(`${thread.objNum}:${thread.gen}`, i);
    }
    threads.push({ value: thread, dict, beads, byPage });
  }
  return { threads, byPage: threadsByPage, beads: beadAt, byKey };
}

/** Look a named destination up in `/Dests` and `/Names /Dests`, as a viewer does. */
function namedLookup(
  doc: PdfDocument,
  catalog: PdfDictValue
): (name: string | Uint8Array) => PdfObject | undefined {
  const dests = doc.derefDict(catalog.get("Dests") ?? null);
  const names = doc.derefDict(catalog.get("Names") ?? null);
  const byKey = new Map<string, PdfObject>();
  for (const [key, value] of names ? readNameTree(doc, names.get("Dests")) : []) {
    const text = bytesToLatin1(key);
    if (!byKey.has(text)) {
      byKey.set(text, value);
    }
  }
  return name => {
    const text = typeof name === "string" ? name : bytesToLatin1(name);
    return typeof name === "string"
      ? (dests?.get(text) ?? byKey.get(text))
      : (byKey.get(text) ?? dests?.get(text));
  };
}

/** The `[key, value]` pairs of a name tree, in tree order. */
function readNameTree(
  doc: PdfDocument,
  root: PdfObject | undefined
): Array<[Uint8Array, PdfObject]> {
  const out: Array<[Uint8Array, PdfObject]> = [];
  for (const [key, value] of readTree(doc, root, "Names")) {
    const resolved = doc.deref(key);
    if (resolved instanceof Uint8Array) {
      out.push([resolved, value]);
    } else if (typeof resolved === "string") {
      out.push([latin1ToBytes(resolved), value]);
    }
  }
  return out;
}

/** The leaf pairs of a name or number tree, walked with an explicit stack; a node reached twice is read once. */
function readTree(
  doc: PdfDocument,
  root: PdfObject | undefined,
  leafKey: "Names" | "Nums"
): Array<[PdfObject, PdfObject]> {
  const out: Array<[PdfObject, PdfObject]> = [];
  const seen = new Set<PdfDictValue>();
  const stack: PdfObject[] = root === undefined ? [] : [root];
  while (stack.length > 0) {
    const node = doc.derefDict(stack.pop()!);
    if (!node || seen.has(node)) {
      continue;
    }
    seen.add(node);
    const leaf = doc.deref(node.get(leafKey) ?? null);
    if (isPdfArray(leaf)) {
      for (let i = 0; i + 1 < leaf.length; i += 2) {
        out.push([leaf[i], leaf[i + 1]]);
      }
    }
    const kids = doc.deref(node.get("Kids") ?? null);
    if (isPdfArray(kids)) {
      for (let i = kids.length - 1; i >= 0; i--) {
        stack.push(kids[i]);
      }
    }
  }
  return out;
}

function bytesToLatin1(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i < bytes.length; i++) {
    out += String.fromCharCode(bytes[i]);
  }
  return out;
}

function latin1ToBytes(text: string): Uint8Array {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    out[i] = text.charCodeAt(i) & 0xff;
  }
  return out;
}

function compareBytes(a: Uint8Array, b: Uint8Array): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) {
    if (a[i] !== b[i]) {
      return a[i] - b[i];
    }
  }
  return a.length - b.length;
}
