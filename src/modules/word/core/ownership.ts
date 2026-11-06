/**
 * Copy-on-write ownership for the in-place edits of `Document.*`.
 *
 * A function outside `Document` that returns a new document may share
 * unchanged arrays, maps and nodes with its input (`{ ...doc, settings }`), so
 * an edit made in place through one of them used to reach the other. Two
 * facts make that safe again without copying on every edit:
 *
 * - **Which document a top-level container belongs to.** `body`, `styles`,
 *   `headers`, … are edited in place only by the document that owns them. A
 *   container nobody has claimed is claimed by the first document edited
 *   through it — a hand-built or freshly read document keeps its own arrays —
 *   unless that document is known to share its tree; any other is copied,
 *   claimed by the copy, and put on the document. Every later edit of that
 *   container is in place, so appending stays O(1).
 * - **Which documents may share nodes.** A function that derives a document
 *   from another calls {@link markDerived}: from then on both the result and
 *   the input may hold nodes the other holds. An edit that changes nodes
 *   inside the tree (replacing text, accepting revisions) first gives such a
 *   document a tree of its own ({@link ownTree}); an edit that only adds to
 *   a container needs only the container.
 *
 * Sharing a node between documents by hand (`{ ...a }`, or one paragraph
 * added to two documents) is the caller's to manage, as it always was.
 */

import type { DocxDocument } from "@word/types";

/** The document allowed to edit each top-level container in place. */
const owners = new WeakMap<object, DocxDocument>();

/** Documents whose nodes another document may also hold. */
const sharing = new WeakSet<DocxDocument>();

function releaseContainers(doc: DocxDocument): void {
  for (const value of Object.values(doc)) {
    if (value !== null && typeof value === "object" && owners.get(value) === doc) {
      owners.delete(value);
    }
  }
}

/**
 * Record that `result` was derived from `inputs` and may share nodes with
 * them. A result that is one of its inputs shares nothing new. Returns `result`.
 */
export function markDerived<T extends DocxDocument>(result: T, ...inputs: DocxDocument[]): T {
  for (const input of inputs) {
    if (input === result) {
      continue;
    }
    releaseContainers(input);
    sharing.add(input);
    sharing.add(result);
  }
  return result;
}

/** Whether `doc` may hold nodes another document also holds. */
export function isSharing(doc: DocxDocument): boolean {
  return sharing.has(doc);
}

/** Record that `doc` owns `container`, which it holds and nothing else does. */
export function claim(doc: DocxDocument, container: object): void {
  owners.set(container, doc);
}

/**
 * The array or map under `key`, which `doc` may now edit in place: the one it
 * holds, when that is its own, or a copy put on `doc` in its place.
 */
export function ownContainer<C extends object>(doc: DocxDocument, key: string, container: C): C {
  const owner = owners.get(container);
  if (owner === doc || (owner === undefined && !sharing.has(doc))) {
    owners.set(container, doc);
    return container;
  }
  const copy = (container instanceof Map
    ? new Map(container)
    : [...(container as unknown as unknown[])]) as unknown as C;
  owners.set(copy, doc);
  (doc as unknown as Record<string, unknown>)[key] = copy;
  return copy;
}

/**
 * Give `doc` a tree no other document holds, before an edit that changes
 * nodes inside it. Free for a document that shares nothing; otherwise one
 * copy of the document's content, after which it shares nothing.
 *
 * **The copy replaces `doc`'s nodes**, so a node taken from `doc` before it
 * (through `Query.*`) no longer belongs to it, and an edit made through that
 * node afterwards does not reach `doc`. Only a document that shares its tree —
 * one a derivation was made from, or made by one — pays this, once.
 *
 * The copy is structural ({@link copyTree}): bytes (`Uint8Array`, `ArrayBuffer`)
 * are shared rather than duplicated — no edit writes into them — so it costs the
 * tree's nodes, not its media, and functions are carried over as they are.
 */
export function ownTree(doc: DocxDocument): void {
  if (!sharing.has(doc)) {
    return;
  }
  const fields = doc as unknown as Record<string, unknown>;
  const seen = new Map<object, unknown>();
  for (const [key, value] of Object.entries(fields)) {
    const copy = copyTree(value, seen);
    fields[key] = copy;
    if (copy !== null && typeof copy === "object") {
      owners.set(copy, doc);
    }
  }
  sharing.delete(doc);
}

/**
 * A copy of the document data in `value` — arrays, objects, `Map`s, `Set`s and
 * `Date`s, from any realm — sharing only bytes (`ArrayBuffer` and its views),
 * which no edit writes into. Kinds are told by their built-in tag, not by
 * `instanceof` or prototype, so a node built in another realm (an iframe, `vm`)
 * is copied like any other rather than shared. `seen` keeps a node reached
 * twice one node in the copy.
 */
function copyTree(value: unknown, seen: Map<object, unknown>): unknown {
  if (value === null || typeof value !== "object") {
    return value;
  }
  const done = seen.get(value);
  if (done !== undefined) {
    return done;
  }
  if (ArrayBuffer.isView(value)) {
    return value;
  }
  switch (Object.prototype.toString.call(value)) {
    case "[object ArrayBuffer]":
    case "[object SharedArrayBuffer]":
      return value;
    case "[object Array]": {
      const copy: unknown[] = [];
      seen.set(value, copy);
      for (const item of value as unknown[]) {
        copy.push(copyTree(item, seen));
      }
      return copy;
    }
    case "[object Map]": {
      const copy = new Map();
      seen.set(value, copy);
      for (const [key, item] of value as Map<unknown, unknown>) {
        copy.set(key, copyTree(item, seen));
      }
      return copy;
    }
    case "[object Set]": {
      const copy = new Set();
      seen.set(value, copy);
      for (const item of value as Set<unknown>) {
        copy.add(copyTree(item, seen));
      }
      return copy;
    }
    case "[object Date]":
      return new Date((value as Date).getTime());
    default: {
      const copy = Object.create(Object.getPrototypeOf(value) as object | null) as Record<
        string,
        unknown
      >;
      seen.set(value, copy);
      for (const [key, item] of Object.entries(value)) {
        copy[key] = copyTree(item, seen);
      }
      return copy;
    }
  }
}
