/**
 * XML Path Query Engine
 *
 * Simplified path expressions for querying XML DOM trees.
 * Covers the most common query patterns without the complexity of full XPath.
 *
 * Supported syntax:
 * - `a/b/c`         — match child `a`, then child `b`, then child `c`
 * - `a/b[@id='1']`  — match child `b` with attribute `id` equal to `"1"`
 * - `a/*`            — wildcard: any element name at that level; valid in any step
 * - `a//c`           — recursive descent: `c` at any depth under `a`
 * - `a/b[0]`         — index: first matching `b` under each parent `a` (0-based, unlike XPath)
 * - `//c`            — a leading `//` searches the whole subtree of the context element
 *
 * Anything else — an absolute path (`/a`), an empty step (`a//`, `a/`), a predicate this does
 * not implement (`[@id]`, `[last()]`), `.` or `..` — is refused with an {@link XmlError}
 * rather than matching nothing. A node is returned once, in document order, even when
 * overlapping `//` steps reach it more than once.
 *
 * Functions:
 * - `query(el, path)`    — first match or undefined
 * - `queryAll(el, path)` — all matches
 */

import { checkAttributeValue, checkElement, checkStringArgument, refuseDom } from "@xml/arguments";
import { isXmlName } from "@xml/chars";
import { CYCLE_CHECK_DEPTH, RECURSION_DEPTH } from "@xml/dom";
import { XmlArgumentError } from "@xml/errors";
import type { XmlElement } from "@xml/types";

function checkQueryArguments(where: string, element: unknown, path: unknown): void {
  checkElement(where, element);
  checkStringArgument(where, "path", path);
}

// =============================================================================
// Path Parsing
// =============================================================================

interface PathStep {
  /** Element name to match, or "*" for wildcard. */
  name: string;
  /** If true, match at any depth (recursive descent `//`). */
  recursive: boolean;
  /** Attribute filter: `[@key='value']`. */
  attrFilter?: { key: string; value: string };
  /** Index filter: `[N]`. */
  index?: number;
}

/**
 * Parse a simplified path expression into steps.
 *
 * Examples:
 *   "a/b/c"            → [a, b, c]
 *   "a//c"             → [a, c(recursive)]
 *   "a/b[@id='1']"     → [a, b(attr:id=1)]
 *   "a/b[0]"           → [a, b(index:0)]
 */
function unsupported(where: string, path: string, why: string): never {
  throw new XmlArgumentError("path", `${where}: unsupported path ${JSON.stringify(path)} — ${why}`);
}

function parsePath(where: string, path: string): PathStep[] {
  const steps: PathStep[] = [];
  const parts = splitPath(where, path);

  let i = 0;
  // A leading "//" searches the whole subtree of the context element.
  let recursive = false;
  if (parts[0] === "") {
    if (parts.length < 3 || parts[1] !== "") {
      unsupported(
        where,
        path,
        'absolute paths are not supported; write "a" for a child or "//a" for any depth'
      );
    }
    recursive = true;
    i = 2;
  }

  while (i < parts.length) {
    const part = parts[i];
    if (part === "") {
      // "a//b": the empty part between the slashes makes the next step recursive.
      if (recursive || i === parts.length - 1) {
        unsupported(where, path, "empty step");
      }
      recursive = true;
      i++;
      continue;
    }
    steps.push(parseStep(where, path, part, recursive));
    recursive = false;
    i++;
  }

  if (steps.length === 0) {
    unsupported(where, path, "no step");
  }
  return steps;
}

/**
 * Split `path` at each `/` outside a predicate, so `a[@href='a/b']` stays one step. An
 * unterminated `[` or quote is refused rather than read as part of a name.
 */
function splitPath(where: string, path: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let inBracket = false;
  let quote = "";
  for (let i = 0; i < path.length; i++) {
    const ch = path[i];
    if (quote !== "") {
      if (ch === quote) {
        quote = "";
      }
    } else if (inBracket) {
      if (ch === "'" || ch === '"') {
        quote = ch;
      } else if (ch === "]") {
        inBracket = false;
      }
    } else if (ch === "[") {
      inBracket = true;
    } else if (ch === "/") {
      parts.push(path.slice(start, i));
      start = i + 1;
    }
  }
  if (quote !== "" || inBracket) {
    unsupported(where, path, `unterminated ${quote !== "" ? "quoted value" : "predicate"}`);
  }
  parts.push(path.slice(start));
  return parts;
}

const STEP_RE = /^([^[\]]+)(?:\[@([^=\]]+)=(?:'([^']*)'|"([^"]*)")\]|\[(\d+)\])?$/u;

function parseStep(where: string, path: string, raw: string, recursive: boolean): PathStep {
  const m = STEP_RE.exec(raw);
  if (!m) {
    unsupported(
      where,
      path,
      `step ${JSON.stringify(raw)}; a step is a name or *, optionally with [@attr='value'] or [index]`
    );
  }

  const name = m[1];
  if (name !== "*" && !isXmlName(name)) {
    unsupported(
      where,
      path,
      `step ${JSON.stringify(raw)}; ${JSON.stringify(name)} is not an XML name or *`
    );
  }
  const step: PathStep = { name, recursive };

  if (m[2] !== undefined) {
    if (!isXmlName(m[2])) {
      unsupported(
        where,
        path,
        `step ${JSON.stringify(raw)}; ${JSON.stringify(m[2])} is not an XML attribute name`
      );
    }
    // Attribute filter: [@key='value'] or [@key="value"]
    step.attrFilter = { key: m[2], value: m[3] ?? m[4] };
  } else if (m[5] !== undefined) {
    // Index filter: [N]
    step.index = parseInt(m[5], 10);
  }

  return step;
}

// =============================================================================
// Step Matching
// =============================================================================

/**
 * One run of a query over a tree its caller gave: the nodes it reaches are checked as they are
 * reached — a node without the fields read is refused, and so is an element among its own
 * descendants, which no walk would leave. Each field is read once, so what is checked is what is used.
 */

/**
 * The most elements one step of a query reaches before it is refused. Far more element objects than a
 * JavaScript heap holds: a tree never reaches it, and a model that shares its elements along
 * exponentially many paths is refused rather than walked for ever.
 */
const MAX_ELEMENTS_PER_STEP = 2 ** 25;

class QueryRun {
  constructor(private readonly where: string) {}

  /** Whether `el`, named `name` (its name as read and checked), is what `step` asks for. */
  matches(el: XmlElement, name: string, step: PathStep): boolean {
    if (step.name !== "*" && name !== step.name) {
      return false;
    }
    const filter = step.attrFilter;
    if (filter === undefined) {
      return true;
    }
    const attributes = (el as { attributes?: unknown }).attributes;
    if (typeof attributes !== "object" || attributes === null || Array.isArray(attributes)) {
      return refuseDom(this.where, "element", "", ".attributes", "an object", attributes);
    }
    // Its own attributes only: one an object inherits is no attribute of the element.
    if (!Object.hasOwn(attributes, filter.key)) {
      return false;
    }
    const value = (attributes as Record<string, unknown>)[filter.key];
    if (typeof value !== "string") {
      checkAttributeValue(this.where, "element", filter.key, value);
    }
    return value === filter.value;
  }

  /** The name of the element {@link elementAt} returned last, read and checked with it. */
  name = "";
  /** The elements this step has reached: see {@link MAX_ELEMENTS_PER_STEP}. */
  private reached = 0;

  /** `el.children`, read once and checked. */
  childList(el: XmlElement): readonly unknown[] {
    const children = (el as { children?: unknown }).children;
    if (!Array.isArray(children)) {
      refuseDom(this.where, "element", "", ".children", "an array", children);
    }
    return children;
  }

  /**
   * `list[i]` if it is an element — its name in {@link name} — or `undefined` for another node; each
   * field read is read once and checked. Inline rather than through `readDom*`: a walk asks this of
   * every node it passes.
   */
  elementAt(list: readonly unknown[], i: number): XmlElement | undefined {
    const child = list[i];
    if (typeof child === "object" && child !== null) {
      const type = (child as { type?: unknown }).type;
      if (type === "element") {
        if (++this.reached > MAX_ELEMENTS_PER_STEP) {
          throw new XmlArgumentError(
            "element",
            `${this.where}: a step reached more than ${MAX_ELEMENTS_PER_STEP} elements: the model shares its elements along more paths than a tree in memory holds`
          );
        }
        const name = (child as { name?: unknown }).name;
        if (typeof name !== "string") {
          refuseDom(this.where, "element", i, ".name", "a string", name);
        }
        this.name = name;
        return child as XmlElement;
      }
      if (type === "text" || type === "cdata") {
        const value = (child as { value?: unknown }).value;
        if (typeof value !== "string") {
          refuseDom(this.where, "element", i, ".value", "a string", value);
        }
        return undefined;
      }
      if (typeof type === "string") {
        return undefined;
      }
      refuseDom(this.where, "element", i, ".type", "a node type", type);
    }
    return refuseDom(this.where, "element", i, "", "an XmlNode", child);
  }

  /**
   * Walk the descendants of `el` in document order, calling `visit` with each element and its name;
   * `visit` returns `false` to stop. `leave` is called as each element's subtree is finished.
   *
   * A query answers for a tree, which is what a parse makes. A model built by hand may hold one element
   * under several parents: the walk then reaches it, and what is under it, along each path, as a walk
   * of the tree the paths spell out would. Two such models it refuses, as no tree is like them: one
   * where an element is among its own descendants, which no walk would leave, and one whose paths
   * outnumber the elements any tree in memory holds ({@link MAX_ELEMENTS_PER_STEP}), as sharing each
   * level twice over makes of a model of linear size. A tree pays for neither: the walk records the
   * elements it is inside only deep below its start — past `RECURSION_DEPTH`, then
   * {@link CYCLE_CHECK_DEPTH} more — which a cycle always reaches.
   */
  walkDescendants(
    el: XmlElement,
    visit: (descendant: XmlElement, name: string) => boolean,
    leave?: (descendant: XmlElement) => void
  ): void {
    this.walkBelow(el, this.childList(el), visit, leave, 0);
  }

  /**
   * Walk the elements in `children`, those of `parent`, and what is below them: recursive — what each
   * element's walk holds is on the call stack, which costs nothing to allocate — then, past
   * `RECURSION_DEPTH`, by {@link walkDeep}. Whether the walk went on to the end.
   */
  private walkBelow(
    parent: XmlElement,
    children: readonly unknown[],
    visit: (descendant: XmlElement, name: string) => boolean,
    leave: ((descendant: XmlElement) => void) | undefined,
    depth: number
  ): boolean {
    if (depth >= RECURSION_DEPTH) {
      return this.walkDeep(parent, children, visit, leave);
    }
    for (let i = 0; i < children.length; i++) {
      const node = this.elementAt(children, i);
      if (node === undefined) {
        continue;
      }
      const name = this.name;
      const below = this.childList(node);
      if (!visit(node, name)) {
        return false;
      }
      if (below.length > 0 && !this.walkBelow(node, below, visit, leave, depth + 1)) {
        return false;
      }
      leave?.(node);
    }
    return true;
  }

  /**
   * {@link walkBelow} with an explicit stack, not recursion, so a subtree of any depth is walked; a
   * cycle, which descends without end, is found here.
   */
  private walkDeep(
    el: XmlElement,
    children: readonly unknown[],
    visit: (descendant: XmlElement, name: string) => boolean,
    leave: ((descendant: XmlElement) => void) | undefined
  ): boolean {
    // Three arrays rather than a frame object per element: a leaf never enters them.
    const nodes: XmlElement[] = [el];
    const lists: (readonly unknown[])[] = [children];
    const next: number[] = [0];
    let top = 0;
    /** The elements the walk is inside, once it is deep enough to be looking for a cycle. */
    let open: Set<XmlElement> | undefined;
    while (top >= 0) {
      const list = lists[top];
      let i = next[top];
      let node: XmlElement | undefined;
      while (i < list.length && node === undefined) {
        node = this.elementAt(list, i++);
      }
      next[top] = i;
      if (node === undefined) {
        const done = nodes[top];
        open?.delete(done);
        top--;
        if (top >= 0) {
          leave?.(done);
        }
        continue;
      }
      const name = this.name;
      const children = this.childList(node);
      if (!visit(node, name)) {
        return false;
      }
      if (children.length === 0) {
        leave?.(node);
        continue;
      }
      if (open === undefined && top >= CYCLE_CHECK_DEPTH) {
        open = new Set(nodes.slice(0, top + 1));
      }
      if (open !== undefined) {
        if (open.has(node)) {
          throw new XmlArgumentError(
            "element",
            `${this.where}: element contains itself: an element is among its own descendants`
          );
        }
        open.add(node);
      }
      top++;
      nodes[top] = node;
      lists[top] = children;
      next[top] = 0;
    }
    return true;
  }

  /**
   * Apply a step to the contexts, in document order, returning the next contexts. `firstOnly` asks
   * for the first result only, where that is known without the rest.
   */
  applyStep(contexts: XmlElement[], step: PathStep, firstOnly: boolean): XmlElement[] {
    this.reached = 0;
    if (!step.recursive) {
      return this.childStep(contexts, step, firstOnly);
    }
    if (step.index === undefined) {
      return this.descendantStep(contexts, step, firstOnly);
    }
    return this.indexedDescendantStep(contexts, step, step.index, firstOnly);
  }

  private childStep(contexts: XmlElement[], step: PathStep, firstOnly: boolean): XmlElement[] {
    const results: XmlElement[] = [];
    // With an index the scan stops at the match it names; asked for the first result only, at the first.
    const wanted = step.index !== undefined ? step.index + 1 : firstOnly ? 1 : Infinity;
    for (const ctx of contexts) {
      const list = this.childList(ctx);
      let count = 0;
      for (let i = 0; i < list.length; i++) {
        const child = this.elementAt(list, i);
        if (child === undefined || !this.matches(child, this.name, step)) {
          continue;
        }
        if (++count === wanted || step.index === undefined) {
          results.push(child);
        }
        if (count === wanted) {
          break;
        }
      }
      if (firstOnly && results.length > 0) {
        break;
      }
    }
    return results;
  }

  private descendantStep(contexts: XmlElement[], step: PathStep, firstOnly: boolean): XmlElement[] {
    // A context inside another adds nothing: its descendants are the outer one's. Walked once
    // each, in document order — the contexts are — so the result is in document order too.
    const results: XmlElement[] = [];
    const inner = contexts.length > 1 ? new Set(contexts) : undefined;
    const covered = inner && new Set<XmlElement>();
    for (const ctx of contexts) {
      if (covered?.has(ctx)) {
        continue;
      }
      let stop = false;
      this.walkDescendants(ctx, (el, name) => {
        if (inner?.has(el)) {
          covered!.add(el);
        }
        if (this.matches(el, name, step)) {
          results.push(el);
          stop = firstOnly;
        }
        return !stop;
      });
      if (stop) {
        break;
      }
    }
    return results;
  }

  /** Per parent: the Nth matching descendant of each context. */
  private indexedDescendantStep(
    contexts: XmlElement[],
    step: PathStep,
    index: number,
    firstOnly: boolean
  ): XmlElement[] {
    const results: XmlElement[] = [];
    // Nested contexts can name one element as each one's Nth: it is one result.
    const picked = contexts.length > 1 ? new Set<XmlElement>() : undefined;
    for (const ctx of contexts) {
      const found = this.nthDescendant(ctx, step, index);
      if (found !== undefined && !picked?.has(found)) {
        picked?.add(found);
        results.push(found);
        // The first context's is the first: an inner context's matches are among the outer's.
        if (firstOnly) {
          break;
        }
      }
    }
    return results;
  }

  /** The `index`-th matching descendant of `ctx`, found by walking no further than it. */
  private nthDescendant(ctx: XmlElement, step: PathStep, index: number): XmlElement | undefined {
    let count = 0;
    let found: XmlElement | undefined;
    this.walkDescendants(ctx, (el, name) => {
      if (this.matches(el, name, step) && count++ === index) {
        found = el;
        return false;
      }
      return true;
    });
    return found;
  }
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Every element at and below `root`, numbered in document order — an element a hand-built model
 * shares between several parents by where it is first reached.
 */
function documentOrder(run: QueryRun, root: XmlElement): Map<XmlElement, number> {
  const order = new Map<XmlElement, number>([[root, 0]]);
  run.walkDescendants(root, el => {
    // An element reached along several paths is numbered where it is first reached.
    if (!order.has(el)) {
      order.set(el, order.size);
    }
    return true;
  });
  return order;
}

/** Internal: run a parsed path against a root element, return all matches — or with `firstOnly`, the first. */
function runQuery(
  where: string,
  element: XmlElement,
  path: string,
  firstOnly = false
): XmlElement[] {
  const steps = parsePath(where, path);
  const run = new QueryRun(where);
  let contexts: XmlElement[] = [element];
  // Whether one context may sit inside another, which only a `//` step produces. A child step then reads an outer
  // context's later children before an inner context's earlier ones (`//a/b` over `<a><b/><a><b/></a><b/></a>`),
  // so its result is put back in document order; a `//` step stays in order by itself, its overlaps removed.
  let mayNest = false;
  let order: Map<XmlElement, number> | undefined;

  for (let k = 0; k < steps.length; k++) {
    const step = steps[k];
    // Only the last step can stop at its first result, and a child step under nested contexts finds
    // its results out of order: those are sorted, so all of them are found.
    const first = firstOnly && k === steps.length - 1 && (step.recursive || !mayNest);
    contexts = run.applyStep(contexts, step, first);
    if (contexts.length === 0) {
      return contexts;
    }
    if (mayNest && !step.recursive && contexts.length > 1) {
      order ??= documentOrder(run, element);
      const rank = order;
      contexts.sort((a, b) => rank.get(a)! - rank.get(b)!);
    }
    if (step.recursive) {
      mayNest = true;
    }
  }

  return contexts;
}

/**
 * Find the first element matching a path expression.
 *
 * @param element - The context element to search from.
 * @param path - Simplified path expression.
 * @returns The first matching element, or undefined.
 *
 * @example
 * ```ts
 * const cell = query(doc.root, "sheetData/row/c[@r='A1']");
 * const value = query(doc.root, "sheetData/row/c/v");
 * const any = query(doc.root, "sheetData//v");  // recursive descent
 * ```
 */
function query(element: XmlElement, path: string): XmlElement | undefined {
  checkQueryArguments("Xml.query", element, path);
  return runQuery("Xml.query", element, path, true)[0];
}

/**
 * Find all elements matching a path expression.
 *
 * @param element - The context element to search from.
 * @param path - Simplified path expression.
 * @returns Array of all matching elements (may be empty).
 *
 * @example
 * ```ts
 * const rows = queryAll(doc.root, "sheetData/row");
 * const cells = queryAll(doc.root, "sheetData/row/c");
 * const allValues = queryAll(doc.root, "sheetData//v");
 * ```
 */
function queryAll(element: XmlElement, path: string): XmlElement[] {
  checkQueryArguments("Xml.queryAll", element, path);
  return runQuery("Xml.queryAll", element, path);
}

export { query, queryAll };
