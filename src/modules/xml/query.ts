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

import { checkElement, checkStringArgument } from "@xml/arguments";
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
function unsupported(path: string, why: string): never {
  throw new XmlArgumentError(
    "path",
    `Xml.query: unsupported path ${JSON.stringify(path)} — ${why}`
  );
}

function parsePath(path: string): PathStep[] {
  const steps: PathStep[] = [];
  const parts = path.split("/");

  let i = 0;
  // A leading "//" searches the whole subtree of the context element.
  let recursive = false;
  if (parts[0] === "") {
    if (parts.length < 3 || parts[1] !== "") {
      unsupported(
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
        unsupported(path, "empty step");
      }
      recursive = true;
      i++;
      continue;
    }
    steps.push(parseStep(path, part, recursive));
    recursive = false;
    i++;
  }

  if (steps.length === 0) {
    unsupported(path, "no step");
  }
  return steps;
}

const STEP_RE =
  /^(\*|[^[\]/@()'"=\s.][^[\]/@()'"=\s]*)(?:\[@([\w:.-]+)=(?:'([^']*)'|"([^"]*)")\]|\[(\d+)\])?$/u;

function parseStep(path: string, raw: string, recursive: boolean): PathStep {
  const m = STEP_RE.exec(raw);
  if (!m) {
    unsupported(
      path,
      `step ${JSON.stringify(raw)}; a step is a name or *, optionally with [@attr='value'] or [index]`
    );
  }

  const step: PathStep = { name: m[1], recursive };

  if (m[2] !== undefined) {
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

function matchesName(el: XmlElement, name: string): boolean {
  if (name === "*") {
    return true;
  }
  return el.name === name;
}

function matchesFilter(el: XmlElement, step: PathStep): boolean {
  if (!matchesName(el, step.name)) {
    return false;
  }
  if (step.attrFilter) {
    return el.attributes[step.attrFilter.key] === step.attrFilter.value;
  }
  return true;
}

/** Get direct child elements matching a step (no index applied). */
function matchChildren(el: XmlElement, step: PathStep): XmlElement[] {
  const results: XmlElement[] = [];
  for (const child of el.children) {
    if (child.type === "element" && matchesFilter(child, step)) {
      results.push(child);
    }
  }
  return results;
}

/** Get all descendant elements matching a step (recursive descent, document order). */
function matchDescendants(el: XmlElement, step: PathStep): XmlElement[] {
  const results: XmlElement[] = [];
  function recurse(node: XmlElement): void {
    for (const child of node.children) {
      if (child.type === "element") {
        if (matchesFilter(child, step)) {
          results.push(child);
        }
        recurse(child);
      }
    }
  }
  recurse(el);
  return results;
}

/** Apply a single step to a set of context elements, returning the next set. */
function applyStep(contexts: XmlElement[], step: PathStep): XmlElement[] {
  const results: XmlElement[] = [];
  // Overlapping contexts (a `//` step below another) reach the same node more than once.
  const seen = step.recursive && contexts.length > 1 ? new Set<XmlElement>() : null;
  const add = (el: XmlElement): void => {
    if (seen === null) {
      results.push(el);
    } else if (!seen.has(el)) {
      seen.add(el);
      results.push(el);
    }
  };

  if (step.index !== undefined) {
    // Per-parent index: for each context, find the Nth match among its children/descendants.
    const idx = step.index;
    for (const ctx of contexts) {
      const matches = step.recursive ? matchDescendants(ctx, step) : matchChildren(ctx, step);
      if (idx >= 0 && idx < matches.length) {
        add(matches[idx]);
      }
    }
  } else {
    for (const ctx of contexts) {
      const matches = step.recursive ? matchDescendants(ctx, step) : matchChildren(ctx, step);
      for (const el of matches) {
        add(el);
      }
    }
  }

  return results;
}

// =============================================================================
// Public API
// =============================================================================

/** Internal: run a parsed path against a root element, return all matches. */
function runQuery(element: XmlElement, path: string): XmlElement[] {
  const steps = parsePath(path);
  let contexts: XmlElement[] = [element];

  for (const step of steps) {
    contexts = applyStep(contexts, step);
    if (contexts.length === 0) {
      return contexts;
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
  return runQuery(element, path)[0];
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
  return runQuery(element, path);
}

export { query, queryAll };
