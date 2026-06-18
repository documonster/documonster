/**
 * PDF bookmark (outline) extractor.
 *
 * Extracts the document outline tree from a PDF's `/Outlines` dictionary.
 * Each outline item has a title, a target page index, and optional children
 * forming a hierarchical bookmark tree.
 *
 * Supports:
 * - Direct destinations (`/Dest` as array or named destination)
 * - Action-based destinations (`/A << /S /GoTo /D ... >>`)
 * - Nested bookmarks (children via `/First`/`/Last` chains)
 * - Circular reference protection
 *
 * @see PDF Reference 1.7, §12.3 - Document-Level Navigation
 */

import type { PdfDocument } from "@pdf/reader/pdf-document";
import type { PdfDictValue, PdfObject } from "@pdf/reader/pdf-parser";
import { isPdfArray, isPdfRef, dictGetName, decodePdfStringBytes } from "@pdf/reader/pdf-parser";
import { getDictStringValue } from "@pdf/reader/reader-utils";

// =============================================================================
// Types
// =============================================================================

/** A bookmark (outline item) extracted from the PDF. */
export interface PdfBookmark {
  /** Bookmark title text */
  title: string;
  /** 0-based page index the bookmark points to (-1 if unresolvable) */
  pageIndex: number;
  /** Child bookmarks (nested outline items) */
  children: PdfBookmark[];
}

// =============================================================================
// Constants
// =============================================================================

/** Maximum depth for recursive outline traversal to prevent stack overflow. */
const MAX_OUTLINE_DEPTH = 100;

/** Maximum number of siblings at any level to prevent infinite /Next chains. */
const MAX_SIBLINGS = 10_000;

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract bookmarks (outlines) from a PDF document.
 *
 * Reads the `/Outlines` dictionary from the catalog and recursively
 * traverses the outline tree following `/First` → `/Next` chains.
 *
 * @param doc - The PDF document
 * @returns Array of top-level bookmarks with nested children
 */
export function extractBookmarks(doc: PdfDocument): PdfBookmark[] {
  try {
    const catalog = doc.getCatalog();
    const outlinesObj = catalog.get("Outlines");
    if (!outlinesObj) {
      return [];
    }

    const outlinesDict = doc.derefDict(outlinesObj);
    if (!outlinesDict) {
      return [];
    }

    // Build a page reference → index map for resolving destinations
    const pageMap = buildPageMap(doc);

    // The outline root's /First points to the first top-level item
    const visited = new Set<number>();
    return collectSiblings(outlinesDict, doc, pageMap, visited, 0);
  } catch {
    return [];
  }
}

// =============================================================================
// Page Map
// =============================================================================

/**
 * Build a map from page object reference identity to 0-based page index.
 *
 * We map by object number since page dicts resolved from different refs
 * will share the same objNum.
 */
function buildPageMap(doc: PdfDocument): Map<number, number> {
  const pages = doc.getPagesWithObjInfo();
  const map = new Map<number, number>();
  for (let i = 0; i < pages.length; i++) {
    const { objNum } = pages[i];
    if (objNum !== 0) {
      map.set(objNum, i);
    }
  }
  return map;
}

// =============================================================================
// Outline Tree Traversal
// =============================================================================

/**
 * Collect the sibling chain starting from the `/First` child of a parent node.
 */
function collectSiblings(
  parentDict: PdfDictValue,
  doc: PdfDocument,
  pageMap: Map<number, number>,
  visited: Set<number>,
  depth: number
): PdfBookmark[] {
  if (depth > MAX_OUTLINE_DEPTH) {
    return [];
  }

  const firstObj = parentDict.get("First");
  if (!firstObj) {
    return [];
  }

  const bookmarks: PdfBookmark[] = [];
  let currentObj: PdfObject | null | undefined = firstObj;
  let count = 0;

  while (currentObj != null && count < MAX_SIBLINGS) {
    count++;

    // Guard against circular references using object numbers
    if (isPdfRef(currentObj)) {
      if (visited.has(currentObj.objNum)) {
        break;
      }
      visited.add(currentObj.objNum);
    }

    const itemDict = doc.derefDict(currentObj);
    if (!itemDict) {
      break;
    }

    const bookmark = parseOutlineItem(itemDict, doc, pageMap, visited, depth);
    if (bookmark) {
      bookmarks.push(bookmark);
    }

    // Follow /Next to the next sibling
    currentObj = itemDict.get("Next");
  }

  return bookmarks;
}

/**
 * Parse a single outline item dictionary into a PdfBookmark.
 */
function parseOutlineItem(
  dict: PdfDictValue,
  doc: PdfDocument,
  pageMap: Map<number, number>,
  visited: Set<number>,
  depth: number
): PdfBookmark | null {
  // Extract title — required per spec
  const title = getOutlineTitle(dict, doc);
  if (!title) {
    return null;
  }

  // Resolve destination to a page index
  const pageIndex = resolveDestination(dict, doc, pageMap);

  // Collect children (nested bookmarks)
  const children = collectSiblings(dict, doc, pageMap, visited, depth + 1);

  return { title, pageIndex, children };
}

// =============================================================================
// Title Extraction
// =============================================================================

/**
 * Extract the title string from an outline item dictionary.
 * The /Title entry is a text string (may be Uint8Array or string).
 */
function getOutlineTitle(dict: PdfDictValue, doc: PdfDocument): string {
  return getDictStringValue(dict, "Title", doc);
}

// =============================================================================
// Destination Resolution
// =============================================================================

/**
 * Resolve an outline item's destination to a 0-based page index.
 *
 * Checks /Dest first, then falls back to /A (action) with /S /GoTo.
 * Returns -1 if the destination cannot be resolved.
 */
function resolveDestination(
  dict: PdfDictValue,
  doc: PdfDocument,
  pageMap: Map<number, number>
): number {
  // 1. Try /Dest (direct destination)
  const destObj = dict.get("Dest");
  if (destObj != null) {
    const pageIndex = resolveDestValue(destObj, doc, pageMap);
    if (pageIndex >= 0) {
      return pageIndex;
    }
  }

  // 2. Try /A (action dictionary) with /S /GoTo
  const actionObj = dict.get("A");
  if (actionObj != null) {
    const actionDict = doc.derefDict(actionObj);
    if (actionDict) {
      const actionType = dictGetName(actionDict, "S");
      if (actionType === "GoTo") {
        const actionDest = actionDict.get("D");
        if (actionDest != null) {
          return resolveDestValue(actionDest, doc, pageMap);
        }
      }
    }
  }

  return -1;
}

/**
 * Resolve a destination value (from /Dest or /A.D) to a page index.
 *
 * Destination formats (PDF Reference 1.7, §12.3.2):
 * - Array: `[pageRef /XYZ left top zoom]`, `[pageRef /Fit]`, etc.
 * - Named string: looked up in the document's /Dests or /Names.Dests
 */
function resolveDestValue(
  destObj: PdfObject,
  doc: PdfDocument,
  pageMap: Map<number, number>
): number {
  const resolved = doc.deref(destObj);
  if (resolved == null) {
    return -1;
  }

  // Array destination: first element is the page reference
  if (isPdfArray(resolved) && resolved.length >= 1) {
    return resolvePageRef(resolved[0], doc, pageMap);
  }

  // Named destination (string) — look up in /Dests or /Names tree
  if (typeof resolved === "string") {
    return resolveNamedDest(resolved, doc, pageMap);
  }

  // Byte string named destination
  if (resolved instanceof Uint8Array) {
    const name = decodePdfStringBytes(resolved);
    return resolveNamedDest(name, doc, pageMap);
  }

  return -1;
}

/**
 * Resolve a page reference (from the first element of a dest array) to a page index.
 */
function resolvePageRef(
  pageObj: PdfObject,
  doc: PdfDocument,
  pageMap: Map<number, number>
): number {
  // If it's a direct reference, use the object number
  if (isPdfRef(pageObj)) {
    const idx = pageMap.get(pageObj.objNum);
    return idx !== undefined ? idx : -1;
  }

  // If it's a page number (integer), use it directly as 0-based index
  if (typeof pageObj === "number" && Number.isInteger(pageObj)) {
    return pageObj;
  }

  return -1;
}

/**
 * Look up a named destination in the catalog's /Dests dictionary
 * or /Names.Dests name tree.
 */
function resolveNamedDest(name: string, doc: PdfDocument, pageMap: Map<number, number>): number {
  const catalog = doc.getCatalog();

  // 1. Try /Dests dictionary (older PDFs)
  const destsObj = catalog.get("Dests");
  if (destsObj != null) {
    const destsDict = doc.derefDict(destsObj);
    if (destsDict) {
      const entry = destsDict.get(name);
      if (entry != null) {
        return resolveDestEntry(entry, doc, pageMap);
      }
    }
  }

  // 2. Try /Names.Dests name tree (PDF 1.2+)
  const namesObj = catalog.get("Names");
  if (namesObj != null) {
    const namesDict = doc.derefDict(namesObj);
    if (namesDict) {
      const destsTreeObj = namesDict.get("Dests");
      if (destsTreeObj != null) {
        const value = lookupNameTree(destsTreeObj, name, doc);
        if (value != null) {
          return resolveDestEntry(value, doc, pageMap);
        }
      }
    }
  }

  return -1;
}

/**
 * Resolve a destination entry value. It may be a dict with /D key,
 * or a direct array destination.
 */
function resolveDestEntry(
  entry: PdfObject,
  doc: PdfDocument,
  pageMap: Map<number, number>
): number {
  const resolved = doc.deref(entry);
  if (resolved == null) {
    return -1;
  }

  // Direct array destination
  if (isPdfArray(resolved) && resolved.length >= 1) {
    return resolvePageRef(resolved[0], doc, pageMap);
  }

  // Dictionary with /D entry (destination dictionary)
  if (resolved instanceof Map) {
    const d = resolved.get("D");
    if (d != null) {
      return resolveDestValue(d, doc, pageMap);
    }
  }

  return -1;
}

/**
 * Look up a key in a PDF name tree.
 *
 * Name trees use either /Names (leaf) or /Kids (intermediate) arrays.
 * /Names is an array of alternating [key, value, key, value, ...] pairs.
 *
 * @see PDF Reference 1.7, §7.9.6 - Name Trees
 */
function lookupNameTree(
  treeObj: PdfObject,
  name: string,
  doc: PdfDocument,
  depth = 0
): PdfObject | null {
  if (depth > MAX_OUTLINE_DEPTH) {
    return null;
  }

  const treeDict = doc.derefDict(treeObj);
  if (!treeDict) {
    return null;
  }

  // Check leaf /Names array
  const namesArr = treeDict.get("Names");
  if (namesArr != null) {
    const resolved = doc.deref(namesArr);
    if (isPdfArray(resolved)) {
      // Alternating [key, value, key, value, ...]
      for (let i = 0; i + 1 < resolved.length; i += 2) {
        const key = doc.deref(resolved[i]);
        let keyStr: string | null = null;
        if (typeof key === "string") {
          keyStr = key;
        } else if (key instanceof Uint8Array) {
          keyStr = decodePdfStringBytes(key);
        }
        if (keyStr === name) {
          return resolved[i + 1];
        }
      }
    }
  }

  // Check intermediate /Kids array
  const kidsArr = treeDict.get("Kids");
  if (kidsArr != null) {
    const resolved = doc.deref(kidsArr);
    if (isPdfArray(resolved)) {
      for (const kid of resolved) {
        const result = lookupNameTree(kid, name, doc, depth + 1);
        if (result != null) {
          return result;
        }
      }
    }
  }

  return null;
}
