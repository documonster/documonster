/**
 * Pivot table & pivot cache check.
 *
 * Each `xl/pivotTables/pivotTableN.xml` declares a `cacheId` that must be
 * registered in the workbook's `<pivotCaches>` list; the cache in turn
 * must back a live pivotCacheDefinition/pivotCacheRecords pair.
 *
 * We do a lightweight cross-reference check: (a) every pivot table
 * advertises a cacheId that resolves to a `<pivotCache>` entry in
 * workbook.xml, and (b) each referenced pivotCacheDefinition's rels point
 * at a pivotCacheRecords file that actually exists.
 */

import type { ValidationContext } from "./context";
import { attrByLocalName, findChildLocal, findChildrenLocal } from "./xml-utils";

const PIVOT_TABLE_RE = /^xl\/pivotTables\/pivotTable\d+\.xml$/;

const WORKBOOK_PATH = "xl/workbook.xml";

export function checkPivot(ctx: ValidationContext): void {
  const declared = collectWorkbookPivotCaches(ctx);
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || !PIVOT_TABLE_RE.test(path)) {
      continue;
    }
    const dom = ctx.readDom(path);
    if (!dom) {
      continue;
    }
    const cacheId = attrByLocalName(dom.root, "cacheId");
    if (cacheId === undefined) {
      ctx.reporter.error("pivot-missing-cacheId", `${path}: missing cacheId attribute`, path);
      continue;
    }
    if (!declared.has(cacheId)) {
      ctx.reporter.error(
        "pivot-cacheId-not-in-workbook",
        `${path}: cacheId="${cacheId}" is not declared in workbook <pivotCaches>`,
        path
      );
    }
  }

  // Each pivotCacheDefinition.rels should point at a pivotCacheRecords file
  // that exists inside the package.
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (
      entry.type === "directory" ||
      !/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(path)
    ) {
      continue;
    }
    const defRels = ctx.readRels(defaultRelsPath(path));
    const recordsRel = defRels.rels.find(r => r.type.includes("pivotCacheRecords"));
    if (recordsRel) {
      const resolved = recordsRel.target.startsWith("/")
        ? recordsRel.target.slice(1)
        : joinPath(parentDir(path), recordsRel.target);
      if (!ctx.has(resolved)) {
        ctx.reporter.error(
          "pivot-cacheRecords-missing",
          `${path}: references missing pivotCacheRecords file ${resolved}`,
          path
        );
      }
    }
  }
}

function collectWorkbookPivotCaches(ctx: ValidationContext): Set<string> {
  const ids = new Set<string>();
  if (!ctx.has(WORKBOOK_PATH)) {
    return ids;
  }
  const dom = ctx.readDom(WORKBOOK_PATH);
  if (!dom) {
    return ids;
  }
  const pivotCaches = findChildLocal(dom.root, "pivotCaches");
  if (!pivotCaches) {
    return ids;
  }
  for (const pc of findChildrenLocal(pivotCaches, "pivotCache")) {
    const id = attrByLocalName(pc, "cacheId");
    if (id !== undefined) {
      ids.add(id);
    }
  }
  return ids;
}

function defaultRelsPath(partPath: string): string {
  const dir = parentDir(partPath);
  const base = partPath.slice(dir.length ? dir.length + 1 : 0);
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

function parentDir(p: string): string {
  const idx = p.lastIndexOf("/");
  return idx === -1 ? "" : p.slice(0, idx);
}

function joinPath(dir: string, target: string): string {
  // Resolve `./foo/bar` or `../baz` relative to `dir`.
  const parts: string[] = dir ? dir.split("/") : [];
  for (const seg of target.split("/")) {
    if (seg === "" || seg === ".") {
      continue;
    }
    if (seg === "..") {
      parts.pop();
      continue;
    }
    parts.push(seg);
  }
  return parts.join("/");
}
