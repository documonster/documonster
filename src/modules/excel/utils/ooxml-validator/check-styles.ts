/**
 * Styles table integrity.
 *
 * Every xf in `cellXfs` references indices into `numFmts`, `fonts`,
 * `fills`, `borders`. Out-of-bounds references produce style corruption
 * on load (Excel falls back to the default style silently, so this is a
 * warning-severity check — it's a strong signal of a broken writer but
 * not a hard "Excel cannot open" issue).
 *
 * `numFmts` uses numeric IDs, including ~100 built-in IDs < 164 that are
 * never declared. We only flag ids >= 164 that have no matching numFmt
 * entry.
 */

import type { ValidationContext } from "./context";
import { attrByLocalName, findChildLocal, findChildrenLocal } from "./xml-utils";

const STYLES_PATH = "xl/styles.xml";

/** IDs < 164 are built-in. Custom numFmts must declare IDs >= 164. */
const CUSTOM_NUMFMT_FIRST_ID = 164;

export function checkStyles(ctx: ValidationContext): void {
  if (!ctx.has(STYLES_PATH)) {
    return;
  }
  const dom = ctx.readDom(STYLES_PATH);
  if (!dom) {
    return;
  }
  const root = dom.root;

  // Collect declared indices.
  const numFmtIds = new Set<number>();
  const numFmtsEl = findChildLocal(root, "numFmts");
  if (numFmtsEl) {
    for (const nf of findChildrenLocal(numFmtsEl, "numFmt")) {
      const idAttr = attrByLocalName(nf, "numFmtId");
      const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
      if (Number.isFinite(id)) {
        numFmtIds.add(id);
      }
    }
  }
  const fontCount = countChildren(root, "fonts", "font");
  const fillCount = countChildren(root, "fills", "fill");
  const borderCount = countChildren(root, "borders", "border");

  const cellXfs = findChildLocal(root, "cellXfs");
  if (!cellXfs) {
    return;
  }
  let xfIdx = -1;
  for (const xf of findChildrenLocal(cellXfs, "xf")) {
    xfIdx++;
    if (ctx.reporter.capped) {
      return;
    }
    const numFmtId = intAttr(xf, "numFmtId");
    if (numFmtId !== undefined && numFmtId >= CUSTOM_NUMFMT_FIRST_ID && !numFmtIds.has(numFmtId)) {
      ctx.reporter.warning(
        "styles-numFmt-missing-for-xf",
        `cellXfs[${xfIdx}] references numFmtId=${numFmtId} which is not declared in <numFmts>`,
        STYLES_PATH
      );
    }
    const fontId = intAttr(xf, "fontId");
    if (fontId !== undefined && (fontId < 0 || fontId >= fontCount)) {
      ctx.reporter.warning(
        "styles-font-index-oob",
        `cellXfs[${xfIdx}] fontId=${fontId} is outside [0, ${fontCount})`,
        STYLES_PATH
      );
    }
    const fillId = intAttr(xf, "fillId");
    if (fillId !== undefined && (fillId < 0 || fillId >= fillCount)) {
      ctx.reporter.warning(
        "styles-fill-index-oob",
        `cellXfs[${xfIdx}] fillId=${fillId} is outside [0, ${fillCount})`,
        STYLES_PATH
      );
    }
    const borderId = intAttr(xf, "borderId");
    if (borderId !== undefined && (borderId < 0 || borderId >= borderCount)) {
      ctx.reporter.warning(
        "styles-border-index-oob",
        `cellXfs[${xfIdx}] borderId=${borderId} is outside [0, ${borderCount})`,
        STYLES_PATH
      );
    }
  }
}

function countChildren(
  root: { children: unknown[] } & { name: string },
  containerName: string,
  childName: string
): number {
  const container = findChildLocal(root as never, containerName);
  if (!container) {
    return 0;
  }
  const countAttr = attrByLocalName(container, "count");
  const n = countAttr !== undefined ? parseInt(countAttr, 10) : NaN;
  if (Number.isFinite(n) && n >= 0) {
    return n;
  }
  return findChildrenLocal(container, childName).length;
}

function intAttr(
  el: { attributes: Record<string, string> } & { name: string },
  localName: string
): number | undefined {
  const raw = attrByLocalName(el as never, localName);
  if (raw === undefined) {
    return undefined;
  }
  const n = parseInt(raw, 10);
  return Number.isFinite(n) ? n : undefined;
}
