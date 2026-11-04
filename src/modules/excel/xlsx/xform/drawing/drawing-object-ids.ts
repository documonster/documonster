import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";

/**
 * The anchor fields that carry a drawing object's `cNvPr/@id`, as the anchor xforms leave them after `prepare`.
 *
 * Kept structural: the anchor models are assembled by several writers, and this only needs to read and set ids.
 */
interface IdentifiedAnchor {
  preserved?: EchoNode;
  picture?: { index?: number };
  graphicFrame?: { index?: number };
  shape?: { kind?: string; cNvPrId?: number };
}

/**
 * Give every object in one drawing part its own `cNvPr/@id`.
 *
 * Each producer numbered its own objects — a picture by its anchor position, a chart by the id it was read with, a
 * form control by its VML shape — and an object kept verbatim from a file brought the id it had there. Nothing checked
 * them against each other, so a text box read as `id="4"` and the fourth picture written beside it shared an id, which
 * is a drawing Excel has to repair.
 *
 * Objects kept verbatim claim their ids first, because their XML refers to those ids — a connector names the shapes it
 * joins by them. One is renumbered only when an earlier object already holds its id, and the connectors in the same
 * object follow it. Everything this library writes is then numbered around them, keeping its planned id when free.
 */
export function assignDrawingObjectIds(anchors: readonly unknown[]): void {
  const used = new Set<number>();
  let next = 1;
  const free = (): number => {
    while (used.has(next)) {
      next++;
    }
    used.add(next);
    return next;
  };
  const claim = (planned: number | undefined): number => {
    if (planned !== undefined && planned > 0 && !used.has(planned)) {
      used.add(planned);
      return planned;
    }
    return free();
  };

  for (const anchor of anchors as IdentifiedAnchor[]) {
    if (anchor?.preserved !== undefined) {
      anchor.preserved = claimPreserved(anchor.preserved, used, free);
    }
  }
  for (const anchor of anchors as IdentifiedAnchor[]) {
    if (anchor === null || anchor === undefined || anchor.preserved !== undefined) {
      continue;
    }
    if (anchor.picture !== undefined) {
      anchor.picture.index = claim(anchor.picture.index);
    } else if (anchor.graphicFrame !== undefined) {
      anchor.graphicFrame.index = claim(anchor.graphicFrame.index);
    } else if (anchor.shape !== undefined) {
      anchor.shape.cNvPrId = claim(anchor.shape.cNvPrId);
    }
  }
}

/** Claim a preserved object's ids, returning a renumbered copy when any of them was taken. */
function claimPreserved(node: EchoNode, used: Set<number>, free: () => number): EchoNode {
  // Almost always nothing collides, so the ids are checked before anything is copied: copying every preserved object
  // on every write was most of the time a drawing-heavy workbook took to write.
  const ids = idsOf(node);
  const unique = new Set(ids);
  if (unique.size === ids.length && ids.every(id => !used.has(id))) {
    for (const id of ids) {
      used.add(id);
    }
    return node;
  }
  // Worked on a copy, so the workbook's own model keeps the ids it was read with and the next write starts from them.
  const copy = structuredClone(node);
  const renumbered = new Map<string, string>();
  const claimIds = (current: EchoNode): void => {
    if (localName(current.tag) === "cNvPr") {
      const id = Number(current.attrs.id);
      if (Number.isInteger(id) && id > 0) {
        if (used.has(id)) {
          const fresh = String(free());
          renumbered.set(String(id), fresh);
          current.attrs.id = fresh;
        } else {
          used.add(id);
        }
      }
    }
    current.children.forEach(claimIds);
  };
  claimIds(copy);
  if (renumbered.size === 0) {
    return node;
  }
  // A connector names the shapes it joins by id, so it follows a shape that moved.
  const followConnectors = (current: EchoNode): void => {
    const local = localName(current.tag);
    if (local === "stCxn" || local === "endCxn") {
      const id = renumbered.get(String(current.attrs.id));
      if (id !== undefined) {
        current.attrs.id = id;
      }
    }
    current.children.forEach(followConnectors);
  };
  followConnectors(copy);
  return copy;
}

/**
 * The ids a preserved object carries, by object. Nothing edits a preserved object in place — it is renumbered on a copy
 * — so they are read once rather than on every write.
 */
const idsByObject = new WeakMap<EchoNode, readonly number[]>();

function idsOf(node: EchoNode): readonly number[] {
  let ids = idsByObject.get(node);
  if (ids === undefined) {
    const found: number[] = [];
    collectIds(node, found);
    ids = found;
    idsByObject.set(node, ids);
  }
  return ids;
}

/** The positive integer `cNvPr/@id`s in `node`, in document order. */
function collectIds(node: EchoNode, ids: number[]): void {
  if (localName(node.tag) === "cNvPr") {
    const id = Number(node.attrs.id);
    if (Number.isInteger(id) && id > 0) {
      ids.push(id);
    }
  }
  for (const child of node.children) {
    collectIds(child, ids);
  }
}

function localName(tag: string): string {
  return tag.slice(tag.indexOf(":") + 1);
}
