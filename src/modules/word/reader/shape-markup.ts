/**
 * What a `wps:wsp` carries beyond the fields of `DrawingShape`.
 *
 * The shape model covers what callers usually edit — geometry preset, a plain
 * fill and line colour, the line width, rotation, text and its vertical
 * anchor. Everything else is kept here, by schema position, for the writer
 * to put back (see `DrawingShapePreserved`).
 */

import type { Mutable } from "@word/core/internal-utils";
import { findRelationshipIdReferences } from "@word/core/relationship-ids";
import { notePreservedLinks } from "@word/reader/drawing-frame";
import { serializeElement } from "@word/reader/parse-utils";
import type { DrawingShape, DrawingShapePreserved } from "@word/types";
import type { XmlElement } from "@xml/types";

type Writable<T> = { -readonly [K in keyof T]: T[K] };

function localName(el: XmlElement): string {
  return el.name.slice(el.name.indexOf(":") + 1);
}

function elements(el: XmlElement): XmlElement[] {
  return el.children.filter((c): c is XmlElement => c.type === "element");
}

/** `el`'s attributes without `omit` and namespace declarations, or undefined when none remain. */
function attrsExcept(el: XmlElement, ...omit: string[]): Record<string, string> | undefined {
  const out: Record<string, string> = {};
  let any = false;
  for (const [k, v] of Object.entries(el.attributes)) {
    if (!omit.includes(k) && k !== "xmlns" && !k.startsWith("xmlns:")) {
      out[k] = v;
      any = true;
    }
  }
  return any ? out : undefined;
}

function serializeAll(els: readonly XmlElement[]): string | undefined {
  return els.length > 0 ? els.map(serializeElement).join("") : undefined;
}

/** `<a:solidFill><a:srgbClr val="…"/></a:solidFill>` exactly — what `fillColor` / `outlineColor` express. */
function isPlainSolidFill(el: XmlElement): boolean {
  const kids = elements(el);
  return (
    localName(el) === "solidFill" &&
    Object.keys(attrsExcept(el) ?? {}).length === 0 &&
    kids.length === 1 &&
    localName(kids[0]!) === "srgbClr" &&
    elements(kids[0]!).length === 0 &&
    Object.keys(attrsExcept(kids[0]!, "val") ?? {}).length === 0
  );
}

function isEmptyNoFill(el: XmlElement): boolean {
  return localName(el) === "noFill" && el.children.length === 0;
}

const FILLS = new Set(["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"]);
const LINE_FILLS = new Set(["noFill", "solidFill", "gradFill", "pattFill"]);

/**
 * The markup of `wspEl` that `shape`'s fields — already read from it — do
 * not express. `undefined` when they express all of it.
 */
export function preserveShapeMarkup(
  wspEl: XmlElement,
  shape: Mutable<DrawingShape>
): DrawingShapePreserved | undefined {
  const kept: Writable<Omit<DrawingShapePreserved, "read">> = {};

  kept.wspAttrs = attrsExcept(wspEl);
  for (const child of elements(wspEl)) {
    switch (localName(child)) {
      case "cNvPr": {
        kept.cNvPrAttrs = attrsExcept(child, "id", "name");
        kept.cNvPrChildrenXml = serializeAll(elements(child));
        if (kept.cNvPrChildrenXml) {
          const ids = new Set(findRelationshipIdReferences(kept.cNvPrChildrenXml));
          if (ids.size > 0) {
            kept.links = [...ids].map(id => ({ id, url: "" }));
            notePreservedLinks(kept.links);
          }
        }
        break;
      }
      case "cNvSpPr":
      case "cNvCnPr":
        if (localName(child) === "cNvCnPr" || child.children.length > 0 || attrsExcept(child)) {
          kept.cNvSpPrXml = serializeElement(child);
        }
        break;
      case "spPr":
        preserveSpPr(child, shape, kept);
        break;
      case "style":
        kept.styleXml = serializeElement(child);
        break;
      case "extLst":
        kept.wspExtLstXml = serializeElement(child);
        break;
      case "txbx": {
        kept.txbxAttrs = attrsExcept(child);
        const extra: XmlElement[] = [];
        for (const part of elements(child)) {
          if (localName(part) === "txbxContent") {
            // Only paragraphs are modelled; anything else keeps the whole
            // content as it was.
            if (elements(part).some(block => localName(block) !== "p")) {
              kept.txbxContentXml = serializeElement(part);
            }
          } else {
            extra.push(part);
          }
        }
        kept.txbxExtLstXml = serializeAll(extra);
        break;
      }
      case "linkedTxbx":
        kept.linkedTxbxXml = serializeElement(child);
        break;
      case "bodyPr":
        kept.bodyPrAttrs = attrsExcept(child, "anchor");
        kept.bodyPrChildrenXml = serializeAll(elements(child));
        break;
    }
  }

  const anyKept = Object.values(kept).some(v => v !== undefined);
  if (!anyKept) {
    return undefined;
  }
  return {
    read: {
      shapeType: shape.shapeType,
      width: shape.width,
      height: shape.height,
      ...(shape.fillColor !== undefined ? { fillColor: shape.fillColor } : {}),
      ...(shape.noFill ? { noFill: true } : {}),
      ...(shape.outlineColor !== undefined ? { outlineColor: shape.outlineColor } : {}),
      ...(shape.noOutline ? { noOutline: true } : {}),
      ...(shape.textContent ? { textContent: shape.textContent } : {})
    },
    ...Object.fromEntries(Object.entries(kept).filter(([, v]) => v !== undefined))
  } as DrawingShapePreserved;
}

function preserveSpPr(
  spPr: XmlElement,
  shape: DrawingShape,
  kept: Writable<Omit<DrawingShapePreserved, "read">>
): void {
  kept.spPrAttrs = attrsExcept(spPr);
  let xfrm: XmlElement | undefined;
  const tail: XmlElement[] = [];
  let lineSeen = false;
  for (const child of elements(spPr)) {
    const name = localName(child);
    if (name === "xfrm") {
      xfrm = child;
    } else if (name === "custGeom") {
      kept.geometryXml = serializeElement(child);
    } else if (name === "prstGeom") {
      const avLst = elements(child).find(c => localName(c) === "avLst");
      const others = elements(child).filter(c => localName(c) !== "avLst");
      if ((avLst && elements(avLst).length > 0) || others.length > 0) {
        kept.geometryXml = serializeElement(child);
      }
    } else if (FILLS.has(name) && !lineSeen) {
      if (!isPlainSolidFill(child) && !isEmptyNoFill(child)) {
        kept.fillXml = serializeElement(child);
      }
    } else if (name === "ln") {
      lineSeen = true;
      kept.lnAttrs = attrsExcept(child, "w");
      const lineTail: XmlElement[] = [];
      for (const part of elements(child)) {
        if (LINE_FILLS.has(localName(part))) {
          if (!isPlainSolidFill(part) && !isEmptyNoFill(part)) {
            kept.lnFillXml = serializeElement(part);
          }
        } else {
          lineTail.push(part);
        }
      }
      kept.lnTailXml = serializeAll(lineTail);
    } else {
      tail.push(child);
    }
  }
  kept.spPrTailXml = serializeAll(tail);

  if (!xfrm) {
    kept.xfrm = null;
    return;
  }
  const off = elements(xfrm).find(c => localName(c) === "off");
  const ext = elements(xfrm).find(c => localName(c) === "ext");
  const offAttrs = off ? attrsExcept(off) : undefined;
  const extDiffers =
    ext !== undefined &&
    (ext.attributes["cx"] !== String(shape.width) || ext.attributes["cy"] !== String(shape.height));
  const offDiffers = offAttrs !== undefined && (offAttrs["x"] !== "0" || offAttrs["y"] !== "0");
  if (offDiffers || extDiffers || !off || !ext) {
    kept.xfrm = {
      ...(off ? { off: offAttrs ?? {} } : {}),
      ...(ext ? { ext: attrsExcept(ext) ?? {} } : {})
    };
  }
}
