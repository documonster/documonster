/**
 * The frame of a modelled drawing — `wp:inline` / `wp:anchor` — written from
 * the drawing's fields and, for a drawing that was read, from the frame
 * markup it was read with (`DrawingFrame`).
 *
 * Every structurally written drawing goes through here: inline and floating
 * pictures, shapes, charts. The writer builds the frame its fields describe
 * twice — as the fields are now and as they were read — and, part by part
 * (each attribute, each child), writes the markup as read where the two
 * agree and the fields where a field was edited. So nothing the fields do
 * not model is lost, and an edit always wins.
 */

import { MODEL_CHECK, WRITTEN, checkAnchorFields, writtenToken } from "@word/core/argument-checks";
import { drawingInteger } from "@word/core/drawing-geometry";
import { rebindRelationshipIds } from "@word/core/relationship-ids";
import { DocxWriteError } from "@word/errors";
import type { DrawingFrame, DrawingFrameFields, PreservedLink } from "@word/types";
import { DEFAULT_RELATIVE_HEIGHT, DEFAULT_WRAP_MARGIN_EMU } from "@word/units";
import type { RenderHelpers } from "@word/writer/render-context";
import { prepareRawMarkup } from "@word/writer/render-context";
import { xmlEncodeAttr, xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** A frame child the fields describe. */
type Slot =
  | "simplePos"
  | "positionH"
  | "positionV"
  | "extent"
  | "effectExtent"
  | "wrap"
  | "cNvGraphicFramePr";

/** The frame a drawing's fields describe. */
export interface FrameModel {
  readonly kind: "inline" | "anchor";
  readonly attrs: Readonly<Record<string, string>>;
  readonly slots: Partial<Record<Slot, string>>;
  readonly name: string;
  readonly descr?: string;
}

/** Which defaults the drawing kind writes when its fields say nothing. */
export type FrameVariant = "picture" | "shape" | "chart";

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

function tag(name: string, attrs: Readonly<Record<string, string>>, inner?: string): string {
  const a = Object.entries(attrs)
    .map(([k, v]) => {
      if (typeof v !== "string") {
        throw new DocxWriteError(`<${name}> ${k}: expected a string, got ${typeof v}`);
      }
      return ` ${k}="${xmlEncodeAttr(v)}"`;
    })
    .join("");
  return inner === undefined ? `<${name}${a}/>` : `<${name}${a}>${inner}</${name}>`;
}

function position(
  name: "wp:positionH" | "wp:positionV",
  pos:
    | { readonly relativeTo?: string; readonly offset?: number; readonly align?: string }
    | undefined,
  defaultRelative: string,
  field: "horizontalPosition" | "verticalPosition"
): string {
  const horizontal = name === "wp:positionH";
  const relativeTo =
    pos &&
    writtenToken(
      pos,
      "relativeTo",
      horizontal ? WRITTEN.horizontalRelative : WRITTEN.verticalRelative,
      `${field}.relativeTo`
    );
  const align =
    pos &&
    writtenToken(
      pos,
      "align",
      horizontal ? WRITTEN.horizontalAlign : WRITTEN.verticalAlign,
      `${field}.align`
    );
  const inner =
    align !== undefined
      ? `<wp:align>${xmlEncodeAttr(align)}</wp:align>`
      : `<wp:posOffset>${drawingInteger(pos?.offset, `${field}.offset`, 0)}</wp:posOffset>`;
  return tag(name, { relativeFrom: relativeTo ?? defaultRelative }, inner);
}

/** A wrap polygon round the whole drawing: tight and through wrapping require one. */
const FULL_POLYGON =
  `<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>` +
  `<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>`;

function wrapElement(wrap: DrawingFrameFields["wrap"]): string {
  if (wrap !== undefined) {
    checkAnchorFields({ wrap: { ...wrap, side: undefined } }, MODEL_CHECK);
  }
  const side = {
    wrapText: (wrap && writtenToken(wrap, "side", WRITTEN.wrapSide, "wrap.side")) ?? "bothSides"
  };
  switch (wrap?.style) {
    case "square":
      return tag("wp:wrapSquare", side);
    case "tight":
      return tag("wp:wrapTight", side, FULL_POLYGON);
    case "through":
      return tag("wp:wrapThrough", side, FULL_POLYGON);
    case "topAndBottom":
      return tag("wp:wrapTopAndBottom", {});
    default:
      return tag("wp:wrapNone", {});
  }
}

const PICTURE_LOCKS = `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS_A}" noChangeAspect="1"/></wp:cNvGraphicFramePr>`;

/** Where a wrap distance is read from: the frame's own field, else its wrap margin. */
function distField(
  fields: DrawingFrameFields,
  own: "distT" | "distB" | "distL" | "distR",
  side: "top" | "bottom" | "left" | "right"
): string {
  return fields[own] === undefined && fields.wrap?.margins?.[side] !== undefined
    ? `wrap.margins.${side}`
    : own;
}

/** The frame `fields` describe, for a drawing of `variant` written as `kind`. */
export function frameModel(
  kind: "inline" | "anchor",
  variant: FrameVariant,
  fields: DrawingFrameFields,
  defaultName: string
): FrameModel {
  const slots: Partial<Record<Slot, string>> = {
    extent: tag("wp:extent", {
      cx: drawingInteger(fields.width, "width", 0),
      cy: drawingInteger(fields.height, "height", 0)
    }),
    effectExtent: tag("wp:effectExtent", { l: "0", t: "0", r: "0", b: "0" }),
    cNvGraphicFramePr: variant === "picture" ? PICTURE_LOCKS : tag("wp:cNvGraphicFramePr", {})
  };
  const base = {
    name: fields.name ?? defaultName,
    ...(fields.altText ? { descr: fields.altText } : {})
  };
  if (kind === "inline") {
    return { kind, attrs: { distT: "0", distB: "0", distL: "0", distR: "0" }, slots, ...base };
  }
  const margins = fields.wrap?.margins;
  const attrs: Record<string, string> = {
    // A distance the frame does not set comes from the wrap margins, and is named there.
    distT: drawingInteger(fields.distT ?? margins?.top, distField(fields, "distT", "top"), 0),
    distB: drawingInteger(fields.distB ?? margins?.bottom, distField(fields, "distB", "bottom"), 0),
    distL: drawingInteger(
      fields.distL ?? margins?.left,
      distField(fields, "distL", "left"),
      DEFAULT_WRAP_MARGIN_EMU
    ),
    distR: drawingInteger(
      fields.distR ?? margins?.right,
      distField(fields, "distR", "right"),
      DEFAULT_WRAP_MARGIN_EMU
    ),
    simplePos: fields.simplePos ? "1" : "0",
    relativeHeight: drawingInteger(
      fields.relativeHeight,
      "relativeHeight",
      DEFAULT_RELATIVE_HEIGHT
    ),
    behindDoc: fields.behindDoc ? "1" : "0",
    locked: fields.lockAnchor ? "1" : "0",
    layoutInCell: fields.layoutInCell === false ? "0" : "1",
    allowOverlap: fields.allowOverlap === false ? "0" : "1"
  };
  slots.simplePos = tag("wp:simplePos", {
    x: drawingInteger(fields.simplePos?.x, "simplePos.x", 0),
    y: drawingInteger(fields.simplePos?.y, "simplePos.y", 0)
  });
  slots.positionH = position(
    "wp:positionH",
    fields.horizontalPosition,
    "column",
    "horizontalPosition"
  );
  slots.positionV = position(
    "wp:positionV",
    fields.verticalPosition,
    "paragraph",
    "verticalPosition"
  );
  slots.wrap = wrapElement(fields.wrap);
  return { kind, attrs, slots, ...base };
}

/** Schema position of each frame child (ECMA-376 CT_Anchor / CT_Inline). */
const ANCHOR_ORDER: Readonly<Record<string, number>> = {
  simplePos: 0,
  positionH: 1,
  positionV: 2,
  extent: 3,
  effectExtent: 4,
  wrapNone: 5,
  wrapSquare: 5,
  wrapTight: 5,
  wrapThrough: 5,
  wrapTopAndBottom: 5,
  docPr: 6,
  cNvGraphicFramePr: 7
};
const INLINE_ORDER: Readonly<Record<string, number>> = {
  extent: 0,
  effectExtent: 1,
  docPr: 2,
  cNvGraphicFramePr: 3
};

function slotOf(name: string): Slot | undefined {
  if (name.startsWith("wrap")) {
    return "wrap";
  }
  return (
    ["simplePos", "positionH", "positionV", "extent", "effectExtent", "cNvGraphicFramePr"] as const
  ).find(s => s === name);
}

/** The relationship ids a frame's hyperlinks get in the part being written. */
function linkRemap(
  links: readonly PreservedLink[] | undefined,
  helpers: RenderHelpers | undefined
): Map<string, string> {
  const remap = new Map<string, string>();
  for (const link of links ?? []) {
    remap.set(
      link.id,
      (helpers?.preservedLinkRId
        ? helpers.preservedLinkRId(link)
        : helpers?.hyperlinkRIds?.get(link)) ?? ""
    );
  }
  return remap;
}

/**
 * Markup that refers to preserved hyperlinks, with each id replaced by the
 * relationship the part being written registered for it.
 */
export function rebindPreservedLinks(
  xml: string,
  links: readonly PreservedLink[] | undefined,
  helpers: RenderHelpers | undefined
): string {
  return links && links.length > 0 ? rebindRelationshipIds(xml, linkRemap(links, helpers)) : xml;
}

/**
 * Open the frame and write everything before `a:graphic`. The caller writes
 * the graphic, then calls {@link closeDrawingFrame}.
 */
export function openDrawingFrame(
  xml: XmlSink,
  now: FrameModel,
  frame: DrawingFrame | undefined,
  then: FrameModel | undefined,
  drawingId: number,
  helpers: RenderHelpers | undefined
): void {
  const kept = frame && then && frame.kind === now.kind ? frame : undefined;
  const element = now.kind === "anchor" ? "wp:anchor" : "wp:inline";
  const order = now.kind === "anchor" ? ANCHOR_ORDER : INLINE_ORDER;
  const docPrIndex = order["docPr"]!;

  // Attributes: as read, except where a field now says otherwise.
  let attrs: Record<string, string> = { ...now.attrs };
  if (kept) {
    attrs = { ...kept.attrs };
    for (const key of new Set([...Object.keys(now.attrs), ...Object.keys(then!.attrs)])) {
      if (now.attrs[key] !== then!.attrs[key]) {
        if (now.attrs[key] === undefined) {
          delete attrs[key];
        } else {
          attrs[key] = now.attrs[key]!;
        }
      }
    }
  }
  xml.openNode(element, attrs);

  const items: { index: number; xml: string }[] = [];
  const docPr = tag(
    "wp:docPr",
    {
      id: xmlValue(drawingId, "id"),
      name: now.name,
      ...(now.descr !== undefined ? { descr: now.descr } : {}),
      ...kept?.docPrAttrs
    },
    kept?.docPrChildrenXml
      ? rebindPreservedLinks(kept.docPrChildrenXml, kept.links, helpers)
      : undefined
  );
  items.push({ index: docPrIndex, xml: docPr });

  if (!kept) {
    for (const [slot, markup] of Object.entries(now.slots)) {
      const index = slot === "wrap" ? order["wrapNone"]! : order[slot]!;
      items.push({ index, xml: markup });
    }
  } else {
    const seen = new Set<Slot>();
    let previous = -1;
    for (const child of kept.children) {
      if (child.afterGraphic) {
        continue;
      }
      const slot = slotOf(child.name);
      const index = order[child.name] ?? previous;
      previous = index;
      if (!slot) {
        items.push({ index, xml: prepareRawMarkup(child.xml, helpers) });
        continue;
      }
      seen.add(slot);
      const current = now.slots[slot];
      if (current === then!.slots[slot]) {
        items.push({ index, xml: prepareRawMarkup(child.xml, helpers) });
      } else if (current !== undefined) {
        items.push({ index, xml: current });
      }
    }
    // A child the frame did not have, that an edited field now calls for.
    for (const [slot, markup] of Object.entries(now.slots) as [Slot, string][]) {
      if (!seen.has(slot) && markup !== then!.slots[slot]) {
        const index = slot === "wrap" ? order["wrapNone"]! : order[slot]!;
        items.push({ index, xml: markup });
      }
    }
  }
  // Stable: children of one schema position keep their order.
  items
    .map((item, i) => ({ ...item, i }))
    .sort((a, b) => a.index - b.index || a.i - b.i)
    .forEach(item => xml.writeRaw(item.xml));
}

/** Write what follows `a:graphic` (the `wp14` relative sizes) and close the frame. */
export function closeDrawingFrame(
  xml: XmlSink,
  now: FrameModel,
  frame: DrawingFrame | undefined,
  helpers: RenderHelpers | undefined
): void {
  if (frame && frame.kind === now.kind) {
    for (const child of frame.children) {
      if (child.afterGraphic) {
        xml.writeRaw(prepareRawMarkup(child.xml, helpers));
      }
    }
  }
  xml.closeNode();
}
