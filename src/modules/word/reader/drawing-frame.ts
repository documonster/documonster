/**
 * The frame of a modelled drawing — `wp:inline` / `wp:anchor` — as read.
 *
 * The drawing's fields model what layout and editing need; everything else
 * about the frame is kept here so the writer can put it back (see
 * `DrawingFrame`). A hyperlink on `docPr` is a relationship of the part, so
 * it is resolved once the part's relationships are known: every link read
 * while a part is parsed is noted ({@link readingPreservedLinks}) and the
 * part resolves exactly those ({@link resolvePreservedLinks}), without
 * walking the model it built.
 */

import { RelType } from "@word/constants";
import { findRelationshipIdReferences } from "@word/core/relationship-ids";
import { serializeElement } from "@word/reader/parse-utils";
import type { ParsedRelationship } from "@word/reader/reader-context";
import type {
  DrawingFrame,
  DrawingFrameChild,
  DrawingFrameFields,
  PreservedLink
} from "@word/types";
import type { XmlElement } from "@xml/types";

/** Where the links being read go — see {@link readingPreservedLinks}. */
let noted: PreservedLink[] | undefined;

/**
 * Run `read` — which must be synchronous, so nothing else is read meanwhile —
 * with every {@link PreservedLink} it creates added to `into`.
 */
export function readingPreservedLinks<T>(into: PreservedLink[], read: () => T): T {
  const previous = noted;
  noted = into;
  try {
    return read();
  } finally {
    noted = previous;
  }
}

/** Record `links`, just created, for the part being read. */
export function notePreservedLinks(links: readonly PreservedLink[]): void {
  if (noted) {
    for (const link of links) {
      noted.push(link);
    }
  }
}

function localName(name: string): string {
  return name.slice(name.indexOf(":") + 1);
}

function attrsExcept(el: XmlElement, omit: readonly string[]): Record<string, string> | undefined {
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

/** The fields of a drawing that model part of its frame. */
const FRAME_FIELDS: readonly (keyof DrawingFrameFields)[] = [
  "width",
  "height",
  "name",
  "altText",
  "horizontalPosition",
  "verticalPosition",
  "wrap",
  "relativeHeight",
  "behindDoc",
  "lockAnchor",
  "layoutInCell",
  "allowOverlap",
  "simplePos",
  "distT",
  "distB",
  "distL",
  "distR"
];

/** A snapshot of `drawing`'s frame fields. */
export function frameFieldsOf(drawing: DrawingFrameFields): DrawingFrameFields {
  const out: Record<string, unknown> = {};
  for (const field of FRAME_FIELDS) {
    if (drawing[field] !== undefined) {
      out[field] = drawing[field];
    }
  }
  return out as DrawingFrameFields;
}

/**
 * The frame `frameEl` (a `wp:inline` or `wp:anchor`) carries, given the
 * fields the drawing was read with. Hyperlinks are recorded by id and
 * resolved later, against the part's relationships.
 */
export function readDrawingFrame(frameEl: XmlElement, read: DrawingFrameFields): DrawingFrame {
  const children: DrawingFrameChild[] = [];
  let docPrAttrs: Record<string, string> | undefined;
  let docPrChildrenXml: string | undefined;
  let links: PreservedLink[] | undefined;
  let afterGraphic = false;
  for (const child of frameEl.children) {
    if (child.type !== "element") {
      continue;
    }
    const name = localName(child.name);
    if (name === "graphic") {
      afterGraphic = true;
    } else if (name === "docPr") {
      docPrAttrs = attrsExcept(child, ["id", "name", "descr"]);
      const kids = child.children.filter((c): c is XmlElement => c.type === "element");
      if (kids.length > 0) {
        docPrChildrenXml = kids.map(serializeElement).join("");
        const ids = new Set(findRelationshipIdReferences(docPrChildrenXml));
        links = [...ids].map(id => ({ id, url: "" }));
        notePreservedLinks(links);
      }
    } else {
      children.push({
        name,
        xml: serializeElement(child),
        ...(afterGraphic ? { afterGraphic } : {})
      });
    }
  }
  return {
    kind: localName(frameEl.name) === "anchor" ? "anchor" : "inline",
    attrs: attrsExcept(frameEl, []) ?? {},
    children,
    ...(docPrAttrs ? { docPrAttrs } : {}),
    ...(docPrChildrenXml ? { docPrChildrenXml } : {}),
    ...(links && links.length > 0 ? { links } : {}),
    read
  };
}

/**
 * Give every hyperlink in `links` — those read from one part, see
 * {@link readingPreservedLinks} — its target, from the part's relationships.
 * A link whose id is not an external hyperlink of the part keeps an empty
 * target, and the writer drops its id.
 */
export function resolvePreservedLinks(
  links: readonly PreservedLink[],
  relMap: ReadonlyMap<string, ParsedRelationship>
): void {
  for (const link of links) {
    if (link.url === "") {
      const rel = relMap.get(link.id);
      if (rel && rel.type === RelType.Hyperlink) {
        (link as { url: string }).url = rel.target;
      }
    }
  }
}
