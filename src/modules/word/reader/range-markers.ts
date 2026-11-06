/**
 * Reading range markers (bookmark, comment, move and custom-XML revision
 * ranges), inside a paragraph or between blocks. See {@link RangeMarker}.
 */

import type { Mutable } from "@word/core/internal-utils";
import { RANGE_MARKER_TYPES } from "@word/core/text-utils";
import { attrVal } from "@word/reader/parse-utils";
import type { BookmarkStart, PermEnd, PermStart, RangeMarker } from "@word/types";
import type { XmlElement } from "@xml/types";

/**
 * Local name of every range marker element, by the names it is read under.
 * Asked of every block, row and cell of a document, so it is one lookup
 * rather than a string rewrite and an allocation per element.
 */
const RANGE_MARKER_NAMES: ReadonlyMap<string, string> = new Map(
  RANGE_MARKER_TYPES.flatMap(name => [
    [name, name],
    [`w:${name}`, name]
  ])
);

/** The shortest range marker name, `permEnd`. */
const SHORTEST_MARKER_NAME = 7;

/** Local name of `el` when it is a range marker, otherwise `undefined`. */
function rangeMarkerName(el: XmlElement): string | undefined {
  // Most elements asked about are `w:p`, `w:tbl` or `w:tc`.
  return el.name.length < SHORTEST_MARKER_NAME ? undefined : RANGE_MARKER_NAMES.get(el.name);
}

/**
 * Parse a range marker element (`name` is its local name). Returns `undefined`
 * when it carries no usable id: without one it cannot be paired with its other
 * end, and fabricating id 0 would collide with every other such marker.
 */
export function parseRangeMarker(el: XmlElement, name: string): RangeMarker | undefined {
  const idAttr = el.attributes["w:id"] ?? el.attributes["id"];
  const id = idAttr !== undefined ? parseInt(idAttr, 10) : NaN;
  if (Number.isNaN(id)) {
    return undefined;
  }
  // Bookmarks and comment ranges are nearly all the markers a document has,
  // so their path is kept short.
  switch (name) {
    case "bookmarkStart":
      return parseBookmarkStart(el, id);
    case "bookmarkEnd":
      return { type: "bookmarkEnd", id };
    case "commentRangeStart":
      return { type: "commentRangeStart", id };
    case "commentRangeEnd":
      return { type: "commentRangeEnd", id };
    default:
      return parseOtherRangeMarker(el, name, id);
  }
}

function parseBookmarkStart(el: XmlElement, id: number): BookmarkStart {
  const attrs = el.attributes;
  const bm: Mutable<BookmarkStart> = {
    type: "bookmarkStart",
    id,
    name: attrs["w:name"] ?? attrs["name"] ?? ""
  };
  const colFirst = attrs["w:colFirst"] ?? attrs["colFirst"];
  if (colFirst !== undefined) {
    bm.colFirst = parseInt(colFirst, 10);
  }
  const colLast = attrs["w:colLast"] ?? attrs["colLast"];
  if (colLast !== undefined) {
    bm.colLast = parseInt(colLast, 10);
  }
  const dcx = attrs["w:displacedByCustomXml"] ?? attrs["displacedByCustomXml"];
  if (dcx === "next" || dcx === "prev") {
    bm.displacedByCustomXml = dcx;
  }
  return bm;
}

function parseOtherRangeMarker(el: XmlElement, name: string, id: number): RangeMarker {
  switch (name) {
    case "permStart":
    case "permEnd": {
      const perm: Mutable<PermStart> | Mutable<PermEnd> =
        name === "permStart" ? { type: "permStart", id } : { type: "permEnd", id };
      const text = (attr: string): string | undefined =>
        el.attributes[`w:${attr}`] ?? el.attributes[attr];
      const dcx = text("displacedByCustomXml");
      if (dcx === "next" || dcx === "prev") {
        perm.displacedByCustomXml = dcx;
      }
      if (perm.type === "permStart") {
        const edGrp = text("edGrp");
        if (edGrp !== undefined) {
          perm.edGrp = edGrp;
        }
        const ed = text("ed");
        if (ed !== undefined) {
          perm.ed = ed;
        }
        const colFirst = text("colFirst");
        if (colFirst !== undefined) {
          perm.colFirst = parseInt(colFirst, 10);
        }
        const colLast = text("colLast");
        if (colLast !== undefined) {
          perm.colLast = parseInt(colLast, 10);
        }
      }
      return perm;
    }
    default: {
      const marker: { type: string; id: number; author?: string; date?: string; name?: string } = {
        type: name,
        id
      };
      const author = attrVal(el, "author");
      if (author) {
        marker.author = author;
      }
      const date = attrVal(el, "date");
      if (date) {
        marker.date = date;
      }
      if (name.startsWith("move")) {
        const mName = attrVal(el, "name");
        if (mName) {
          marker.name = mName;
        }
      }
      return marker as RangeMarker;
    }
  }
}

/** Whether `el` is a range marker element. */
export function isRangeMarkerElement(el: XmlElement): boolean {
  return rangeMarkerName(el) !== undefined;
}

/** Parse `el` into `out` if it is a range marker; `false` when it is not one. */
export function takeRangeMarker(el: XmlElement, out: unknown[]): boolean {
  const name = rangeMarkerName(el);
  if (!name) {
    return false;
  }
  const marker = parseRangeMarker(el, name);
  if (marker) {
    out.push(marker);
  }
  return true;
}
