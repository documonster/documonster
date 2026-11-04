/**
 * DOCX Writers - Range markers
 *
 * Renders bookmark, comment, permission, move and custom-XML revision range
 * markers (`EG_RangeMarkupElements`), wherever they appear: between blocks,
 * among paragraph children, inside hyperlinks and inside complex fields.
 */

import type {
  BookmarkEnd,
  BookmarkStart,
  CommentRangeEnd,
  CommentRangeStart,
  PermEnd,
  PermStart,
  RangeMarker,
  SmartTagStart
} from "@word/types";
import type { RenderHelpers } from "@word/writer/render-context";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render a bookmark start. */
function renderBookmarkStart(xml: XmlSink, bm: BookmarkStart, helpers?: RenderHelpers): void {
  const attrs: Record<string, string> = {
    "w:id": xmlValue(
      helpers?.bookmarkIds ? helpers.bookmarkIds.start(bm.id, bm.name) : bm.id,
      "w:id"
    ),
    "w:name": bm.name
  };
  if (bm.colFirst !== undefined) {
    attrs["w:colFirst"] = xmlValue(bm.colFirst, "w:colFirst");
  }
  if (bm.colLast !== undefined) {
    attrs["w:colLast"] = xmlValue(bm.colLast, "w:colLast");
  }
  if (bm.displacedByCustomXml) {
    attrs["w:displacedByCustomXml"] = bm.displacedByCustomXml;
  }
  xml.leafNode("w:bookmarkStart", attrs);
}

/** Render a bookmark end. */
function renderBookmarkEnd(xml: XmlSink, bm: BookmarkEnd, helpers?: RenderHelpers): void {
  const id = helpers?.bookmarkIds ? helpers.bookmarkIds.end(bm.id, bm.name) : bm.id;
  if (id !== undefined) {
    xml.leafNode("w:bookmarkEnd", { "w:id": xmlValue(id, "w:id") });
  }
}

/** Render a range-permission start. */
function renderPermStart(xml: XmlSink, perm: PermStart, helpers?: RenderHelpers): void {
  const attrs: Record<string, string> = {
    "w:id": xmlValue(helpers?.permIds ? helpers.permIds.start(perm.id) : perm.id, "w:id")
  };
  if (perm.edGrp !== undefined) {
    attrs["w:edGrp"] = perm.edGrp;
  }
  if (perm.ed !== undefined) {
    attrs["w:ed"] = perm.ed;
  }
  if (perm.colFirst !== undefined) {
    attrs["w:colFirst"] = xmlValue(perm.colFirst, "w:colFirst");
  }
  if (perm.colLast !== undefined) {
    attrs["w:colLast"] = xmlValue(perm.colLast, "w:colLast");
  }
  if (perm.displacedByCustomXml) {
    attrs["w:displacedByCustomXml"] = perm.displacedByCustomXml;
  }
  xml.leafNode("w:permStart", attrs);
}

/** Render a range-permission end; one that closes no start is dropped, as for bookmarks. */
function renderPermEnd(xml: XmlSink, perm: PermEnd, helpers?: RenderHelpers): void {
  const id = helpers?.permIds ? helpers.permIds.end(perm.id) : perm.id;
  if (id === undefined) {
    return;
  }
  const attrs: Record<string, string> = { "w:id": xmlValue(id, "w:id") };
  if (perm.displacedByCustomXml) {
    attrs["w:displacedByCustomXml"] = perm.displacedByCustomXml;
  }
  xml.leafNode("w:permEnd", attrs);
}

/** Render comment range start. */
function renderCommentRangeStart(xml: XmlSink, cr: CommentRangeStart): void {
  xml.leafNode("w:commentRangeStart", { "w:id": xmlValue(cr.id, "w:id") });
}

/** Render comment range end. */
function renderCommentRangeEnd(xml: XmlSink, cr: CommentRangeEnd): void {
  xml.leafNode("w:commentRangeEnd", { "w:id": xmlValue(cr.id, "w:id") });
}

/**
 * The wrapper elements span markers have opened in the content being written,
 * innermost last. One per sequence of paragraph content (a paragraph, a
 * hyperlink); see {@link closeOpenSpans}.
 */
export type OpenSpans = ("bidi" | "smartTag")[];

/**
 * Render one range marker. A span marker opens or closes its wrapper element
 * and needs `spans`, the wrappers open in the content it sits in: without it
 * (between blocks, where no span can be) it is skipped, and an end that closes
 * no open wrapper of its kind is skipped too.
 */
export function renderRangeMarker(
  xml: XmlSink,
  marker: RangeMarker,
  helpers?: RenderHelpers,
  spans?: OpenSpans
): void {
  switch (marker.type) {
    case "bidiStart":
      if (spans) {
        xml.openNode(marker.kind === "override" ? "w:bdo" : "w:dir", { "w:val": marker.direction });
        spans.push("bidi");
      }
      return;
    case "smartTagStart":
      if (spans) {
        renderSmartTagOpen(xml, marker);
        spans.push("smartTag");
      }
      return;
    case "bidiEnd":
    case "smartTagEnd":
      if (spans && spans[spans.length - 1] === (marker.type === "bidiEnd" ? "bidi" : "smartTag")) {
        spans.pop();
        xml.closeNode();
      }
      return;
    case "bookmarkStart":
      renderBookmarkStart(xml, marker, helpers);
      return;
    case "bookmarkEnd":
      renderBookmarkEnd(xml, marker, helpers);
      return;
    case "permStart":
      renderPermStart(xml, marker, helpers);
      return;
    case "permEnd":
      renderPermEnd(xml, marker, helpers);
      return;
    case "commentRangeStart":
      renderCommentRangeStart(xml, marker);
      return;
    case "commentRangeEnd":
      renderCommentRangeEnd(xml, marker);
      return;
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd": {
      const attrs: Record<string, string> = { "w:id": xmlValue(marker.id, "w:id") };
      if (marker.author) {
        attrs["w:author"] = marker.author;
      }
      if (marker.date) {
        attrs["w:date"] = marker.date;
      }
      if (marker.name) {
        attrs["w:name"] = marker.name;
      }
      xml.leafNode(`w:${marker.type}`, attrs);
      return;
    }
    default: {
      // Custom-XML revision ranges.
      const attrs: Record<string, string> = { "w:id": xmlValue(marker.id, "w:id") };
      if (marker.author) {
        attrs["w:author"] = marker.author;
      }
      if (marker.date) {
        attrs["w:date"] = marker.date;
      }
      xml.leafNode(`w:${marker.type}`, attrs);
    }
  }
}

/**
 * Render range markers as siblings of the content around them, which the
 * schema allows wherever `EG_RangeMarkupElements` appears.
 */
export function renderRangeMarkers(
  xml: XmlSink,
  markers: readonly RangeMarker[] | undefined,
  helpers?: RenderHelpers,
  spans?: OpenSpans
): void {
  // Asked of every paragraph, child, row and cell written; nearly all have none.
  if (markers === undefined) {
    return;
  }
  for (const marker of markers) {
    renderRangeMarker(xml, marker, helpers, spans);
  }
}

/** Close the wrappers still open at the end of their content (a span left unpaired by editing). */
export function closeOpenSpans(xml: XmlSink, spans: OpenSpans): void {
  while (spans.length > 0) {
    spans.pop();
    xml.closeNode();
  }
}

function renderSmartTagOpen(xml: XmlSink, tag: SmartTagStart): void {
  const attrs: Record<string, string> = {};
  if (tag.uri !== undefined) {
    attrs["w:uri"] = tag.uri;
  }
  attrs["w:element"] = tag.element;
  xml.openNode("w:smartTag", attrs);
  if (tag.attributes && tag.attributes.length > 0) {
    xml.openNode("w:smartTagPr");
    for (const attr of tag.attributes) {
      const a: Record<string, string> = {};
      if (attr.uri !== undefined) {
        a["w:uri"] = attr.uri;
      }
      a["w:name"] = attr.name;
      a["w:val"] = attr.val;
      xml.leafNode("w:attr", a);
    }
    xml.closeNode();
  }
}
