/**
 * DOCX Writers - Section Properties
 *
 * Renders w:sectPr element including page size, margins, columns,
 * headers/footers, page numbering, page borders, vertical alignment, etc.
 */

import type { XmlSink } from "@xml/types";

import {
  DEFAULT_PAGE_WIDTH,
  DEFAULT_PAGE_HEIGHT,
  DEFAULT_MARGIN,
  DEFAULT_HEADER_FOOTER_MARGIN,
  DEFAULT_COLUMN_SPACE
} from "../constants";
import type {
  SectionProperties,
  SectionColumns,
  HeaderFooterRef,
  PageBorders,
  FootnoteProperties,
  EndnoteProperties
} from "../types";
import { renderBorderElement } from "./run-writer";

/** Render footnote/endnote properties. */
function renderNoteProperties(
  xml: XmlSink,
  tagName: string,
  props: FootnoteProperties | EndnoteProperties
): void {
  xml.openNode(tagName);
  if (props.position) {
    xml.leafNode("w:pos", { "w:val": props.position });
  }
  if (props.numFmt) {
    xml.leafNode("w:numFmt", { "w:val": props.numFmt });
  }
  if (props.numStart !== undefined) {
    xml.leafNode("w:numStart", { "w:val": String(props.numStart) });
  }
  if (props.numRestart) {
    xml.leafNode("w:numRestart", { "w:val": props.numRestart });
  }
  xml.closeNode();
}

/** Render a header or footer reference. */
function renderHeaderFooterRef(xml: XmlSink, tagName: string, ref: HeaderFooterRef): void {
  xml.leafNode(tagName, { "w:type": ref.type, "r:id": ref.rId });
}

/** Render column layout. */
function renderColumns(xml: XmlSink, cols: SectionColumns): void {
  const attrs: Record<string, string> = {};
  if (cols.space !== undefined) {
    attrs["w:space"] = String(cols.space);
  }
  if (cols.count !== undefined) {
    attrs["w:num"] = String(cols.count);
  }
  if (cols.equalWidth !== undefined) {
    attrs["w:equalWidth"] = cols.equalWidth ? "1" : "0";
  }
  if (cols.separator) {
    attrs["w:sep"] = "1";
  }

  if (cols.columns && cols.columns.length > 0) {
    xml.openNode("w:cols", attrs);
    for (const col of cols.columns) {
      const colAttrs: Record<string, string> = { "w:w": String(col.width) };
      if (col.space !== undefined) {
        colAttrs["w:space"] = String(col.space);
      }
      xml.leafNode("w:col", colAttrs);
    }
    xml.closeNode();
  } else {
    xml.leafNode("w:cols", attrs);
  }
}

/** Render page borders (w:pgBorders). */
function renderPageBorders(xml: XmlSink, borders: PageBorders): void {
  const attrs: Record<string, string> = {};
  if (borders.display) {
    attrs["w:display"] = borders.display;
  }
  if (borders.offsetFrom) {
    attrs["w:offsetFrom"] = borders.offsetFrom;
  }
  if (borders.zOrder) {
    attrs["w:zOrder"] = borders.zOrder;
  }
  xml.openNode("w:pgBorders", Object.keys(attrs).length > 0 ? attrs : undefined);
  if (borders.top) {
    renderBorderElement(xml, "w:top", borders.top);
  }
  if (borders.left) {
    renderBorderElement(xml, "w:left", borders.left);
  }
  if (borders.bottom) {
    renderBorderElement(xml, "w:bottom", borders.bottom);
  }
  if (borders.right) {
    renderBorderElement(xml, "w:right", borders.right);
  }
  xml.closeNode();
}

/** Render w:sectPr element. */
export function renderSectionProperties(
  xml: XmlSink,
  sect: SectionProperties,
  insidePropertyChange = false
): void {
  xml.openNode("w:sectPr");

  // Header references
  if (sect.headers) {
    for (const ref of sect.headers) {
      renderHeaderFooterRef(xml, "w:headerReference", ref);
    }
  }

  // Footer references
  if (sect.footers) {
    for (const ref of sect.footers) {
      renderHeaderFooterRef(xml, "w:footerReference", ref);
    }
  }

  // Section break type
  if (sect.breakType) {
    xml.leafNode("w:type", { "w:val": sect.breakType });
  }

  // Page size
  {
    const ps = sect.pageSize;
    const w = ps?.width ?? DEFAULT_PAGE_WIDTH;
    const h = ps?.height ?? DEFAULT_PAGE_HEIGHT;
    const attrs: Record<string, string> = { "w:w": String(w), "w:h": String(h) };
    if (ps?.orientation === "landscape") {
      attrs["w:orient"] = "landscape";
    }
    xml.leafNode("w:pgSz", attrs);
  }

  // Page margins
  {
    const m = sect.margins;
    xml.leafNode("w:pgMar", {
      "w:top": String(m?.top ?? DEFAULT_MARGIN),
      "w:right": String(m?.right ?? DEFAULT_MARGIN),
      "w:bottom": String(m?.bottom ?? DEFAULT_MARGIN),
      "w:left": String(m?.left ?? DEFAULT_MARGIN),
      "w:header": String(m?.header ?? DEFAULT_HEADER_FOOTER_MARGIN),
      "w:footer": String(m?.footer ?? DEFAULT_HEADER_FOOTER_MARGIN),
      "w:gutter": String(m?.gutter ?? 0)
    });
  }

  // Page borders
  if (sect.pageBorders) {
    renderPageBorders(xml, sect.pageBorders);
  }

  // Columns
  if (sect.columns) {
    renderColumns(xml, sect.columns);
  } else {
    xml.leafNode("w:cols", { "w:space": String(DEFAULT_COLUMN_SPACE) });
  }

  // Title page (different first page header/footer)
  if (sect.titlePage) {
    xml.leafNode("w:titlePg");
  }

  // Page numbering
  if (sect.pageNumbering) {
    const attrs: Record<string, string> = {};
    if (sect.pageNumbering.start !== undefined) {
      attrs["w:start"] = String(sect.pageNumbering.start);
    }
    if (sect.pageNumbering.format) {
      attrs["w:fmt"] = sect.pageNumbering.format;
    }
    xml.leafNode("w:pgNumType", attrs);
  }

  // Line numbers
  if (sect.lineNumbers) {
    const attrs: Record<string, string> = {};
    if (sect.lineNumbers.countBy !== undefined) {
      attrs["w:countBy"] = String(sect.lineNumbers.countBy);
    }
    if (sect.lineNumbers.start !== undefined) {
      attrs["w:start"] = String(sect.lineNumbers.start);
    }
    if (sect.lineNumbers.restart) {
      attrs["w:restart"] = sect.lineNumbers.restart;
    }
    if (sect.lineNumbers.distance !== undefined) {
      attrs["w:distance"] = String(sect.lineNumbers.distance);
    }
    xml.leafNode("w:lnNumType", attrs);
  }

  // Vertical alignment
  if (sect.verticalAlign) {
    xml.leafNode("w:vAlign", { "w:val": sect.verticalAlign });
  }

  // Text direction
  if (sect.textDirection) {
    xml.leafNode("w:textDirection", { "w:val": sect.textDirection });
  }

  // Section-level bidi (RTL)
  if (sect.bidi) {
    xml.leafNode("w:bidi");
  }

  // RTL gutter
  if (sect.rtlGutter) {
    xml.leafNode("w:rtlGutter");
  }

  // Form protection
  if (sect.formProtection) {
    xml.leafNode("w:formProt", { "w:val": "1" });
  }

  // Footnote properties
  if (sect.footnoteProperties) {
    renderNoteProperties(xml, "w:footnotePr", sect.footnoteProperties);
  }

  // Endnote properties
  if (sect.endnoteProperties) {
    renderNoteProperties(xml, "w:endnotePr", sect.endnoteProperties);
  }

  // Document grid
  if (sect.docGrid) {
    const gridAttrs: Record<string, string> = {
      "w:linePitch": String(sect.docGrid.linePitch ?? 360)
    };
    if (sect.docGrid.charSpace !== undefined) {
      gridAttrs["w:charSpace"] = String(sect.docGrid.charSpace);
    }
    if (sect.docGrid.type) {
      gridAttrs["w:type"] = sect.docGrid.type;
    }
    xml.leafNode("w:docGrid", gridAttrs);
  } else {
    xml.leafNode("w:docGrid", { "w:linePitch": "360" });
  }

  // Section property change (track changes)
  if (!insidePropertyChange && sect.propertyChange) {
    const rev = sect.propertyChange.revision;
    const attrs: Record<string, string> = {
      "w:id": String(rev.id),
      "w:author": rev.author
    };
    if (rev.date) {
      attrs["w:date"] = rev.date;
    }
    xml.openNode("w:sectPrChange", attrs);
    if (sect.propertyChange.previousProperties) {
      renderSectionProperties(xml, sect.propertyChange.previousProperties, true);
    } else {
      xml.openNode("w:sectPr");
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.closeNode();
}
