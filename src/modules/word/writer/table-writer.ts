/**
 * DOCX Writers - Table
 *
 * Renders w:tbl, w:tr, w:tc elements with full property support
 * including floating tables, table look, cell spacing, and revisions.
 */

import type { XmlSink } from "@xml/types";

import type {
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TableRowProperties,
  TableCellProperties,
  TableBorders,
  TableWidth,
  TableCellMargins,
  TableLook,
  TableFloat
} from "../types";
import { renderBorder, renderParagraph } from "./paragraph-writer";
import type { RenderHelpers } from "./render-context";
import { renderShading } from "./run-writer";

/** Render a table width element. */
function renderTableWidth(xml: XmlSink, tagName: string, tw: TableWidth): void {
  xml.leafNode(tagName, { "w:w": String(tw.value), "w:type": tw.type });
}

/** Render table borders. */
function renderTableBorders(xml: XmlSink, tagName: string, borders: TableBorders): void {
  xml.openNode(tagName);
  if (borders.top) {
    renderBorder(xml, "w:top", borders.top);
  }
  if (borders.left) {
    renderBorder(xml, "w:left", borders.left);
  }
  if (borders.bottom) {
    renderBorder(xml, "w:bottom", borders.bottom);
  }
  if (borders.right) {
    renderBorder(xml, "w:right", borders.right);
  }
  if (borders.insideH) {
    renderBorder(xml, "w:insideH", borders.insideH);
  }
  if (borders.insideV) {
    renderBorder(xml, "w:insideV", borders.insideV);
  }
  if (borders.start) {
    renderBorder(xml, "w:start", borders.start);
  }
  if (borders.end) {
    renderBorder(xml, "w:end", borders.end);
  }
  if (borders.tl2br) {
    renderBorder(xml, "w:tl2br", borders.tl2br);
  }
  if (borders.tr2bl) {
    renderBorder(xml, "w:tr2bl", borders.tr2bl);
  }
  xml.closeNode();
}

/** Render table cell margins. */
function renderCellMargins(xml: XmlSink, tagName: string, margins: TableCellMargins): void {
  xml.openNode(tagName);
  if (margins.top) {
    renderTableWidth(xml, "w:top", margins.top);
  }
  if (margins.left) {
    renderTableWidth(xml, "w:left", margins.left);
  }
  if (margins.bottom) {
    renderTableWidth(xml, "w:bottom", margins.bottom);
  }
  if (margins.right) {
    renderTableWidth(xml, "w:right", margins.right);
  }
  if (margins.start) {
    renderTableWidth(xml, "w:start", margins.start);
  }
  if (margins.end) {
    renderTableWidth(xml, "w:end", margins.end);
  }
  xml.closeNode();
}

/** Render table look (conditional formatting flags). */
function renderTableLook(xml: XmlSink, look: TableLook): void {
  // Per ECMA-376, w:tblLook has individual attributes (w:firstRow etc.) as the
  // authoritative representation. w:val is a legacy 4-hex-digit bitmask kept
  // for compatibility with older consumers.
  //
  // We write individual attributes AND compute w:val to match them. This
  // ensures both representations agree (no conflict).
  const attrs: Record<string, string> = {};
  let val = 0;

  if (look.firstRow !== undefined) {
    attrs["w:firstRow"] = look.firstRow ? "1" : "0";
    if (look.firstRow) {
      val |= 0x0020;
    }
  }
  if (look.lastRow !== undefined) {
    attrs["w:lastRow"] = look.lastRow ? "1" : "0";
    if (look.lastRow) {
      val |= 0x0040;
    }
  }
  if (look.firstColumn !== undefined) {
    attrs["w:firstColumn"] = look.firstColumn ? "1" : "0";
    if (look.firstColumn) {
      val |= 0x0080;
    }
  }
  if (look.lastColumn !== undefined) {
    attrs["w:lastColumn"] = look.lastColumn ? "1" : "0";
    if (look.lastColumn) {
      val |= 0x0100;
    }
  }
  if (look.noHBand !== undefined) {
    attrs["w:noHBand"] = look.noHBand ? "1" : "0";
    if (look.noHBand) {
      val |= 0x0200;
    }
  }
  if (look.noVBand !== undefined) {
    attrs["w:noVBand"] = look.noVBand ? "1" : "0";
    if (look.noVBand) {
      val |= 0x0400;
    }
  }

  attrs["w:val"] = val.toString(16).toUpperCase().padStart(4, "0");
  xml.leafNode("w:tblLook", attrs);
}

/** Render floating table properties (w:tblpPr). */
function renderTableFloat(xml: XmlSink, float: TableFloat): void {
  const attrs: Record<string, string> = {};
  if (float.horizontalAnchor) {
    attrs["w:horzAnchor"] = float.horizontalAnchor;
  }
  if (float.verticalAnchor) {
    attrs["w:vertAnchor"] = float.verticalAnchor;
  }
  if (float.absoluteHorizontalPosition !== undefined) {
    attrs["w:tblpX"] = String(float.absoluteHorizontalPosition);
  }
  if (float.absoluteVerticalPosition !== undefined) {
    attrs["w:tblpY"] = String(float.absoluteVerticalPosition);
  }
  if (float.relativeHorizontalPosition) {
    attrs["w:tblpXSpec"] = float.relativeHorizontalPosition;
  }
  if (float.relativeVerticalPosition) {
    attrs["w:tblpYSpec"] = float.relativeVerticalPosition;
  }
  if (float.topFromText !== undefined) {
    attrs["w:topFromText"] = String(float.topFromText);
  }
  if (float.bottomFromText !== undefined) {
    attrs["w:bottomFromText"] = String(float.bottomFromText);
  }
  if (float.leftFromText !== undefined) {
    attrs["w:leftFromText"] = String(float.leftFromText);
  }
  if (float.rightFromText !== undefined) {
    attrs["w:rightFromText"] = String(float.rightFromText);
  }
  // Note: w:tblOverlap is a separate sibling element of w:tblpPr, not an attribute.
  // It's rendered separately by the caller (renderTableProperties).
  xml.leafNode("w:tblpPr", attrs);
}

/** Render w:tblPr. */
function renderTableProperties(
  xml: XmlSink,
  tPr: TableProperties,
  insidePropertyChange = false
): void {
  xml.openNode("w:tblPr");

  if (tPr.style) {
    xml.leafNode("w:tblStyle", { "w:val": tPr.style });
  }

  // Floating table must come before width
  if (tPr.float) {
    renderTableFloat(xml, tPr.float);
    // w:tblOverlap is a separate sibling element
    if (tPr.float.overlap) {
      xml.leafNode("w:tblOverlap", { "w:val": tPr.float.overlap });
    }
  }

  if (tPr.width) {
    renderTableWidth(xml, "w:tblW", tPr.width);
  }

  if (tPr.alignment) {
    xml.leafNode("w:jc", { "w:val": tPr.alignment });
  }

  if (tPr.cellSpacing) {
    renderTableWidth(xml, "w:tblCellSpacing", tPr.cellSpacing);
  }

  if (tPr.indent !== undefined) {
    xml.leafNode("w:tblInd", { "w:w": String(tPr.indent), "w:type": "dxa" });
  }

  if (tPr.borders) {
    renderTableBorders(xml, "w:tblBorders", tPr.borders);
  }

  if (tPr.shading) {
    renderShading(xml, tPr.shading);
  }

  if (tPr.layout) {
    xml.leafNode("w:tblLayout", { "w:type": tPr.layout });
  }

  if (tPr.cellMargins) {
    renderCellMargins(xml, "w:tblCellMar", tPr.cellMargins);
  }

  if (tPr.look) {
    renderTableLook(xml, tPr.look);
  }

  if (tPr.visuallyRightToLeft) {
    xml.leafNode("w:bidiVisual");
  }

  // Accessibility: caption and description
  if (tPr.caption !== undefined) {
    xml.leafNode("w:tblCaption", { "w:val": tPr.caption });
  }
  if (tPr.description !== undefined) {
    xml.leafNode("w:tblDescription", { "w:val": tPr.description });
  }

  // Table property change (track changes) - must be last child
  if (!insidePropertyChange && tPr.propertyChange) {
    const rev = tPr.propertyChange.revision;
    const attrs: Record<string, string> = {
      "w:id": String(rev.id),
      "w:author": rev.author
    };
    if (rev.date) {
      attrs["w:date"] = rev.date;
    }
    xml.openNode("w:tblPrChange", attrs);
    if (tPr.propertyChange.previousProperties) {
      renderTableProperties(xml, tPr.propertyChange.previousProperties, true);
    } else {
      xml.openNode("w:tblPr");
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.closeNode();
}

/** Render w:trPr. */
function renderTableRowProperties(
  xml: XmlSink,
  rPr: TableRowProperties,
  insidePropertyChange = false
): void {
  xml.openNode("w:trPr");

  // Per ECMA-376 CT_TrPr schema order:
  // cnfStyle → divId → gridBefore → gridAfter → wBefore → wAfter → cantSplit →
  // trHeight → tblHeader → tblCellSpacing → jc → hidden → ins/del
  if (rPr.cnfStyle) {
    xml.leafNode("w:cnfStyle", { "w:val": rPr.cnfStyle });
  }

  if (rPr.gridBefore !== undefined) {
    xml.leafNode("w:gridBefore", { "w:val": String(rPr.gridBefore) });
  }
  if (rPr.gridAfter !== undefined) {
    xml.leafNode("w:gridAfter", { "w:val": String(rPr.gridAfter) });
  }
  if (rPr.widthBefore) {
    renderTableWidth(xml, "w:wBefore", rPr.widthBefore);
  }
  if (rPr.widthAfter) {
    renderTableWidth(xml, "w:wAfter", rPr.widthAfter);
  }

  if (rPr.cantSplit) {
    xml.leafNode("w:cantSplit");
  }

  if (rPr.height) {
    const attrs: Record<string, string> = { "w:val": String(rPr.height.value) };
    if (rPr.height.rule) {
      attrs["w:hRule"] = rPr.height.rule;
    }
    xml.leafNode("w:trHeight", attrs);
  }

  if (rPr.tableHeader) {
    xml.leafNode("w:tblHeader");
  }

  if (rPr.cellSpacing) {
    renderTableWidth(xml, "w:tblCellSpacing", rPr.cellSpacing);
  }

  if (rPr.hidden) {
    xml.leafNode("w:hidden");
  }

  // Track changes: inserted row
  if (rPr.inserted) {
    const attrs: Record<string, string> = {
      "w:id": String(rPr.inserted.revision.id),
      "w:author": rPr.inserted.revision.author
    };
    if (rPr.inserted.revision.date) {
      attrs["w:date"] = rPr.inserted.revision.date;
    }
    xml.leafNode("w:ins", attrs);
  }

  // Track changes: deleted row
  if (rPr.deleted) {
    const attrs: Record<string, string> = {
      "w:id": String(rPr.deleted.revision.id),
      "w:author": rPr.deleted.revision.author
    };
    if (rPr.deleted.revision.date) {
      attrs["w:date"] = rPr.deleted.revision.date;
    }
    xml.leafNode("w:del", attrs);
  }

  // Row property change (track changes)
  if (!insidePropertyChange && rPr.propertyChange) {
    const rev = rPr.propertyChange.revision;
    const attrs: Record<string, string> = {
      "w:id": String(rev.id),
      "w:author": rev.author
    };
    if (rev.date) {
      attrs["w:date"] = rev.date;
    }
    xml.openNode("w:trPrChange", attrs);
    if (rPr.propertyChange.previousProperties) {
      renderTableRowProperties(xml, rPr.propertyChange.previousProperties, true);
    } else {
      xml.openNode("w:trPr");
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.closeNode();
}

/** Render w:tcPr. */
function renderTableCellProperties(
  xml: XmlSink,
  cPr: TableCellProperties,
  insidePropertyChange = false
): void {
  xml.openNode("w:tcPr");

  // Per CT_TcPr order: cnfStyle → tcW → gridSpan → hMerge → vMerge → tcBorders →
  // shd → noWrap → tcMar → textDirection → tcFitText → vAlign → hideMark
  if (cPr.cnfStyle) {
    xml.leafNode("w:cnfStyle", { "w:val": cPr.cnfStyle });
  }

  if (cPr.width) {
    renderTableWidth(xml, "w:tcW", cPr.width);
  } else {
    // OOXML's CT_TcPr.tcW is technically optional, but Word treats a
    // missing <w:tcW> as a hard schema violation: it triggers the
    // "content unreadable" repair dialog and silently drops the table.
    // Emit a w:type="auto" placeholder so Word lays the cell out using
    // the table-level grid widths.
    xml.leafNode("w:tcW", { "w:w": "0", "w:type": "auto" });
  }

  if (cPr.gridSpan !== undefined && cPr.gridSpan > 1) {
    xml.leafNode("w:gridSpan", { "w:val": String(cPr.gridSpan) });
  }

  if (cPr.verticalMerge !== undefined) {
    if (cPr.verticalMerge === "restart") {
      xml.leafNode("w:vMerge", { "w:val": "restart" });
    } else {
      xml.leafNode("w:vMerge");
    }
  }

  if (cPr.borders) {
    renderTableBorders(xml, "w:tcBorders", cPr.borders);
  }

  if (cPr.shading) {
    renderShading(xml, cPr.shading);
  }

  if (cPr.noWrap) {
    xml.leafNode("w:noWrap");
  }

  if (cPr.margins) {
    renderCellMargins(xml, "w:tcMar", cPr.margins);
  }

  if (cPr.textDirection) {
    xml.leafNode("w:textDirection", { "w:val": cPr.textDirection });
  }

  if (cPr.fitText) {
    xml.leafNode("w:tcFitText");
  }

  if (cPr.verticalAlign) {
    xml.leafNode("w:vAlign", { "w:val": cPr.verticalAlign });
  }

  if (cPr.hideMark) {
    xml.leafNode("w:hideMark");
  }

  // Cell-level revisions (cellIns/cellDel/cellMerge)
  if (cPr.inserted) {
    const attrs: Record<string, string> = {
      "w:id": String(cPr.inserted.revision.id),
      "w:author": cPr.inserted.revision.author
    };
    if (cPr.inserted.revision.date) {
      attrs["w:date"] = cPr.inserted.revision.date;
    }
    xml.leafNode("w:cellIns", attrs);
  }
  if (cPr.deleted) {
    const attrs: Record<string, string> = {
      "w:id": String(cPr.deleted.revision.id),
      "w:author": cPr.deleted.revision.author
    };
    if (cPr.deleted.revision.date) {
      attrs["w:date"] = cPr.deleted.revision.date;
    }
    xml.leafNode("w:cellDel", attrs);
  }
  if (cPr.cellMerge) {
    const attrs: Record<string, string> = {
      "w:vMerge": cPr.cellMerge.vMerge,
      "w:id": String(cPr.cellMerge.revision.id),
      "w:author": cPr.cellMerge.revision.author
    };
    if (cPr.cellMerge.revision.date) {
      attrs["w:date"] = cPr.cellMerge.revision.date;
    }
    xml.leafNode("w:cellMerge", attrs);
  }

  // Cell property change (track changes)
  if (!insidePropertyChange && cPr.propertyChange) {
    const rev = cPr.propertyChange.revision;
    const attrs: Record<string, string> = {
      "w:id": String(rev.id),
      "w:author": rev.author
    };
    if (rev.date) {
      attrs["w:date"] = rev.date;
    }
    xml.openNode("w:tcPrChange", attrs);
    if (cPr.propertyChange.previousProperties) {
      renderTableCellProperties(xml, cPr.propertyChange.previousProperties, true);
    } else {
      xml.openNode("w:tcPr");
      xml.closeNode();
    }
    xml.closeNode();
  }

  xml.closeNode();
}

/** Render a table cell. */
function renderTableCell(xml: XmlSink, cell: TableCell, helpers?: RenderHelpers): void {
  xml.openNode("w:tc");

  // OOXML CT_Tc requires a <w:tcPr> with at least <w:tcW> for Word to
  // accept the cell. If the model didn't supply any cell properties we
  // synthesise a minimal <w:tcPr><w:tcW w:type="auto"/></w:tcPr>; if it
  // did, renderTableCellProperties handles the missing-tcW fallback.
  if (cell.properties) {
    renderTableCellProperties(xml, cell.properties);
  } else {
    xml.openNode("w:tcPr");
    xml.leafNode("w:tcW", { "w:w": "0", "w:type": "auto" });
    xml.closeNode();
  }

  // OOXML §17.4.66 (CT_Tc) requires every table cell to *end* with a
  // <w:p>. An empty cell needs at least one <w:p>, and a cell whose last
  // block is a nested <w:tbl> must have a trailing <w:p> after it — Word
  // (and LibreOffice) reject the document otherwise.
  if (cell.content.length === 0) {
    xml.openNode("w:p");
    xml.closeNode();
  } else {
    for (const block of cell.content) {
      if (block.type === "paragraph") {
        renderParagraph(xml, block, helpers);
      } else if (block.type === "table") {
        renderTable(xml, block, helpers);
      }
    }
    const last = cell.content[cell.content.length - 1];
    if (last && last.type === "table") {
      xml.openNode("w:p");
      xml.closeNode();
    }
  }

  xml.closeNode();
}

/** Render a table row. */
function renderTableRow(xml: XmlSink, row: TableRow, helpers?: RenderHelpers): void {
  xml.openNode("w:tr");

  // Table property exception (w:tblPrEx) must come before w:trPr per schema
  if (row.properties?.tblPrEx) {
    renderTablePropertiesEx(xml, row.properties.tblPrEx);
  }

  if (row.properties) {
    renderTableRowProperties(xml, row.properties);
  }

  for (const cell of row.cells) {
    renderTableCell(xml, cell, helpers);
  }

  xml.closeNode();
}

/** Render w:tblPrEx (table-level property exceptions applied to a row). */
function renderTablePropertiesEx(xml: XmlSink, ex: TableProperties): void {
  xml.openNode("w:tblPrEx");
  // Output the subset of tblPr that can appear in tblPrEx
  if (ex.width) {
    renderTableWidth(xml, "w:tblW", ex.width);
  }
  if (ex.alignment) {
    xml.leafNode("w:jc", { "w:val": ex.alignment });
  }
  if (ex.cellSpacing) {
    renderTableWidth(xml, "w:tblCellSpacing", ex.cellSpacing);
  }
  if (ex.indent !== undefined) {
    xml.leafNode("w:tblInd", { "w:w": String(ex.indent), "w:type": "dxa" });
  }
  if (ex.borders) {
    renderTableBorders(xml, "w:tblBorders", ex.borders);
  }
  if (ex.shading) {
    renderShading(xml, ex.shading);
  }
  if (ex.layout) {
    xml.leafNode("w:tblLayout", { "w:type": ex.layout });
  }
  if (ex.cellMargins) {
    renderCellMargins(xml, "w:tblCellMar", ex.cellMargins);
  }
  if (ex.look) {
    renderTableLook(xml, ex.look);
  }
  xml.closeNode();
}

/** Render a w:tbl element. */
export function renderTable(xml: XmlSink, table: Table, helpers?: RenderHelpers): void {
  xml.openNode("w:tbl");

  if (table.properties) {
    renderTableProperties(xml, table.properties);
  }

  // Column grid. ECMA-376 §17.4.49 requires <w:tblGrid> with at least one
  // <w:gridCol>. When the caller did not supply explicit column widths we
  // synthesise the grid by inferring the column count from the widest row
  // (counting gridSpan + gridBefore + gridAfter) and dividing the table
  // width evenly. This keeps the package valid for Word / LibreOffice
  // even when authors used the convenience builders that don't compute
  // a grid up front.
  xml.openNode("w:tblGrid");
  if (table.columnWidths && table.columnWidths.length > 0) {
    for (const w of table.columnWidths) {
      xml.leafNode("w:gridCol", { "w:w": String(w) });
    }
  } else {
    const columnCount = Math.max(1, computeTableColumnCount(table));
    const totalWidth = inferTableWidthInTwips(table.properties);
    const colWidth = Math.max(1, Math.floor(totalWidth / columnCount));
    for (let i = 0; i < columnCount; i++) {
      xml.leafNode("w:gridCol", { "w:w": String(colWidth) });
    }
  }
  xml.closeNode();

  for (const row of table.rows) {
    renderTableRow(xml, row, helpers);
  }

  xml.closeNode();
}

/**
 * Count the number of grid columns the table requires.
 *
 * Each cell occupies `gridSpan` columns (default 1); each row may also
 * declare `gridBefore` / `gridAfter` to leave empty grid units on its sides.
 * The grid must be wide enough for every row, so we take the maximum.
 */
function computeTableColumnCount(table: Table): number {
  let max = 1;
  for (const row of table.rows) {
    let count = row.properties?.gridBefore ?? 0;
    for (const cell of row.cells) {
      count += cell.properties?.gridSpan ?? 1;
    }
    count += row.properties?.gridAfter ?? 0;
    if (count > max) {
      max = count;
    }
  }
  return max;
}

/**
 * Pick a sensible total width (in twips) for a synthesised tblGrid.
 *
 * - If the table specifies a `dxa` width, use it directly.
 * - If it specifies a `pct` width, treat it as a fraction of a default
 *   page text area (US Letter minus 1" margins = 9 360 twips). A 5000/pct
 *   value (= 100%) maps to the full text area.
 * - Otherwise default to the same 9 360 twips.
 */
function inferTableWidthInTwips(props?: Table["properties"]): number {
  const TEXT_AREA_TWIPS = 9360;
  const w = props?.width;
  if (!w) {
    return TEXT_AREA_TWIPS;
  }
  if (w.type === "dxa" && typeof w.value === "number" && w.value > 0) {
    return w.value;
  }
  if (w.type === "pct" && typeof w.value === "number" && w.value > 0) {
    return Math.round((w.value / 5000) * TEXT_AREA_TWIPS);
  }
  return TEXT_AREA_TWIPS;
}
