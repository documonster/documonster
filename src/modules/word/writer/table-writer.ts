/**
 * DOCX Writers - Table
 *
 * Renders w:tbl, w:tr, w:tc elements with full property support
 * including floating tables, table look, cell spacing, and revisions.
 */

import { WRITTEN, writtenToken } from "@word/core/argument-checks";
import {
  EDGE_NAMES,
  NARROWER_PROPERTIES,
  ROW_ONLY_PROPERTIES,
  oneEdgeName
} from "@word/core/change-records";
import {
  atNode,
  refuseAbsentFields,
  refuseTableDiagonals,
  refuseTableOnlyProperties,
  unwritableNode,
  within,
  writableList
} from "@word/core/node-path";
import type {
  BlockContentControl,
  Table,
  TableRow,
  TableCell,
  TableProperties,
  TablePropertyExceptions,
  TableRowProperties,
  TableCellProperties,
  TableBorders,
  TableWidth,
  TableCellMargins,
  TableLook,
  TableFloat
} from "@word/types";
import { renderBorder, renderParagraph } from "@word/writer/paragraph-writer";
import { renderRangeMarkers } from "@word/writer/range-marker-writer";
import type { RenderHelpers } from "@word/writer/render-context";
import { revisionAttributes } from "@word/writer/render-context";
import { renderShading } from "@word/writer/run-writer";
import { emitInSchemaOrder, type OrderedChild } from "@word/writer/schema-order";
import { renderSdtProperties } from "@word/writer/sdt-writer";
import { xmlValue } from "@xml/encode";
import type { XmlSink } from "@xml/types";

/** Render a table width element. */
function renderTableWidth(xml: XmlSink, tagName: string, tw: TableWidth, at: string): void {
  within(at, () =>
    xml.leafNode(tagName, {
      "w:w": xmlValue(tw.value, "w:w"),
      "w:type": writtenToken(tw, "type", WRITTEN.tableWidthType, "type")
    })
  );
}

/** Render table borders (`w:tblBorders` or `w:tcBorders`), in their content model's order. */
export function renderTableBorders(
  xml: XmlSink,
  tagName: "w:tblBorders" | "w:tcBorders",
  borders: TableBorders
): void {
  if (tagName === "w:tblBorders") {
    refuseTableDiagonals(borders);
  }
  refuseBothEdgeNames(borders);
  xml.openNode(tagName);
  const sides =
    tagName === "w:tcBorders"
      ? ([
          "top",
          "left",
          "bottom",
          "right",
          "insideH",
          "insideV",
          "start",
          "end",
          "tl2br",
          "tr2bl"
        ] as const)
      : (["top", "left", "bottom", "right", "insideH", "insideV", "start", "end"] as const);
  emitInSchemaOrder(
    tagName === "w:tcBorders" ? "w:CT_TcBorders" : "w:CT_TblBorders",
    sides.map((side): OrderedChild => {
      const border = borders[side];
      return !!border && [`w:${side}`, () => renderBorder(xml, `w:${side}`, border)];
    })
  );
  xml.closeNode();
}

/** Refuse borders or margins naming one edge twice (`EDGE_NAMES`): Office's schema takes one name per edge. */
function refuseBothEdgeNames(set: TableBorders | TableCellMargins): void {
  for (const [first, second] of EDGE_NAMES) {
    if (set[first] !== undefined && set[second] !== undefined) {
      unwritableNode(second, oneEdgeName(first), set[second]);
    }
  }
}

/** Render table cell margins (`w:tblCellMar` or `w:tcMar`), in their content model's order. */
function renderCellMargins(
  xml: XmlSink,
  tagName: "w:tblCellMar" | "w:tcMar",
  margins: TableCellMargins
): void {
  refuseBothEdgeNames(margins);
  xml.openNode(tagName);
  emitInSchemaOrder(
    tagName === "w:tcMar" ? "w:CT_TcMar" : "w:CT_TblCellMar",
    (["top", "left", "bottom", "right", "start", "end"] as const).map((side): OrderedChild => {
      const width = margins[side];
      return !!width && [`w:${side}`, () => renderTableWidth(xml, `w:${side}`, width, side)];
    })
  );
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
    attrs["w:tblpX"] = xmlValue(float.absoluteHorizontalPosition, "w:tblpX");
  }
  if (float.absoluteVerticalPosition !== undefined) {
    attrs["w:tblpY"] = xmlValue(float.absoluteVerticalPosition, "w:tblpY");
  }
  if (float.relativeHorizontalPosition) {
    attrs["w:tblpXSpec"] = float.relativeHorizontalPosition;
  }
  if (float.relativeVerticalPosition) {
    attrs["w:tblpYSpec"] = float.relativeVerticalPosition;
  }
  if (float.topFromText !== undefined) {
    attrs["w:topFromText"] = xmlValue(float.topFromText, "w:topFromText");
  }
  if (float.bottomFromText !== undefined) {
    attrs["w:bottomFromText"] = xmlValue(float.bottomFromText, "w:bottomFromText");
  }
  if (float.leftFromText !== undefined) {
    attrs["w:leftFromText"] = xmlValue(float.leftFromText, "w:leftFromText");
  }
  if (float.rightFromText !== undefined) {
    attrs["w:rightFromText"] = xmlValue(float.rightFromText, "w:rightFromText");
  }
  // Note: w:tblOverlap is a separate sibling element of w:tblpPr, not an attribute.
  // It's rendered separately by the caller (renderTableProperties).
  xml.leafNode("w:tblpPr", attrs);
}

/**
 * Render w:tblPr. `base` is `CT_TblPrBase` — a table style's, a conditional format's, or the previous properties in
 * `w:tblPrChange` — which has no change record of its own.
 */
export function renderTableProperties(xml: XmlSink, tPr: TableProperties, base = false): void {
  if (base) {
    refuseAbsentFields(tPr, ["propertyChange"], NARROWER_PROPERTIES.baseTable);
  }
  xml.openNode("w:tblPr");
  const alignment = writtenToken(tPr, "alignment", WRITTEN.tableAlignment, "alignment");
  const layout = writtenToken(tPr, "layout", WRITTEN.tableLayout, "layout");
  emitInSchemaOrder("w:CT_TblPr", [
    !!tPr.style && ["w:tblStyle", () => xml.leafNode("w:tblStyle", { "w:val": tPr.style! })],
    !!tPr.float && ["w:tblpPr", () => renderTableFloat(xml, tPr.float!)],
    // w:tblOverlap is a separate sibling element of w:tblpPr, not an attribute.
    !!tPr.float?.overlap && [
      "w:tblOverlap",
      () => xml.leafNode("w:tblOverlap", { "w:val": tPr.float!.overlap! })
    ],
    !!tPr.visuallyRightToLeft && ["w:bidiVisual", () => xml.leafNode("w:bidiVisual")],
    tPr.rowBandSize !== undefined && [
      "w:tblStyleRowBandSize",
      () => xml.leafNode("w:tblStyleRowBandSize", { "w:val": xmlValue(tPr.rowBandSize!, "w:val") })
    ],
    tPr.colBandSize !== undefined && [
      "w:tblStyleColBandSize",
      () => xml.leafNode("w:tblStyleColBandSize", { "w:val": xmlValue(tPr.colBandSize!, "w:val") })
    ],
    ...tablePropertiesExceptionChildren(xml, tPr, alignment, layout),
    tPr.caption !== undefined && [
      "w:tblCaption",
      () => xml.leafNode("w:tblCaption", { "w:val": tPr.caption! })
    ],
    tPr.description !== undefined && [
      "w:tblDescription",
      () => xml.leafNode("w:tblDescription", { "w:val": tPr.description! })
    ],
    !base && !!tPr.propertyChange && ["w:tblPrChange", () => renderTablePropertyChange(xml, tPr)]
  ]);
  xml.closeNode();
}

/**
 * The children `w:tblPr` shares with `w:tblPrEx` (`CT_TblPrExBase`): width, justification, spacing, indent,
 * borders, shading, layout, margins, look.
 */
function tablePropertiesExceptionChildren(
  xml: XmlSink,
  tPr: TablePropertyExceptions,
  alignment: string | undefined,
  layout: string | undefined
): OrderedChild[] {
  return [
    !!tPr.width && ["w:tblW", () => renderTableWidth(xml, "w:tblW", tPr.width!, "width")],
    alignment !== undefined && ["w:jc", () => xml.leafNode("w:jc", { "w:val": alignment })],
    !!tPr.cellSpacing && [
      "w:tblCellSpacing",
      () => renderTableWidth(xml, "w:tblCellSpacing", tPr.cellSpacing!, "cellSpacing")
    ],
    tPr.indent !== undefined && [
      "w:tblInd",
      () => xml.leafNode("w:tblInd", { "w:w": xmlValue(tPr.indent!, "w:w"), "w:type": "dxa" })
    ],
    !!tPr.borders && [
      "w:tblBorders",
      () => within("borders", () => renderTableBorders(xml, "w:tblBorders", tPr.borders!))
    ],
    !!tPr.shading && ["w:shd", () => within("shading", () => renderShading(xml, tPr.shading!))],
    layout !== undefined && [
      "w:tblLayout",
      () => xml.leafNode("w:tblLayout", { "w:type": layout })
    ],
    !!tPr.cellMargins && [
      "w:tblCellMar",
      () => within("cellMargins", () => renderCellMargins(xml, "w:tblCellMar", tPr.cellMargins!))
    ],
    !!tPr.look && ["w:tblLook", () => renderTableLook(xml, tPr.look!)]
  ];
}

/** `w:tblPrChange`, holding the previous `w:tblPr`. */
function renderTablePropertyChange(xml: XmlSink, tPr: TableProperties): void {
  xml.openNode("w:tblPrChange", revisionAttributes(tPr.propertyChange!.revision));
  if (tPr.propertyChange!.previousProperties) {
    const previous = tPr.propertyChange!.previousProperties;
    within("propertyChange.previousProperties", () => renderTableProperties(xml, previous, true));
  } else {
    xml.openNode("w:tblPr");
    xml.closeNode();
  }
  xml.closeNode();
}

/** Render w:trPr — a row's, a table style's or a conditional format's — in `CT_TrPr` order. */
export function renderTableRowProperties(
  xml: XmlSink,
  rPr: TableRowProperties,
  insidePropertyChange = false
): void {
  if (insidePropertyChange) {
    refuseAbsentFields(
      rPr,
      [...ROW_ONLY_PROPERTIES, "propertyChange"],
      NARROWER_PROPERTIES.previousRow
    );
  }
  xml.openNode("w:trPr");
  const { height, propertyChange } = rPr;
  emitInSchemaOrder("w:CT_TrPr", [
    !!rPr.cnfStyle && ["w:cnfStyle", () => xml.leafNode("w:cnfStyle", { "w:val": rPr.cnfStyle! })],
    rPr.gridBefore !== undefined && [
      "w:gridBefore",
      () => xml.leafNode("w:gridBefore", { "w:val": xmlValue(rPr.gridBefore!, "w:val") })
    ],
    rPr.gridAfter !== undefined && [
      "w:gridAfter",
      () => xml.leafNode("w:gridAfter", { "w:val": xmlValue(rPr.gridAfter!, "w:val") })
    ],
    !!rPr.widthBefore && [
      "w:wBefore",
      () => renderTableWidth(xml, "w:wBefore", rPr.widthBefore!, "widthBefore")
    ],
    !!rPr.widthAfter && [
      "w:wAfter",
      () => renderTableWidth(xml, "w:wAfter", rPr.widthAfter!, "widthAfter")
    ],
    !!rPr.cantSplit && ["w:cantSplit", () => xml.leafNode("w:cantSplit")],
    !!height && [
      "w:trHeight",
      () => {
        const attrs: Record<string, string> = { "w:val": xmlValue(height.value, "w:val") };
        const rule = writtenToken(height, "rule", WRITTEN.rowHeightRule, "height.rule");
        if (rule !== undefined) {
          attrs["w:hRule"] = rule;
        }
        xml.leafNode("w:trHeight", attrs);
      }
    ],
    !!rPr.tableHeader && ["w:tblHeader", () => xml.leafNode("w:tblHeader")],
    !!rPr.cellSpacing && [
      "w:tblCellSpacing",
      () => renderTableWidth(xml, "w:tblCellSpacing", rPr.cellSpacing!, "cellSpacing")
    ],
    !!rPr.hidden && ["w:hidden", () => xml.leafNode("w:hidden")],
    // Track changes: an inserted or deleted row.
    !!rPr.inserted && [
      "w:ins",
      () => xml.leafNode("w:ins", revisionAttributes(rPr.inserted!.revision))
    ],
    !!rPr.deleted && [
      "w:del",
      () => xml.leafNode("w:del", revisionAttributes(rPr.deleted!.revision))
    ],
    !insidePropertyChange &&
      !!propertyChange && [
        "w:trPrChange",
        () => {
          xml.openNode("w:trPrChange", revisionAttributes(propertyChange.revision));
          if (propertyChange.previousProperties) {
            const previous = propertyChange.previousProperties;
            within("propertyChange.previousProperties", () =>
              renderTableRowProperties(xml, previous, true)
            );
          } else {
            xml.openNode("w:trPr");
            xml.closeNode();
          }
          xml.closeNode();
        }
      ]
  ]);
  xml.closeNode();
}

/** How {@link renderTableCellProperties} writes a `w:tcPr` that is not a table cell's own. */
export interface CellPropertiesOptions {
  /** A `w:tcPr` inside `w:tcPrChange`: no change record of its own. */
  readonly insidePropertyChange?: boolean;
  /**
   * Write `<w:tcW w:w="0" w:type="auto"/>` when the properties have no width. A table cell needs it (see below);
   * a style's `w:tcPr` does not, and writing one there would set every cell the style formats to auto width.
   */
  readonly widthPlaceholder?: boolean;
}

/** Render w:tcPr — a cell's, a table style's or a conditional format's — in `CT_TcPr` order. */
export function renderTableCellProperties(
  xml: XmlSink,
  cPr: TableCellProperties,
  { insidePropertyChange = false, widthPlaceholder = true }: CellPropertiesOptions = {}
): void {
  if (insidePropertyChange) {
    refuseAbsentFields(cPr, ["propertyChange"], NARROWER_PROPERTIES.previousCell);
  }
  xml.openNode("w:tcPr");
  const textDirection = writtenToken(cPr, "textDirection", WRITTEN.textDirection, "textDirection");
  const verticalAlign = writtenToken(
    cPr,
    "verticalAlign",
    WRITTEN.cellVerticalAlign,
    "verticalAlign"
  );
  const { propertyChange } = cPr;
  emitInSchemaOrder("w:CT_TcPr", [
    !!cPr.cnfStyle && ["w:cnfStyle", () => xml.leafNode("w:cnfStyle", { "w:val": cPr.cnfStyle! })],
    cPr.width
      ? ["w:tcW", () => renderTableWidth(xml, "w:tcW", cPr.width!, "width")]
      : // OOXML's CT_TcPr.tcW is technically optional, but Word treats a
        // missing <w:tcW> as a hard schema violation: it triggers the
        // "content unreadable" repair dialog and silently drops the table.
        // Emit a w:type="auto" placeholder so Word lays the cell out using
        // the table-level grid widths.
        widthPlaceholder && [
          "w:tcW",
          () => xml.leafNode("w:tcW", { "w:w": "0", "w:type": "auto" })
        ],
    cPr.gridSpan !== undefined &&
      cPr.gridSpan > 1 && [
        "w:gridSpan",
        () => xml.leafNode("w:gridSpan", { "w:val": xmlValue(cPr.gridSpan!, "w:val") })
      ],
    cPr.verticalMerge !== undefined && [
      "w:vMerge",
      () =>
        xml.leafNode(
          "w:vMerge",
          cPr.verticalMerge === "restart" ? { "w:val": "restart" } : undefined
        )
    ],
    !!cPr.borders && [
      "w:tcBorders",
      () => within("borders", () => renderTableBorders(xml, "w:tcBorders", cPr.borders!))
    ],
    !!cPr.shading && ["w:shd", () => within("shading", () => renderShading(xml, cPr.shading!))],
    !!cPr.noWrap && ["w:noWrap", () => xml.leafNode("w:noWrap")],
    !!cPr.margins && [
      "w:tcMar",
      () => within("margins", () => renderCellMargins(xml, "w:tcMar", cPr.margins!))
    ],
    textDirection !== undefined && [
      "w:textDirection",
      () => xml.leafNode("w:textDirection", { "w:val": textDirection })
    ],
    !!cPr.fitText && ["w:tcFitText", () => xml.leafNode("w:tcFitText")],
    verticalAlign !== undefined && [
      "w:vAlign",
      () => xml.leafNode("w:vAlign", { "w:val": verticalAlign })
    ],
    !!cPr.hideMark && ["w:hideMark", () => xml.leafNode("w:hideMark")],
    // Cell-level revisions (cellIns/cellDel/cellMerge): one `xsd:choice`.
    !!cPr.inserted && [
      "w:cellIns",
      () => xml.leafNode("w:cellIns", revisionAttributes(cPr.inserted!.revision))
    ],
    !!cPr.deleted && [
      "w:cellDel",
      () => xml.leafNode("w:cellDel", revisionAttributes(cPr.deleted!.revision))
    ],
    !!cPr.cellMerge && [
      "w:cellMerge",
      () =>
        xml.leafNode("w:cellMerge", {
          "w:vMerge": cPr.cellMerge!.vMerge,
          ...revisionAttributes(cPr.cellMerge!.revision)
        })
    ],
    !insidePropertyChange &&
      !!propertyChange && [
        "w:tcPrChange",
        () => {
          xml.openNode("w:tcPrChange", revisionAttributes(propertyChange.revision));
          if (propertyChange.previousProperties) {
            const previous = propertyChange.previousProperties;
            within("propertyChange.previousProperties", () =>
              // No placeholder width: what rejecting the change restores is what the cell had.
              renderTableCellProperties(xml, previous, {
                insidePropertyChange: true,
                widthPlaceholder: false
              })
            );
          } else {
            xml.openNode("w:tcPr");
            xml.closeNode();
          }
          xml.closeNode();
        }
      ]
  ]);
  xml.closeNode();
}

/**
 * The cell's content controls that fit its content, ordered to open: by
 * start, the outer (longer) first. One out of range, holding no block (it
 * has nowhere to be written in a cell), or overlapping another without
 * nesting is left out.
 */
function cellControls(cell: TableCell): BlockContentControl[] {
  const n = cell.content.length;
  const sorted = (cell.contentControls ?? [])
    .filter(
      c =>
        Number.isInteger(c.start) &&
        Number.isInteger(c.end) &&
        c.start >= 0 &&
        c.end <= n &&
        c.start < c.end
    )
    .slice()
    .sort((a, b) => a.start - b.start || b.end - a.end);
  const out: BlockContentControl[] = [];
  const stack: BlockContentControl[] = [];
  for (const control of sorted) {
    while (stack.length > 0 && stack[stack.length - 1]!.end <= control.start) {
      stack.pop();
    }
    const parent = stack[stack.length - 1];
    if (parent && control.end > parent.end) {
      continue;
    }
    out.push(control);
    stack.push(control);
  }
  return out;
}

/** Render a table cell. */
function renderTableCell(xml: XmlSink, cell: TableCell, helpers?: RenderHelpers): void {
  xml.openNode("w:tc");

  // OOXML CT_Tc requires a <w:tcPr> with at least <w:tcW> for Word to
  // accept the cell. If the model didn't supply any cell properties we
  // synthesise a minimal <w:tcPr><w:tcW w:type="auto"/></w:tcPr>; if it
  // did, renderTableCellProperties handles the missing-tcW fallback.
  if (cell.properties) {
    const properties = cell.properties;
    within("properties", () => renderTableCellProperties(xml, properties));
  } else {
    xml.openNode("w:tcPr");
    xml.leafNode("w:tcW", { "w:w": "0", "w:type": "auto" });
    xml.closeNode();
  }

  // OOXML §17.4.66 (CT_Tc) requires every table cell to *end* with a
  // <w:p>. An empty cell needs at least one <w:p>, and a cell whose last
  // block is a nested <w:tbl> must have a trailing <w:p> after it — Word
  // (and LibreOffice) reject the document otherwise.
  if (writableList(cell.content, "content").length === 0) {
    xml.openNode("w:p");
    xml.closeNode();
  } else {
    const controls = cellControls(cell);
    const open: number[] = [];
    const closeUpTo = (at: number): void => {
      while (open.length > 0 && open[open.length - 1]! <= at) {
        open.pop();
        xml.closeNode(); // w:sdtContent
        xml.closeNode(); // w:sdt
      }
    };
    let next = 0;
    cell.content.forEach((block, i) => {
      closeUpTo(i);
      while (next < controls.length && controls[next]!.start === i) {
        const control = controls[next++]!;
        xml.openNode("w:sdt");
        renderSdtProperties(xml, control.properties);
        xml.openNode("w:sdtContent");
        open.push(control.end);
      }
      try {
        if (block?.type === "paragraph") {
          renderParagraph(xml, block, helpers);
        } else if (block?.type === "table") {
          renderTable(xml, block, helpers);
        } else {
          unwritableNode(
            "type",
            "a paragraph or a table",
            (block as { type?: unknown } | null)?.type
          );
        }
      } catch (error) {
        throw atNode(`content[${i}]`, error);
      }
    });
    closeUpTo(cell.content.length);
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
    const ex = row.properties.tblPrEx;
    within("properties.tblPrEx", () => renderTablePropertiesEx(xml, ex));
  }

  if (row.properties) {
    const properties = row.properties;
    within("properties", () => renderTableRowProperties(xml, properties));
  }

  for (let c = 0; c < row.cells.length; c++) {
    const cell = row.cells[c]!;
    try {
      if (typeof cell !== "object" || cell === null) {
        unwritableNode("", "a table cell", cell);
      }
      renderRangeMarkers(xml, cell.rangeMarkersBefore, helpers);
      renderTableCell(xml, cell, helpers);
    } catch (error) {
      throw atNode(`cells[${c}]`, error);
    }
  }
  renderRangeMarkers(xml, row.rangeMarkersAfter, helpers);

  xml.closeNode();
}

/**
 * Render w:tblPrEx, the table properties a row overrides (`CT_TblPrEx`) — or, `previous`, the ones a
 * `w:tblPrExChange` holds (`CT_TblPrExBase`), which have no change record of their own.
 */
function renderTablePropertiesEx(
  xml: XmlSink,
  ex: TablePropertyExceptions,
  previous = false
): void {
  refuseTableOnlyProperties(ex);
  if (previous) {
    refuseAbsentFields(ex, ["propertyChange"], NARROWER_PROPERTIES.previousExceptions);
  }
  const { propertyChange } = ex;
  xml.openNode("w:tblPrEx");
  emitInSchemaOrder(previous ? "w:CT_TblPrExBase" : "w:CT_TblPrEx", [
    ...tablePropertiesExceptionChildren(
      xml,
      ex,
      writtenToken(ex, "alignment", WRITTEN.tableAlignment, "alignment"),
      writtenToken(ex, "layout", WRITTEN.tableLayout, "layout")
    ),
    !previous &&
      !!propertyChange && [
        "w:tblPrExChange",
        () => {
          xml.openNode("w:tblPrExChange", revisionAttributes(propertyChange.revision));
          if (propertyChange.previousProperties) {
            const before = propertyChange.previousProperties;
            within("propertyChange.previousProperties", () =>
              renderTablePropertiesEx(xml, before, true)
            );
          } else {
            xml.openNode("w:tblPrEx");
            xml.closeNode();
          }
          xml.closeNode();
        }
      ]
  ]);
  xml.closeNode();
}

/** Render a w:tbl element. */
export function renderTable(xml: XmlSink, table: Table, helpers?: RenderHelpers): void {
  const rows = writableList(table.rows, "rows");
  rows.forEach((row, r) => {
    if (typeof row !== "object" || row === null) {
      unwritableNode(`rows[${r}]`, "a table row", row);
    }
    writableList(row.cells, `rows[${r}].cells`);
  });
  renderRangeMarkers(xml, table.rangeMarkersBefore, helpers);
  xml.openNode("w:tbl");

  // `CT_Tbl` requires `w:tblPr`, even with nothing in it: a table without properties was written without the element,
  // which the schema refuses.
  const properties = table.properties ?? {};
  within("properties", () => renderTableProperties(xml, properties));

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
      xml.leafNode("w:gridCol", { "w:w": xmlValue(w, "w:w") });
    }
  } else {
    const columnCount = Math.max(1, computeTableColumnCount(table));
    const totalWidth = inferTableWidthInTwips(table.properties);
    const colWidth = Math.max(1, Math.floor(totalWidth / columnCount));
    for (let i = 0; i < columnCount; i++) {
      xml.leafNode("w:gridCol", { "w:w": xmlValue(colWidth, "w:w") });
    }
  }
  xml.closeNode();

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r]!;
    try {
      renderRangeMarkers(xml, row.rangeMarkersBefore, helpers);
      renderTableRow(xml, row, helpers);
    } catch (error) {
      throw atNode(`rows[${r}]`, error);
    }
  }
  renderRangeMarkers(xml, table.rangeMarkersAfter, helpers);

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
