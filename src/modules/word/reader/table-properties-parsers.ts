/**
 * DOCX Reader - Table Properties Parsers
 *
 * Pure parsers for `w:tblPr`, `w:tblBorders`, `w:tcMar` (cell margins).
 * Used by both the inline table parser and the styles parser.
 */

import { setAsideValue, unrecognizedEntry } from "@utils/unrecognized-values";
import { WRITTEN, readToken } from "@word/core/argument-checks";
import { EDGE_NAMES } from "@word/core/change-records";
import type { Mutable } from "@word/core/internal-utils";
import {
  attrInt,
  attrVal,
  findChildNs,
  localName,
  safeParseInt,
  serializeElement
} from "@word/reader/parse-utils";
import {
  parseBorder,
  parseRevisionInfo,
  parseRowHeight,
  parseShading,
  parseTableWidth
} from "@word/reader/properties-parsers";
import type {
  TableBorders,
  TableCellMargins,
  TableCellProperties,
  TableFloat,
  TableLook,
  TableProperties,
  TablePropertyExceptions,
  TableRowProperties
} from "@word/types";
import type { XmlElement } from "@xml/types";

/**
 * A `w:tblBorders` (`container: "table"`) or `w:tcBorders`.
 *
 * Only a cell has diagonals: `CT_TblBorders` has no `tl2br`/`tr2bl`. One found in a table's borders is set aside
 * rather than typed, as not written back (`writtenBack: false`) — `Document.listUnrecognizedValues` lists it, and
 * the document check reports it — because the
 * writer has nowhere valid to put it back and refuses a table diagonal a caller sets. Typing it made a file this
 * library wrote before that rule impossible to save.
 */
function parseTableBorders(el: XmlElement, container: "table" | "cell" = "cell"): TableBorders {
  const borders: Mutable<TableBorders> = {};
  for (const side of [
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
  ] as const) {
    const sideEl = findChildNs(el, side);
    if (!sideEl) {
      continue;
    }
    if (container === "table" && (side === "tl2br" || side === "tr2bl")) {
      setAsideValue(borders, side, serializeElement(sideEl), { writtenBack: false });
      continue;
    }
    borders[side] = parseBorder(sideEl);
  }
  return keepFirstEdgeName(borders, el);
}

function parseTableCellMargins(el: XmlElement): TableCellMargins {
  const margins: Mutable<TableCellMargins> = {};
  for (const side of ["top", "left", "bottom", "right", "start", "end"] as const) {
    const sideEl = findChildNs(el, side);
    if (sideEl) {
      margins[side] = parseTableWidth(sideEl);
    }
  }
  return keepFirstEdgeName(margins, el);
}

/**
 * `set`, read from `el`, with one name per edge (`EDGE_NAMES`): of an edge a file names twice, the name it gives
 * first is kept and the other set aside as not written back — the writer refuses both, and Office takes one.
 */
function keepFirstEdgeName<T extends TableBorders | TableCellMargins>(set: T, el: XmlElement): T {
  for (const pair of EDGE_NAMES) {
    if (set[pair[0]] === undefined || set[pair[1]] === undefined) {
      continue;
    }
    const first = el.children.find(
      child =>
        child.type === "element" && (pair as readonly string[]).includes(localName(child.name))
    );
    const second =
      first?.type === "element" && localName(first.name) === pair[0] ? pair[1] : pair[0];
    const secondEl = findChildNs(el, second)!;
    delete (set as Record<string, unknown>)[second];
    setAsideValue(set, second, serializeElement(secondEl), { writtenBack: false });
  }
  return set;
}

/**
 * A `w:tblPr`. In a style, and as a change record's previous properties, it is `CT_TblPrBase` (`base`), which has
 * no `w:tblPrChange`: one found there is set aside as not written back, as a diagonal in a table's borders is — the
 * writer has nowhere valid to put it, and refuses one a caller sets.
 */
function parseTableProperties(el: XmlElement, { base = false } = {}): TableProperties {
  const props: Mutable<TableProperties> = {};

  const styleEl = findChildNs(el, "tblStyle");
  if (styleEl) {
    props.style = attrVal(styleEl, "val");
  }

  const wEl = findChildNs(el, "tblW");
  if (wEl) {
    props.width = parseTableWidth(wEl);
  }

  const jcEl = findChildNs(el, "jc");
  if (jcEl) {
    props.alignment = readToken(
      attrVal(jcEl, "val"),
      WRITTEN.tableAlignment,
      props,
      "alignment"
    ) as TableProperties["alignment"];
  }

  const indEl = findChildNs(el, "tblInd");
  if (indEl) {
    props.indent = safeParseInt(indEl.attributes["w:w"] ?? indEl.attributes["w"], 0);
  }

  const bordersEl = findChildNs(el, "tblBorders");
  if (bordersEl) {
    props.borders = parseTableBorders(bordersEl, "table");
  }

  const layoutEl = findChildNs(el, "tblLayout");
  if (layoutEl) {
    props.layout = readToken(
      attrVal(layoutEl, "type"),
      WRITTEN.tableLayout,
      props,
      "layout"
    ) as TableProperties["layout"];
  }

  const cellMarEl = findChildNs(el, "tblCellMar");
  if (cellMarEl) {
    props.cellMargins = parseTableCellMargins(cellMarEl);
  }

  // TableLook
  const lookEl = findChildNs(el, "tblLook");
  if (lookEl) {
    const look: Mutable<TableLook> = {};

    // Read individual attributes first (authoritative when explicit "0"/"1")
    const readFlag = (name: string): boolean | undefined => {
      const v = attrVal(lookEl, name);
      if (v === "1" || v === "true") {
        return true;
      }
      if (v === "0" || v === "false") {
        return false;
      }
      return undefined;
    };

    const firstRow = readFlag("firstRow");
    const lastRow = readFlag("lastRow");
    const firstColumn = readFlag("firstColumn");
    const lastColumn = readFlag("lastColumn");
    const noHBand = readFlag("noHBand");
    const noVBand = readFlag("noVBand");

    if (firstRow !== undefined) {
      look.firstRow = firstRow;
    }
    if (lastRow !== undefined) {
      look.lastRow = lastRow;
    }
    if (firstColumn !== undefined) {
      look.firstColumn = firstColumn;
    }
    if (lastColumn !== undefined) {
      look.lastColumn = lastColumn;
    }
    if (noHBand !== undefined) {
      look.noHBand = noHBand;
    }
    if (noVBand !== undefined) {
      look.noVBand = noVBand;
    }

    // Fall back to w:val bitmask ONLY if no individual attrs were specified
    if (Object.keys(look).length === 0) {
      const val = attrVal(lookEl, "val");
      if (val) {
        const v = parseInt(val, 16);
        if (v & 0x0020) {
          look.firstRow = true;
        }
        if (v & 0x0040) {
          look.lastRow = true;
        }
        if (v & 0x0080) {
          look.firstColumn = true;
        }
        if (v & 0x0100) {
          look.lastColumn = true;
        }
        if (v & 0x0200) {
          look.noHBand = true;
        }
        if (v & 0x0400) {
          look.noVBand = true;
        }
      }
    }

    if (Object.keys(look).length > 0) {
      props.look = look;
    }
  }

  // TableFloat
  const tblpPrEl = findChildNs(el, "tblpPr");
  if (tblpPrEl) {
    const tf: Mutable<TableFloat> = {};
    const hAnchor = attrVal(tblpPrEl, "horzAnchor");
    if (hAnchor) {
      tf.horizontalAnchor = hAnchor as TableFloat["horizontalAnchor"];
    }
    const vAnchor = attrVal(tblpPrEl, "vertAnchor");
    if (vAnchor) {
      tf.verticalAnchor = vAnchor as TableFloat["verticalAnchor"];
    }
    const tblpX = attrInt(tblpPrEl, "tblpX");
    if (tblpX !== undefined) {
      tf.absoluteHorizontalPosition = tblpX;
    }
    const tblpY = attrInt(tblpPrEl, "tblpY");
    if (tblpY !== undefined) {
      tf.absoluteVerticalPosition = tblpY;
    }
    const tblpXSpec = attrVal(tblpPrEl, "tblpXSpec");
    if (tblpXSpec) {
      tf.relativeHorizontalPosition = tblpXSpec as TableFloat["relativeHorizontalPosition"];
    }
    const tblpYSpec = attrVal(tblpPrEl, "tblpYSpec");
    if (tblpYSpec) {
      tf.relativeVerticalPosition = tblpYSpec as TableFloat["relativeVerticalPosition"];
    }
    const topFromText = attrInt(tblpPrEl, "topFromText");
    if (topFromText !== undefined) {
      tf.topFromText = topFromText;
    }
    const bottomFromText = attrInt(tblpPrEl, "bottomFromText");
    if (bottomFromText !== undefined) {
      tf.bottomFromText = bottomFromText;
    }
    const leftFromText = attrInt(tblpPrEl, "leftFromText");
    if (leftFromText !== undefined) {
      tf.leftFromText = leftFromText;
    }
    const rightFromText = attrInt(tblpPrEl, "rightFromText");
    if (rightFromText !== undefined) {
      tf.rightFromText = rightFromText;
    }
    const overlap = attrVal(tblpPrEl, "overlap");
    if (overlap) {
      tf.overlap = overlap as TableFloat["overlap"];
    }
    props.float = tf;
  }

  // w:tblOverlap is a separate sibling element of w:tblpPr (value "never"|"overlap")
  const tblOverlapEl = findChildNs(el, "tblOverlap");
  if (tblOverlapEl && props.float) {
    const v = attrVal(tblOverlapEl, "val");
    if (v === "never" || v === "overlap") {
      (props.float as Mutable<TableFloat>).overlap = v;
    }
  }

  // Cell spacing
  const csEl = findChildNs(el, "tblCellSpacing");
  if (csEl) {
    props.cellSpacing = parseTableWidth(csEl);
  }

  // Bidi visual
  if (findChildNs(el, "bidiVisual")) {
    props.visuallyRightToLeft = true;
  }

  // Shading
  const shdEl = findChildNs(el, "shd");
  if (shdEl) {
    props.shading = parseShading(shdEl);
  }

  // Accessibility: caption and description
  const captionEl = findChildNs(el, "tblCaption");
  if (captionEl) {
    props.caption = attrVal(captionEl, "val");
  }
  const descEl = findChildNs(el, "tblDescription");
  if (descEl) {
    props.description = attrVal(descEl, "val");
  }

  // Table property change
  const tblPrChangeEl = findChildNs(el, "tblPrChange");
  if (tblPrChangeEl && base) {
    setAsideValue(props, "propertyChange", serializeElement(tblPrChangeEl), { writtenBack: false });
  } else if (tblPrChangeEl) {
    const rev = parseRevisionInfo(tblPrChangeEl);
    if (rev) {
      const prev = findChildNs(tblPrChangeEl, "tblPr");
      props.propertyChange = {
        revision: rev,
        previousProperties: prev ? parseTableProperties(prev, { base: true }) : undefined
      };
    }
  }

  return props;
}

/**
 * A `w:tcPr` — a cell's, a table style's or a conditional format's. One reading for all three, so a style's cell
 * properties keep what a cell's do: a conditional format used to read only its borders and shading and lost the rest
 * on the way back. Inside `w:tcPrChange` (`withChange: false`) the schema has no change record (`CT_TcPrInner`), so
 * one found there is set aside as not written back — which is also what bounds the recursion; the previous properties
 * were dropped before, for fear of an unbounded one.
 */
function parseTableCellProperties(
  tcPrEl: XmlElement,
  { withChange = true } = {}
): TableCellProperties {
  const p: Mutable<TableCellProperties> = {};
  const wEl = findChildNs(tcPrEl, "tcW");
  if (wEl) {
    p.width = parseTableWidth(wEl);
  }
  const gsEl = findChildNs(tcPrEl, "gridSpan");
  if (gsEl) {
    p.gridSpan = attrInt(gsEl, "val");
  }
  const vmEl = findChildNs(tcPrEl, "vMerge");
  if (vmEl) {
    p.verticalMerge = (attrVal(vmEl, "val") ?? "continue") as TableCellProperties["verticalMerge"];
  }
  const bordersEl = findChildNs(tcPrEl, "tcBorders");
  if (bordersEl) {
    p.borders = parseTableBorders(bordersEl);
  }
  const shdEl = findChildNs(tcPrEl, "shd");
  if (shdEl) {
    p.shading = parseShading(shdEl);
  }
  const vAlignEl = findChildNs(tcPrEl, "vAlign");
  if (vAlignEl) {
    p.verticalAlign = readToken(
      attrVal(vAlignEl, "val"),
      WRITTEN.cellVerticalAlign,
      p,
      "verticalAlign"
    ) as TableCellProperties["verticalAlign"];
  }
  if (findChildNs(tcPrEl, "noWrap")) {
    p.noWrap = true;
  }
  const textDirEl = findChildNs(tcPrEl, "textDirection");
  if (textDirEl) {
    p.textDirection = readToken(
      attrVal(textDirEl, "val"),
      WRITTEN.textDirection,
      p,
      "textDirection"
    ) as TableCellProperties["textDirection"];
  }
  const marginsEl = findChildNs(tcPrEl, "tcMar");
  if (marginsEl) {
    p.margins = parseTableCellMargins(marginsEl);
  }
  const cnfEl = findChildNs(tcPrEl, "cnfStyle");
  if (cnfEl) {
    p.cnfStyle = attrVal(cnfEl, "val");
  }
  if (findChildNs(tcPrEl, "hideMark")) {
    p.hideMark = true;
  }
  if (findChildNs(tcPrEl, "tcFitText")) {
    p.fitText = true;
  }
  const cellInsEl = findChildNs(tcPrEl, "cellIns");
  if (cellInsEl) {
    const rev = parseRevisionInfo(cellInsEl);
    if (rev) {
      p.inserted = { revision: rev };
    }
  }
  const cellDelEl = findChildNs(tcPrEl, "cellDel");
  if (cellDelEl) {
    const rev = parseRevisionInfo(cellDelEl);
    if (rev) {
      p.deleted = { revision: rev };
    }
  }
  const cellMergeEl = findChildNs(tcPrEl, "cellMerge");
  if (cellMergeEl) {
    const vMerge = attrVal(cellMergeEl, "vMerge");
    const rev = parseRevisionInfo(cellMergeEl);
    if (rev && (vMerge === "cont" || vMerge === "rest")) {
      p.cellMerge = { vMerge, revision: rev };
    }
  }
  const tcPrChangeEl = findChildNs(tcPrEl, "tcPrChange");
  if (tcPrChangeEl && !withChange) {
    setAsideValue(p, "propertyChange", serializeElement(tcPrChangeEl), { writtenBack: false });
  } else if (tcPrChangeEl) {
    const rev = parseRevisionInfo(tcPrChangeEl);
    if (rev) {
      const prev = findChildNs(tcPrChangeEl, "tcPr");
      p.propertyChange = {
        revision: rev,
        ...(prev
          ? { previousProperties: parseTableCellProperties(prev, { withChange: false }) }
          : {})
      };
    }
  }
  return p;
}

/**
 * A `w:trPr` — a row's, a table style's or a conditional format's. A conditional format used to read only its
 * height, and the previous properties of a row's change only three fields. Inside `w:trPrChange` (`withChange:
 * false`, `CT_TrPrBase`) a change record, an insertion and a deletion are set aside as not written back.
 */
function parseTableRowProperties(
  trPrEl: XmlElement,
  { withChange = true } = {}
): TableRowProperties {
  const p: Mutable<TableRowProperties> = {};
  const heightEl = findChildNs(trPrEl, "trHeight");
  if (heightEl) {
    p.height = parseRowHeight(heightEl);
  }
  if (findChildNs(trPrEl, "tblHeader")) {
    p.tableHeader = true;
  }
  if (findChildNs(trPrEl, "cantSplit")) {
    p.cantSplit = true;
  }
  if (findChildNs(trPrEl, "hidden")) {
    p.hidden = true;
  }
  const csEl = findChildNs(trPrEl, "tblCellSpacing");
  if (csEl) {
    p.cellSpacing = parseTableWidth(csEl);
  }
  // A row's insertion and deletion are `CT_TrPr`'s; its previous properties (`CT_TrPrBase`) have neither.
  for (const [name, field] of [
    ["ins", "inserted"],
    ["del", "deleted"]
  ] as const) {
    const el = findChildNs(trPrEl, name);
    const rev = el && withChange ? parseRevisionInfo(el) : undefined;
    if (el && !withChange) {
      setAsideValue(p, field, serializeElement(el), { writtenBack: false });
    } else if (rev) {
      p[field] = { revision: rev };
    }
  }
  const gbEl = findChildNs(trPrEl, "gridBefore");
  if (gbEl) {
    p.gridBefore = attrInt(gbEl, "val");
  }
  const gaEl = findChildNs(trPrEl, "gridAfter");
  if (gaEl) {
    p.gridAfter = attrInt(gaEl, "val");
  }
  const wbEl = findChildNs(trPrEl, "wBefore");
  if (wbEl) {
    p.widthBefore = parseTableWidth(wbEl);
  }
  const waEl = findChildNs(trPrEl, "wAfter");
  if (waEl) {
    p.widthAfter = parseTableWidth(waEl);
  }
  const cnfEl = findChildNs(trPrEl, "cnfStyle");
  if (cnfEl) {
    p.cnfStyle = attrVal(cnfEl, "val");
  }
  const trPrChangeEl = findChildNs(trPrEl, "trPrChange");
  if (trPrChangeEl && !withChange) {
    setAsideValue(p, "propertyChange", serializeElement(trPrChangeEl), { writtenBack: false });
  } else if (trPrChangeEl) {
    const rev = parseRevisionInfo(trPrChangeEl);
    if (rev) {
      const prevTrPr = findChildNs(trPrChangeEl, "trPr");
      p.propertyChange = {
        revision: rev,
        previousProperties: prevTrPr
          ? parseTableRowProperties(prevTrPr, { withChange: false })
          : undefined
      };
    }
  }
  return p;
}

/** What `w:tblPrEx` holds of the table properties (`CT_TblPrExBase`). */
const EXCEPTION_FIELDS = [
  "width",
  "alignment",
  "cellSpacing",
  "indent",
  "borders",
  "shading",
  "layout",
  "cellMargins",
  "look"
] as const satisfies readonly (keyof TablePropertyExceptions & keyof TableProperties)[];

/**
 * The children of `w:tblPr` that `w:tblPrEx` does not have: properties of the whole table, which a row cannot
 * override. Pinned against the schema by a test.
 */
export const TABLE_ONLY_ELEMENTS = [
  "tblStyle",
  "tblpPr",
  "tblOverlap",
  "bidiVisual",
  "tblStyleRowBandSize",
  "tblStyleColBandSize",
  "tblCaption",
  "tblDescription",
  "tblPrChange"
] as const;

/**
 * A row's `w:tblPrEx`. Its children are read as a `w:tblPr`'s are, and what is kept is what the schema lets a row
 * override — with any value set aside on them. A whole-table property found here (`w:tblStyle`, `w:tblpPr`, …) is
 * set aside as not written back, under its element's name: the model used to take a table's properties here, so one
 * was typed on read and then dropped on write without a word. Inside `w:tblPrExChange` (`withChange: false`,
 * `CT_TblPrExBase`) a change record is set aside the same way.
 */
function parseTablePropertyExceptions(
  el: XmlElement,
  { withChange = true } = {}
): TablePropertyExceptions {
  const table = parseTableProperties(el, { base: true });
  const exceptions: Mutable<TablePropertyExceptions> = {};
  for (const field of EXCEPTION_FIELDS) {
    if (table[field] !== undefined) {
      (exceptions as Record<string, unknown>)[field] = table[field];
    }
    const kept = unrecognizedEntry(table, field);
    if (kept !== undefined) {
      const { value, ...options } = kept;
      setAsideValue(exceptions, field, value, options);
    }
  }
  for (const name of TABLE_ONLY_ELEMENTS) {
    const child = findChildNs(el, name);
    if (child) {
      setAsideValue(exceptions, name, serializeElement(child), { writtenBack: false });
    }
  }
  const changeEl = findChildNs(el, "tblPrExChange");
  if (changeEl && !withChange) {
    setAsideValue(exceptions, "propertyChange", serializeElement(changeEl), { writtenBack: false });
  } else if (changeEl) {
    const revision = parseRevisionInfo(changeEl);
    if (revision) {
      const previous = findChildNs(changeEl, "tblPrEx");
      exceptions.propertyChange = {
        revision,
        ...(previous
          ? { previousProperties: parseTablePropertyExceptions(previous, { withChange: false }) }
          : {})
      };
    }
  }
  return exceptions;
}

export {
  parseTableBorders,
  parseTableCellMargins,
  parseTableCellProperties,
  parseTableProperties,
  parseTablePropertyExceptions,
  parseTableRowProperties
};
