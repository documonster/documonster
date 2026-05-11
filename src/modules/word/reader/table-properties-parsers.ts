/**
 * DOCX Reader - Table Properties Parsers
 *
 * Pure parsers for `w:tblPr`, `w:tblBorders`, `w:tcMar` (cell margins).
 * Used by both the inline table parser and the styles parser.
 */

import type { XmlElement } from "@xml/types";

import { type Mutable } from "../core/internal-utils";
import type {
  TableBorders,
  TableCellMargins,
  TableFloat,
  TableLook,
  TableProperties
} from "../types";
import { attrInt, attrVal, findChildNs } from "./parse-utils";
import {
  parseBorder,
  parseRevisionInfo,
  parseShading,
  parseTableWidth
} from "./properties-parsers";

function parseTableBorders(el: XmlElement): TableBorders {
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
    if (sideEl) {
      borders[side] = parseBorder(sideEl);
    }
  }
  return borders;
}

function parseTableCellMargins(el: XmlElement): TableCellMargins {
  const margins: Mutable<TableCellMargins> = {};
  for (const side of ["top", "left", "bottom", "right", "start", "end"] as const) {
    const sideEl = findChildNs(el, side);
    if (sideEl) {
      margins[side] = parseTableWidth(sideEl);
    }
  }
  return margins;
}

function parseTableProperties(el: XmlElement): TableProperties {
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
    props.alignment = attrVal(jcEl, "val") as TableProperties["alignment"];
  }

  const indEl = findChildNs(el, "tblInd");
  if (indEl) {
    props.indent = parseInt(indEl.attributes["w:w"] ?? indEl.attributes["w"] ?? "0", 10);
  }

  const bordersEl = findChildNs(el, "tblBorders");
  if (bordersEl) {
    props.borders = parseTableBorders(bordersEl);
  }

  const layoutEl = findChildNs(el, "tblLayout");
  if (layoutEl) {
    props.layout = attrVal(layoutEl, "type") as TableProperties["layout"];
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
  if (tblPrChangeEl) {
    const rev = parseRevisionInfo(tblPrChangeEl);
    if (rev) {
      const prev = findChildNs(tblPrChangeEl, "tblPr");
      props.propertyChange = {
        revision: rev,
        previousProperties: prev ? parseTableProperties(prev) : undefined
      };
    }
  }

  return props;
}

export { parseTableBorders, parseTableCellMargins, parseTableProperties };
