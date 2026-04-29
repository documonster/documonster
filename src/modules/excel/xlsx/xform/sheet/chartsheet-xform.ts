/**
 * ChartsheetXform — parse and render OOXML chartsheet XML.
 *
 * A chartsheet is a simple sheet type that contains only a chart.
 * Structure:
 * ```xml
 * <chartsheet xmlns="...">
 *   <sheetPr/>
 *   <sheetViews><sheetView .../></sheetViews>
 *   <pageMargins .../>
 *   <drawing r:id="rId1"/>
 * </chartsheet>
 * ```
 */

import { BaseXform } from "@excel/xlsx/xform/base-xform";
import type { XmlSink } from "@xml/types";

export interface ChartsheetModel {
  /** Sheet number (positional index in the XLSX archive) */
  sheetNo: number;
  /** Sheet name (from workbook.xml) */
  name: string;
  /** Sheet ID (from workbook.xml) */
  id: number;
  /** Relationship ID linking to this chartsheet from workbook.xml.rels */
  rId?: string;
  /** Sheet visibility state */
  state?: "visible" | "hidden" | "veryHidden";
  /** Tab selected */
  tabSelected?: boolean;
  /** Zoom scale */
  zoomScale?: number;
  /** Page margins */
  pageMargins?: {
    l?: number;
    r?: number;
    t?: number;
    b?: number;
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
  };
  /** Chartsheet print options. */
  printOptions?: {
    horizontalCentered?: boolean;
    verticalCentered?: boolean;
    headings?: boolean;
    gridLines?: boolean;
    gridLinesSet?: boolean;
  };
  /** Chartsheet page setup. */
  pageSetup?: {
    paperSize?: number;
    scale?: number;
    firstPageNumber?: number;
    fitToWidth?: number;
    fitToHeight?: number;
    pageOrder?: "downThenOver" | "overThenDown";
    orientation?: "default" | "portrait" | "landscape";
    usePrinterDefaults?: boolean;
    blackAndWhite?: boolean;
    draft?: boolean;
    cellComments?: "none" | "asDisplayed" | "atEnd";
    errors?: "displayed" | "blank" | "dash" | "NA";
    horizontalDpi?: number;
    verticalDpi?: number;
    copies?: number;
  };
  /** Drawing relationship reference */
  drawing?: { rId: string };
  /** Relationships parsed from the chartsheet .rels file */
  relationships?: any[];
  /** Drawing part name without extension (e.g. drawing2) */
  drawingName?: string;
  /** Classic chart number displayed by this chartsheet */
  chartNumber?: number;
  /** ChartEx number displayed by this chartsheet */
  chartExNumber?: number;
}

const CHARTSHEET_ATTRIBUTES = {
  xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
};

class ChartsheetXform extends BaseXform<ChartsheetModel> {
  private inSheetView = false;

  get tag(): string {
    return "chartsheet";
  }

  render(xmlStream: XmlSink, model?: ChartsheetModel): void {
    const m = model ?? this.model;
    if (!m) {
      return;
    }

    xmlStream.openNode("chartsheet", CHARTSHEET_ATTRIBUTES);

    // sheetViews
    xmlStream.openNode("sheetViews");
    const svAttrs: Record<string, string> = { workbookViewId: "0" };
    if (m.tabSelected) {
      svAttrs.tabSelected = "1";
    }
    if (m.zoomScale !== undefined) {
      svAttrs.zoomScale = String(m.zoomScale);
    }
    xmlStream.leafNode("sheetView", svAttrs);
    xmlStream.closeNode();

    // printOptions
    if (m.printOptions) {
      xmlStream.leafNode("printOptions", booleanAttrs(m.printOptions));
    }

    // pageMargins
    if (m.pageMargins) {
      const pm = m.pageMargins;
      xmlStream.leafNode("pageMargins", {
        left: pm.left !== undefined ? String(pm.left) : pm.l !== undefined ? String(pm.l) : "0.7",
        right:
          pm.right !== undefined ? String(pm.right) : pm.r !== undefined ? String(pm.r) : "0.7",
        top: pm.top !== undefined ? String(pm.top) : pm.t !== undefined ? String(pm.t) : "0.75",
        bottom:
          pm.bottom !== undefined ? String(pm.bottom) : pm.b !== undefined ? String(pm.b) : "0.75",
        header: pm.header !== undefined ? String(pm.header) : "0.3",
        footer: pm.footer !== undefined ? String(pm.footer) : "0.3"
      });
    }

    if (m.pageSetup) {
      xmlStream.leafNode("pageSetup", definedAttrs(m.pageSetup));
    }

    // drawing
    if (m.drawing) {
      xmlStream.leafNode("drawing", { "r:id": m.drawing.rId });
    }

    xmlStream.closeNode();
  }

  parseOpen(node: any): boolean {
    const { name } = node;
    const attrs = node.attributes || {};

    switch (name) {
      case "chartsheet":
        this.model = {
          sheetNo: 0,
          name: "",
          id: 0
        };
        break;
      case "sheetView":
        this.inSheetView = true;
        if (this.model) {
          if (attrs.tabSelected === "1") {
            this.model.tabSelected = true;
          }
          if (attrs.zoomScale) {
            this.model.zoomScale = parseInt(attrs.zoomScale, 10);
          }
        }
        break;
      case "pageMargins":
        if (this.model) {
          this.model.pageMargins = {
            l: attrs.left !== undefined ? parseFloat(attrs.left) : undefined,
            r: attrs.right !== undefined ? parseFloat(attrs.right) : undefined,
            t: attrs.top !== undefined ? parseFloat(attrs.top) : undefined,
            b: attrs.bottom !== undefined ? parseFloat(attrs.bottom) : undefined,
            left: attrs.left !== undefined ? parseFloat(attrs.left) : undefined,
            right: attrs.right !== undefined ? parseFloat(attrs.right) : undefined,
            top: attrs.top !== undefined ? parseFloat(attrs.top) : undefined,
            bottom: attrs.bottom !== undefined ? parseFloat(attrs.bottom) : undefined,
            header: attrs.header !== undefined ? parseFloat(attrs.header) : undefined,
            footer: attrs.footer !== undefined ? parseFloat(attrs.footer) : undefined
          };
        }
        break;
      case "printOptions":
        if (this.model) {
          this.model.printOptions = {
            horizontalCentered: parseBool(attrs.horizontalCentered),
            verticalCentered: parseBool(attrs.verticalCentered),
            headings: parseBool(attrs.headings),
            gridLines: parseBool(attrs.gridLines),
            gridLinesSet: parseBool(attrs.gridLinesSet)
          };
        }
        break;
      case "pageSetup":
        if (this.model) {
          this.model.pageSetup = {
            paperSize: parseNumber(attrs.paperSize),
            scale: parseNumber(attrs.scale),
            firstPageNumber: parseNumber(attrs.firstPageNumber),
            fitToWidth: parseNumber(attrs.fitToWidth),
            fitToHeight: parseNumber(attrs.fitToHeight),
            pageOrder: attrs.pageOrder,
            orientation: attrs.orientation,
            usePrinterDefaults: parseBool(attrs.usePrinterDefaults),
            blackAndWhite: parseBool(attrs.blackAndWhite),
            draft: parseBool(attrs.draft),
            cellComments: attrs.cellComments,
            errors: attrs.errors,
            horizontalDpi: parseNumber(attrs.horizontalDpi),
            verticalDpi: parseNumber(attrs.verticalDpi),
            copies: parseNumber(attrs.copies)
          };
        }
        break;
      case "drawing":
        if (this.model && attrs["r:id"]) {
          this.model.drawing = { rId: attrs["r:id"] };
        }
        break;
      default:
        break;
    }
    return true;
  }

  parseText(): void {
    // no text content in chartsheet elements
  }

  parseClose(name: string): boolean {
    switch (name) {
      case "chartsheet":
        return false;
      case "sheetView":
        this.inSheetView = false;
        return true;
      default:
        return true;
    }
  }
}

function booleanAttrs(model: Record<string, boolean | undefined>): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const [key, value] of Object.entries(model)) {
    if (value !== undefined) {
      attrs[key] = value ? "1" : "0";
    }
  }
  return attrs;
}

function definedAttrs(
  model: Record<string, string | number | boolean | undefined>
): Record<string, string> {
  const attrs: Record<string, string> = {};
  for (const [key, value] of Object.entries(model)) {
    if (value !== undefined) {
      attrs[key] = typeof value === "boolean" ? (value ? "1" : "0") : String(value);
    }
  }
  return attrs;
}

function parseBool(value: string | undefined): boolean | undefined {
  return value === undefined ? undefined : value === "1" || value === "true";
}

function parseNumber(value: string | undefined): number | undefined {
  return value === undefined ? undefined : Number(value);
}

export { ChartsheetXform };
