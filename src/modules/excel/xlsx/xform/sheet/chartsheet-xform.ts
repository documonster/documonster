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
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
    header?: number;
    footer?: number;
  };
  /** Drawing relationship reference */
  drawing?: { rId: string };
  /** Relationships parsed from the chartsheet .rels file */
  relationships?: any[];
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

    // pageMargins
    if (m.pageMargins) {
      const pm = m.pageMargins;
      xmlStream.leafNode("pageMargins", {
        left: pm.left !== undefined ? String(pm.left) : "0.7",
        right: pm.right !== undefined ? String(pm.right) : "0.7",
        top: pm.top !== undefined ? String(pm.top) : "0.75",
        bottom: pm.bottom !== undefined ? String(pm.bottom) : "0.75",
        header: pm.header !== undefined ? String(pm.header) : "0.3",
        footer: pm.footer !== undefined ? String(pm.footer) : "0.3"
      });
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
            left: attrs.left !== undefined ? parseFloat(attrs.left) : undefined,
            right: attrs.right !== undefined ? parseFloat(attrs.right) : undefined,
            top: attrs.top !== undefined ? parseFloat(attrs.top) : undefined,
            bottom: attrs.bottom !== undefined ? parseFloat(attrs.bottom) : undefined,
            header: attrs.header !== undefined ? parseFloat(attrs.header) : undefined,
            footer: attrs.footer !== undefined ? parseFloat(attrs.footer) : undefined
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

export { ChartsheetXform };
