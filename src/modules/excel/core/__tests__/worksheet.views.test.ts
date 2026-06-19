import { columnCollapsed, columnOutlineLevel, columnSetOutlineLevel } from "@excel/core/column";
import { rowCollapsed, rowSetOutlineLevel } from "@excel/core/row";
import { getColumn } from "@excel/core/worksheet";
import { Cell, Workbook, Worksheet } from "@excel/index";
import type { WorksheetViewFrozen, WorksheetViewSplit } from "@excel/types";
import { describe, it, expect } from "vitest";

describe("Worksheet", () => {
  describe("Views", () => {
    // =========================================================================
    // Column Outline / Collapsed
    // =========================================================================

    it("adjusts collapsed property of columns based on outlineLevel", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet1");

      const col1 = getColumn(ws, 1);
      const col2 = getColumn(ws, 2);
      const col3 = getColumn(ws, 3);
      expect(columnCollapsed(col1)).toBe(false);
      expect(columnCollapsed(col2)).toBe(false);
      expect(columnCollapsed(col3)).toBe(false);

      columnSetOutlineLevel(col1, 0);
      columnSetOutlineLevel(col2, 1);
      columnSetOutlineLevel(col3, 2);
      expect(columnCollapsed(col1)).toBe(false);
      expect(columnCollapsed(col2)).toBe(true);
      expect(columnCollapsed(col3)).toBe(true);

      ws.properties.outlineLevelCol = 2;
      expect(columnCollapsed(col1)).toBe(false);
      expect(columnCollapsed(col2)).toBe(false);
      expect(columnCollapsed(col3)).toBe(true);

      ws.properties.outlineLevelCol = 3;
      expect(columnCollapsed(col1)).toBe(false);
      expect(columnCollapsed(col2)).toBe(false);
      expect(columnCollapsed(col3)).toBe(false);
    });

    // =========================================================================
    // Row Outline / Collapsed
    // =========================================================================

    it("adjusts collapsed property of rows based on outlineLevel", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet1");

      const row1 = Worksheet.getRow(ws, 1);
      const row2 = Worksheet.getRow(ws, 2);
      const row3 = Worksheet.getRow(ws, 3);
      expect(rowCollapsed(row1)).toBe(false);
      expect(rowCollapsed(row2)).toBe(false);
      expect(rowCollapsed(row3)).toBe(false);

      rowSetOutlineLevel(row1, 0);
      rowSetOutlineLevel(row2, 1);
      rowSetOutlineLevel(row3, 2);
      expect(rowCollapsed(row1)).toBe(false);
      expect(rowCollapsed(row2)).toBe(true);
      expect(rowCollapsed(row3)).toBe(true);

      ws.properties.outlineLevelRow = 2;
      expect(rowCollapsed(row1)).toBe(false);
      expect(rowCollapsed(row2)).toBe(false);
      expect(rowCollapsed(row3)).toBe(true);

      ws.properties.outlineLevelRow = 3;
      expect(rowCollapsed(row1)).toBe(false);
      expect(rowCollapsed(row2)).toBe(false);
      expect(rowCollapsed(row3)).toBe(false);
    });

    it("sets outline levels via column headers", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "sheet1");

      Worksheet.setColumns(ws, [
        { key: "id", width: 10, outlineLevel: 1 },
        { key: "name", width: 32, outlineLevel: 2 },
        { key: "dob", width: 10, outlineLevel: 3 }
      ]);

      expect(columnOutlineLevel(getColumn(ws, 1))).toBe(1);
      expect(columnOutlineLevel(getColumn(ws, 2))).toBe(2);
      expect(columnOutlineLevel(getColumn(ws, 3))).toBe(3);
    });

    // =========================================================================
    // Frozen Panes
    // =========================================================================

    it("sets frozen view (split at a specific cell)", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "frozen");

      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 2 }];

      expect(ws.views.length).toBe(1);
      const view = ws.views[0] as Partial<WorksheetViewFrozen>;
      expect(view.state).toBe("frozen");
      expect(view.xSplit).toBe(1);
      expect(view.ySplit).toBe(2);
    });

    it("frozen view survives XLSX round-trip", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "frozen");
      ws.views = [{ state: "frozen", xSplit: 0, ySplit: 1, activeCell: "A2" }];
      Cell.setValue(ws, "A1", "header");

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      const ws2 = Workbook.getWorksheet(wb2, "frozen")!;
      expect(ws2.views.length).toBe(1);
      const view = ws2.views[0] as Partial<WorksheetViewFrozen>;
      expect(view.state).toBe("frozen");
      expect(view.ySplit).toBe(1);
    });

    // =========================================================================
    // Split Panes
    // =========================================================================

    it("sets split view", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "split");

      ws.views = [{ state: "split", xSplit: 2000, ySplit: 3000 }];

      const view = ws.views[0] as Partial<WorksheetViewSplit>;
      expect(view.state).toBe("split");
      expect(view.xSplit).toBe(2000);
      expect(view.ySplit).toBe(3000);
    });

    // =========================================================================
    // View Options
    // =========================================================================

    it("sets showGridLines and showRowColHeaders", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "options");

      ws.views = [{ showGridLines: false, showRowColHeaders: false }];

      expect(ws.views[0].showGridLines).toBe(false);
      expect(ws.views[0].showRowColHeaders).toBe(false);
    });

    it("sets zoom factor", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "zoom");

      ws.views = [{ zoomScale: 150, zoomScaleNormal: 150 }];

      expect(ws.views[0].zoomScale).toBe(150);
      expect(ws.views[0].zoomScaleNormal).toBe(150);
    });

    it("sets right-to-left view", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "rtl");

      ws.views = [{ rightToLeft: true }];

      expect(ws.views[0].rightToLeft).toBe(true);
    });

    // =========================================================================
    // Worksheet State
    // =========================================================================

    it("default state is visible", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      expect(ws.state).toBe("visible");
    });

    it("state can be set to hidden", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test", { state: "hidden" });

      expect(ws.state).toBe("hidden");
    });

    it("state can be set to veryHidden", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test", { state: "veryHidden" });

      expect(ws.state).toBe("veryHidden");
    });

    it("hidden state survives XLSX round-trip", async () => {
      const wb = Workbook.create();
      Workbook.addWorksheet(wb, "visible");
      Workbook.addWorksheet(wb, "hidden", { state: "hidden" });

      const buffer = await Workbook.toBuffer(wb);
      const wb2 = Workbook.create();
      await Workbook.read(wb2, buffer);

      expect(Workbook.getWorksheet(wb2, "visible")!.state).toBe("visible");
      expect(Workbook.getWorksheet(wb2, "hidden")!.state).toBe("hidden");
    });

    // =========================================================================
    // Page Setup
    // =========================================================================

    it("page setup properties can be set", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      ws.pageSetup.orientation = "landscape";
      ws.pageSetup.paperSize = 9; // A4
      ws.pageSetup.margins = {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3
      };

      expect(ws.pageSetup.orientation).toBe("landscape");
      expect(ws.pageSetup.paperSize).toBe(9);
      expect(ws.pageSetup.margins!.left).toBe(0.7);
    });

    it("fitToPage is enabled when fitToWidth/fitToHeight set without explicit scale", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test", {
        pageSetup: { fitToWidth: 1, fitToHeight: 1 }
      });

      expect(ws.pageSetup.fitToWidth).toBe(1);
      expect(ws.pageSetup.fitToHeight).toBe(1);
      expect(ws.pageSetup.fitToPage).toBe(true);
    });

    // =========================================================================
    // Header / Footer
    // =========================================================================

    it("header and footer can be set", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "test");

      ws.headerFooter.oddHeader = "&CPage &P of &N";
      ws.headerFooter.oddFooter = "&LConfidential&RDate: &D";

      expect(ws.headerFooter.oddHeader).toBe("&CPage &P of &N");
      expect(ws.headerFooter.oddFooter).toBe("&LConfidential&RDate: &D");
    });
  });
});
