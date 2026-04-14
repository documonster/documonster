import type { WorksheetViewFrozen, WorksheetViewSplit } from "@excel/types";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

describe("Worksheet", () => {
  describe("Views", () => {
    // =========================================================================
    // Column Outline / Collapsed
    // =========================================================================

    it("adjusts collapsed property of columns based on outlineLevel", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet1");

      const col1 = ws.getColumn(1);
      const col2 = ws.getColumn(2);
      const col3 = ws.getColumn(3);
      expect(col1.collapsed).toBe(false);
      expect(col2.collapsed).toBe(false);
      expect(col3.collapsed).toBe(false);

      col1.outlineLevel = 0;
      col2.outlineLevel = 1;
      col3.outlineLevel = 2;
      expect(col1.collapsed).toBe(false);
      expect(col2.collapsed).toBe(true);
      expect(col3.collapsed).toBe(true);

      ws.properties.outlineLevelCol = 2;
      expect(col1.collapsed).toBe(false);
      expect(col2.collapsed).toBe(false);
      expect(col3.collapsed).toBe(true);

      ws.properties.outlineLevelCol = 3;
      expect(col1.collapsed).toBe(false);
      expect(col2.collapsed).toBe(false);
      expect(col3.collapsed).toBe(false);
    });

    // =========================================================================
    // Row Outline / Collapsed
    // =========================================================================

    it("adjusts collapsed property of rows based on outlineLevel", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet1");

      const row1 = ws.getRow(1);
      const row2 = ws.getRow(2);
      const row3 = ws.getRow(3);
      expect(row1.collapsed).toBe(false);
      expect(row2.collapsed).toBe(false);
      expect(row3.collapsed).toBe(false);

      row1.outlineLevel = 0;
      row2.outlineLevel = 1;
      row3.outlineLevel = 2;
      expect(row1.collapsed).toBe(false);
      expect(row2.collapsed).toBe(true);
      expect(row3.collapsed).toBe(true);

      ws.properties.outlineLevelRow = 2;
      expect(row1.collapsed).toBe(false);
      expect(row2.collapsed).toBe(false);
      expect(row3.collapsed).toBe(true);

      ws.properties.outlineLevelRow = 3;
      expect(row1.collapsed).toBe(false);
      expect(row2.collapsed).toBe(false);
      expect(row3.collapsed).toBe(false);
    });

    it("sets outline levels via column headers", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("sheet1");

      ws.columns = [
        { key: "id", width: 10, outlineLevel: 1 },
        { key: "name", width: 32, outlineLevel: 2 },
        { key: "dob", width: 10, outlineLevel: 3 }
      ];

      expect(ws.getColumn(1).outlineLevel).toBe(1);
      expect(ws.getColumn(2).outlineLevel).toBe(2);
      expect(ws.getColumn(3).outlineLevel).toBe(3);
    });

    // =========================================================================
    // Frozen Panes
    // =========================================================================

    it("sets frozen view (split at a specific cell)", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("frozen");

      ws.views = [{ state: "frozen", xSplit: 1, ySplit: 2 }];

      expect(ws.views.length).toBe(1);
      const view = ws.views[0] as Partial<WorksheetViewFrozen>;
      expect(view.state).toBe("frozen");
      expect(view.xSplit).toBe(1);
      expect(view.ySplit).toBe(2);
    });

    it("frozen view survives XLSX round-trip", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("frozen");
      ws.views = [{ state: "frozen", xSplit: 0, ySplit: 1, activeCell: "A2" }];
      ws.getCell("A1").value = "header";

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      const ws2 = wb2.getWorksheet("frozen")!;
      expect(ws2.views.length).toBe(1);
      const view = ws2.views[0] as Partial<WorksheetViewFrozen>;
      expect(view.state).toBe("frozen");
      expect(view.ySplit).toBe(1);
    });

    // =========================================================================
    // Split Panes
    // =========================================================================

    it("sets split view", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("split");

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
      const wb = new Workbook();
      const ws = wb.addWorksheet("options");

      ws.views = [{ showGridLines: false, showRowColHeaders: false }];

      expect(ws.views[0].showGridLines).toBe(false);
      expect(ws.views[0].showRowColHeaders).toBe(false);
    });

    it("sets zoom factor", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("zoom");

      ws.views = [{ zoomScale: 150, zoomScaleNormal: 150 }];

      expect(ws.views[0].zoomScale).toBe(150);
      expect(ws.views[0].zoomScaleNormal).toBe(150);
    });

    it("sets right-to-left view", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("rtl");

      ws.views = [{ rightToLeft: true }];

      expect(ws.views[0].rightToLeft).toBe(true);
    });

    // =========================================================================
    // Worksheet State
    // =========================================================================

    it("default state is visible", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      expect(ws.state).toBe("visible");
    });

    it("state can be set to hidden", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test", { state: "hidden" });

      expect(ws.state).toBe("hidden");
    });

    it("state can be set to veryHidden", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test", { state: "veryHidden" });

      expect(ws.state).toBe("veryHidden");
    });

    it("hidden state survives XLSX round-trip", async () => {
      const wb = new Workbook();
      wb.addWorksheet("visible");
      wb.addWorksheet("hidden", { state: "hidden" });

      const buffer = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(buffer);

      expect(wb2.getWorksheet("visible")!.state).toBe("visible");
      expect(wb2.getWorksheet("hidden")!.state).toBe("hidden");
    });

    // =========================================================================
    // Page Setup
    // =========================================================================

    it("page setup properties can be set", () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

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
      const wb = new Workbook();
      const ws = wb.addWorksheet("test", {
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
      const wb = new Workbook();
      const ws = wb.addWorksheet("test");

      ws.headerFooter.oddHeader = "&CPage &P of &N";
      ws.headerFooter.oddFooter = "&LConfidential&RDate: &D";

      expect(ws.headerFooter.oddHeader).toBe("&CPage &P of &N");
      expect(ws.headerFooter.oddFooter).toBe("&LConfidential&RDate: &D");
    });
  });
});
