import { extractAll } from "@archive/unzip/extract";
import { expectValidXlsx } from "@excel/__tests__/helpers/expect-valid-xlsx";
import {
  formCheckboxSetText,
  formCheckboxVmlAnchor,
  isFormCheckbox
} from "@excel/core/form-control";
import {
  addFormCheckbox,
  getFormCheckboxes,
  getSheetModel,
  setSheetModel
} from "@excel/core/worksheet";
import { Cell, Workbook } from "@excel/index";
import { describe, it, expect } from "vitest";

describe("Form Control Checkbox", () => {
  describe("FormCheckbox class via worksheet", () => {
    it("should create FormCheckbox with range string", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", {
        checked: true,
        link: "C1",
        text: "Test Checkbox"
      });

      expect(checkbox.model.tl.col).toBe(0);
      expect(checkbox.model.tl.row).toBe(0);
      expect(checkbox.model.br.col).toBe(1);
      expect(checkbox.model.br.row).toBe(1);
      expect(checkbox.model.checked).toBe("Checked");
      expect(checkbox.model.link).toBe("$C$1");
      expect(checkbox.model.text).toBe("Test Checkbox");
    });

    it("should create FormCheckbox with range object", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(
        ws,
        { startCol: 0, startRow: 0, endCol: 1, endRow: 1 },
        { checked: false }
      );

      expect(checkbox.model.tl.col).toBe(0);
      expect(checkbox.model.tl.row).toBe(0);
      expect(checkbox.model.br.col).toBe(1);
      expect(checkbox.model.br.row).toBe(1);
      expect(checkbox.model.checked).toBe("Unchecked");
    });

    it("should default to unchecked", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", {});
      expect(checkbox.model.checked).toBe("Unchecked");
    });

    it("should convert link cell to absolute reference", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", { link: "D5" });
      expect(checkbox.model.link).toBe("$D$5");
    });

    it("should handle already absolute link cell", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", { link: "$E$10" });
      expect(checkbox.model.link).toBe("$E$10");
    });

    it("should generate correct anchor for VML", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", {});
      const model = checkbox.model;

      // Anchor should have from and to positions
      expect(model.tl).toBeDefined();
      expect(model.tl.col).toBe(0);
      expect(model.tl.row).toBe(0);
      expect(model.br.col).toBe(1);
      expect(model.br.row).toBe(1);
    });

    it("should generate VML anchor string", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", {});
      const anchorString = formCheckboxVmlAnchor(checkbox.model);

      // Format: col1, colOff1, row1, rowOff1, col2, colOff2, row2, rowOff2
      expect(anchorString).toMatch(/^\d+,\s*\d+,\s*\d+,\s*\d+,\s*\d+,\s*\d+,\s*\d+,\s*\d+$/);
    });

    it("should serialize to model correctly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", {
        checked: true,
        link: "C1",
        text: "My Checkbox"
      });

      const model = checkbox.model;
      expect(model.checked).toBe("Checked");
      expect(model.link).toBe("$C$1");
      expect(model.text).toBe("My Checkbox");
      expect(model.tl).toBeDefined();
      expect(model.br).toBeDefined();
    });
  });

  describe("Worksheet Form Control API", () => {
    it("should add form checkbox to worksheet", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", { checked: true });

      expect(ws.formControls).toHaveLength(1);
      expect(isFormCheckbox(ws.formControls[0])).toBe(true);
      expect(ws.formControls[0].model.checked).toBe("Checked");
    });

    it("should add multiple form checkboxes", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", { checked: true, text: "Option 1" });
      addFormCheckbox(ws, "A3:B4", { checked: false, text: "Option 2" });
      addFormCheckbox(ws, "A5:B6", { checked: true, text: "Option 3" });

      expect(ws.formControls).toHaveLength(3);
      expect(ws.formControls[0].model.text).toBe("Option 1");
      expect(ws.formControls[1].model.text).toBe("Option 2");
      expect(ws.formControls[2].model.text).toBe("Option 3");
    });

    it("should return the created checkbox from addFormCheckbox", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", { text: "Test" });

      expect(isFormCheckbox(checkbox)).toBe(true);
      expect(checkbox.model.text).toBe("Test");
    });

    it("should get form checkboxes", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", {});
      addFormCheckbox(ws, "C1:D2", {});

      const checkboxes = getFormCheckboxes(ws);
      expect(checkboxes).toHaveLength(2);
    });
  });

  describe("XLSX Serialization", () => {
    it("should write form checkbox to xlsx", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", {
        checked: true,
        link: "C1",
        text: "Test Checkbox"
      });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      // Should contain unified VML drawing file
      expect(entries.has("xl/drawings/vmlDrawing1.vml")).toBe(true);

      // Should contain a DrawingML drawing part (Excel typically requires this for form controls)
      expect(entries.has("xl/drawings/drawing1.xml")).toBe(true);
      expect(entries.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(true);

      // Should contain ctrlProp file
      expect(entries.has("xl/ctrlProps/ctrlProp1.xml")).toBe(true);

      // Should have worksheet relationships
      expect(entries.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);

      // Should include <controls> entries that reference ctrlProp relationships
      const sheetEntry = entries.get("xl/worksheets/sheet1.xml");
      expect(sheetEntry).toBeDefined();
      const sheetXml = new TextDecoder().decode(sheetEntry!.data);
      expect(sheetXml).toContain("<controls");
      expect(sheetXml).toContain("<control");
      expect(sheetXml).toContain("<drawing");

      const controlRidMatch = sheetXml.match(/<control[^>]*\br:id="(rId\d+)"/);
      const shapeIdMatch = sheetXml.match(/<control[^>]*\bshapeId="(\d+)"/);
      expect(controlRidMatch).toBeTruthy();
      expect(shapeIdMatch).toBeTruthy();
      const controlRid = controlRidMatch![1];
      const shapeId = Number(shapeIdMatch![1]);
      expect(shapeId).toBe(1025);

      const relsEntry = entries.get("xl/worksheets/_rels/sheet1.xml.rels");
      expect(relsEntry).toBeDefined();
      const relsXml = new TextDecoder().decode(relsEntry!.data);
      expect(relsXml).toContain(`Id="${controlRid}"`);
      expect(relsXml).toContain("/relationships/ctrlProp");
    });

    it("should write multiple checkboxes to xlsx", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", { checked: true });
      addFormCheckbox(ws, "A3:B4", { checked: false });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      // Should contain unified VML drawing file
      expect(entries.has("xl/drawings/vmlDrawing1.vml")).toBe(true);

      // Should contain ctrlProp files for each checkbox
      expect(entries.has("xl/ctrlProps/ctrlProp1.xml")).toBe(true);
      expect(entries.has("xl/ctrlProps/ctrlProp2.xml")).toBe(true);
    });

    it("should generate valid VML structure", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", {
        checked: true,
        text: "Test Checkbox"
      });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      const vmlEntry = entries.get("xl/drawings/vmlDrawing1.vml");
      expect(vmlEntry).toBeDefined();

      const vmlContent = new TextDecoder().decode(vmlEntry!.data);

      // Check VML structure
      expect(vmlContent).toContain("<xml");
      expect(vmlContent).toContain("v:shapetype");
      expect(vmlContent).toContain("v:shape");
      expect(vmlContent).toContain("x:ClientData");
      expect(vmlContent).toContain('ObjectType="Checkbox"');
      // These fields are important to prevent Excel from repairing the sheet
      // by dropping legacy control parts.
      expect(vmlContent).toContain("x:MoveWithCells");
      expect(vmlContent).toContain("x:SizeWithCells");
      expect(vmlContent).toContain("x:LockText");
      expect(vmlContent).toContain("x:Row");
      expect(vmlContent).toContain("x:Column");
      expect(vmlContent).toContain("x:Anchor");
      expect(vmlContent).toContain("<x:Checked>");
    });

    it("issue-35 regression: should write Excel-compatible legacy checkboxes", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "J2:K3", { link: "D6", checked: false, text: "J2:K3" });
      addFormCheckbox(ws, "J4:J4", { link: "D7", checked: false, text: "J4:J4" });
      addFormCheckbox(ws, "J5:J5", { link: "D8", checked: false, text: "J5:J5" });
      addFormCheckbox(ws, "J6:J6", { link: "D9", checked: false, text: "J6:J6" });
      addFormCheckbox(ws, "J7", { link: "D10", checked: false, text: "J7" });
      addFormCheckbox(ws, "J8", { link: "D11", checked: false, text: "J8" });
      addFormCheckbox(ws, "J9", { link: "D12", checked: false, text: "J9" });

      // Ensure linked cells exist in sheet data
      Cell.setValue(ws, "D6", false);
      Cell.setValue(ws, "D7", false);
      Cell.setValue(ws, "D8", false);
      Cell.setValue(ws, "D9", false);
      Cell.setValue(ws, "D10", false);
      Cell.setValue(ws, "D11", false);
      Cell.setValue(ws, "D12", false);

      const buffer = await Workbook.toBuffer(wb);

      // Gate on strict OOXML wiring and ordering rules we learned from Excel repair logs.
      await expectValidXlsx(buffer);

      const entries = await extractAll(buffer);

      // Key parts that prevent Excel Online / desktop Excel from repairing sheet1.xml
      expect(entries.has("xl/drawings/vmlDrawing1.vml")).toBe(true);
      expect(entries.has("xl/drawings/drawing1.xml")).toBe(true);
      expect(entries.has("xl/drawings/_rels/drawing1.xml.rels")).toBe(true);
      expect(entries.has("xl/worksheets/_rels/sheet1.xml.rels")).toBe(true);
      expect(entries.has("xl/ctrlProps/ctrlProp1.xml")).toBe(true);

      const sheetXml = new TextDecoder().decode(entries.get("xl/worksheets/sheet1.xml")!.data);

      // Worksheet ordering constraint: legacyDrawing must come before controls.
      const legacyIdx = sheetXml.indexOf("<legacyDrawing");
      const controlsIdx = sheetXml.indexOf("<controls");
      expect(legacyIdx).not.toBe(-1);
      expect(controlsIdx).not.toBe(-1);
      expect(legacyIdx).toBeLessThan(controlsIdx);

      // Excel-typical structure for legacy controls on modern clients: x14 AlternateContent + controlPr/anchor.
      expect(sheetXml).toContain("mc:AlternateContent");
      expect(sheetXml).toContain('Requires="x14"');
      expect(sheetXml).toContain("<controlPr");
      expect(sheetXml).toContain("<anchor");
      expect(sheetXml).toContain("<drawing");

      const drawingXml = new TextDecoder().decode(entries.get("xl/drawings/drawing1.xml")!.data);
      expect(drawingXml).toContain('Requires="a14"');
      expect(drawingXml).toContain("a14:compatExt");
      expect(drawingXml).toContain("_x0000_s1025");
    });

    it("should generate valid ctrlProp structure", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", {
        checked: true,
        link: "C1"
      });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      const ctrlPropEntry = entries.get("xl/ctrlProps/ctrlProp1.xml");
      expect(ctrlPropEntry).toBeDefined();

      const ctrlPropContent = new TextDecoder().decode(ctrlPropEntry!.data);

      // Check ctrlProp structure
      expect(ctrlPropContent).toContain('<?xml version="1.0"');
      expect(ctrlPropContent).toContain("formControlPr");
      expect(ctrlPropContent).toContain('objectType="CheckBox"');
      expect(ctrlPropContent).toContain('checked="1"');
      expect(ctrlPropContent).toContain('fmlaLink="$C$1"');
    });

    it("should include proper content types", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      addFormCheckbox(ws, "A1:B2", { checked: true });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      const contentTypesEntry = entries.get("[Content_Types].xml");
      expect(contentTypesEntry).toBeDefined();

      const contentTypesContent = new TextDecoder().decode(contentTypesEntry!.data);

      // Should include ctrlProp content type
      expect(contentTypesContent).toContain("vnd.ms-excel.controlproperties+xml");
    });

    it("should not create VML files when no form controls exist", async () => {
      const wb = Workbook.create();
      Workbook.addWorksheet(wb, "Sheet1");

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      // Should not contain VML drawing file for form controls
      expect(entries.has("xl/drawings/vmlDrawingFC1.vml")).toBe(false);

      // Should not contain ctrlProp file
      expect(entries.has("xl/ctrlProps/ctrlProp1.xml")).toBe(false);
    });

    it("should handle worksheet with notes and form controls", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      // Add a note (which also uses VML)
      Cell.setNote(ws, "D1", { texts: [{ text: "This is a note" }] });

      // Add a form checkbox
      addFormCheckbox(ws, "A1:B2", { checked: true });

      const buffer = await Workbook.toBuffer(wb);
      const entries = await extractAll(buffer);

      // Both notes and form controls should be in a single unified VML file
      // This follows Excel's native behavior
      expect(entries.has("xl/drawings/vmlDrawing1.vml")).toBe(true);
      expect(entries.has("xl/drawings/vmlDrawingFC1.vml")).toBe(false);

      // Verify the unified VML contains both shapetypes
      const vmlEntry = entries.get("xl/drawings/vmlDrawing1.vml")!;
      const vmlContent = new TextDecoder().decode(vmlEntry.data);

      // Should contain shapetype for notes (id="_x0000_t202")
      expect(vmlContent).toContain("_x0000_t202");
      // Should contain shapetype for checkboxes (id="_x0000_t201")
      expect(vmlContent).toContain("_x0000_t201");
      // Should contain note shape
      expect(vmlContent).toContain('type="#_x0000_t202"');
      // Should contain checkbox shape
      expect(vmlContent).toContain('type="#_x0000_t201"');
    });
  });

  describe("Edge cases", () => {
    it("should handle large range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:Z100", {});
      expect(checkbox.model.tl.col).toBe(0);
      expect(checkbox.model.tl.row).toBe(0);
      expect(checkbox.model.br.col).toBe(25); // Z is 26th column, 0-based is 25
      expect(checkbox.model.br.row).toBe(99);
    });

    it("should handle single cell range", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:A1", {});
      expect(checkbox.model.tl.col).toBe(0);
      expect(checkbox.model.tl.row).toBe(0);
      // Single-cell ranges are treated like single-cell references for sizing
      expect(checkbox.model.br.col).toBe(2);
      expect(checkbox.model.br.row).toBe(1);
    });

    it("should handle single cell reference", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1", {});
      expect(checkbox.model.tl.col).toBe(0);
      expect(checkbox.model.tl.row).toBe(0);
      expect(checkbox.model.br.col).toBe(2);
      expect(checkbox.model.br.row).toBe(1);
    });

    it("should handle empty text", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", { text: "" });
      expect(checkbox.model.text).toBe("");
    });

    it("should handle unicode text", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");

      const checkbox = addFormCheckbox(ws, "A1:B2", { text: "测试复选框 🔲" });
      expect(checkbox.model.text).toBe("测试复选框 🔲");
    });
  });

  // ===========================================================================
  // Model round-trip via importSheet (regression for form controls being
  // silently dropped on the deserialise path).
  // ===========================================================================
  describe("worksheet model round-trip", () => {
    it("preserves form controls through worksheet.model getter+setter", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Source");
      addFormCheckbox(ws, "B2:C3", {
        text: "Round trip",
        link: "A1",
        checked: true,
        noThreeD: false,
        print: true
      });

      const wb2 = Workbook.create();
      const ws2 = Workbook.addWorksheet(wb2, "Target");
      setSheetModel(ws2, getSheetModel(ws));

      expect(ws2.formControls.length).toBe(1);
      const fc = ws2.formControls[0];
      expect(isFormCheckbox(fc)).toBe(true);
      expect(fc.model.text).toBe("Round trip");
      expect(fc.model.link).toBe("$A$1");
      expect(fc.model.checked).toBe("Checked");
      expect(fc.model.noThreeD).toBe(false);
      expect(fc.model.print).toBe(true);
      expect(fc.model.tl.col).toBe(1);
      expect(fc.model.br.col).toBe(2);
    });

    it("workbook.importSheet copies form controls to the new sheet", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Source");
      addFormCheckbox(ws, "D4:E5", { text: "Imported", checked: false });

      const imported = Workbook.importSheet(wb, ws, "ImportedSheet");
      expect(imported.formControls.length).toBe(1);
      expect(imported.formControls[0].model.text).toBe("Imported");
      expect(isFormCheckbox(imported.formControls[0])).toBe(true);
      // Mutating the imported control must not bleed back to the source.
      formCheckboxSetText(imported.formControls[0], "Mutated");
      expect(ws.formControls[0].model.text).toBe("Imported");
    });
  });
});
