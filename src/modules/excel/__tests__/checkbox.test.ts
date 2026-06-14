import { extractAll } from "@archive/unzip/extract";
import { testUtils } from "@excel/__tests__/shared";
import {
  cellEffectiveType,
  cellGetValue,
  cellSetValue,
  cellType,
  cellToString,
  cellGetModel,
  cellToCsvString
} from "@excel/cell";
import { Enums } from "@excel/enums";
import { Cell, Workbook } from "@excel/index";
import { getXlsxIo } from "@excel/workbook";
import { getCell } from "@excel/worksheet";
import { StylesXform } from "@excel/xlsx/xform/style/styles-xform";
import { describe, it, expect } from "vitest";

describe("Checkbox Feature", () => {
  describe("Cell checkbox value", () => {
    it("should set checkbox value using object syntax", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      // Set checkbox values
      Cell.setValue(ws, "A1", { checkbox: true });
      Cell.setValue(ws, "A2", { checkbox: false });

      const cell1 = getCell(ws, "A1");
      const cell2 = getCell(ws, "A2");

      expect(cellGetValue(cell1)).toEqual({ checkbox: true });
      expect(cellGetValue(cell2)).toEqual({ checkbox: false });
      expect(cellType(cell1)).toBe(Enums.ValueType.Checkbox);
      expect(cellType(cell2)).toBe(Enums.ValueType.Checkbox);
    });

    it("should convert checkbox to boolean in effectiveType", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      Cell.setValue(ws, "A1", { checkbox: true });

      const cell = getCell(ws, "A1");
      expect(cellEffectiveType(cell)).toBe(Enums.ValueType.Boolean);
    });

    it("should convert checkbox to CSV correctly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      Cell.setValue(ws, "A1", { checkbox: true });
      Cell.setValue(ws, "A2", { checkbox: false });

      expect(cellToCsvString(getCell(ws, "A1"))).toBe(1);
      expect(cellToCsvString(getCell(ws, "A2"))).toBe(0);
    });

    it("should convert checkbox to string correctly", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      Cell.setValue(ws, "A1", { checkbox: true });
      Cell.setValue(ws, "A2", { checkbox: false });

      expect(cellToString(getCell(ws, "A1"))).toBe("true");
      expect(cellToString(getCell(ws, "A2"))).toBe("false");
    });
  });

  describe("Checkbox serialization", () => {
    it("should write and read checkbox values correctly", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      // Set various checkbox values
      Cell.setValue(ws, "A1", { checkbox: true });
      Cell.setValue(ws, "A2", { checkbox: false });
      Cell.setValue(ws, "B1", { checkbox: true });
      Cell.setValue(ws, "B2", { checkbox: false });

      // Write to buffer
      const buffer = await getXlsxIo(wb).writeBuffer();

      // Read back
      const wb2 = Workbook.create();
      await getXlsxIo(wb2).load(buffer);
      const ws2 = Workbook.getWorksheet(wb2, "Checkboxes")!;

      // Verify values - checkboxes should be read back as booleans
      // because the checkbox metadata is in the style
      expect(Cell.getValue(ws2, "A1")).toBe(true);
      expect(Cell.getValue(ws2, "A2")).toBe(false);
      expect(Cell.getValue(ws2, "B1")).toBe(true);
      expect(Cell.getValue(ws2, "B2")).toBe(false);
    });

    it("should include featurePropertyBag in workbook when checkboxes are used", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      Cell.setValue(ws, "A1", { checkbox: true });

      const buffer = await getXlsxIo(wb).writeBuffer();

      // Check that the zip contains featurePropertyBag
      const entries = await extractAll(buffer);

      expect(entries.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(true);

      const featureBag = entries.get("xl/featurePropertyBag/featurePropertyBag.xml");
      expect(featureBag).toBeDefined();
      const featureBagText = new TextDecoder().decode(featureBag!.data);
      const normalizedFeatureBag = featureBagText.replace(/>\s+</g, "><").trim();
      expect(normalizedFeatureBag).toContain(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      );
      expect(normalizedFeatureBag).toContain(
        '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag">'
      );
      expect(normalizedFeatureBag).toContain('<bag type="Checkbox"/>');
      expect(normalizedFeatureBag).toContain(
        '<bag type="XFControls"><bagId k="CellControl">0</bagId></bag>'
      );
      expect(normalizedFeatureBag).toContain(
        '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag>'
      );
      expect(normalizedFeatureBag).toContain(
        '<bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag>'
      );
      expect(normalizedFeatureBag).toContain("</FeaturePropertyBags>");

      // Check Content_Types.xml includes featurePropertyBag
      const contentTypes = entries.get("[Content_Types].xml");
      expect(contentTypes).toBeDefined();
      const contentTypesText = new TextDecoder().decode(contentTypes!.data);
      expect(contentTypesText).toContain("featurePropertyBag");
      expect(contentTypesText).toContain("application/vnd.ms-excel.featurepropertybag+xml");

      // Check workbook rels includes featurePropertyBag relationship
      const workbookRels = entries.get("xl/_rels/workbook.xml.rels");
      expect(workbookRels).toBeDefined();
      const workbookRelsText = new TextDecoder().decode(workbookRels!.data);
      expect(workbookRelsText).toContain("FeaturePropertyBag");
      expect(workbookRelsText).toContain("featurePropertyBag/featurePropertyBag.xml");
    });

    it("should not include featurePropertyBag when no checkboxes are used", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "NoCheckboxes");

      Cell.setValue(ws, "A1", "Normal text");
      Cell.setValue(ws, "A2", 123);
      Cell.setValue(ws, "A3", true); // regular boolean

      const buffer = await getXlsxIo(wb).writeBuffer();

      const entries = await extractAll(buffer);

      expect(entries.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(false);

      const contentTypes = entries.get("[Content_Types].xml");
      expect(contentTypes).toBeDefined();
      const contentTypesText = new TextDecoder().decode(contentTypes!.data);
      expect(contentTypesText).not.toContain("featurepropertybag");
    });

    it("should apply checkbox style to checkbox cells", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Checkboxes");

      Cell.setValue(ws, "A1", { checkbox: true });
      Cell.setValue(ws, "A2", { checkbox: false });

      const buffer = await getXlsxIo(wb).writeBuffer();

      const entries = await extractAll(buffer);

      // Check styles.xml includes checkbox style with extLst
      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);

      // Should have xfComplement extension
      expect(stylesText).toContain("xfpb:xfComplement");
      expect(stylesText).toContain("C7286773-470A-42A8-94C5-96B5CB345126");
      // The <ext> element should not switch to a default namespace; it should only declare xfpb
      expect(stylesText).not.toContain(
        'xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"'
      );

      // Check worksheet uses checkbox style
      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);

      // Both cells should have same style (checkbox style)
      // and t="b" for boolean type
      expect(sheet1Text).toContain('t="b"');
    });

    it("should merge checkbox with user-provided style", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "StyledCheckbox");

      const cell = getCell(ws, "A1");
      cellSetValue(cell, { checkbox: true });
      cell.style.font = { bold: true };

      // Sanity check: style is present on the cell model before serialization
      expect(cellGetModel(cell).style?.font?.bold).toBe(true);

      const buffer = await getXlsxIo(wb).writeBuffer();
      const entries = await extractAll(buffer);

      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);

      // Get the actual styleId used by the checkbox cell
      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);
      const a1Match = sheet1Text.match(/<c\b[^>]*\br="A1"[^>]*>/);
      expect(a1Match).toBeTruthy();
      const styleIdMatch = a1Match![0].match(/\bs="(\d+)"/);
      expect(styleIdMatch).toBeTruthy();
      const styleId = Number(styleIdMatch![1]);
      expect(Number.isFinite(styleId)).toBe(true);

      // Extract the xf at that styleId from cellXfs
      const cellXfsMatch = stylesText.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
      expect(cellXfsMatch).toBeTruthy();
      const xfNodes = cellXfsMatch![1].match(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g) || [];
      expect(xfNodes.length).toBeGreaterThan(styleId);
      const xfXml = xfNodes[styleId];
      expect(xfXml).toContain("xfpb:xfComplement");

      // Capture the fontId from that xf
      const fontIdMatch = xfXml.match(/\bfontId="(\d+)"/);
      expect(fontIdMatch).toBeTruthy();
      const checkboxFontId = Number(fontIdMatch![1]);

      // Ensure the referenced font is bold
      const fontsMatch = stylesText.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
      expect(fontsMatch).toBeTruthy();
      const fontNodes = fontsMatch![1].match(/<font>[\s\S]*?<\/font>/g) || [];
      expect(fontNodes.length).toBeGreaterThan(checkboxFontId);
      expect(fontNodes[checkboxFontId]).toContain("<b/");
    });

    it("should merge checkbox with fill/border/numFmt/alignment", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "StyledCheckbox2");

      const cell = getCell(ws, "A1");
      cellSetValue(cell, { checkbox: true });
      cell.style.font = testUtils.styles.fonts.broadwayRedOutline20;
      cell.style.border = testUtils.styles.borders.doubleRed;
      cell.style.fill = testUtils.styles.fills.blueWhiteHGrad;
      cell.style.alignment = testUtils.styles.namedAlignments.middleCentre;
      cell.style.numFmt = testUtils.styles.numFmts.numFmt1;

      const buffer = await getXlsxIo(wb).writeBuffer();
      const entries = await extractAll(buffer);

      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);
      const a1CellTag = sheet1Text.match(/<c\b[^>]*\br="A1"[^>]*>/);
      expect(a1CellTag).toBeTruthy();
      const styleIdMatch = a1CellTag![0].match(/\bs="(\d+)"/);
      expect(styleIdMatch).toBeTruthy();
      const styleId = Number(styleIdMatch![1]);

      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);

      const cellXfsMatch = stylesText.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
      expect(cellXfsMatch).toBeTruthy();
      const xfNodes = cellXfsMatch![1].match(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g) || [];
      expect(xfNodes.length).toBeGreaterThan(styleId);
      const xfXml = xfNodes[styleId];

      // Checkbox ext
      expect(xfXml).toContain("xfpb:xfComplement");

      // User style should be applied
      expect(xfXml).toContain('applyFont="1"');
      expect(xfXml).toContain('applyFill="1"');
      expect(xfXml).toContain('applyBorder="1"');
      expect(xfXml).toContain('applyNumberFormat="1"');
      const numFmtIdMatch = xfXml.match(/\bnumFmtId="(\d+)"/);
      expect(numFmtIdMatch).toBeTruthy();
      expect(Number(numFmtIdMatch![1])).toBeGreaterThan(0);
      expect(xfXml).toContain('applyAlignment="1"');
      expect(xfXml).toContain("<alignment");
    });

    it("should not leak checkbox extension to non-checkbox cells", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "LeakCheck");

      const sharedStyle: any = {
        font: testUtils.styles.fonts.broadwayRedOutline20
      };

      const a1 = getCell(ws, "A1");
      cellSetValue(a1, true);
      a1.style = sharedStyle;

      const a2 = getCell(ws, "A2");
      cellSetValue(a2, { checkbox: true });
      a2.style = sharedStyle;

      const buffer = await getXlsxIo(wb).writeBuffer();
      const entries = await extractAll(buffer);

      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);

      const a1Tag = sheet1Text.match(/<c\b[^>]*\br="A1"[^>]*>/);
      const a2Tag = sheet1Text.match(/<c\b[^>]*\br="A2"[^>]*>/);
      expect(a1Tag).toBeTruthy();
      expect(a2Tag).toBeTruthy();

      const a1StyleIdMatch = a1Tag![0].match(/\bs="(\d+)"/);
      const a2StyleIdMatch = a2Tag![0].match(/\bs="(\d+)"/);
      expect(a1StyleIdMatch).toBeTruthy();
      expect(a2StyleIdMatch).toBeTruthy();

      const a1StyleId = Number(a1StyleIdMatch![1]);
      const a2StyleId = Number(a2StyleIdMatch![1]);
      expect(a1StyleId).not.toBe(a2StyleId);

      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);
      const cellXfsMatch = stylesText.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
      expect(cellXfsMatch).toBeTruthy();
      const xfNodes = cellXfsMatch![1].match(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g) || [];
      expect(xfNodes.length).toBeGreaterThan(Math.max(a1StyleId, a2StyleId));

      const xfA1 = xfNodes[a1StyleId];
      const xfA2 = xfNodes[a2StyleId];

      expect(xfA1).not.toContain("xfpb:xfComplement");
      expect(xfA2).toContain("xfpb:xfComplement");
    });

    it("should support multiple checkbox styles in one sheet", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "MultiCheckboxStyles");

      const a1 = getCell(ws, "A1");
      cellSetValue(a1, { checkbox: true });
      a1.style.font = { bold: true };

      const a2 = getCell(ws, "A2");
      cellSetValue(a2, { checkbox: false });
      a2.style.fill = testUtils.styles.fills.blueWhiteHGrad;

      const buffer = await getXlsxIo(wb).writeBuffer();
      const entries = await extractAll(buffer);

      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);

      const a1Tag = sheet1Text.match(/<c\b[^>]*\br="A1"[^>]*>/);
      const a2Tag = sheet1Text.match(/<c\b[^>]*\br="A2"[^>]*>/);
      expect(a1Tag).toBeTruthy();
      expect(a2Tag).toBeTruthy();

      const a1StyleIdMatch = a1Tag![0].match(/\bs="(\d+)"/);
      const a2StyleIdMatch = a2Tag![0].match(/\bs="(\d+)"/);
      expect(a1StyleIdMatch).toBeTruthy();
      expect(a2StyleIdMatch).toBeTruthy();
      const a1StyleId = Number(a1StyleIdMatch![1]);
      const a2StyleId = Number(a2StyleIdMatch![1]);
      expect(a1StyleId).not.toBe(a2StyleId);

      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);
      const cellXfsMatch = stylesText.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
      expect(cellXfsMatch).toBeTruthy();
      const xfNodes = cellXfsMatch![1].match(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g) || [];
      expect(xfNodes.length).toBeGreaterThan(Math.max(a1StyleId, a2StyleId));

      const xfA1 = xfNodes[a1StyleId];
      const xfA2 = xfNodes[a2StyleId];
      expect(xfA1).toContain("xfpb:xfComplement");
      expect(xfA2).toContain("xfpb:xfComplement");

      // A1 should retain bold font
      const a1FontIdMatch = xfA1.match(/\bfontId="(\d+)"/);
      expect(a1FontIdMatch).toBeTruthy();
      const a1FontId = Number(a1FontIdMatch![1]);
      const fontsMatch = stylesText.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
      expect(fontsMatch).toBeTruthy();
      const fontNodes = fontsMatch![1].match(/<font>[\s\S]*?<\/font>/g) || [];
      expect(fontNodes.length).toBeGreaterThan(a1FontId);
      expect(fontNodes[a1FontId]).toContain("<b/");

      // A2 should retain non-default fill (gradient)
      const a2FillIdMatch = xfA2.match(/\bfillId="(\d+)"/);
      expect(a2FillIdMatch).toBeTruthy();
      const a2FillId = Number(a2FillIdMatch![1]);
      expect(a2FillId).toBeGreaterThan(0);
      const fillsMatch = stylesText.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/);
      expect(fillsMatch).toBeTruthy();
      const fillNodes = fillsMatch![1].match(/<fill>[\s\S]*?<\/fill>/g) || [];
      expect(fillNodes.length).toBeGreaterThan(a2FillId);
      expect(fillNodes[a2FillId]).toContain("<gradientFill");
    });

    it("should dedupe identical checkbox styles", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "DedupeCheckboxStyles");

      // Use the exact same style values for both checkbox cells
      const style: any = {
        border: testUtils.styles.borders.doubleRed,
        numFmt: testUtils.styles.numFmts.numFmt1
      };

      const a1 = getCell(ws, "A1");
      cellSetValue(a1, { checkbox: true });
      a1.style = style;

      const a2 = getCell(ws, "A2");
      cellSetValue(a2, { checkbox: false });
      a2.style = style;

      const buffer = await getXlsxIo(wb).writeBuffer();
      const entries = await extractAll(buffer);

      const sheet1 = entries.get("xl/worksheets/sheet1.xml");
      expect(sheet1).toBeDefined();
      const sheet1Text = new TextDecoder().decode(sheet1!.data);

      const a1Tag = sheet1Text.match(/<c\b[^>]*\br="A1"[^>]*>/);
      const a2Tag = sheet1Text.match(/<c\b[^>]*\br="A2"[^>]*>/);
      expect(a1Tag).toBeTruthy();
      expect(a2Tag).toBeTruthy();

      const a1StyleIdMatch = a1Tag![0].match(/\bs="(\d+)"/);
      const a2StyleIdMatch = a2Tag![0].match(/\bs="(\d+)"/);
      expect(a1StyleIdMatch).toBeTruthy();
      expect(a2StyleIdMatch).toBeTruthy();
      const a1StyleId = Number(a1StyleIdMatch![1]);
      const a2StyleId = Number(a2StyleIdMatch![1]);
      expect(a1StyleId).toBe(a2StyleId);

      const styles = entries.get("xl/styles.xml");
      expect(styles).toBeDefined();
      const stylesText = new TextDecoder().decode(styles!.data);
      const cellXfsMatch = stylesText.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
      expect(cellXfsMatch).toBeTruthy();
      const xfNodes = cellXfsMatch![1].match(/<xf\b[^>]*\/>|<xf\b[\s\S]*?<\/xf>/g) || [];
      expect(xfNodes.length).toBeGreaterThan(a1StyleId);

      const xfXml = xfNodes[a1StyleId];
      expect(xfXml).toContain("xfpb:xfComplement");
      expect(xfXml).toContain('applyBorder="1"');
      expect(xfXml).toContain('applyNumberFormat="1"');
    });

    it("should be able to generate checkbox+font in StylesXform", () => {
      const styles = new StylesXform(true);
      styles.addStyleModel({ font: { bold: true } }, Enums.ValueType.Checkbox);

      const stylesXml = styles.toXml(styles.model);
      expect(stylesXml).toContain("xfpb:xfComplement");
      expect(stylesXml).toContain("<b/");
    });
  });

  describe("Checkbox with mixed content", () => {
    it("should handle worksheet with both checkboxes and regular values", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Mixed");

      Cell.setValue(ws, "A1", "Header");
      Cell.setValue(ws, "A2", { checkbox: true });
      Cell.setValue(ws, "A3", { checkbox: false });
      Cell.setValue(ws, "B1", 123);
      Cell.setValue(ws, "B2", true); // regular boolean
      Cell.setValue(ws, "C1", { checkbox: true });

      const buffer = await getXlsxIo(wb).writeBuffer();

      const entries = await extractAll(buffer);

      expect(entries.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(true);

      // Read back and verify
      const wb2 = Workbook.create();
      await getXlsxIo(wb2).load(buffer);
      const ws2 = Workbook.getWorksheet(wb2, "Mixed")!;

      expect(Cell.getValue(ws2, "A1")).toBe("Header");
      expect(Cell.getValue(ws2, "A2")).toBe(true);
      expect(Cell.getValue(ws2, "A3")).toBe(false);
      expect(Cell.getValue(ws2, "B1")).toBe(123);
      expect(Cell.getValue(ws2, "B2")).toBe(true);
      expect(Cell.getValue(ws2, "C1")).toBe(true);
    });

    it("should handle multiple worksheets with checkboxes", async () => {
      const wb = Workbook.create();

      const ws1 = Workbook.addWorksheet(wb, "Sheet1");
      Cell.setValue(ws1, "A1", { checkbox: true });

      const ws2 = Workbook.addWorksheet(wb, "Sheet2");
      Cell.setValue(ws2, "B2", { checkbox: false });

      const buffer = await getXlsxIo(wb).writeBuffer();

      const entries = await extractAll(buffer);

      // Should only have one featurePropertyBag file for the entire workbook
      const featureBagPaths = [...entries.keys()].filter(
        path => path === "xl/featurePropertyBag/featurePropertyBag.xml"
      );
      expect(featureBagPaths.length).toBe(1);
    });
  });

  describe("Edge cases", () => {
    it("should handle empty workbook without checkboxes", async () => {
      const wb = Workbook.create();
      Workbook.addWorksheet(wb, "Empty");

      const buffer = await getXlsxIo(wb).writeBuffer();

      const entries = await extractAll(buffer);

      expect(entries.has("xl/featurePropertyBag/featurePropertyBag.xml")).toBe(false);
    });

    it("should distinguish between regular boolean and checkbox", () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Test");

      Cell.setValue(ws, "A1", true); // regular boolean
      Cell.setValue(ws, "A2", { checkbox: true }); // checkbox

      expect(Cell.getType(ws, "A1")).toBe(Enums.ValueType.Boolean);
      expect(Cell.getType(ws, "A2")).toBe(Enums.ValueType.Checkbox);
    });
  });
});
