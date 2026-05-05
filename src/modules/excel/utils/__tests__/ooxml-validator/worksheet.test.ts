/**
 * Worksheet structural checks: child element ordering, cell r= consistency,
 * merge region overlap, style/SST index bounds, shared-formula masters.
 *
 * These are the four new structural checks that make up the bulk of the
 * refactor's new value — they catch the most common causes of "Excel
 * needs to repair" and "cannot open" errors.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage } from "./fixtures";

function sheetXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
${body}
</worksheet>`;
}

describe("ooxml-validator / worksheet — child ordering", () => {
  it("accepts a canonical ordering", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetPr/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
  <mergeCells count="0"/>
  <pageMargins left="0" right="0" top="0" bottom="0" header="0" footer="0"/>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-child-out-of-order")).toBe(false);
  });

  it("flags mergeCells before sheetData", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <mergeCells count="0"/>
  <sheetData/>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-child-out-of-order")).toBe(true);
  });

  it("flags pageMargins after legacyDrawing", async () => {
    // picture is at rank 34, pageMargins at rank 21.
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData/>
  <legacyDrawing r:id="rIdX"/>
  <pageMargins left="0" right="0" top="0" bottom="0" header="0" footer="0"/>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-child-out-of-order")).toBe(true);
  });

  it("tolerates unknown extension elements", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData/>
  <extLst><ext uri="some-ext"/></extLst>
  <ns:unknownExtension/>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-child-out-of-order")).toBe(false);
  });
});

describe("ooxml-validator / worksheet — cell ref consistency", () => {
  it("flags cell r= row that doesn't match enclosing row", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="B2"><v>10</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-ref-row-mismatch")).toBe(true);
  });

  it("flags cell missing r= attribute", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c><v>10</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-ref-missing")).toBe(true);
  });

  it("flags invalid cell r= attribute", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="not-a-cell"><v>10</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-ref-invalid")).toBe(true);
  });

  it("flags row index > 1048576", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="9999999"><c r="A9999999"><v>10</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-row-index-out-of-bounds")).toBe(true);
  });
});

describe("ooxml-validator / worksheet — merge overlap", () => {
  it("flags overlapping merge cells", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData/>
  <mergeCells count="2">
    <mergeCell ref="A1:C3"/>
    <mergeCell ref="B2:D4"/>
  </mergeCells>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-merge-overlap")).toBe(true);
  });

  it("accepts non-overlapping merge cells", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData/>
  <mergeCells count="2">
    <mergeCell ref="A1:B2"/>
    <mergeCell ref="C3:D4"/>
  </mergeCells>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-merge-overlap")).toBe(false);
  });

  it("flags invalid merge range", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData/>
  <mergeCells count="1">
    <mergeCell ref="not-a-range"/>
  </mergeCells>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-merge-invalid-range")).toBe(true);
  });
});

describe("ooxml-validator / worksheet — style index", () => {
  it("flags cell s= that exceeds cellXfs count", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1" s="99"><v>10</v></c></row>
  </sheetData>
    `);
    // baseParts has cellXfs count=1
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-style-index-oob")).toBe(true);
  });

  it("accepts cell s= within cellXfs bounds", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1" s="0"><v>10</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-style-index-oob")).toBe(false);
  });
});

describe("ooxml-validator / worksheet — shared strings index", () => {
  it("flags cell t=s with index beyond SST size", async () => {
    const parts = baseParts();
    // baseParts SST has uniqueCount=0; index 0 is already invalid.
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-sst-index-oob")).toBe(true);
  });

  it("accepts cell t=s with valid index", async () => {
    const parts = baseParts();
    parts["xl/sharedStrings.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>hello</t></si>
</sst>`;
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-cell-sst-index-oob")).toBe(false);
  });
});

describe("ooxml-validator / worksheet — shared formula", () => {
  it("flags follower without a master", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" si="0"/><v>10</v></c></row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>11</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-sharedFormula-master-missing")).toBe(true);
  });

  it("flags duplicate master for same si", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">SUM(B:B)</f><v>10</v></c></row>
    <row r="2"><c r="A2"><f t="shared" ref="A2:A4" si="0">SUM(C:C)</f><v>11</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(report.problems.some(p => p.kind === "sheet-sharedFormula-duplicate-master")).toBe(true);
  });

  it("accepts a correctly-formed shared-formula master+followers", async () => {
    const parts = baseParts();
    parts["xl/worksheets/sheet1.xml"] = sheetXml(`
  <sheetData>
    <row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">SUM(B:B)</f><v>10</v></c></row>
    <row r="2"><c r="A2"><f t="shared" si="0"/><v>11</v></c></row>
    <row r="3"><c r="A3"><f t="shared" si="0"/><v>12</v></c></row>
  </sheetData>
    `);
    const report = await validateXlsxBuffer(buildPackage(parts));
    expect(
      report.problems.some(
        p =>
          p.kind === "sheet-sharedFormula-master-missing" ||
          p.kind === "sheet-sharedFormula-duplicate-master"
      )
    ).toBe(false);
  });
});
