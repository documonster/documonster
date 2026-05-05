/**
 * Table part checks — negative samples for every `table-*` kind.
 *
 * Each test builds a minimal table fixture with the exact writer-bug
 * pattern the check is meant to catch, asserts the validator surfaces
 * the matching kind, and then a positive baseline proves the
 * "canonical" shape doesn't regress to false positives.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith } from "./fixtures";

// -----------------------------------------------------------------------------
// Fixture helpers
// -----------------------------------------------------------------------------

function buildTablePackage(tableXml: string): Uint8Array {
  const parts = baseParts();
  parts["xl/tables/table1.xml"] = tableXml;
  parts["[Content_Types].xml"] = contentTypesWith([
    {
      partName: "/xl/workbook.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    },
    {
      partName: "/xl/styles.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
    },
    {
      partName: "/xl/sharedStrings.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
    },
    {
      partName: "/xl/theme/theme1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.theme+xml"
    },
    {
      partName: "/xl/worksheets/sheet1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    },
    {
      partName: "/xl/tables/table1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    }
  ]);
  return buildPackage(parts);
}

// -----------------------------------------------------------------------------
// Tests
// -----------------------------------------------------------------------------

describe("ooxml-validator / table — filterColumn redundancy", () => {
  it("flags a three-column table where every column has a bare filterColumn", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C10">
  <autoFilter ref="A1:C10">
    <filterColumn colId="0" hiddenButton="1"/>
    <filterColumn colId="1" hiddenButton="1"/>
    <filterColumn colId="2" hiddenButton="1"/>
  </autoFilter>
  <tableColumns count="3">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="C"/>
  </tableColumns>
</table>`)
    );
    expect(report.problems.some(p => p.kind === "table-filterColumn-redundant-per-column")).toBe(
      true
    );
  });

  it("accepts a table where only one column has a filterColumn (intentional hide)", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C10">
  <autoFilter ref="A1:C10">
    <filterColumn colId="1" hiddenButton="1"/>
  </autoFilter>
  <tableColumns count="3">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="C"/>
  </tableColumns>
</table>`)
    );
    expect(report.problems.some(p => p.kind === "table-filterColumn-redundant-per-column")).toBe(
      false
    );
  });

  it("accepts a table whose filterColumns carry real filter state", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B10">
  <autoFilter ref="A1:B10">
    <filterColumn colId="0"><filters><filter val="X"/></filters></filterColumn>
    <filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="5"/></customFilters></filterColumn>
  </autoFilter>
  <tableColumns count="2">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
  </tableColumns>
</table>`)
    );
    expect(report.problems.some(p => p.kind === "table-filterColumn-redundant-per-column")).toBe(
      false
    );
  });
});

describe("ooxml-validator / table — totalsRowFormula with built-in function", () => {
  it("flags totalsRowFunction=sum alongside a totalsRowFormula", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B11" totalsRowCount="1">
  <autoFilter ref="A1:B10"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A" totalsRowLabel="Total"/>
    <tableColumn id="2" name="B" totalsRowFunction="sum">
      <totalsRowFormula>SUBTOTAL(109,T1[B])</totalsRowFormula>
    </tableColumn>
  </tableColumns>
</table>`)
    );
    expect(
      report.problems.some(p => p.kind === "table-totalsRowFormula-with-builtin-function")
    ).toBe(true);
  });

  it("accepts totalsRowFunction=custom with a totalsRowFormula", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:B11" totalsRowCount="1">
  <autoFilter ref="A1:B10"/>
  <tableColumns count="2">
    <tableColumn id="1" name="A" totalsRowLabel="Total"/>
    <tableColumn id="2" name="B" totalsRowFunction="custom">
      <totalsRowFormula>MEDIAN(T1[B])</totalsRowFormula>
    </tableColumn>
  </tableColumns>
</table>`)
    );
    expect(
      report.problems.some(p => p.kind === "table-totalsRowFormula-with-builtin-function")
    ).toBe(false);
  });
});

describe("ooxml-validator / table — autoFilter ref vs totalsRow", () => {
  it("flags autoFilter covering the totals row", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C11" totalsRowCount="1">
  <autoFilter ref="A1:C11"/>
  <tableColumns count="3">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="C"/>
  </tableColumns>
</table>`)
    );
    expect(report.problems.some(p => p.kind === "table-autoFilter-covers-totalsRow")).toBe(true);
  });

  it("accepts autoFilter that stops one row above the totals row", async () => {
    const report = await validateXlsxBuffer(
      buildTablePackage(`<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C11" totalsRowCount="1">
  <autoFilter ref="A1:C10"/>
  <tableColumns count="3">
    <tableColumn id="1" name="A"/>
    <tableColumn id="2" name="B"/>
    <tableColumn id="3" name="C"/>
  </tableColumns>
</table>`)
    );
    expect(report.problems.some(p => p.kind === "table-autoFilter-covers-totalsRow")).toBe(false);
  });
});
