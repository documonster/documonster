/**
 * ChartEx sidecar stub-detection tests.
 *
 * Pure-stub chartStyle / chartColorStyle files (id attribute only, no
 * children) make Excel 2016+ drop the parent chartEx + drawing — but
 * only for sidecars referenced from a `chartEx` chart part. Classic
 * chart stubs are lenient and must NOT be flagged.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith, relsWith } from "./fixtures";

function buildChartExWithSidecars(
  styleXml: string,
  colorsXml: string,
  referencedFromChartEx: boolean
): Uint8Array {
  const parts = baseParts();
  parts["xl/charts/chartEx1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
  <cx:chartData><cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data></cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0"><cx:dataId val="0"/></cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>
</cx:chartSpace>`;
  parts["xl/charts/style1.xml"] = styleXml;
  parts["xl/charts/colors1.xml"] = colorsXml;

  if (referencedFromChartEx) {
    parts["xl/charts/_rels/chartEx1.xml.rels"] = relsWith([
      {
        id: "rId1",
        type: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
        target: "style1.xml"
      },
      {
        id: "rId2",
        type: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
        target: "colors1.xml"
      }
    ]);
  }

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
      partName: "/xl/charts/chartEx1.xml",
      contentType: "application/vnd.ms-office.chartex+xml"
    },
    {
      partName: "/xl/charts/style1.xml",
      contentType: "application/vnd.ms-office.chartstyle+xml"
    },
    {
      partName: "/xl/charts/colors1.xml",
      contentType: "application/vnd.ms-office.chartcolorstyle+xml"
    }
  ]);
  return buildPackage(parts);
}

const STUB_STYLE = `<?xml version="1.0" encoding="UTF-8"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="395"/>`;

const STUB_COLORS = `<?xml version="1.0" encoding="UTF-8"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"/>`;

const POPULATED_STYLE = `<?xml version="1.0" encoding="UTF-8"?>
<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="395">
  <cs:axisTitle/><cs:categoryAxis/><cs:chartArea/><cs:dataLabel/><cs:dataLabelCallout/>
</cs:chartStyle>`;

const POPULATED_COLORS = `<?xml version="1.0" encoding="UTF-8"?>
<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">
  <a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/>
  <a:schemeClr val="accent4"/><a:schemeClr val="accent5"/><a:schemeClr val="accent6"/>
</cs:colorStyle>`;

describe("ooxml-validator / chartEx sidecar stub detection", () => {
  it("flags stub <cs:chartStyle id='395'/> when referenced from a chartEx", async () => {
    const report = await validateXlsxBuffer(
      buildChartExWithSidecars(STUB_STYLE, POPULATED_COLORS, true)
    );
    expect(report.problems.some(p => p.kind === "chartEx-chartStyle-stub-form")).toBe(true);
  });

  it("flags stub <cs:colorStyle id='10' meth='cycle'/> when referenced from a chartEx", async () => {
    const report = await validateXlsxBuffer(
      buildChartExWithSidecars(POPULATED_STYLE, STUB_COLORS, true)
    );
    expect(report.problems.some(p => p.kind === "chartEx-chartColorStyle-stub-form")).toBe(true);
  });

  it("accepts populated chartStyle + chartColorStyle referenced from a chartEx", async () => {
    const report = await validateXlsxBuffer(
      buildChartExWithSidecars(POPULATED_STYLE, POPULATED_COLORS, true)
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chartEx-chartStyle-stub-form" ||
          p.kind === "chartEx-chartColorStyle-stub-form"
      )
    ).toBe(false);
  });

  it("does NOT flag stub sidecars when they are NOT referenced from a chartEx (classic chart)", async () => {
    // Sidecars exist in the zip but no chartEx part points at them.
    // Classic-chart stubs are lenient — we must not raise a false positive.
    const report = await validateXlsxBuffer(
      buildChartExWithSidecars(STUB_STYLE, STUB_COLORS, false)
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chartEx-chartStyle-stub-form" ||
          p.kind === "chartEx-chartColorStyle-stub-form"
      )
    ).toBe(false);
  });
});
