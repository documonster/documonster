/**
 * Chart & ChartEx internal-structure checks.
 *
 * Negative samples for every `chartEx-*` problem kind the validator
 * emits. Each test seeds a minimal fixture with the exact broken
 * pattern the writer used to emit, then asserts the validator produces
 * the matching `kind`. Positive baseline tests cover the "schema
 * conformant" shape so we do not accidentally flag valid output.
 */

import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { describe, expect, it } from "vitest";

import { baseParts, buildPackage, contentTypesWith } from "./fixtures";

// -----------------------------------------------------------------------------
// ChartEx fixture helpers
// -----------------------------------------------------------------------------

/**
 * Wrap a chartEx `<cx:chartSpace>` body in the minimum scaffolding
 * needed for the validator to find it — the part itself, the override
 * content-type, and a drawing rel that points at it (so the validator
 * doesn't fail on a missing parent anchor first).
 */
function buildChartExPackage(chartExInner: string): Uint8Array {
  const parts = baseParts();

  // Add a drawing that hosts the chartEx.
  parts["xl/drawings/drawing1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
      <mc:Choice Requires="cx1" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2014/chartex">
        <xdr:graphicFrame macro="">
          <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
          <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
        </xdr:graphicFrame>
      </mc:Choice>
      <mc:Fallback><xdr:sp/></mc:Fallback>
    </mc:AlternateContent>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
  parts["xl/charts/chartEx1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">
${chartExInner}
</cx:chartSpace>`;

  // Content types and rels.
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
    { partName: "/xl/charts/chartEx1.xml", contentType: "application/vnd.ms-office.chartex+xml" },
    {
      partName: "/xl/drawings/drawing1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
    }
  ]);
  return buildPackage(parts);
}

function minimalChartExBody(): string {
  // A valid baseline chartEx with one data entry, one axis, one series.
  // ChartEx requires `<cx:f>` to point at a hidden defined name, not
  // directly at a worksheet range, so the baseline uses the
  // `_xlchart.v1.N` convention Excel itself writes.
  return `
  <cx:chartData>
    <cx:data id="0">
      <cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim>
    </cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`;
}

// -----------------------------------------------------------------------------
// Tests
// -----------------------------------------------------------------------------

describe("ooxml-validator / chart — baseline", () => {
  it("accepts a minimal schema-valid chartEx", async () => {
    const report = await validateXlsxBuffer(buildChartExPackage(minimalChartExBody()));
    const chartProblems = report.problems.filter(p => p.kind.startsWith("chart"));
    expect(chartProblems).toEqual([]);
  });
});

describe("ooxml-validator / chart — chartEx schema violations", () => {
  it("flags a series with more than one <cx:dataId>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:strDim type="cat"><cx:f>_xlchart.v1.0</cx:f></cx:strDim></cx:data>
    <cx:data id="1"><cx:numDim type="val"><cx:f>_xlchart.v1.1</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:dataId val="1"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-series-too-many-dataId")).toBe(true);
  });

  it("flags <cx:axisId>N</cx:axisId> text-content form", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:axisId>2</cx:axisId>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-typed-element-text-form")).toBe(true);
  });

  it("flags <cx:binCount>N</cx:binCount> text-content form", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:binning><cx:binCount>12</cx:binCount></cx:binning>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-typed-element-text-form")).toBe(true);
  });

  it("flags <cx:auto/> element", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:binning><cx:auto/></cx:binning>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-invalid-auto-element")).toBe(true);
  });

  it("flags <cx:paretoLine> inside <cx:layoutPr>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:paretoLine/>
          </cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-paretoLine-in-layoutPr")).toBe(true);
  });

  it("flags <cx:title><cx:layout/>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:title pos="t" align="ctr" overlay="0"><cx:layout/></cx:title>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-title-direct-layout")).toBe(true);
  });
});

describe("ooxml-validator / chart — Tier 2 semantic checks", () => {
  it("flags axis missing both pos and type", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
        <cx:axis id="0"/>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-axis-missing-pos-and-type")).toBe(true);
  });

  it("accepts axis with pos only", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
        <cx:axis id="0" pos="b"/>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-axis-missing-pos-and-type")).toBe(false);
  });

  it("flags <cx:f> pointing at a direct sheet range", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$1:$A$3</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(
      report.problems.some(p => p.kind === "chartEx-f-uses-direct-range-not-defined-name")
    ).toBe(true);
  });

  it("accepts <cx:f> pointing at a hidden defined name (_xlchart.v1.N)", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="clusteredColumn" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(
      report.problems.some(p => p.kind === "chartEx-f-uses-direct-range-not-defined-name")
    ).toBe(false);
  });

  it("flags waterfall series missing <cx:layoutPr><cx:subtotals>", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="waterfall" hidden="0">
          <cx:dataId val="0"/>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-waterfall-missing-subtotals")).toBe(true);
  });

  it("accepts waterfall series with empty <cx:subtotals/> marker", async () => {
    const report = await validateXlsxBuffer(
      buildChartExPackage(`
  <cx:chartData>
    <cx:data id="0"><cx:numDim type="val"><cx:f>_xlchart.v1.0</cx:f></cx:numDim></cx:data>
  </cx:chartData>
  <cx:chart>
    <cx:plotArea>
      <cx:plotAreaRegion>
        <cx:series layoutId="waterfall" hidden="0">
          <cx:dataId val="0"/>
          <cx:layoutPr><cx:subtotals/></cx:layoutPr>
        </cx:series>
      </cx:plotAreaRegion>
    </cx:plotArea>
  </cx:chart>`)
    );
    expect(report.problems.some(p => p.kind === "chartEx-waterfall-missing-subtotals")).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Classic chart ECMA-376 schema conformance
//
// Every test seeds a hand-crafted `xl/charts/chart1.xml` body wrapped
// in the standard chartSpace scaffolding; the package ships a drawing
// that hosts the chart so the validator's relationship checks pass
// first, then the chart-specific schema checks run. Each test
// asserts the presence/absence of the relevant problem kind so
// regressions in the writer's element ordering / enum handling /
// numeric-range clamping surface as a failing test rather than a
// silently-corrupt xlsx.
// -----------------------------------------------------------------------------

/**
 * Wrap an inner `<c:chart>…</c:chart>` snippet in the full chart-space
 * scaffolding and return a zip buffer the validator can ingest.
 */
function buildClassicChartPackage(chartInner: string): Uint8Array {
  const parts = baseParts();
  parts["xl/drawings/drawing1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
  parts["xl/charts/chart1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>${chartInner}</c:chart>
</c:chartSpace>`;
  parts["xl/drawings/_rels/drawing1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
  parts["xl/worksheets/_rels/sheet1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  parts["xl/worksheets/sheet1.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`;
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
      partName: "/xl/charts/chart1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    },
    {
      partName: "/xl/drawings/drawing1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
    }
  ]);
  return buildPackage(parts);
}

/** Minimal bar chart the other tests mutate one element at a time. */
function minimalBarChartBody(overrides: { plotArea?: string; extraAxes?: string }): string {
  const plotArea =
    overrides.plotArea ??
    `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>
      ${overrides.extraAxes ?? ""}
    </c:plotArea>`;
  return `${plotArea}<c:plotVisOnly val="1"/>`;
}

describe("ooxml-validator / chart — classic schema baseline", () => {
  it("accepts a minimal schema-valid bar chart", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    const problems = report.problems.filter(p => p.kind.startsWith("chart-"));
    expect(problems).toEqual([]);
  });
});

describe("ooxml-validator / chart — classic child order", () => {
  it("flags <c:separator> appearing after <c:showLeaderLines> in <c:dLbls>", async () => {
    // Historical writer bug: `_renderDataLabels` emitted
    // showLeaderLines before separator. ECMA-376 CT_DLbls mandates
    // separator first, then showLeaderLines, then leaderLines.
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showCatName val="1"/>
            <c:showLeaderLines val="1"/>
            <c:separator> • </c:separator>
          </c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-child-out-of-order")).toBe(true);
  });

  it("accepts separator before showLeaderLines in <c:dLbls>", async () => {
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showCatName val="1"/>
            <c:separator> • </c:separator>
            <c:showLeaderLines val="1"/>
          </c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-child-out-of-order")).toBe(false);
  });

  it("flags <c:scaling> with <c:min> before <c:max>", async () => {
    // ECMA-376 CT_Scaling: logBase → orientation → max → min. The
    // previous writer order (min before max) triggered "Repaired
    // Records: Drawing" on Excel open and broke LibreOffice strict
    // mode outright.
    const plotArea = minimalBarChartBody({
      extraAxes: `<c:valAx>
        <c:axId val="3"/>
        <c:scaling><c:orientation val="minMax"/><c:min val="0"/><c:max val="100"/></c:scaling>
        <c:axPos val="r"/>
        <c:crossAx val="1"/>
      </c:valAx>`
    });
    const report = await validateXlsxBuffer(buildClassicChartPackage(plotArea), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-child-out-of-order")).toBe(true);
  });

  it("accepts <c:scaling> with <c:max> before <c:min>", async () => {
    const plotArea = minimalBarChartBody({
      extraAxes: `<c:valAx>
        <c:axId val="3"/>
        <c:scaling><c:orientation val="minMax"/><c:max val="100"/><c:min val="0"/></c:scaling>
        <c:axPos val="r"/>
        <c:crossAx val="1"/>
      </c:valAx>`
    });
    const report = await validateXlsxBuffer(buildClassicChartPackage(plotArea), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-child-out-of-order")).toBe(false);
  });

  it("accepts catAx before valAx inside plotArea (choice group)", async () => {
    // CT_PlotArea's axis elements live in a
    // `<xsd:choice maxOccurs="unbounded">` group — their relative
    // order is unconstrained. Neither catAx-first nor valAx-first
    // should trigger a violation.
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-child-out-of-order")).toBe(false);
  });
});

describe("ooxml-validator / chart — classic required children", () => {
  it("flags a barChart missing the second <c:axId>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-missing-required-child")).toBe(true);
  });

  it("flags a bar3DChart with only 2 <c:axId> children (needs 3)", async () => {
    const plotArea = `<c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:bar3DChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-missing-required-child")).toBe(true);
  });

  it("flags a barChart with 3 <c:axId> children (max 2)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-wrong-child-count")).toBe(true);
  });

  it("flags an axis missing <c:crossAx>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("crossAx")
      )
    ).toBe(true);
  });
});

describe("ooxml-validator / chart — classic enum values", () => {
  it("flags <c:barDir val='foo'> with an invalid enum value", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="foo"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-invalid-enum-value")).toBe(true);
  });

  it("accepts <c:barDir val='col'>", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-invalid-enum-value")).toBe(false);
  });

  it("flags <c:orientation val='reversed'> (valid values are minMax / maxMin)", async () => {
    const plotArea = minimalBarChartBody({
      extraAxes: `<c:valAx>
        <c:axId val="3"/>
        <c:scaling><c:orientation val="reversed"/></c:scaling>
        <c:axPos val="r"/>
        <c:crossAx val="1"/>
      </c:valAx>`
    });
    const report = await validateXlsxBuffer(buildClassicChartPackage(plotArea), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-invalid-enum-value")).toBe(true);
  });
});

describe("ooxml-validator / chart — classic numeric ranges", () => {
  it("flags <c:holeSize val='100'> (schema caps at 90)", async () => {
    const plotArea = `<c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:holeSize val="100"/>
      </c:doughnutChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-value-out-of-range")).toBe(true);
  });

  it("flags <c:firstSliceAng val='-10'> (schema range 0-360)", async () => {
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:firstSliceAng val="-10"/>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-value-out-of-range")).toBe(true);
  });

  it("flags <c:overlap val='200'> (schema range -100..100)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:overlap val="200"/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-value-out-of-range")).toBe(true);
  });

  it("accepts <c:holeSize val='60'>", async () => {
    const plotArea = `<c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:holeSize val="60"/>
      </c:doughnutChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-value-out-of-range")).toBe(false);
  });
});

describe("ooxml-validator / chart — classic CT_DLbl choice exclusivity", () => {
  it("flags a <c:dLbl> that has both <c:delete> and a display-flag child", async () => {
    // CT_DLbl is `idx, choice(delete | (layout…separator)), extLst?`.
    // A dLbl that emits both branches simultaneously is schema-invalid
    // — some Excel builds strip the label wholesale.
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:dLbls>
            <c:dLbl>
              <c:idx val="0"/>
              <c:delete val="1"/>
              <c:showVal val="1"/>
            </c:dLbl>
          </c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("choice")
      )
    ).toBe(true);
  });

  it("accepts a <c:dLbl> with only <c:delete val='1'/>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:dLbls>
            <c:dLbl>
              <c:idx val="0"/>
              <c:delete val="1"/>
            </c:dLbl>
          </c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("choice")
      )
    ).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Context-aware classic chart rules
//
// These tests exercise the per-chart-type restrictions Excel applies
// at read time — the constraints the generic schema tables cannot
// express because the same child element gets validated differently
// based on its enclosing chart type.
// -----------------------------------------------------------------------------

describe("ooxml-validator / chart — per-chart-type c:dLblPos allow-list", () => {
  it("flags <c:dLblPos> inside a <c:doughnutChart> series (any value)", async () => {
    // Doughnut series `c:dLbls` may not contain `c:dLblPos` at all —
    // Excel's reader strips the entire `drawing*.xml` part when any
    // value is present, even `bestFit`.
    const plotArea = `<c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="bestFit"/>
            <c:showVal val="1"/>
          </c:dLbls>
        </c:ser>
      </c:doughnutChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chart-invalid-enum-value" &&
          p.message.includes("doughnutChart") &&
          p.message.includes("does not accept")
      )
    ).toBe(true);
  });

  it("flags <c:dLblPos val='t'> inside a <c:barChart> (bar only allows ctr/inBase/inEnd/outEnd)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls>
            <c:dLbl>
              <c:idx val="0"/>
              <c:dLblPos val="t"/>
              <c:showVal val="1"/>
            </c:dLbl>
          </c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("barChart")
      )
    ).toBe(true);
  });

  it("flags <c:dLblPos val='outEnd'> inside a <c:lineChart> (line only allows ctr/l/r/t/b)", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showVal val="1"/>
          </c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("lineChart")
      )
    ).toBe(true);
  });

  it("accepts <c:dLblPos val='outEnd'> inside a <c:barChart>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="outEnd"/>
            <c:showVal val="1"/>
          </c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("barChart")
      )
    ).toBe(false);
  });

  it("accepts <c:dLblPos val='bestFit'> inside a <c:pieChart>", async () => {
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls>
            <c:dLblPos val="bestFit"/>
            <c:showPercent val="1"/>
          </c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("pieChart")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — forbidden children per chart type", () => {
  it("flags <c:trendline> inside a pie-family series", async () => {
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:trendline><c:trendlineType val="linear"/></c:trendline>
        </c:ser>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-forbidden-child" && p.message.includes("trendline")
      )
    ).toBe(true);
  });

  it("flags <c:errBars> inside a doughnut series", async () => {
    const plotArea = `<c:plotArea>
      <c:doughnutChart>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errBarType val="both"/>
            <c:errValType val="fixedVal"/>
            <c:val val="1"/>
          </c:errBars>
        </c:ser>
      </c:doughnutChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-forbidden-child" && p.message.includes("errBars"))
    ).toBe(true);
  });

  it("flags a surface chart with group-level <c:dLbls>", async () => {
    const plotArea = `<c:plotArea>
      <c:surfaceChart>
        <c:wireframe val="0"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:dLbls><c:showVal val="1"/></c:dLbls>
        <c:axId val="1"/><c:axId val="2"/><c:axId val="3"/>
      </c:surfaceChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-forbidden-child" && p.message.includes("dLbls"))
    ).toBe(true);
  });

  it("accepts <c:trendline> inside a line series", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:trendline><c:trendlineType val="linear"/></c:trendline>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-forbidden-child")).toBe(false);
  });
});

describe("ooxml-validator / chart — errBars cardinality + direction", () => {
  it("flags a scatter series with 3 <c:errBars> (schema max 2)", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
          <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
          <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-wrong-child-count" && p.message.includes("errBars")
      )
    ).toBe(true);
  });

  it("flags two <c:errBars> with the same <c:errDir> on a scatter series", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
          <c:errBars><c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="stdDev"/><c:val val="2"/></c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-duplicate-errBars-direction")).toBe(true);
  });

  it("accepts two <c:errBars> on a scatter series with distinct x/y directions", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
          <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-duplicate-errBars-direction")).toBe(false);
    expect(report.problems.some(p => p.kind === "chart-wrong-child-count")).toBe(false);
  });

  it("flags a line series with 2 <c:errBars> (non-scatter/bubble caps at 1)", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:val val="1"/></c:errBars>
          <c:errBars><c:errDir val="y"/><c:errBarType val="plus"/><c:errValType val="stdDev"/><c:val val="2"/></c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-wrong-child-count" && p.message.includes("errBars")
      )
    ).toBe(true);
  });
});

describe("ooxml-validator / chart — series count + idx/order uniqueness", () => {
  it("flags a stockChart with only 2 series (schema requires 3 or 4)", async () => {
    const plotArea = `<c:plotArea>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:ser><c:idx val="1"/><c:order val="1"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("stockChart")
      )
    ).toBe(true);
  });

  it("flags a stockChart with 5 series (schema max 4)", async () => {
    const plotArea = `<c:plotArea>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:ser><c:idx val="1"/><c:order val="1"/></c:ser>
        <c:ser><c:idx val="2"/><c:order val="2"/></c:ser>
        <c:ser><c:idx val="3"/><c:order val="3"/></c:ser>
        <c:ser><c:idx val="4"/><c:order val="4"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-wrong-child-count" && p.message.includes("stockChart")
      )
    ).toBe(true);
  });

  it("accepts a stockChart with 3 series (HLC)", async () => {
    const plotArea = `<c:plotArea>
      <c:stockChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:ser><c:idx val="1"/><c:order val="1"/></c:ser>
        <c:ser><c:idx val="2"/><c:order val="2"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:stockChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.message.includes("stockChart"))).toBe(false);
  });

  it("flags a barChart with 0 series (schema requires ≥1)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("<c:ser>")
      )
    ).toBe(true);
  });

  it("flags two series in the same group sharing <c:idx val>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:ser><c:idx val="0"/><c:order val="1"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-duplicate-series-idx")).toBe(true);
  });

  it("flags two series in the same group sharing <c:order val>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:ser><c:idx val="1"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-duplicate-series-order")).toBe(true);
  });
});

describe("ooxml-validator / chart — data reference structure", () => {
  it("flags a <c:numRef> missing the required <c:f>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
                <c:pt idx="1"><c:v>2</c:v></c:pt>
                <c:pt idx="2"><c:v>3</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("<c:numRef>")
      )
    ).toBe(true);
  });

  it("flags a <c:pt idx='5'> when <c:ptCount val='3'> declares only 3 points", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$A$1:$A$3</c:f>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="5"><c:v>99</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-pt-idx-out-of-range")).toBe(true);
  });

  it("flags a <c:pt> inside <c:numCache> missing <c:v>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$A$1:$A$1</c:f>
              <c:numCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"/>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("<c:v>")
      )
    ).toBe(true);
  });

  it("accepts a well-formed <c:numRef> with <c:f> and consistent cache", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$A$1:$A$3</c:f>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>30</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chart-missing-required-child" &&
          (p.message.includes("numRef") || p.message.includes("<c:v>"))
      )
    ).toBe(false);
    expect(report.problems.some(p => p.kind === "chart-pt-idx-out-of-range")).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Chart-type and series child-order rules
//
// Every CT_*Chart and CT_*Ser element has a canonical child
// sequence per ECMA-376 Part 1 §21.2.x. Third-party validators and
// LibreOffice strict mode refuse out-of-order output; Excel
// tolerates most mis-orderings but flags the file with "Repaired
// Records: Drawing" on open.
// -----------------------------------------------------------------------------

describe("ooxml-validator / chart — chart-type child order", () => {
  it("flags <c:barChart> with <c:axId> before <c:barDir>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:axId val="1"/><c:axId val="2"/>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("barChart")
      )
    ).toBe(true);
  });

  it("flags <c:lineChart> with <c:axId> before <c:grouping>", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:axId val="1"/><c:axId val="2"/>
        <c:grouping val="standard"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:lineChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("lineChart")
      )
    ).toBe(true);
  });

  it("flags <c:doughnutChart> with <c:holeSize> before <c:firstSliceAng>", async () => {
    const plotArea = `<c:plotArea>
      <c:doughnutChart>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:holeSize val="60"/>
        <c:firstSliceAng val="30"/>
      </c:doughnutChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("doughnutChart")
      )
    ).toBe(true);
  });

  it("flags <c:ofPieChart> with <c:splitPos> before <c:splitType>", async () => {
    const plotArea = `<c:plotArea>
      <c:ofPieChart>
        <c:ofPieType val="pie"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:splitPos val="10"/>
        <c:splitType val="pos"/>
      </c:ofPieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("ofPieChart")
      )
    ).toBe(true);
  });

  it("accepts <c:ofPieChart> with canonical child order", async () => {
    const plotArea = `<c:plotArea>
      <c:ofPieChart>
        <c:ofPieType val="pie"/>
        <c:varyColors val="1"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:splitType val="pos"/>
        <c:splitPos val="10"/>
      </c:ofPieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("ofPieChart")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — series child order per type", () => {
  it("flags <c:barSer> with <c:cat> before <c:dLbls>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:cat><c:strRef><c:f>Sheet1!$A$1:$A$3</c:f></c:strRef></c:cat>
          <c:dLbls><c:showVal val="1"/></c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("CT_BarSer")
      )
    ).toBe(true);
  });

  it("flags <c:scatterSer> with <c:yVal> before <c:xVal>", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$3</c:f></c:numRef></c:yVal>
          <c:xVal><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:xVal>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("CT_ScatterSer")
      )
    ).toBe(true);
  });

  it("flags <c:pieSer> with <c:val> before <c:dLbls>", async () => {
    const plotArea = `<c:plotArea>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:val>
          <c:dLbls><c:showPercent val="1"/></c:dLbls>
        </c:ser>
      </c:pieChart>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("CT_PieSer")
      )
    ).toBe(true);
  });

  it("accepts a <c:barSer> with canonical child order", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("CT_BarSer")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — axId / crossAx resolution", () => {
  it("flags a chart whose series <c:axId> does not match any axis", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="99"/><c:axId val="100"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-axid-unresolved" && p.message.includes("99"))
    ).toBe(true);
  });

  it("flags an axis whose <c:crossAx> does not match any axis", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="999"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-axid-unresolved" && p.message.includes("crossAx"))
    ).toBe(true);
  });

  it("accepts a chart whose axId / crossAx references all resolve", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-axid-unresolved")).toBe(false);
  });
});

describe("ooxml-validator / chart — errBars conditional children", () => {
  it("flags <c:errBars> with errValType='fixedVal' missing <c:val>", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="fixedVal"/>
          </c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chart-missing-required-child" &&
          p.message.includes("errValType") &&
          p.message.includes("fixedVal")
      )
    ).toBe(true);
  });

  it("flags <c:errBars> with errValType='cust' missing <c:plus> or <c:minus>", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="cust"/>
            <c:plus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:plus>
          </c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("cust")
      )
    ).toBe(true);
  });

  it("flags <c:errBars> with errValType='cust' containing a forbidden <c:val>", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="cust"/>
            <c:plus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:plus>
            <c:minus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:minus>
            <c:val val="1"/>
          </c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-forbidden-child" && p.message.includes("cust"))
    ).toBe(true);
  });

  it("flags <c:errBars> with errValType='percentage' containing <c:plus>", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="percentage"/>
            <c:plus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:plus>
            <c:val val="15"/>
          </c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-forbidden-child" && p.message.includes("percentage")
      )
    ).toBe(true);
  });

  it("accepts a well-formed cust errBars with plus + minus", async () => {
    const plotArea = `<c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="marker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:errBars>
            <c:errDir val="y"/>
            <c:errBarType val="both"/>
            <c:errValType val="cust"/>
            <c:plus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:plus>
            <c:minus><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:minus>
          </c:errBars>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:scatterChart>
      <c:valAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:valAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p =>
          (p.kind === "chart-missing-required-child" || p.kind === "chart-forbidden-child") &&
          p.message.includes("errValType")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — CT_Tx choice exclusivity", () => {
  it("flags a <c:tx> containing both <c:strRef> and <c:rich>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>
            <c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef>
            <c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>X</a:t></a:r></a:p></c:rich>
          </c:tx>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p =>
          p.kind === "chart-child-out-of-order" &&
          p.message.includes("<c:tx>") &&
          p.message.includes("branches")
      )
    ).toBe(true);
  });

  it("accepts a <c:tx> with a single branch (<c:rich>)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>
            <c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Series 1</a:t></a:r></a:p></c:rich>
          </c:tx>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-child-out-of-order" && p.message.includes("<c:tx>")
      )
    ).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// DrawingML / numFmt / text-body sub-schemas — `<a:srgbClr>` hex
// format, `<a:schemeClr>` enum, `<c:numFmt formatCode>` required,
// `<c:rich>` structure.
// -----------------------------------------------------------------------------

describe("ooxml-validator / chart — <a:srgbClr> hex format", () => {
  it("flags <a:srgbClr val='GG0000'> (non-hex character)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:srgbClr val="GG0000"/></a:solidFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("srgbClr")
      )
    ).toBe(true);
  });

  it("flags <a:srgbClr val='1234'> (wrong length)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:srgbClr val="1234"/></a:solidFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-invalid-enum-value" && p.message.includes("1234"))
    ).toBe(true);
  });

  it("accepts <a:srgbClr val='4472C4'> (canonical 6-digit hex)", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(minimalBarChartBody({})), {
      includeWarnings: true
    });
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("srgbClr")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — <a:schemeClr val> theme slot enum", () => {
  it("flags <a:schemeClr val='accent9'> (only accent1..6 exist)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:schemeClr val="accent9"/></a:solidFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("schemeClr")
      )
    ).toBe(true);
  });

  it("accepts <a:schemeClr val='accent1'>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-invalid-enum-value" && p.message.includes("schemeClr")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — <c:numFmt formatCode> required", () => {
  it("flags <c:numFmt> with no formatCode attribute", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls><c:numFmt sourceLinked="0"/><c:showVal val="1"/></c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("formatCode")
      )
    ).toBe(true);
  });

  it("accepts <c:numFmt formatCode='$#,##0'>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:dLbls><c:numFmt formatCode="$#,##0"/><c:showVal val="1"/></c:dLbls>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("formatCode")
      )
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — <c:rich> text-body structure", () => {
  it("flags <c:rich> missing <a:bodyPr>", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>
            <c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>X</a:t></a:r></a:p></c:rich>
          </c:tx>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("<a:bodyPr>")
      )
    ).toBe(true);
  });

  it("flags <c:rich> with no <a:p> paragraphs", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>
            <c:rich><a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></c:rich>
          </c:tx>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-missing-required-child" && p.message.includes("<a:p>")
      )
    ).toBe(true);
  });
});

describe("ooxml-validator / chart — series child whitelist", () => {
  it("flags <c:bubbleSize> inside a <c:barChart> series", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:bubbleSize><c:numRef><c:f>Sheet1!$A$1:$A$3</c:f></c:numRef></c:bubbleSize>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-forbidden-child" && p.message.includes("bubbleSize")
      )
    ).toBe(true);
  });

  it("flags <c:explosion> inside a <c:lineChart> series", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:explosion val="15"/>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-forbidden-child" && p.message.includes("explosion")
      )
    ).toBe(true);
  });

  it("flags <c:smooth> inside a <c:barChart> series", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:smooth val="1"/>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-forbidden-child" && p.message.includes("smooth"))
    ).toBe(true);
  });

  it("accepts <c:smooth> inside a <c:lineChart> series", async () => {
    const plotArea = `<c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:smooth val="1"/>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(p => p.kind === "chart-forbidden-child" && p.message.includes("smooth"))
    ).toBe(false);
  });
});

describe("ooxml-validator / chart — axis units & tick skips", () => {
  it("flags <c:majorUnit val='0'> (schema requires > 0)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="l"/>
        <c:crossAx val="1"/>
        <c:majorUnit val="0"/>
      </c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-value-out-of-range" && p.message.includes("majorUnit")
      )
    ).toBe(true);
  });

  it("flags <c:tickLblSkip val='0'> (schema requires >= 1)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
        <c:tickLblSkip val="0"/>
      </c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-value-out-of-range" && p.message.includes("tickLblSkip")
      )
    ).toBe(true);
  });

  it("flags <a:alpha val='150000'> (schema caps at 100000)", async () => {
    const plotArea = `<c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:srgbClr val="4472C4"><a:alpha val="150000"/></a:srgbClr></a:solidFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>`;
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(`${plotArea}<c:plotVisOnly val="1"/>`),
      { includeWarnings: true }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-value-out-of-range" && p.message.includes("alpha")
      )
    ).toBe(true);
  });
});

// -----------------------------------------------------------------------------
// Cross-part references
//
// Checks that span the chart XML boundary and validate links into
// sibling parts (theme1.xml for scheme colours, workbook.xml for
// defined names). These require the helper to swap in a theme /
// workbook that actually has what the chart references.
// -----------------------------------------------------------------------------

/** Build a chart package with a custom theme1.xml body. */
function buildClassicChartPackageWithTheme(chartInner: string, themeXml: string): Uint8Array {
  const parts = baseParts();
  parts["xl/theme/theme1.xml"] = themeXml;
  parts["xl/drawings/drawing1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
  parts["xl/charts/chart1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>${chartInner}</c:chart>
</c:chartSpace>`;
  parts["xl/drawings/_rels/drawing1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
  parts["xl/worksheets/_rels/sheet1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
  parts["xl/worksheets/sheet1.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`;
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
      partName: "/xl/charts/chart1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    },
    {
      partName: "/xl/drawings/drawing1.xml",
      contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
    }
  ]);
  return buildPackage(parts);
}

/** Full Office-default theme with all 12 clrScheme slots. */
function fullTheme(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;
}

/** Theme with just accent1 — used to verify slot-resolution flags missing slots. */
function partialTheme(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Partial">
  <a:themeElements>
    <a:clrScheme name="Partial">
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
    </a:clrScheme>
  </a:themeElements>
</a:theme>`;
}

describe("ooxml-validator / chart — theme scheme-colour slot resolution", () => {
  const chartWithSchemeClr = (val: string): string => `<c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:solidFill><a:schemeClr val="${val}"/></a:solidFill>
        </c:spPr>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
  </c:plotArea><c:plotVisOnly val="1"/>`;

  it("flags <a:schemeClr val='accent2'> when theme only declares accent1", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackageWithTheme(chartWithSchemeClr("accent2"), partialTheme()),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-theme-missing-schemeClr-slot")).toBe(true);
  });

  it("accepts <a:schemeClr val='accent1'> when theme declares it", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackageWithTheme(chartWithSchemeClr("accent1"), fullTheme()),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-theme-missing-schemeClr-slot")).toBe(false);
  });

  it("accepts workbook-facing aliases (bg1 → lt1, tx1 → dk1) via clrMap default", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackageWithTheme(chartWithSchemeClr("bg1"), fullTheme()),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-theme-missing-schemeClr-slot")).toBe(false);
  });

  it("always accepts <a:schemeClr val='phClr'> (placeholder)", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackageWithTheme(chartWithSchemeClr("phClr"), partialTheme()),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-theme-missing-schemeClr-slot")).toBe(false);
  });
});

describe("ooxml-validator / chart — <c:f> formula syntax", () => {
  const chartWithFormula = (f: string): string => `<c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:val>
          <c:numRef>
            <c:f>${f}</c:f>
          </c:numRef>
        </c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
  </c:plotArea><c:plotVisOnly val="1"/>`;

  it("flags an empty <c:f/>", async () => {
    const report = await validateXlsxBuffer(buildClassicChartPackage(chartWithFormula("")), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(true);
  });

  it("flags an unbalanced formula like 'Sheet1!(A1:B10'", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("Sheet1!(A1:B10")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(true);
  });

  it("accepts a canonical qualified range 'Sales!$B$2:$B$7'", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("Sales!$B$2:$B$7")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts a structured table reference 'Transactions[Revenue]'", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("Transactions[Revenue]")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts a structured table ref with specifier 'Table1[[#Headers],[Col]]'", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("Table1[[#Headers],[Col]]")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts a quoted-sheet reference 'Sales Data'!$B$2:$B$7", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("'Sales Data'!$B$2:$B$7")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts the reserved '_xlchart.v1.0' hidden name", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("_xlchart.v1.0")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts sheet-qualified '_xl' reserved names ('Sheet1!_xlchart.v1.0')", async () => {
    // Excel sometimes stores the hidden-name reference qualified by
    // the owning sheet — e.g. `Sales!_xlchart.v1.0`. Treat as
    // reserved on either side of the `!`.
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("Sheet1!_xlchart.v1.0")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
    expect(report.problems.some(p => p.kind === "chart-f-undefined-name")).toBe(false);
  });

  it("accepts a multi-range scatter formula wrapped in parens", async () => {
    // `(Sheet1!$A$1:$A$5,Sheet1!$C$1:$C$5)` is the canonical
    // multi-range form scatter / combo charts emit when a single
    // series draws data from two discontiguous ranges.
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("(Sheet1!$A$1:$A$5,Sheet1!$C$1:$C$5)")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });

  it("accepts a function-call formula ('SUM(A1:A10)')", async () => {
    const report = await validateXlsxBuffer(
      buildClassicChartPackage(chartWithFormula("SUM(A1:A10)")),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-invalid-syntax")).toBe(false);
  });
});

describe("ooxml-validator / chart — <c:f> defined-name resolution", () => {
  /** Build a package where the workbook declares the given defined names. */
  function buildWithDefinedNames(chartF: string, definedNames: string[]): Uint8Array {
    const parts = baseParts();
    const names = definedNames
      .map(n => `<definedName name="${n}">Sheet1!$A$1:$A$3</definedName>`)
      .join("");
    parts["xl/workbook.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
  <definedNames>${names}</definedNames>
</workbook>`;
    parts["xl/drawings/drawing1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;
    parts["xl/charts/chart1.xml"] = `<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:plotArea>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:ser>
        <c:idx val="0"/><c:order val="0"/>
        <c:val><c:numRef><c:f>${chartF}</c:f></c:numRef></c:val>
      </c:ser>
      <c:axId val="1"/><c:axId val="2"/>
    </c:barChart>
    <c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
    <c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
  </c:plotArea><c:plotVisOnly val="1"/></c:chart>
</c:chartSpace>`;
    parts["xl/drawings/_rels/drawing1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`;
    parts["xl/worksheets/_rels/sheet1.xml.rels"] = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
    parts["xl/worksheets/sheet1.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData/>
  <drawing r:id="rId1"/>
</worksheet>`;
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
        partName: "/xl/charts/chart1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
      },
      {
        partName: "/xl/drawings/drawing1.xml",
        contentType: "application/vnd.openxmlformats-officedocument.drawing+xml"
      }
    ]);
    return buildPackage(parts);
  }

  it("flags <c:f>UndefinedName</c:f> when no matching definedName exists", async () => {
    const report = await validateXlsxBuffer(buildWithDefinedNames("UndefinedName", []), {
      includeWarnings: true
    });
    expect(
      report.problems.some(
        p => p.kind === "chart-f-undefined-name" && p.message.includes("UndefinedName")
      )
    ).toBe(true);
  });

  it("accepts <c:f>Sales</c:f> when workbook declares a matching <definedName>", async () => {
    const report = await validateXlsxBuffer(buildWithDefinedNames("Sales", ["Sales"]), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-f-undefined-name")).toBe(false);
  });

  it("flags <c:f>Sheet1!Missing</c:f> when the sheet-scoped defined name is absent", async () => {
    const report = await validateXlsxBuffer(
      buildWithDefinedNames("Sheet1!Missing", ["OtherName"]),
      {
        includeWarnings: true
      }
    );
    expect(
      report.problems.some(
        p => p.kind === "chart-f-undefined-name" && p.message.includes("Missing")
      )
    ).toBe(true);
  });

  it("accepts <c:f>Sheet1!LocalName</c:f> when the name is declared", async () => {
    const report = await validateXlsxBuffer(
      buildWithDefinedNames("Sheet1!LocalName", ["LocalName"]),
      { includeWarnings: true }
    );
    expect(report.problems.some(p => p.kind === "chart-f-undefined-name")).toBe(false);
  });

  it("accepts reserved built-ins like '_xlnm.Print_Area' without a declared <definedName>", async () => {
    const report = await validateXlsxBuffer(buildWithDefinedNames("_xlnm.Print_Area", []), {
      includeWarnings: true
    });
    expect(report.problems.some(p => p.kind === "chart-f-undefined-name")).toBe(false);
  });
});
