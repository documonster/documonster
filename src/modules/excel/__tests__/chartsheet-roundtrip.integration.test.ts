/**
 * Chartsheet round-trip integration tests.
 *
 * Single home for chartsheet-specific fidelity: multiple chartsheets,
 * hidden state, page setup / margins / view settings, tab metadata,
 * and the chartsheet → drawing → chart relationship chain.
 *
 * Real-fixture round-trip lives in
 * `workbook-chart-roundtrip.integration.test.ts`. Other chart fidelity
 * dimensions live in `chart-mutation-roundtrip`,
 * `chart-combo-axis-matrix`, and `pivot-chart-roundtrip`.
 */

import { extractAll } from "@archive/unzip/extract";
import { installChartSupport } from "@excel/chart/install";
import { Workbook } from "@excel/workbook";
import { beforeAll, describe, expect, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";
import { buildChartsheetFixtures, type SyntheticFixture } from "./helpers/synthetic-fixtures";
import { entryText, loadRoundTrip } from "./helpers/zip-text";

installChartSupport();

let chartsheetFixtures: SyntheticFixture[];

beforeAll(async () => {
  chartsheetFixtures = await buildChartsheetFixtures();
});

describe("Chartsheet round-trip", () => {
  it("preserves multiple chartsheets (pie, funnel, combo, hidden) plus a Data worksheet", async () => {
    const [fixture] = chartsheetFixtures;
    const { wb, bytes, entries } = await loadRoundTrip(fixture.bytes);

    // High-level model: 1 worksheet + 4 chartsheets.
    expect(wb.worksheets.length).toBe(1);
    expect(wb.chartsheets.length).toBe(4);
    expect(wb.chartsheets.map(cs => cs.name).sort()).toEqual([
      "Combo Chart Sheet",
      "Funnel Chart Sheet",
      "Hidden Sheet",
      "Pie Chart Sheet"
    ]);

    // OOXML parts: 4 chartsheets + 1 worksheet.
    const chartsheetPaths = [...entries.keys()]
      .filter(p => /^xl\/chartsheets\/sheet\d+[.]xml$/.test(p))
      .sort();
    expect(chartsheetPaths.length).toBe(4);
    expect([...entries.keys()].filter(p => /^xl\/worksheets\/sheet\d+[.]xml$/.test(p)).length).toBe(
      1
    );

    // Package must validate cleanly.
    await expectValidXlsx(bytes);

    // Each chartsheet xml references a drawing relationship.
    for (const path of chartsheetPaths) {
      const xml = entryText(entries, path)!;
      expect(xml, `${path} drawing ref`).toMatch(/<drawing\s/);
      expect(xml, `${path} chartsheet root`).toContain("<chartsheet");
    }

    // Each chartsheet has its own drawing rels file pointing at a chart part.
    for (const path of chartsheetPaths) {
      const stem = path
        .split("/")
        .pop()!
        .replace(/[.]xml$/, "");
      const relsPath = `xl/chartsheets/_rels/${stem}.xml.rels`;
      const relsXml = entryText(entries, relsPath);
      expect(relsXml, relsPath).toBeDefined();
      expect(relsXml!).toContain("../drawings/");
    }
  });

  it("preserves hidden state, zoomScale, zoomToFit, pageSetup, pageMargins on the configured chartsheets", async () => {
    const [fixture] = chartsheetFixtures;
    const { wb, entries } = await loadRoundTrip(fixture.bytes);

    const pie = wb.chartsheets.find(cs => cs.name === "Pie Chart Sheet")!;
    // tabSelected omitted/false → undefined after round-trip (writer
    // skips emitting `<sheetView tabSelected="0"/>` since 0 is the
    // OOXML default).
    expect(pie.tabSelected ?? false).toBe(false);
    expect(pie.zoomScale).toBe(80);
    expect(pie.zoomToFit).toBe(true);
    expect(pie.pageMargins).toMatchObject({
      l: 0.7,
      r: 0.7,
      t: 0.75,
      b: 0.75,
      header: 0.3,
      footer: 0.3
    });
    expect(pie.pageSetup?.orientation).toBe("landscape");
    expect(pie.pageSetup?.paperSize).toBe(9);

    const hidden = wb.chartsheets.find(cs => cs.name === "Hidden Sheet")!;
    expect(hidden.state).toBe("hidden");

    // The XML of the Pie chartsheet should also carry the page settings.
    // Locate by index rather than name (write order matches addChartsheet
    // call order in the fixture, so sheet1 = Pie).
    const sheet1Xml = entryText(entries, "xl/chartsheets/sheet1.xml")!;
    expect(sheet1Xml).toMatch(/<pageMargins[^>]*left="0\.7"/);
    expect(sheet1Xml).toMatch(/<pageSetup[^>]*orientation="landscape"/);
    expect(sheet1Xml).toMatch(/<pageSetup[^>]*paperSize="9"/);
    expect(sheet1Xml).toMatch(/zoomScale="80"/);
    expect(sheet1Xml).toMatch(/zoomToFit="1"/);

    // The Hidden chartsheet's hidden state lives on workbook.xml's
    // <sheets><sheet state="hidden" ...> entry.
    const workbookXml = entryText(entries, "xl/workbook.xml")!;
    expect(workbookXml).toMatch(/name="Hidden Sheet"[^>]*state="hidden"/);
  });

  it("emits the right content types and chartsheet relationships for each sheet", async () => {
    const [fixture] = chartsheetFixtures;
    const { entries } = await loadRoundTrip(fixture.bytes);

    const contentTypes = entryText(entries, "[Content_Types].xml")!;
    for (let i = 1; i <= 4; i++) {
      expect(contentTypes, `chartsheet${i} content type`).toContain(
        `/xl/chartsheets/sheet${i}.xml`
      );
    }

    // The workbook rels must reference each chartsheet.
    const workbookRels = entryText(entries, "xl/_rels/workbook.xml.rels")!;
    for (let i = 1; i <= 4; i++) {
      expect(workbookRels, `workbook rels chartsheet${i}`).toContain(`chartsheets/sheet${i}.xml`);
    }
  });

  it("repeated round-trips remain stable for chartsheet-only workbooks", async () => {
    const [fixture] = chartsheetFixtures;
    let bytes = fixture.bytes;
    for (let i = 0; i < 3; i++) {
      const wb = new Workbook();
      await wb.xlsx.load(bytes);
      expect(wb.chartsheets.length, `pass ${i + 1}`).toBe(4);
      bytes = new Uint8Array(await wb.xlsx.writeBuffer());
      await expectValidXlsx(bytes, { label: `pass ${i + 1}` });
    }
  });

  it("a high-level mutation on a chartsheet's chart re-renders without dropping page metadata", async () => {
    const [fixture] = chartsheetFixtures;
    const wb = new Workbook();
    await wb.xlsx.load(fixture.bytes);
    const pie = wb.chartsheets.find(cs => cs.name === "Pie Chart Sheet")!;
    const chart = pie.chart;
    expect(chart, "pie chartsheet chart accessor").toBeDefined();
    chart!.title = "Mutated Pie Title";

    const out = new Uint8Array(await wb.xlsx.writeBuffer());
    const entries = await extractAll(out);
    await expectValidXlsx(out);

    const sheet1Xml = entryText(entries, "xl/chartsheets/sheet1.xml")!;
    expect(sheet1Xml, "page metadata still present").toMatch(/<pageSetup/);
    expect(sheet1Xml, "page metadata still present").toMatch(/<pageMargins/);

    const wb2 = new Workbook();
    await wb2.xlsx.load(out);
    const pie2 = wb2.chartsheets.find(cs => cs.name === "Pie Chart Sheet")!;
    expect(pie2.chart).toBeDefined();
    expect(pie2.chart!.title).toBe("Mutated Pie Title");
    expect(pie2.zoomScale).toBe(80);
  });
});
