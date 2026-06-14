/**
 * Pivot chart round-trip integration tests.
 *
 * Single home for pivot-specific chart fidelity: multiple pivot
 * caches, multi-value PivotTables, full dropZone option matrix,
 * `refreshOnOpen`, and the c14:pivotOptions extension envelope.
 */

import { extractAll } from "@archive/unzip/extract";
import { installChartSupport } from "@excel/chart/install";
import { Workbook } from "@excel/index";
import { beforeAll, describe, expect, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";
import { buildPivotChartFixtures, type SyntheticFixture } from "./helpers/synthetic-fixtures";
import { entryText, type EntryMap } from "./helpers/zip-text";

installChartSupport();

const PIVOT_OPT_EXT_URI = "{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}";

let pivotChartFixtures: SyntheticFixture[];

beforeAll(async () => {
  pivotChartFixtures = await buildPivotChartFixtures();
});

async function loadAndWrite(bytes: Uint8Array): Promise<{ bytes: Uint8Array; entries: EntryMap }> {
  const wb = Workbook.create();
  await Workbook.loadXlsx(wb, bytes);
  const out = new Uint8Array(await Workbook.toXlsxBuffer(wb));
  return { bytes: out, entries: await extractAll(out) };
}

describe("Pivot chart round-trip", () => {
  it("preserves a multi-value pivot chart with all dropZone options through round-trip", async () => {
    const fixtures = pivotChartFixtures;
    const fixture = fixtures.find(f => f.id === "pivot-chart-multi-value")!;
    expect(fixture).toBeDefined();
    const { bytes, entries } = await loadAndWrite(fixture.bytes);
    await expectValidXlsx(bytes);

    const chartPath = [...entries.keys()].find(p => /^xl\/charts\/chart\d+[.]xml$/.test(p))!;
    const chartXml = entryText(entries, chartPath)!;
    expect(chartXml, "pivot source").toContain("<c:pivotSource>");
    expect(chartXml, "c14 pivotOptions extension uri").toContain(PIVOT_OPT_EXT_URI);
    expect(chartXml, "c14 pivotOptions element").toContain("<c14:pivotOptions>");
    expect(chartXml).toContain('<c14:dropZonesVisible val="1"/>');
    expect(chartXml).toContain('<c14:dropZoneCategories val="1"/>');
    expect(chartXml).toContain('<c14:dropZoneData val="1"/>');
    expect(chartXml).toContain('<c14:dropZoneSeries val="1"/>');

    // refreshOnOpen lives on pivotCacheDefinition.xml, not on the chart.
    const cacheDefPath = [...entries.keys()].find(p =>
      /^xl\/pivotCache\/pivotCacheDefinition\d+[.]xml$/.test(p)
    )!;
    expect(cacheDefPath, "pivot cache definition part").toBeDefined();
    const cacheDefXml = entryText(entries, cacheDefPath)!;
    expect(cacheDefXml, "refreshOnLoad").toMatch(/refreshOnLoad="1"/);
  });

  it("preserves two independent pivot caches each driving a separate pivot chart", async () => {
    const fixtures = pivotChartFixtures;
    const fixture = fixtures.find(f => f.id === "pivot-chart-multi-cache")!;
    const { bytes, entries } = await loadAndWrite(fixture.bytes);
    await expectValidXlsx(bytes);

    // Two separate cache definitions and two pivot tables.
    const cacheDefs = [...entries.keys()].filter(p =>
      /^xl\/pivotCache\/pivotCacheDefinition\d+[.]xml$/.test(p)
    );
    const cacheRecs = [...entries.keys()].filter(p =>
      /^xl\/pivotCache\/pivotCacheRecords\d+[.]xml$/.test(p)
    );
    const pivots = [...entries.keys()].filter(p =>
      /^xl\/pivotTables\/pivotTable\d+[.]xml$/.test(p)
    );
    const charts = [...entries.keys()].filter(p => /^xl\/charts\/chart\d+[.]xml$/.test(p));

    expect(cacheDefs.length, "two cache defs").toBe(2);
    expect(cacheRecs.length, "two cache record parts").toBe(2);
    expect(pivots.length, "two pivot tables").toBe(2);
    expect(charts.length, "two pivot charts").toBe(2);

    // Each pivot chart references c:pivotSource.
    for (const path of charts) {
      const xml = entryText(entries, path)!;
      expect(xml, `${path} pivot source`).toContain("<c:pivotSource>");
    }

    // Workbook rels references both cache definitions.
    const workbookRels = entryText(entries, "xl/_rels/workbook.xml.rels")!;
    for (let i = 1; i <= 2; i++) {
      expect(workbookRels, `cache${i} rel`).toContain(`pivotCacheDefinition${i}.xml`);
    }
  });

  it("repeated load → write cycles keep the c14:pivotOptions extension stable", async () => {
    const fixtures = pivotChartFixtures;
    const fixture = fixtures.find(f => f.id === "pivot-chart-multi-value")!;
    let bytes = fixture.bytes;
    for (let i = 0; i < 3; i++) {
      const wb = Workbook.create();
      await Workbook.loadXlsx(wb, bytes);
      bytes = new Uint8Array(await Workbook.toXlsxBuffer(wb));
      const entries = await extractAll(bytes);
      const chartPath = [...entries.keys()].find(p => /^xl\/charts\/chart\d+[.]xml$/.test(p))!;
      const xml = entryText(entries, chartPath)!;
      expect(xml, `pass ${i + 1} pivotOptions uri`).toContain(PIVOT_OPT_EXT_URI);
      expect(xml, `pass ${i + 1} dropZonesVisible`).toContain('<c14:dropZonesVisible val="1"/>');
      await expectValidXlsx(bytes, { label: `pass ${i + 1}` });
    }
  });

  it("audit baseline passes for every pivot chart fixture", async () => {
    const fixtures = pivotChartFixtures;
    for (const fixture of fixtures) {
      const { bytes } = await loadAndWrite(fixture.bytes);
      await expectValidXlsx(bytes, { label: fixture.id });
    }
  });
});
