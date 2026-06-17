import { extractAll } from "@archive/unzip/extract";
/**
 * Combo / axis matrix integration tests.
 *
 * Single home for cross-cutting axis behaviour that the per-preset
 * matrix in `chart-synthetic-corpus.integration.test.ts` cannot
 * exercise — log-scale value axis, multiple secondary axes, axis-id
 * uniqueness, scatter+line combos, and 3D combo groups.
 */
import { Chart, Workbook } from "@excel/index";
import type { WorkbookData } from "@excel/workbook-core";
import { getCharts } from "@excel/worksheet";
import { beforeAll, describe, expect, it } from "vitest";

import { expectValidXlsx } from "./helpers/expect-valid-xlsx";
import { buildComboAxisFixtures, type SyntheticFixture } from "./helpers/synthetic-fixtures";
import { entryText, type EntryMap } from "./helpers/zip-text";

let comboAxisFixtures: SyntheticFixture[];

beforeAll(async () => {
  comboAxisFixtures = await buildComboAxisFixtures();
});

async function loadFixture(bytes: Uint8Array): Promise<{
  wb: WorkbookData;
  out: Uint8Array;
  entries: EntryMap;
  chartXml: string;
}> {
  const wb = Workbook.create();
  await Workbook.loadXlsx(wb, bytes);
  const out = new Uint8Array(await Workbook.toXlsxBuffer(wb));
  const entries = await extractAll(out);
  return { wb, out, entries, chartXml: entryText(entries, "xl/charts/chart1.xml")! };
}

/**
 * Extract every `<c:axId val="..."/>` that lives inside a recognised
 * axis-definition wrapper (`c:catAx`, `c:valAx`, `c:dateAx`, `c:serAx`).
 * The "axis-id uniqueness" property we care about is that no two axis
 * **definitions** share the same id — the cross-references emitted via
 * `<c:crossAx>` reuse those ids on purpose.
 */
function extractAxisDefIds(chartXml: string): string[] {
  return [
    ...chartXml.matchAll(
      /<c:(?:catAx|valAx|dateAx|serAx)>[\s\S]*?<c:axId val="(\d+)"\/>[\s\S]*?<\/c:(?:catAx|valAx|dateAx|serAx)>/g
    )
  ].map(m => m[1]);
}

describe("Chart combo / axis matrix", () => {
  it("preserves stacked bar + secondary line combo through round-trip", async () => {
    const fixtures = comboAxisFixtures;
    const fixture = fixtures.find(f => f.id === "combo-stacked-secondary")!;
    expect(fixture).toBeDefined();
    const { out, chartXml } = await loadFixture(fixture.bytes);
    await expectValidXlsx(out);
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("<c:lineChart>");
    expect(chartXml).toMatch(/<c:grouping val="stacked"\/>/);
    // Secondary axis pair: at least 2 c:catAx and 2 c:valAx blocks.
    const catAxCount = (chartXml.match(/<c:catAx>/g) ?? []).length;
    const valAxCount = (chartXml.match(/<c:valAx>/g) ?? []).length;
    expect(catAxCount, "secondary catAx").toBeGreaterThanOrEqual(2);
    expect(valAxCount, "secondary valAx").toBeGreaterThanOrEqual(2);
  });

  it("preserves scatter + line combo with secondary axis through round-trip", async () => {
    const fixtures = comboAxisFixtures;
    const fixture = fixtures.find(f => f.id === "combo-scatter-line")!;
    const { out, chartXml } = await loadFixture(fixture.bytes);
    await expectValidXlsx(out);
    expect(chartXml).toContain("<c:scatterChart>");
    expect(chartXml).toContain("<c:lineChart>");
  });

  it("preserves logarithmic value axis (logBase=10) through round-trip", async () => {
    const fixtures = comboAxisFixtures;
    const fixture = fixtures.find(f => f.id === "axis-log")!;
    const { wb, out, chartXml } = await loadFixture(fixture.bytes);
    await expectValidXlsx(out);
    expect(chartXml).toMatch(/<c:logBase val="10"\/>/);
    expect(chartXml).toMatch(/<c:min val="1"\/>/);
    expect(chartXml).toMatch(/<c:max val="100"\/>/);

    const chart = getCharts(Workbook.getWorksheet(wb, "Data")!)[0];
    const valueAxis = Chart.chartModel(chart)?.chart.plotArea.axes.find(a => a.axisType === "val");
    expect(valueAxis, "value axis").toBeDefined();
    expect(valueAxis!.scaling?.logBase).toBe(10);
    expect(valueAxis!.scaling?.min).toBe(1);
    expect(valueAxis!.scaling?.max).toBe(100);
  });

  it("preserves three combo groups (bar primary, line+area secondary) through round-trip", async () => {
    const fixtures = comboAxisFixtures;
    const fixture = fixtures.find(f => f.id === "combo-three-groups-shared-secondary")!;
    const { out, chartXml } = await loadFixture(fixture.bytes);
    await expectValidXlsx(out);
    expect(chartXml).toContain("<c:barChart>");
    expect(chartXml).toContain("<c:lineChart>");
    expect(chartXml).toContain("<c:areaChart>");

    // axis ids must remain unique across the document (Excel rejects
    // duplicates outright). Collect every <c:axId val="N"/> within an
    // axis-def wrapper and assert each axis def has a unique id.
    const axisIds = extractAxisDefIds(chartXml);
    expect(axisIds.length).toBeGreaterThanOrEqual(3);
    expect(new Set(axisIds).size, `axis ids unique: ${axisIds.join(",")}`).toBe(axisIds.length);
  });

  it("preserves 3D bar + 2D line combo (with serAx for the 3D group) through round-trip", async () => {
    const fixtures = comboAxisFixtures;
    const fixture = fixtures.find(f => f.id === "combo-3d-bar-line")!;
    const { out, chartXml } = await loadFixture(fixture.bytes);
    await expectValidXlsx(out);
    expect(chartXml).toContain("<c:bar3DChart>");
    expect(chartXml).toContain("<c:lineChart>");
    // 3D bar groups require a serAx alongside catAx/valAx.
    expect(chartXml).toContain("<c:serAx>");
  });

  it("axis ids never collide between primary and secondary axes across all combo fixtures", async () => {
    const fixtures = comboAxisFixtures;
    for (const fixture of fixtures) {
      const { chartXml } = await loadFixture(fixture.bytes);
      const allAxisIds = [...chartXml.matchAll(/<c:axId val="(\d+)"\/>/g)].map(m => m[1]);
      // Every <c:axId> should appear an even number of times overall
      // (each axis id is referenced by both ends of the cross-reference
      // pair: from its own definition and from its sibling's
      // <c:crossAx> via <c:axId>). The unique-id property is what we
      // care about — a duplicated definition of the same id under two
      // different axis types crashes Excel.
      const axisDefBlocks = extractAxisDefIds(chartXml);
      expect(
        new Set(axisDefBlocks).size,
        `${fixture.id}: axis-def ids ${axisDefBlocks.join(",")}`
      ).toBe(axisDefBlocks.length);
      // Sanity: at least one axis def in non-pie variants.
      expect(allAxisIds.length, fixture.id).toBeGreaterThan(0);
    }
  });

  it("audit baseline passes for every combo/axis fixture", async () => {
    const fixtures = comboAxisFixtures;
    for (const fixture of fixtures) {
      const wb = Workbook.create();
      await Workbook.loadXlsx(wb, fixture.bytes);
      const out = new Uint8Array(await Workbook.toXlsxBuffer(wb));
      await expectValidXlsx(out, { label: fixture.id });
    }
  });
});
