/**
 * A Word chart's scheme colours — and its automatic series colours, which are the theme's
 * accents — follow the document's theme (`doc.theme.colorScheme`), on every route that draws
 * a Word chart with the Excel engine: `Pdf.fromDocx`, `WordExcel.toChartSvg` and the
 * `renderChart` callback of `Layout.toSvg`.
 */
import { decompressPdfContent } from "@pdf/__tests__/test-helpers";
import { Pdf } from "@pdf/index";
import { WordExcel } from "@word/excel";
import { Layout } from "@word/index";
import type { Chart, DocumentTheme, DocxDocument } from "@word/types";
import { describe, expect, it } from "vitest";

/** The document theme with accent1 changed from the default `4472C4`. */
const THEME: DocumentTheme = {
  name: "Custom",
  colorScheme: {
    name: "Custom",
    colors: {
      dk1: "000000",
      lt1: "FFFFFF",
      dk2: "44546A",
      lt2: "E7E6E6",
      accent1: "C00000",
      accent2: "ED7D31",
      accent3: "A5A5A5",
      accent4: "FFC000",
      accent5: "5B9BD5",
      accent6: "70AD47",
      hlink: "0563C1",
      folHlink: "954F72"
    }
  },
  fontScheme: { name: "Office", majorFont: "Calibri Light", minorFont: "Calibri" }
};

const CHART: Chart = {
  type: "bar",
  title: "Sales",
  series: [{ name: "S", categories: ["A", "B"], values: [1, 2] }]
};

function doc(theme?: DocumentTheme): DocxDocument {
  return {
    body: [{ type: "chart", chart: CHART }],
    ...(theme ? { theme } : {})
  } as DocxDocument;
}

/** `C00000` as PDF fill operands. */
const ACCENT1_RG = /^0\.7529 0 0 rg$/m;
/** The default theme's accent1, `4472C4`. */
const DEFAULT_RG = /^0\.2667 0\.4471 0\.7686 rg$/m;

describe("Word charts follow the document's theme", () => {
  it("Pdf.fromDocx", async () => {
    const themed = decompressPdfContent(await Pdf.fromDocx(doc(THEME)));
    expect(themed).toMatch(ACCENT1_RG);
    expect(themed).not.toMatch(DEFAULT_RG);
    // With no theme of its own, the default Office theme.
    expect(decompressPdfContent(await Pdf.fromDocx(doc()))).toMatch(DEFAULT_RG);
  });

  it("WordExcel.toChartSvg takes the theme's colours", () => {
    expect(WordExcel.toChartSvg(CHART).toUpperCase()).toContain('FILL="#4472C4"');
    const themed = WordExcel.toChartSvg(CHART, { themeColors: THEME.colorScheme.colors });
    expect(themed.toUpperCase()).toContain('FILL="#C00000"');
    expect(() => WordExcel.toChartSvg(CHART, { themeColors: { accent1: "red" } } as never)).toThrow(
      /options\.themeColors\.accent1/
    );
  });

  it("Layout.toSvg hands its renderChart callback the document's theme colours", () => {
    const seen: unknown[] = [];
    Layout.toSvg(doc(THEME), {
      renderChart: (chart, context) => {
        seen.push(context?.themeColors);
        const source = chart.source as { type?: string; chart?: Chart } | undefined;
        return source?.type === "chart" && source.chart
          ? WordExcel.toChartSvg(source.chart, { themeColors: context?.themeColors })
          : undefined;
      }
    });
    expect(seen).toEqual([THEME.colorScheme.colors]);
  });
});
