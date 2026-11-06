import { Draw, DrawArgumentError } from "@draw/index";
import { ChartClassic, ExcelChartOptionsError } from "@excel/chart/index";
/**
 * One background check (`@utils/background-check`): `Draw.toSvg` / `Draw.toRgba`, the chart renderers
 * (`ChartRenderOptions.background`) and Word's page renderer (`RenderOptions.background`) all
 * call `checkBackgroundWith`, each with its own error factory — so each throws its own class,
 * naming the same field, in the same words.
 */
import { checkBackgroundWith } from "@utils/background-check";
import { Document, DocxArgumentError, Layout } from "@word/index";
import { describe, expect, it } from "vitest";

function thrown(run: () => unknown): Error & { field?: string } {
  try {
    run();
  } catch (error) {
    return error as Error & { field?: string };
  }
  throw new Error("expected a refusal");
}

/** The message after the call's own prefix. */
const words = (error: Error, where: string) => error.message.replace(`${where}: `, "");

const BAD: readonly unknown[] = [
  "not-a-colour",
  42,
  [1, 2],
  { r: 2, g: 0, b: 0, a: 1 },
  { r: 0, g: 0, b: 0 }
];

describe("checkBackgroundWith", () => {
  it("passes what Draw accepts and hands the caller's factory the field, the expectation and the value", () => {
    const calls: unknown[][] = [];
    const fail = (where: string, field: string, expected: string, value: unknown): never => {
      calls.push([where, field, expected, value]);
      throw new Error("refused");
    };
    for (const good of [
      "#fff",
      "rgb(1 2 3 / 0.5)",
      "transparent",
      "none",
      { r: 0, g: 1, b: 0.5, a: 1 }
    ]) {
      expect(() => checkBackgroundWith("X", "options.background", good, fail)).not.toThrow();
    }
    expect(() =>
      checkBackgroundWith("X", "options.background", { r: 0, g: 0, b: 0, a: 7 }, fail)
    ).toThrow();
    expect(calls).toEqual([["X", "options.background.a", "a number from 0 to 1", 7]]);
  });

  it("Draw, the chart renderers and Word refuse the same values with the same words, each in its own class", () => {
    const doc = Document.create();
    Document.addParagraph(doc, "x");
    const model = ChartClassic.createModel({
      type: "bar",
      series: [{ name: "S", categories: "Sheet1!$A$1:$A$2", values: "Sheet1!$B$1:$B$2" }]
    });
    for (const bad of BAD) {
      const draw = thrown(() =>
        Draw.toSvg({ width: 1, height: 1, children: [] }, { background: bad as never })
      );
      const chart = thrown(() => ChartClassic.toSvg(model, { background: bad as never }));
      const word = thrown(() => Layout.toSvg(doc, { background: bad as never }));
      expect(draw).toBeInstanceOf(DrawArgumentError);
      expect(chart).toBeInstanceOf(ExcelChartOptionsError);
      expect(word).toBeInstanceOf(DocxArgumentError);
      expect(chart.field).toBe(draw.field);
      expect(word.field).toBe(draw.field);
      expect(words(chart, "ChartClassic.toSvg"), String(bad)).toBe(words(draw, "Draw.toSvg"));
      expect(words(word, "Layout.toSvg"), String(bad)).toBe(words(draw, "Draw.toSvg"));
    }
  });
});
