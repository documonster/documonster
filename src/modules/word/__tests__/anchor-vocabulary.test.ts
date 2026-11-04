/**
 * The anchor and chart enumerations (`ST_RelFromH/V`, `ST_AlignH/V`,
 * `ST_WrapText`, `ST_LegendPos`, `ST_DLblPos`) are checked where they are
 * passed and where they are written, and a read keeps what it cannot type.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { Build, Document, Io, Package, Validation } from "../index";
import type { ChartContent, DocxDocument, FloatingImage } from "../types";

const PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);
const decoder = new TextDecoder();

function refused(label: string, run: () => unknown, field: string): void {
  let error: unknown;
  try {
    run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  expect((error as DocxArgumentError).field, label).toBe(field);
}

async function documentXml(doc: DocxDocument): Promise<string> {
  return decoder.decode((await Package.readPart(await Io.toBuffer(doc), "word/document.xml"))!);
}

describe("anchor enumerations", () => {
  it("are refused where they are passed", () => {
    const d = Document.create();
    const image = (options: object) => () =>
      Document.addFloatingImage(d, PNG, "png", 9525, 9525, options as never);
    refused(
      "h relativeTo",
      image({ horizontalPosition: { relativeTo: "bogus" } }),
      "options.horizontalPosition.relativeTo"
    );
    refused(
      "h align",
      image({ horizontalPosition: { align: "bogus" } }),
      "options.horizontalPosition.align"
    );
    refused(
      "v relativeTo",
      image({ verticalPosition: { relativeTo: "bogus" } }),
      "options.verticalPosition.relativeTo"
    );
    refused(
      "v align",
      image({ verticalPosition: { align: "bogus" } }),
      "options.verticalPosition.align"
    );
    refused("wrap string", image({ wrap: "square" }), "options.wrap");
    refused("wrap style", image({ wrap: { style: "bogus" } }), "options.wrap.style");
    refused("wrap side", image({ wrap: { style: "square", side: "bogus" } }), "options.wrap.side");
    refused(
      "Build.floatingImage",
      () =>
        Build.floatingImage({ rId: "r", width: 1, height: 1, wrap: { style: "bogus" as never } }),
      "options.wrap.style"
    );
    refused(
      "Build.shape",
      () =>
        Build.shape({
          shapeType: "rect",
          width: 1,
          height: 1,
          horizontalPosition: { relativeTo: "bogus" as never }
        }),
      "options.horizontalPosition.relativeTo"
    );
    refused(
      "Build.chart legend",
      () => Build.chart({ type: "bar", series: [], legend: "bogus" as never }),
      "options.legend"
    );
    expect(Document.getContentCount(d)).toBe(0);
  });

  it("are refused when a hand-built model is written", async () => {
    const d = Document.create();
    const image = Build.floatingImage({ rId: "r", width: 1, height: 1 });
    Document.addContent(d, {
      ...image,
      horizontalPosition: { relativeTo: "bogus" }
    } as unknown as FloatingImage);
    await expect(Io.toBuffer(d)).rejects.toThrow(DocxArgumentError);
  });

  it("an unknown value read from a file is kept, not typed", async () => {
    const d = Document.create();
    Document.addFloatingImage(d, PNG, "png", 9525, 9525, {
      horizontalPosition: { relativeTo: "page", align: "center" },
      wrap: { style: "square", side: "left" }
    });
    const xml = await documentXml(d);
    const bogus = xml
      .replace('relativeFrom="page"', 'relativeFrom="bogusH"')
      .replace('wrapText="left"', 'wrapText="bogusW"');
    expect(bogus).not.toBe(xml);
    const bytes = await Io.toBuffer(d);
    const read = await Io.read(
      await Package.edit(bytes, [
        { type: "replacePartText", path: "word/document.xml", text: bogus }
      ])
    );
    const img = read.body.find((b): b is FloatingImage => b.type === "floatingImage")!;
    expect(img.horizontalPosition?.relativeTo).toBeUndefined();
    expect(img.wrap?.side).toBeUndefined();
    const issues = Validation.validate(read).issues.filter(i => i.rule === "unrecognized-value");
    expect(issues.length).toBeGreaterThanOrEqual(2);
    const again = await documentXml(read);
    expect(again).toContain('relativeFrom="bogusH"');
    expect(again).toContain('wrapText="bogusW"');
  });
});

describe("chart enumerations", () => {
  const series = [{ name: "s", categories: ["a", "b"], values: [1, 2] }];

  async function roundTrip(chart: ChartContent): Promise<ChartContent> {
    const d = Document.create();
    Document.addContent(d, chart);
    const read = await Io.read(await Io.toBuffer(d));
    return read.body.find((b): b is ChartContent => b.type === "chart")!;
  }

  it("reads a data-label position back as it was set", async () => {
    for (const position of ["outsideEnd", "center", "insideEnd", "bestFit"] as const) {
      const chart = Build.chart({ type: "bar", series: [{ ...series[0]!, showDataLabels: true }] });
      const withLabels: ChartContent = {
        ...chart,
        chart: { ...chart.chart, dataLabels: { showValue: true, position } }
      };
      expect((await roundTrip(withLabels)).chart.dataLabels?.position, position).toBe(position);
    }
  });

  it("keeps a chart without a legend without one", async () => {
    const read = await roundTrip(Build.chart({ type: "bar", series, legend: "none" }));
    expect(read.chart.legend).toBe("none");
    const again = await roundTrip(read);
    expect(again.chart.legend).toBe("none");
  });
});
