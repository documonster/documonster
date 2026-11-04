/**
 * A write-time model check reports `field` as the exact path a caller follows
 * in the document (or, for the streaming writer, from its `element`
 * parameter) to reach the value it refused: walking the field on the object
 * the caller passed lands on the bad value itself. Before, the writer named
 * descriptive paths (`body[0].paragraph.alignment`, `table width type`) that
 * led nowhere.
 */

import { describe, it, expect } from "vitest";

import { Build, Document, DocxArgumentError, Io, Streaming } from "../index";
import type { BodyContent, DocxDocument } from "../index";

const BAD = "bogus";

/** Follow `field` (`a.b[0]["key"].c`) through objects, arrays and Maps. */
function walk(root: unknown, field: string): unknown {
  const steps = field.match(/[^.[\]]+|\[\d+\]|\["(?:[^"\\]|\\.)*"\]/g) ?? [];
  let at: unknown = root;
  for (const step of steps) {
    const key = step.startsWith("[")
      ? step.startsWith('["')
        ? (JSON.parse(step.slice(1, -1)) as string)
        : Number(step.slice(1, -1))
      : step;
    at =
      at instanceof Map ? at.get(key) : (at as Record<string | number, unknown> | undefined)?.[key];
  }
  return at;
}

async function refusedAt(doc: DocxDocument, bad: unknown): Promise<string> {
  const error = await Io.toBuffer(doc).catch((e: unknown) => e);
  expect(error).toBeInstanceOf(DocxArgumentError);
  const field = (error as DocxArgumentError).field;
  const reached = walk(doc, field);
  if (typeof bad === "number" && Number.isNaN(bad)) {
    expect(Number.isNaN(reached), field).toBe(true);
  } else {
    expect(reached, field).toBe(bad);
  }
  // The message names the same path.
  expect((error as Error).message).toContain(field);
  return field;
}

function doc(...body: unknown[]): DocxDocument {
  const d = Document.create();
  for (const block of body) {
    Document.addContent(d, block as BodyContent);
  }
  return d;
}

const text = (props?: object) => ({ ...Build.text("x"), properties: props }) as never;
const para = (props?: object, children: unknown[] = [Build.text("x")]) =>
  ({ ...Build.paragraph(children as never), properties: props }) as never;

describe("write-time model checks name the exact path", () => {
  it("paragraph properties", async () => {
    expect(await refusedAt(doc(para({ alignment: BAD })), BAD)).toBe(
      "body[0].properties.alignment"
    );
    expect(await refusedAt(doc(para({ textDirection: BAD })), BAD)).toBe(
      "body[0].properties.textDirection"
    );
    expect(await refusedAt(doc(para({ frame: { dropCap: BAD } })), BAD)).toBe(
      "body[0].properties.frame.dropCap"
    );
    expect(
      await refusedAt(
        doc(
          para({
            tabs: [
              { type: "left", position: 1 },
              { type: BAD, position: 2 }
            ]
          })
        ),
        BAD
      )
    ).toBe("body[0].properties.tabs[1].type");
    expect(await refusedAt(doc(para({ shading: { pattern: BAD } })), BAD)).toBe(
      "body[0].properties.shading.pattern"
    );
    expect(await refusedAt(doc(para({ spacing: { line: 240, lineRule: BAD } })), BAD)).toBe(
      "body[0].properties.spacing.lineRule"
    );
    expect(await refusedAt(doc(para({ markRunProperties: { highlight: BAD } })), BAD)).toBe(
      "body[0].properties.markRunProperties.highlight"
    );
  });

  it("run properties", async () => {
    expect(await refusedAt(doc(para(undefined, [text({ highlight: BAD })])), BAD)).toBe(
      "body[0].children[0].properties.highlight"
    );
    expect(
      await refusedAt(doc(para(undefined, [Build.text("a"), text({ underline: BAD })])), BAD)
    ).toBe("body[0].children[1].properties.underline");
    expect(await refusedAt(doc(para(undefined, [text({ underline: { style: BAD } })])), BAD)).toBe(
      "body[0].children[0].properties.underline.style"
    );
    expect(await refusedAt(doc(para(undefined, [text({ shading: { pattern: BAD } })])), BAD)).toBe(
      "body[0].children[0].properties.shading.pattern"
    );
    expect(await refusedAt(doc(para(undefined, [text({ font: { hint: BAD } })])), BAD)).toBe(
      "body[0].children[0].properties.font.hint"
    );
  });

  it("tables, rows and cells", async () => {
    const table = (
      tableProps?: object,
      rowProps?: object,
      cellProps?: object,
      cellContent?: unknown[]
    ) =>
      ({
        ...Build.table([
          {
            ...Build.row([
              {
                ...Build.cell("a"),
                ...(cellContent ? { content: cellContent } : {}),
                properties: cellProps
              } as never
            ]),
            properties: rowProps
          }
        ]),
        properties: tableProps
      }) as never;
    expect(await refusedAt(doc(table({ alignment: BAD })), BAD)).toBe(
      "body[0].properties.alignment"
    );
    expect(await refusedAt(doc(table({ layout: BAD })), BAD)).toBe("body[0].properties.layout");
    expect(await refusedAt(doc(table({ width: { value: 1, type: BAD } })), BAD)).toBe(
      "body[0].properties.width.type"
    );
    expect(await refusedAt(doc(table(undefined, { height: { value: 1, rule: BAD } })), BAD)).toBe(
      "body[0].rows[0].properties.height.rule"
    );
    expect(await refusedAt(doc(table(undefined, { tblPrEx: { layout: BAD } })), BAD)).toBe(
      "body[0].rows[0].properties.tblPrEx.layout"
    );
    expect(await refusedAt(doc(table(undefined, undefined, { verticalAlign: BAD })), BAD)).toBe(
      "body[0].rows[0].cells[0].properties.verticalAlign"
    );
    expect(
      await refusedAt(doc(table(undefined, undefined, { width: { value: 1, type: BAD } })), BAD)
    ).toBe("body[0].rows[0].cells[0].properties.width.type");
    expect(
      await refusedAt(doc(table(undefined, undefined, undefined, [para({ alignment: BAD })])), BAD)
    ).toBe("body[0].rows[0].cells[0].content[0].properties.alignment");
  });

  it("section properties, final and in a paragraph", async () => {
    const final = doc(para());
    Document.setSectionProperties(final, {});
    (final as { sectionProperties: object }).sectionProperties = { breakType: BAD };
    expect(await refusedAt(final, BAD)).toBe("sectionProperties.breakType");
    const grid = doc(para());
    (grid as { sectionProperties: object }).sectionProperties = { docGrid: { type: BAD } };
    expect(await refusedAt(grid, BAD)).toBe("sectionProperties.docGrid.type");
    expect(await refusedAt(doc(para({ sectionProperties: { verticalAlign: BAD } })), BAD)).toBe(
      "body[0].properties.sectionProperties.verticalAlign"
    );
  });

  it("drawings: a floating image and a shape", async () => {
    const image = Build.floatingImage({ rId: "r", width: 1, height: 1 });
    expect(await refusedAt(doc({ ...image, horizontalPosition: { relativeTo: BAD } }), BAD)).toBe(
      "body[0].horizontalPosition.relativeTo"
    );
    expect(await refusedAt(doc({ ...image, wrap: { style: "square", side: BAD } }), BAD)).toBe(
      "body[0].wrap.side"
    );
    expect(await refusedAt(doc({ ...image, width: Number.NaN }), Number.NaN)).toBe("body[0].width");
    expect(
      await refusedAt(
        doc({ ...image, wrap: { style: "square", margins: { top: Number.NaN } } }),
        Number.NaN
      )
    ).toBe("body[0].wrap.margins.top");
    expect(
      await refusedAt(doc({ ...image, verticalPosition: { offset: Number.NaN } }), Number.NaN)
    ).toBe("body[0].verticalPosition.offset");
    const shape = Build.rect(1, 1);
    expect(await refusedAt(doc({ ...shape, rotation: Number.NaN }), Number.NaN)).toBe(
      "body[0].rotation"
    );
  });

  it("an inline image in a run", async () => {
    const d = Document.create();
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    Document.addImage(d, png, "png", 9525, 9525);
    const block = d.body[0] as unknown as { children: { content: object[] }[] };
    const image = block.children[0]!.content[0]!;
    (block.children[0]!.content as object[])[0] = { ...image, srcRect: { l: Number.NaN } };
    expect(await refusedAt(d, Number.NaN)).toBe("body[0].children[0].content[0].srcRect.l");
  });

  it("a chart", async () => {
    const chart = Build.chart({ type: "bar", series: [] });
    expect(await refusedAt(doc({ ...chart, chart: { ...chart.chart, legend: BAD } }), BAD)).toBe(
      "body[0].chart.legend"
    );
    expect(
      await refusedAt(
        doc({ ...chart, chart: { ...chart.chart, dataLabels: { position: BAD } } }),
        BAD
      )
    ).toBe("body[0].chart.dataLabels.position");
  });

  it("a header", async () => {
    const d = doc(para());
    Document.setHeader(d, "default", { children: [para({ alignment: BAD })] });
    const field = await refusedAt(d, BAD);
    expect(field).toMatch(/^headers\["[^"]+"\]\.content\.children\[0\]\.properties\.alignment$/);
  });

  it("a hyperlink with nested children, by the real nested path", async () => {
    const run = (props?: object) => ({ ...Build.text("x"), properties: props });
    const inner = (props?: object) => ({
      ...Build.hyperlink("in", { anchor: "b" }),
      children: [run(props)]
    });
    const link = (children: object[], nestedChildren: object[]) => ({
      ...Build.hyperlink("x", { anchor: "a" }),
      children,
      nestedChildren
    });
    // A run after a nested child: `children[1]`, not its place (2) in the merged order.
    expect(
      await refusedAt(
        doc(
          para(undefined, [link([run(), run({ underline: BAD })], [{ before: 1, child: inner() }])])
        ),
        BAD
      )
    ).toBe("body[0].children[0].children[1].properties.underline");
    // Inside the nested child itself.
    expect(
      await refusedAt(
        doc(para(undefined, [link([run()], [{ before: 0, child: inner({ highlight: BAD }) }])])),
        BAD
      )
    ).toBe("body[0].children[0].nestedChildren[0].child.children[0].properties.highlight");
  });

  it("the glossary's building blocks, from the document root", async () => {
    const d = doc(para());
    (d as { glossary?: object }).glossary = {
      blocks: [
        { name: "ok", gallery: "autoText", content: [para()] },
        { name: "bad", gallery: "autoText", content: [para(), para({ alignment: BAD })] }
      ]
    };
    expect(await refusedAt(d, BAD)).toBe("glossary.blocks[1].content[1].properties.alignment");
  });

  it("styles, which the glossary's styles part writes too, from the document root", async () => {
    const d = doc(para());
    Document.useDefaultStyles(d);
    const styles = d.styles as object[];
    const i = styles.length;
    (d as unknown as { styles: object[] }).styles = [
      ...styles,
      { styleId: "Bad", type: "paragraph", name: "Bad", paragraphProperties: { alignment: BAD } }
    ];
    (d as { glossary?: object }).glossary = {
      blocks: [{ name: "ok", gallery: "autoText", content: [para()] }]
    };
    expect(await refusedAt(d, BAD)).toBe(`styles[${i}].paragraphProperties.alignment`);
  });

  it("the streaming writer, from its element parameter", () => {
    const writer = Streaming.createWriter();
    const element = para({ alignment: BAD });
    let error: unknown;
    try {
      writer.add(element);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(DocxArgumentError);
    const field = (error as DocxArgumentError).field;
    expect(field).toBe("element.properties.alignment");
    expect(walk({ element }, field)).toBe(BAD);
  });
});
