import { unzip } from "@archive/read-archive";
import { XmlError } from "@xml/errors";
import { describe, expect, it } from "vitest";

import {
  Build,
  Document,
  DocxArgumentError,
  DocxInvalidDrawingError,
  Io,
  type BodyContent,
  type DocxDocument
} from "../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
const BAD_LITERALS =
  /="(?:undefined|null|NaN|-?Infinity|\[object Object\])"|>(?:undefined|null|NaN|-?Infinity|\[object Object\])</;

async function documentXml(doc: ReturnType<typeof Document.create>): Promise<string> {
  const bytes = await Io.toBuffer(doc);
  for await (const entry of unzip(bytes).entries()) {
    if (entry.path === "word/document.xml") {
      return new TextDecoder().decode(await entry.bytes());
    }
  }
  throw new Error("no document part");
}

function fieldOf(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(DocxInvalidDrawingError);
    return (err as DocxInvalidDrawingError).field;
  }
  throw new Error("expected a DocxInvalidDrawingError");
}

describe("drawing arguments at the API boundary", () => {
  it("Build.rect rejects a size passed as an object (the R2i corruption)", () => {
    const size = { width: 1905000, height: 762000 };
    expect(fieldOf(() => Build.rect(size as never, undefined as never))).toBe("width");
    expect(fieldOf(() => Build.rect(1905000, undefined as never))).toBe("height");
  });

  it("every shape builder checks its size", () => {
    expect(fieldOf(() => Build.shape({ shapeType: "rect", width: NaN, height: 1 }))).toBe(
      "options.width"
    );
    expect(fieldOf(() => Build.ellipse(1, -5))).toBe("height");
    expect(fieldOf(() => Build.roundRect(Infinity, 1))).toBe("width");
    expect(fieldOf(() => Build.line("1" as never, 1))).toBe("width");
    expect(fieldOf(() => Build.arrow("right", 1, {} as never))).toBe("height");
    expect(fieldOf(() => Build.flowchartShape("process", null as never, 1))).toBe("width");
    expect(fieldOf(() => Build.callout("rect", 1, NaN))).toBe("height");
    expect(fieldOf(() => Build.star(5, -1, 1))).toBe("width");
    expect(
      fieldOf(() => Build.drawingShape({ shapeType: "rect", width: {} as never, height: 1 }))
    ).toBe("width");
  });

  it("checks positions, wrap distances, rotation and outline width", () => {
    expect(fieldOf(() => Build.rect(1, 1, { horizontalPosition: { offset: "10" as never } }))).toBe(
      "options.horizontalPosition.offset"
    );
    expect(fieldOf(() => Build.rect(1, 1, { verticalPosition: { offset: NaN } }))).toBe(
      "options.verticalPosition.offset"
    );
    expect(fieldOf(() => Build.rect(1, 1, { rotation: {} as never }))).toBe("options.rotation");
    expect(fieldOf(() => Build.rect(1, 1, { outline: { width: NaN } }))).toBe(
      "options.outline.width"
    );
    expect(fieldOf(() => Build.floatingImage({ rId: "r", width: 1, height: 1, distL: -1 }))).toBe(
      "options.distL"
    );
    expect(
      fieldOf(() =>
        Build.floatingImage({
          rId: "r",
          width: 1,
          height: 1,
          wrap: { style: "square", margins: { top: NaN } }
        })
      )
    ).toBe("options.wrap.margins.top");
    expect(
      fieldOf(() =>
        Build.floatingImage({ rId: "r", width: 1, height: 1, simplePos: { x: 0, y: {} as never } })
      )
    ).toBe("options.simplePos.y");
  });

  it("checks the effect geometry an advanced shape serialises", () => {
    expect(
      fieldOf(() =>
        Build.shape({
          shapeType: "rect",
          width: 1,
          height: 1,
          effects: { glow: { color: "FF0000", radius: {} as never } }
        })
      )
    ).toBe("options.effects.glow.radius");
    expect(
      fieldOf(() =>
        Build.shape({
          shapeType: "rect",
          width: 1,
          height: 1,
          fill: { type: "gradient", stops: [{ position: NaN, color: "FF0000" }] }
        })
      )
    ).toBe("options.fill.stops[0].position");
  });

  it("Document helpers check sizes", () => {
    const doc = Document.create();
    expect(fieldOf(() => Document.addImage(doc, PNG, "png", { cx: 1 } as never, 1))).toBe("width");
    expect(fieldOf(() => Document.addFloatingImage(doc, PNG, "png", 1, NaN))).toBe("height");
    expect(fieldOf(() => Document.addTextBox(doc, "x", { width: {} as never }))).toBe(
      "options.width"
    );
    expect(fieldOf(() => Document.addTextBox(doc, "x", { height: -1 }))).toBe("options.height");
    expect(doc.images ?? []).toHaveLength(0);
    expect(doc.body).toHaveLength(0);
  });

  it("every valid baseline form keeps working", async () => {
    const doc = Document.create();
    Document.addContent(doc, Build.rect(1905000, 762000, { wrap: { style: "square" } }));
    Document.addContent(doc, Build.ellipse(0, 0));
    Document.addContent(
      doc,
      Build.rect(1005840.0000000001, 762000, {
        rotation: -5400000,
        horizontalPosition: { relativeTo: "page", offset: -12700 },
        verticalPosition: { relativeTo: "page", align: "top" },
        outline: { width: 12700, color: "000000" }
      })
    );
    Document.addTextBox(doc, "box", { width: 4000.5, height: 900 });
    Document.addImage(doc, PNG, "png", 914400, 914400);
    Document.addFloatingImage(doc, PNG, "png", 914400, 914400, { distL: 0, rotation: 60000 });
    const xml = await documentXml(doc);
    expect(xml).not.toMatch(BAD_LITERALS);
    // Fractional EMU is written as the integer the schema requires.
    expect(xml).toContain('cx="1005840"');
    expect(xml).toContain("<wp:posOffset>-12700</wp:posOffset>");
  });
});

describe("drawing numbers at write time (hand-built models)", () => {
  function shapeModel(patch: Record<string, unknown>): BodyContent {
    return { ...Build.rect(1905000, 762000), ...patch } as BodyContent;
  }

  it.each([
    ["width", { width: { width: 1 } }],
    ["height", { height: undefined }],
    ["horizontalPosition.offset", { horizontalPosition: { relativeTo: "page", offset: NaN } }],
    ["distT", { distT: Infinity }],
    ["rotation", { rotation: "90" }],
    ["outlineWidth", { outlineWidth: {} }]
  ])("refuses %s instead of writing a corrupt package", async (field, patch) => {
    const doc = Document.create();
    Document.addContent(doc, shapeModel(patch));
    const err = await Io.toBuffer(doc).then(
      () => undefined,
      (e: unknown) => e
    );
    expect(err).toBeInstanceOf(DocxInvalidDrawingError);
    // The writer names the node's path in front of the field.
    expect((err as DocxInvalidDrawingError).field).toBe(`body[0].${field}`);
  });

  it("refuses an inline image whose size is not a number", async () => {
    const doc = Document.create();
    Document.addImage(doc, PNG, "png", 914400, 914400);
    const para = doc.body[0] as unknown as { children: { content: { width: unknown }[] }[] };
    para.children[0]!.content[0]!.width = "914400";
    await expect(Io.toBuffer(doc)).rejects.toBeInstanceOf(DocxInvalidDrawingError);
  });
});

describe("wrong-typed model values elsewhere in a document", () => {
  const OBJ = {} as never;
  // The builders refuse these where they are given (argument-checks.test.ts); a model built by
  // hand reaches the writer unchecked, which must still refuse to write it.
  type Writable = { -readonly [K in keyof DocxDocument]: DocxDocument[K] };
  it.each<[string, (doc: ReturnType<typeof Document.create>) => void]>([
    [
      "a run size",
      d =>
        Document.addContent(d, {
          type: "paragraph",
          children: [{ properties: { size: OBJ }, content: [{ type: "text", text: "x" }] }]
        })
    ],
    [
      "a paragraph spacing",
      d =>
        Document.addContent(d, {
          type: "paragraph",
          properties: { spacing: { before: NaN } },
          children: [{ content: [{ type: "text", text: "x" }] }]
        })
    ],
    [
      "a page size",
      d => {
        (d as Writable).sectionProperties = { pageSize: { width: OBJ, height: 15840 } };
      }
    ],
    [
      "a default font size",
      d => {
        (d as Writable).docDefaults = { runProperties: { size: OBJ } };
      }
    ]
  ])("%s is a typed error at write time, not a corrupt package", async (_, apply) => {
    const doc = Document.create();
    apply(doc);
    // The XML writer's refusal comes back as the Word API's own error, naming
    // where the value is; the XmlError is its cause.
    const error = await Io.toBuffer(doc).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(DocxArgumentError);
    expect((error as Error).cause).toBeInstanceOf(XmlError);
  });
});
