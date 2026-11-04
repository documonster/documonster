/**
 * Arguments are checked where they are given: a value with no valid form in the file is a
 * `PdfArgumentError` naming the call and the field, thrown before the call changes
 * anything — not a `0` written by `pdfNumber`, a black fill read off a colour string, or a
 * page no viewer opens.
 */
import { Pdf, PdfArgumentError, isPdfError } from "@pdf/index";
import { describe, expect, it } from "vitest";

type Call = () => unknown;

function page() {
  return new Pdf.Builder().addPage();
}

/** Run `call` and return the `PdfArgumentError` it threw. */
function argumentError(call: Call): PdfArgumentError {
  try {
    call();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect(isPdfError(err)).toBe(true);
    return err as PdfArgumentError;
  }
  throw new Error("expected a PdfArgumentError, nothing was thrown");
}

async function argumentRejection(call: () => Promise<unknown>): Promise<PdfArgumentError> {
  try {
    await call();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return err as PdfArgumentError;
  }
  throw new Error("expected a PdfArgumentError, nothing was thrown");
}

const RECT = [10, 10, 100, 40] as [number, number, number, number];

describe("PdfPageBuilder argument checks", () => {
  // Each of these was accepted and written as a document that showed something else.
  const invalid: Array<[string, string, Call]> = [
    ["addPage", "options.width", () => new Pdf.Builder().addPage({ width: NaN })],
    ["addPage", "options.height", () => new Pdf.Builder().addPage({ height: 0 })],
    ["addPage", "options.width", () => new Pdf.Builder().addPage({ width: -100 })],
    ["drawText", "options.fontSize", () => page().drawText("x", { x: 1, y: 1, fontSize: NaN })],
    ["drawText", "options.fontSize", () => page().drawText("x", { x: 1, y: 1, fontSize: 0 })],
    ["drawText", "options.x", () => page().drawText("x", { x: NaN, y: 1 })],
    ["drawText", "options.y", () => page().drawText("x", { x: 1, y: Infinity })],
    ["drawText", "text", () => page().drawText(42 as unknown as string, { x: 1, y: 1 })],
    [
      "drawText",
      "options.color.r",
      () => page().drawText("x", { x: 1, y: 1, color: { r: 255, g: 0, b: 0 } })
    ],
    [
      "drawText",
      "options.anchor",
      () => page().drawText("x", { x: 1, y: 1, anchor: "center" as "middle" })
    ],
    ["drawText", "options.maxWidth", () => page().drawText("x", { x: 1, y: 1, maxWidth: -1 })],
    ["drawText", "options.lineHeight", () => page().drawText("x", { x: 1, y: 1, lineHeight: NaN })],
    ["drawText", "options.rotation", () => page().drawText("x", { x: 1, y: 1, rotation: NaN })],
    [
      "drawText",
      "options.charSpacing",
      () => page().drawText("x", { x: 1, y: 1, charSpacing: NaN })
    ],
    ["measureText", "options.fontSize", () => page().measureText("x", { fontSize: -2 })],
    [
      "drawRect",
      "options.height",
      () => page().drawRect({ x: 0, y: 0, width: -5, height: Infinity })
    ],
    [
      "drawRect",
      "options.fill",
      () =>
        page().drawRect({
          x: 0,
          y: 0,
          width: 5,
          height: 5,
          fill: "notacolor" as unknown as { r: number; g: number; b: number }
        })
    ],
    [
      "drawRect",
      "options.stroke.a",
      () => page().drawRect({ x: 0, y: 0, width: 5, height: 5, stroke: { r: 0, g: 0, b: 0, a: 2 } })
    ],
    [
      "drawRect",
      "options.lineWidth",
      () => page().drawRect({ x: 0, y: 0, width: 5, height: 5, lineWidth: -1 })
    ],
    [
      "drawRect",
      "options.borderRadius",
      () => page().drawRect({ x: 0, y: 0, width: 5, height: 5, borderRadius: NaN })
    ],
    ["drawCircle", "options.r", () => page().drawCircle({ cx: 0, cy: 0, r: -1 })],
    ["drawCircle", "options.cx", () => page().drawCircle({ cx: NaN, cy: 0, r: 1 })],
    ["drawEllipse", "options.ry", () => page().drawEllipse({ cx: 0, cy: 0, rx: 1, ry: NaN })],
    ["drawLine", "options.x2", () => page().drawLine({ x1: 0, y1: 0, x2: NaN, y2: 0 })],
    [
      "drawLine",
      "options.dashPattern",
      () => page().drawLine({ x1: 0, y1: 0, x2: 1, y2: 1, dashPattern: [0, 0] })
    ],
    [
      "drawLine",
      "options.dashPattern[1]",
      () => page().drawLine({ x1: 0, y1: 0, x2: 1, y2: 1, dashPattern: [3, -1] })
    ],
    [
      "drawLine",
      "options.lineCap",
      () => page().drawLine({ x1: 0, y1: 0, x2: 1, y2: 1, lineCap: "flat" as "butt" })
    ],
    [
      "drawPath",
      "ops[1].y",
      () =>
        page().drawPath([
          { op: "move", x: 0, y: 0 },
          { op: "line", x: 1, y: NaN }
        ])
    ],
    [
      "drawPath",
      "ops[0].op",
      () =>
        page().drawPath([
          { op: "arc", x: 0, y: 0 } as unknown as { op: "move"; x: number; y: number }
        ])
    ],
    [
      "drawPath",
      "options.fillRule",
      () => page().drawPath([{ op: "move", x: 0, y: 0 }], { fillRule: "winding" as "nonzero" })
    ],
    ["drawSvgPath", "options.lineWidth", () => page().drawSvgPath("M0 0 L1 1", { lineWidth: NaN })],
    ["drawSvg", "options.x", () => page().drawSvg({ svg: "<svg/>", x: NaN, y: 0 })],
    [
      "drawImage",
      "options.data",
      () =>
        page().drawImage({
          data: "bytes" as unknown as Uint8Array,
          format: "png",
          x: 0,
          y: 0,
          width: 1,
          height: 1
        })
    ],
    [
      "drawImage",
      "options.width",
      () =>
        page().drawImage({
          data: new Uint8Array(8),
          format: "png",
          x: 0,
          y: 0,
          width: NaN,
          height: 1
        })
    ],
    [
      "addAnnotation",
      "options.type",
      () =>
        page().addAnnotation({ type: "Circle", rect: RECT } as unknown as Parameters<
          Pdf.PageBuilder["addAnnotation"]
        >[0])
    ],
    [
      "addAnnotation",
      "options.rect[2]",
      () => page().addAnnotation({ type: "Highlight", rect: [0, 0, NaN, 1] })
    ],
    [
      "addAnnotation",
      "options.quadPoints",
      () => page().addAnnotation({ type: "Highlight", rect: RECT, quadPoints: [1, 2, 3] })
    ],
    [
      "addAnnotation",
      "options.stampName",
      () =>
        page().addAnnotation({ type: "Stamp", rect: RECT, stampName: "Top Secret" as "TopSecret" })
    ],
    [
      "addAnnotation",
      "options.iconName",
      () => page().addAnnotation({ type: "Text", rect: RECT, iconName: "Star" as "Note" })
    ],
    [
      "addAnnotation",
      "options.fontSize",
      () => page().addAnnotation({ type: "FreeText", rect: RECT, contents: "x", fontSize: 0 })
    ],
    [
      "addAnnotation",
      "options.contents",
      () =>
        page().addAnnotation({ type: "FreeText", rect: RECT } as unknown as Parameters<
          Pdf.PageBuilder["addAnnotation"]
        >[0])
    ],
    [
      "addFormField",
      "options.name",
      () => page().addFormField({ type: "text", name: "", rect: RECT })
    ],
    [
      "addFormField",
      "options.maxLength",
      () => page().addFormField({ type: "text", name: "f", rect: RECT, maxLength: 2.5 })
    ],
    [
      "addFormField",
      "options.type",
      () =>
        page().addFormField({ type: "slider", name: "f", rect: RECT } as unknown as Parameters<
          Pdf.PageBuilder["addFormField"]
        >[0])
    ],
    [
      "addFormField",
      "options.options[1]",
      () =>
        page().addFormField({
          type: "dropdown",
          name: "f",
          rect: RECT,
          options: ["a", 2 as unknown as string]
        })
    ],
    [
      "addFormField",
      "options.selected",
      () =>
        page().addFormField({
          type: "radio",
          name: "f",
          buttons: [{ rect: RECT, value: "a" }],
          selected: "b"
        })
    ],
    [
      "addFormField",
      "options.buttons",
      () => page().addFormField({ type: "radio", name: "f", buttons: [] })
    ]
  ];

  it.each(invalid)("%s refuses %s", (method, field, call) => {
    const err = argumentError(call);
    expect(err.field).toBe(field);
    // The message starts with the call as the caller wrote it.
    const owner = method === "addPage" ? "Pdf.Builder" : "Pdf.PageBuilder";
    expect(err.message.startsWith(`${owner}#${method}: `)).toBe(true);
    expect(err.message).toContain(field);
  });

  it("changes nothing when it refuses", () => {
    const doc = new Pdf.Builder();
    const p = doc.addPage();
    const before = p.getContentStream().toString();
    argumentError(() =>
      p.drawRect({ x: 0, y: 0, width: 1, height: 1, fill: { r: 2, g: 0, b: 0 } })
    );
    argumentError(() => p.drawText("x", { x: 0, y: 0, fontSize: NaN }));
    argumentError(() => p.addFormField({ type: "text", name: "", rect: RECT }));
    expect(p.getContentStream().toString()).toBe(before);
    argumentError(() => doc.addPage({ width: 0 }));
    expect(doc.pages).toHaveLength(1);
  });

  it("keeps accepting boundary values the file can express", async () => {
    const doc = new Pdf.Builder();
    const p = doc.addPage({ width: 0.5, height: 14400 });
    p.drawText("", { x: 0, y: 0 });
    p.drawText("x", { x: -10, y: 0, fontSize: 0.1, charSpacing: -1, rotation: 720 });
    // A negative width or height draws the rectangle in the other direction, as `re` does.
    p.drawRect({ x: 0, y: 0, width: -5, height: -5, lineWidth: 0 });
    p.drawRect({ x: 0, y: 0, width: 1, height: 1, fill: { r: 0, g: 1, b: 1, a: 0 } });
    p.drawCircle({ cx: 0, cy: 0, r: 0 });
    p.drawLine({ x1: 0, y1: 0, x2: 1, y2: 1, dashPattern: [] });
    p.drawLine({ x1: 0, y1: 0, x2: 1, y2: 1, dashPattern: [0, 2] });
    p.drawPath([]);
    p.addAnnotation({ type: "Highlight", rect: RECT, quadPoints: [0, 0, 1, 0, 0, 1, 1, 1] });
    p.addFormField({ type: "text", name: "f", rect: RECT, maxLength: 0 });
    p.addFormField({
      type: "radio",
      name: "r",
      buttons: [{ rect: RECT, value: "a" }],
      selected: ""
    });
    expect((await doc.toBuffer()).length).toBeGreaterThan(0);
  });

  it("writes a radio export value as an escaped name", async () => {
    const doc = new Pdf.Builder();
    doc.addPage().addFormField({
      type: "radio",
      name: "choice",
      buttons: [{ rect: RECT, value: "Option A" }],
      selected: "Option A"
    });
    const text = new TextDecoder("latin1").decode(await doc.toBuffer());
    expect(text).toContain("/Option#20A");
    expect(text).not.toContain("/Option A");
  });

  it("checks EditorPage the same way, since it draws through a page builder", async () => {
    const source = new Pdf.Builder();
    source.addPage();
    const editor = Pdf.Editor.read(await source.toBuffer());
    const err = argumentError(() =>
      editor.getPage(0).drawRect({ x: 0, y: 0, width: NaN, height: 1 })
    );
    expect(err.field).toBe("options.width");
    expect(err.message.startsWith("Pdf.EditorPage#drawRect: ")).toBe(true);
    // A page added to the editor is a page builder, and says so.
    const added = argumentError(() => editor.addPage().drawText("x", { x: NaN, y: 0 }));
    expect(added.message.startsWith("Pdf.PageBuilder#drawText: ")).toBe(true);
  });
});

describe("PdfDocumentBuilder and PdfEditor argument checks", () => {
  it.each<[string, Call]>([
    ["collectionIndex", () => new Pdf.Builder().embedFont(new Uint8Array(4), -1)],
    ["metadata.creationDate", () => new Pdf.Builder().setMetadata({ creationDate: new Date(NaN) })],
    ["metadata.title", () => new Pdf.Builder().setMetadata({ title: 5 as unknown as string })],
    [
      "encryption.ownerPassword",
      () => new Pdf.Builder().setEncryption({} as { ownerPassword: string })
    ],
    ["families[0]", () => new Pdf.Builder().preferSystemFonts([1 as unknown as string])],
    ["language", () => new Pdf.Builder().setTextLanguage("zh" as "ja")],
    ["options.textLanguage", () => new Pdf.Builder({ textLanguage: "en" as "ja" })],
    ["pageIndex", () => new Pdf.Builder().addBookmark("t", 1.5)],
    ["pageIndex", () => new Pdf.Builder().addBookmark("t", -1)],
    ["parent", () => new Pdf.Builder().addBookmark("t", 0, NaN)],
    ["options.fontSize", () => new Pdf.Builder().generateTableOfContents({ fontSize: NaN })],
    ["options.indent", () => new Pdf.Builder().generateTableOfContents({ indent: -1 })],
    [
      "options.certificate",
      () =>
        new Pdf.Builder().sign({
          certificate: "pem" as unknown as Uint8Array,
          privateKey: new Uint8Array(1)
        })
    ]
  ])("Builder refuses %s", (field, call) => {
    expect(argumentError(call).field).toBe(field);
  });

  it("keeps a RangeError for a bookmark parent that is an integer out of range", () => {
    expect(() => new Pdf.Builder().addBookmark("t", 0, 3)).toThrow(RangeError);
  });

  it("refuses bad editor arguments before recording an edit", async () => {
    const source = new Pdf.Builder();
    source.addPage();
    const bytes = await source.toBuffer();
    const editor = Pdf.Editor.read(bytes);
    expect(argumentError(() => editor.addPage({ width: NaN })).field).toBe("options.width");
    expect(argumentError(() => editor.embedFont(new Uint8Array(4), 0.5)).field).toBe(
      "collectionIndex"
    );
    expect(argumentError(() => editor.setFormField("a", 1 as unknown as string)).field).toBe(
      "value"
    );
    expect(
      argumentError(() => editor.setFormFields({ a: "x", b: null as unknown as string })).field
    ).toBe('fields["b"]');
    expect(
      argumentError(() => Pdf.Editor.read(bytes, { password: 1 as unknown as string })).field
    ).toBe("options.password");
    expect(
      (await argumentRejection(() => editor.toBuffer({ encryption: "none" as "keep" }))).field
    ).toBe("options.encryption");
    expect(
      (await argumentRejection(() => editor.toBuffer({ incremental: "yes" as "required" }))).field
    ).toBe("options.incremental");
    expect(
      (
        await argumentRejection(() =>
          editor.splitPages({ invalidateSignatures: 1 as unknown as boolean })
        )
      ).field
    ).toBe("options.invalidateSignatures");
    // Nothing above was recorded: the editor still writes the original unchanged.
    expect(await editor.toBuffer({ incremental: true })).toEqual(bytes);
  });
});

describe("converter option checks", () => {
  const rows = [["a", 1]];

  it.each<[string, Record<string, unknown>]>([
    ["pageSize", { pageSize: "B9" }],
    ["pageSize.width", { pageSize: { width: NaN, height: 100 } }],
    ["pageSize.height", { pageSize: { width: 100, height: 0 } }],
    ["orientation", { orientation: "sideways" }],
    ["margins.top", { margins: { top: -1 } }],
    ["margins.left", { margins: { left: NaN } }],
    ["scale", { scale: NaN }],
    ["fitToWidth", { fitToWidth: NaN }],
    ["repeatRows", { repeatRows: 1.5 }],
    ["repeatCols", { repeatCols: -1 }],
    ["pageOrder", { pageOrder: "random" }],
    ["errors", { errors: "hidden" }],
    ["cellComments", { cellComments: "all" }],
    ["gridLineColor", { gridLineColor: "grey" }],
    ["defaultFontSize", { defaultFontSize: 0 }],
    ["headerFooter.date", { headerFooter: { date: new Date("nope") } }],
    ["creationDate", { creationDate: "2024-01-01" }],
    ["font", { font: "Arial" }],
    ["textLanguage", { textLanguage: "fr" }],
    ["encryption.ownerPassword", { encryption: { userPassword: "u" } }],
    ["watermark.type", { watermark: { type: "logo" } }],
    ["watermark.text", { watermark: { type: "text" } }],
    ["watermark.opacity", { watermark: { type: "text", text: "D", opacity: NaN } }],
    ["watermark.color.g", { watermark: { type: "text", text: "D", color: { r: 0, g: 9, b: 0 } } }],
    [
      "watermark.position.x",
      { watermark: { type: "text", text: "D", position: { x: NaN, y: 0 } } }
    ],
    ["watermark.pages[0]", { watermark: { type: "text", text: "D", pages: [0] } }],
    ["watermark.placement", { watermark: { type: "text", text: "D", placement: "behind" } }],
    ["watermark.data", { watermark: { type: "image", format: "png" } }]
  ])("fromTable refuses %s", async (field, options) => {
    const err = await argumentRejection(() =>
      Pdf.fromTable(rows, options as Parameters<typeof Pdf.fromTable>[1])
    );
    expect(err.field).toBe(`options.${field}`);
    expect(err.message.startsWith("Pdf.fromTable: ")).toBe(true);
  });

  it("keeps clamping finite values it has always clamped", async () => {
    const bytes = await Pdf.fromTable(rows, {
      scale: 10,
      fitToWidth: 1.7,
      margins: { top: 0 },
      repeatRows: false,
      gridLineColor: "#D0D0D0",
      watermark: { type: "text", text: "D", opacity: 3, fontSize: -4, pages: [1] }
    });
    expect(bytes.length).toBeGreaterThan(0);
  });

  it("checks fromDocx, fromChart and createDrawSurface options before loading a bridge", async () => {
    const doc = {} as Parameters<typeof Pdf.fromDocx>[0];
    expect((await argumentRejection(() => Pdf.fromDocx(doc, { pageWidth: NaN }))).field).toBe(
      "options.pageWidth"
    );
    expect((await argumentRejection(() => Pdf.fromDocx(doc, { marginTop: -1 }))).field).toBe(
      "options.marginTop"
    );
    const chart = {} as Parameters<typeof Pdf.fromChart>[0];
    expect((await argumentRejection(() => Pdf.fromChart(chart, { rasterScale: 0 }))).field).toBe(
      "options.rasterScale"
    );
    const p = page();
    expect(
      argumentError(() => Pdf.createDrawSurface(p, { x: 0, y: 0, width: NaN, height: 1 })).field
    ).toBe("rect.width");
    expect(
      argumentError(() =>
        Pdf.createDrawSurface(p, { x: 0, y: 0, width: 1, height: 1 }, { scale: 0 })
      ).field
    ).toBe("options.scale");
  });
});
