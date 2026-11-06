/**
 * Boundary contracts: what the builder, editor, converters, signing helpers and reader do
 * with an argument that has no valid reading, and with state a caller keeps changing after
 * handing it over. Each case was a silent wrong file, a lost value or a raw `TypeError`.
 */
import { describe, expect, it } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import {
  Pdf,
  PdfArgumentError,
  PdfPasswordError,
  PdfSignaturePlaceholderMissingError,
  PdfStructureError
} from "../index";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGP8z8Dwn4GBgYGJAQoAKu8BBf8M8UgAAAAASUVORK5CYII="
  ),
  c => c.charCodeAt(0)
);

/** A JPEG header — SOI, an optional APP14 "Adobe", SOF0 — with `components` channels. */
function jpegHeader(components: 1 | 3 | 4, adobe = false): Uint8Array {
  const app14 = adobe
    ? [0xff, 0xee, 0x00, 0x0e, 0x41, 0x64, 0x6f, 0x62, 0x65, 0, 100, 0, 0, 0, 0, 2]
    : [];
  const sof = [0xff, 0xc0, 0x00, 8 + components * 3, 8, 0, 2, 0, 3, components];
  for (let i = 0; i < components; i++) {
    sof.push(i + 1, 0x11, 0);
  }
  return new Uint8Array([0xff, 0xd8, ...app14, ...sof, 0xff, 0xd9]);
}

const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

function argumentError(run: () => unknown): PdfArgumentError {
  try {
    run();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return err as PdfArgumentError;
  }
  throw new Error("expected a PdfArgumentError");
}

async function rejection(run: () => Promise<unknown>): Promise<Error> {
  try {
    await run();
  } catch (err) {
    return err as Error;
  }
  throw new Error("expected a rejection");
}

async function onePage(): Promise<Uint8Array> {
  const b = new Pdf.Builder();
  b.addPage().drawText("x", { x: 10, y: 10 });
  return b.toBuffer();
}

describe("images are checked when they are given", () => {
  it("refuses bytes that are not the declared format", () => {
    const page = new Pdf.Builder().addPage();
    // PNG bytes declared as JPEG used to be written as a 1×1 DCTDecode image.
    expect(
      argumentError(() =>
        page.drawImage({ data: PNG, format: "jpeg", x: 0, y: 0, width: 1, height: 1 })
      ).field
    ).toBe("options.data");
    expect(
      argumentError(() =>
        page.drawImage({
          data: new Uint8Array([1, 2, 3, 4]),
          format: "png",
          x: 0,
          y: 0,
          width: 1,
          height: 1
        })
      ).field
    ).toBe("options.data");
    expect(
      argumentError(() =>
        page.drawImage({ data: PNG, format: "gif" as "png", x: 0, y: 0, width: 1, height: 1 })
      ).field
    ).toBe("options.format");
  });

  it("writes a JPEG in the colour space its frame header declares", async () => {
    const spaces: string[] = [];
    for (const [components, adobe] of [
      [1, false],
      [3, false],
      [4, true]
    ] as const) {
      const b = new Pdf.Builder();
      b.addPage().drawImage({
        data: jpegHeader(components, adobe),
        format: "jpeg",
        x: 0,
        y: 0,
        width: 3,
        height: 2
      });
      const text = latin1(await b.toBuffer());
      const dict = /<<[^<>]*\/Subtype \/Image[^<>]*>>/.exec(text)?.[0] ?? "";
      expect(dict).toContain("/Width 3");
      expect(dict).toContain("/Height 2");
      spaces.push(
        /\/ColorSpace (\/\w+)/.exec(dict)?.[1] + (dict.includes("/Decode [1 0") ? "+decode" : "")
      );
    }
    // Every JPEG used to be declared /DeviceRGB, which a grey or CMYK one does not decode as.
    expect(spaces).toEqual(["/DeviceGray", "/DeviceRGB", "/DeviceCMYK+decode"]);
  });

  it("embeds a PNG on an incremental update as decoded pixels, as a full rewrite does", async () => {
    const source = await onePage();
    const sizes: number[] = [];
    for (const incremental of [false, true]) {
      const editor = Pdf.Editor.read(source);
      editor.getPage(0).drawImage({ data: PNG, format: "png", x: 0, y: 0, width: 10, height: 10 });
      const out = await editor.toBuffer({ incremental });
      const image = (await Pdf.extract(out)).pages[0].images[0];
      sizes.push(image.data.length);
    }
    // 2×2 RGB is 12 bytes; the update used to write the 77-byte PNG file as the pixels.
    expect(sizes).toEqual([12, 12]);
  });
});

describe("the builder keeps a copy of what it checked", () => {
  it("is not changed by a later write to the caller's metadata or image options", async () => {
    const b = new Pdf.Builder();
    const metadata = { title: "One" };
    b.setMetadata(metadata);
    const options = { data: PNG, format: "png" as const, x: 0, y: 0, width: 5, height: 5 };
    b.addPage().drawImage(options);
    (metadata as { title: unknown }).title = 42;
    (options as { format: unknown }).format = "gif";
    const result = await Pdf.extract(await b.toBuffer());
    expect(result.metadata.title).toBe("One");
    expect(result.pages[0].images).toHaveLength(1);
  });

  it("is not changed by a later write to a form field's rect or options", async () => {
    const b = new Pdf.Builder();
    const options = ["A", "B"];
    b.addPage().addFormField({
      type: "dropdown",
      name: "pick",
      rect: [0, 0, 10, 10],
      options,
      value: "A"
    });
    options.push("C");
    const field = (await Pdf.extract(await b.toBuffer())).formFields[0];
    expect(field.options).toEqual(["A", "B"]);
  });
});

describe("form-field names identify one field", () => {
  it("refuses a second field with the same name, on one page or across pages", async () => {
    const page = new Pdf.Builder().addPage();
    page.addFormField({ type: "text", name: "dup", rect: [0, 0, 10, 10] });
    expect(
      argumentError(() =>
        page.addFormField({ type: "checkbox", name: "dup", rect: [0, 20, 10, 30] })
      ).field
    ).toBe("options.name");

    const b = new Pdf.Builder();
    b.addPage().addFormField({ type: "text", name: "dup", rect: [0, 0, 10, 10] });
    const second = b.addPage();
    // Refused at the call that makes the clash, not when the document is written.
    expect(
      argumentError(() => second.addFormField({ type: "text", name: "dup", rect: [0, 0, 10, 10] }))
        .field
    ).toBe("options.name");
    expect(await b.toBuffer()).toBeInstanceOf(Uint8Array);
  });

  it("refuses an editor field named like one the document has", async () => {
    const b = new Pdf.Builder();
    b.addPage().addFormField({ type: "text", name: "name", rect: [0, 0, 10, 10] });
    const editor = Pdf.Editor.read(await b.toBuffer());
    expect(
      argumentError(() =>
        editor.getPage(0).addFormField({ type: "text", name: "name", rect: [0, 20, 10, 30] })
      ).field
    ).toBe("options.name");
  });
});

describe("Editor.setFormField", () => {
  async function form(): Promise<Uint8Array> {
    const b = new Pdf.Builder();
    const page = b.addPage();
    page.addFormField({ type: "text", name: "name", rect: [0, 0, 100, 20], value: "Alice" });
    page.addFormField({ type: "checkbox", name: "agree", rect: [0, 30, 10, 40] });
    page.addFormField({
      type: "radio",
      name: "size",
      buttons: [
        { rect: [0, 50, 10, 60], value: "S" },
        { rect: [20, 50, 30, 60], value: "M" }
      ]
    });
    return b.toBuffer();
  }

  it("refuses a field the document does not have", async () => {
    // It used to be dropped at write time without a word.
    const editor = Pdf.Editor.read(await form());
    expect(argumentError(() => editor.setFormField("nope", "x")).field).toBe("fieldName");
    expect(argumentError(() => editor.setFormFields({ name: "Bob", nope: "x" })).field).toBe(
      'fields["nope"]'
    );
    // Nothing was set by the call that threw.
    expect(
      (await Pdf.extract(await editor.toBuffer())).formFields.find(f => f.name === "name")?.value
    ).toBe("Alice");
  });

  it("refuses a state a checkbox or radio group does not have", async () => {
    const editor = Pdf.Editor.read(await form());
    expect(() => editor.setFormField("agree", "banana")).toThrow(PdfArgumentError);
    expect(() => editor.setFormField("size", "XL")).toThrow(/"S", "M"/);
    editor.setFormField("agree", "Yes").setFormField("size", "M").setFormField("name", "Bob");
    const values = Object.fromEntries(
      (await Pdf.extract(await editor.toBuffer())).formFields.map(f => [f.name, f.value])
    );
    expect(values).toMatchObject({ agree: "Yes", size: "M", name: "Bob" });
  });

  for (const incremental of [false, true]) {
    it(`sets a button's state and its value as names (${incremental ? "incremental" : "full"})`, async () => {
      // A checkbox set to "Yes" used to keep /AS /Off — drawn unchecked — with /V a string,
      // and a radio group's value never reached the group.
      const editor = Pdf.Editor.read(await form());
      editor.setFormField("agree", "Yes").setFormField("size", "M");
      const out = await editor.toBuffer({ incremental });
      const text = latin1(out);
      expect(text).toMatch(/\/AS \/Yes/);
      expect(text).toMatch(/\/AS \/M/);
      expect(text).not.toMatch(/\/V \(Yes\)/);
      const values = Object.fromEntries(
        (await Pdf.extract(out)).formFields.map(f => [f.name, f.value])
      );
      expect(values).toMatchObject({ agree: "Yes", size: "M" });
    });
  }
});

describe("Editor page lists", () => {
  it("splitPages takes { pageIndices } with the write options", async () => {
    const b = new Pdf.Builder();
    b.addPage().drawText("A", { x: 10, y: 10 });
    b.addPage().drawText("B", { x: 10, y: 10 });
    const editor = Pdf.Editor.read(await b.toBuffer());
    const parts = await editor.splitPages({ pageIndices: [1], encryption: "keep" });
    expect(parts).toHaveLength(1);
    expect((await Pdf.extract(parts[0])).text).toBe("B");
    // The deprecated positional form still works.
    expect(await editor.splitPages([0, 1])).toHaveLength(2);
  });

  it("refuses `pages`, Pdf.extract's 1-based list, instead of ignoring it", async () => {
    const editor = Pdf.Editor.read(await onePage());
    expect((await rejection(() => editor.splitPages({ pages: [0] } as never))).message).toMatch(
      /pageIndices/
    );
    expect(() => editor.copyPagesFrom(new Uint8Array(), { pages: [0] } as never)).toThrow(
      /pageIndices/
    );
  });
});

describe("converter and builder arguments", () => {
  it("checks fromDocx options before the bridge reads them", async () => {
    const doc = { body: [] } as unknown as Parameters<typeof Pdf.fromDocx>[0];
    // Each was a TypeError from inside the bridge, or silently accepted.
    for (const [options, field] of [
      [{ textLanguage: "en" }, "options.textLanguage"],
      [{ onWarning: 42 }, "options.onWarning"],
      [{ preferSystemFonts: "Arial" }, "options.preferSystemFonts"],
      [{ defaultFont: 42 }, "options.defaultFont"],
      [{ chartRenderer: "x" }, "options.chartRenderer"]
    ] as const) {
      const err = await rejection(() => Pdf.fromDocx(doc, options as never));
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(field);
    }
    expect(await rejection(() => Pdf.fromDocx(null as never))).toBeInstanceOf(PdfArgumentError);
  });

  it("checks fromExcel and fromChart handles and options", async () => {
    expect(await rejection(() => Pdf.fromExcel(null as never))).toBeInstanceOf(PdfArgumentError);
    const err = await rejection(() => Pdf.fromExcel({} as never, { recalculate: 42 as never }));
    expect((err as PdfArgumentError).field).toBe("options.recalculate");
    expect(await rejection(() => Pdf.fromChart(null as never))).toBeInstanceOf(PdfArgumentError);
    const named = await rejection(() => Pdf.fromChart({} as never, { title: 42 as never }));
    // Named for the call that received it, not the builder it is passed on to.
    expect(named.message).toMatch(/^Pdf\.fromChart: options\.title/);
  });

  it("checks fromTable cells, styles and columns", async () => {
    for (const [input, field] of [
      [[["a", { a: 1 }, "b"]], "input[0][1]"],
      [[["a", () => 1]], "input[0][1]"],
      [[[Number.NaN]], "input[0][0]"],
      [[[new Date("x")]], "input[0][0]"],
      [[[{ value: "x", fontColor: "red" }]], "input[0][0].fontColor"],
      [[[{ value: "x", fontSize: Number.NaN }]], "input[0][0].fontSize"],
      [[[{ value: "x", align: "middle" }]], "input[0][0].align"],
      [{ data: [["x"]], columns: [{ width: -10 }] }, "input.columns[0].width"],
      [{ data: [["x"]], columns: [0] }, "input.columns[0]"],
      [null, "input"]
    ] as const) {
      const err = await rejection(() => Pdf.fromTable(input as never));
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(field);
    }
    const ok = await Pdf.fromTable([["a", 1, true, null, { value: "b", fontColor: "FFFF0000" }]]);
    expect((await Pdf.extract(ok)).text).toContain("b");
  });

  it("checks the builder's handler, font bytes and PDF/A level", () => {
    const b = new Pdf.Builder();
    expect(argumentError(() => b.onWarning(42 as never)).field).toBe("handler");
    expect(argumentError(() => b.embedFont(null as never)).field).toBe("source");
    expect(argumentError(() => b.setPdfACompliance("2b" as never)).field).toBe("level");
  });

  it("names the parent of an out-of-range bookmark sensibly", () => {
    // It said "out of range (0..-1)" when there was no bookmark.
    expect(() => new Pdf.Builder().addBookmark("a", 0, 0)).toThrow(/none has been added yet/);
  });
});

describe("createDrawSurface", () => {
  it("refuses a box with no area and a page that is not a page builder", () => {
    const page = new Pdf.Builder().addPage();
    expect(
      argumentError(() => Pdf.createDrawSurface(page, { x: 0, y: 0, width: -5, height: 5 })).field
    ).toBe("rect.width");
    expect(
      argumentError(() => Pdf.createDrawSurface({} as never, { x: 0, y: 0, width: 5, height: 5 }))
        .field
    ).toBe("page");
  });
});

describe("the page builder classes are not constructible", () => {
  it("has no public constructor", () => {
    // @ts-expect-error -- the constructor is private: a page comes from addPage/getPage.
    const make = () => new Pdf.PageBuilder();
    expect(typeof make).toBe("function");
    // @ts-expect-error -- likewise for an editor page.
    const makeEditorPage = () => new Pdf.EditorPage();
    expect(typeof makeEditorPage).toBe("function");
  });
});

describe("Pdf.extract arguments", () => {
  it("reports a wrong type as a wrong type, not as a corrupt file", async () => {
    const bytes = await onePage();
    for (const [run, field] of [
      // Any ArrayBufferView and an ArrayBuffer are read as bytes; anything else is not bytes.
      [() => Pdf.extract("%PDF-1.7" as never), "data"],
      [() => Pdf.extract(null as never), "data"],
      [() => Pdf.extract(bytes, { pages: [1.5] }), "options.pages[0]"],
      [() => Pdf.extract(bytes, { pages: [0] }), "options.pages[0]"],
      [() => Pdf.extract(bytes, { pages: "1" as never }), "options.pages"],
      [() => Pdf.extract(bytes, { extractText: "no" as never }), "options.extractText"],
      [() => Pdf.extract(bytes, null as never), "options"]
    ] as const) {
      const err = await rejection(run);
      expect(err).toBeInstanceOf(PdfArgumentError);
      expect((err as PdfArgumentError).field).toBe(field);
    }
  });

  it("reports the page count and the rotation even without metadata", async () => {
    const editor = Pdf.Editor.read(await onePage());
    editor.rotatePage(0, 90);
    const bytes = await editor.toBuffer();
    const result = await Pdf.extract(bytes, { extractMetadata: false });
    // It used to say pageCount 0 for a one-page file.
    expect(result.metadata.pageCount).toBe(1);
    expect(result.pages[0].rotation).toBe(90);
  });

  it("raises PdfPasswordError, a PdfStructureError, for a wrong password", async () => {
    const b = new Pdf.Builder();
    b.addPage();
    b.setEncryption({ ownerPassword: "o", userPassword: "u" });
    const bytes = await b.toBuffer();
    for (const run of [
      () => Pdf.extract(bytes, { password: "bad" }),
      () => Pdf.verifySignatures(bytes),
      async () => Pdf.Editor.read(bytes)
    ]) {
      const err = await rejection(run);
      expect(err).toBeInstanceOf(PdfPasswordError);
      expect(err).toBeInstanceOf(PdfStructureError);
    }
  });
});

describe("signing material is checked when it is given", () => {
  it("refuses a certificate or key that is not DER X.509 / PKCS#8 RSA", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate();
    const b = new Pdf.Builder();
    // These used to pass `sign()` and fail inside `toBuffer()` with a TypeError or a raw
    // node:crypto error.
    expect(
      argumentError(() => b.sign({ certificate: new Uint8Array([1, 2, 3]), privateKey })).field
    ).toBe("options.certificate");
    expect(
      argumentError(() => b.sign({ certificate, privateKey: new Uint8Array([0x30, 0x00]) })).field
    ).toBe("options.privateKey");
    expect(argumentError(() => b.sign({ certificate, privateKey: certificate })).field).toBe(
      "options.privateKey"
    );
    b.addPage();
    b.sign({ certificate, privateKey, name: "Ann", reason: "Approval" });
    const [report] = await Pdf.verifySignatures(await b.toBuffer());
    expect(report.intact).toBe(true);
    // Read back under the name it was written with.
    expect(report.reason).toBe("Approval");
  });

  it("checks the low-level helpers' arguments", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate();
    const bytes = await onePage();
    expect(
      await rejection(() => Pdf.Signature.fillPlaceholder(bytes, undefined as never))
    ).toBeInstanceOf(PdfArgumentError);
    const missing = await rejection(() =>
      Pdf.Signature.fillPlaceholder(bytes, { certificate, privateKey })
    );
    expect(missing).toBeInstanceOf(PdfSignaturePlaceholderMissingError);
    expect(missing).toBeInstanceOf(PdfStructureError);
    expect(
      argumentError(() => Pdf.Signature.buildDictPlaceholder({ name: 42 as never })).field
    ).toBe("options.name");
    expect(argumentError(() => Pdf.Signature.parseAsn1(null as never)).field).toBe("data");
    // Indefinite length and a length past the end used to parse as an empty node.
    expect(() => Pdf.Signature.parseAsn1(new Uint8Array([0x30, 0x80, 1]))).toThrow(
      PdfStructureError
    );
    expect(() => Pdf.Signature.parseAsn1(new Uint8Array([0x30, 0x05, 1]))).toThrow(
      PdfStructureError
    );
    expect(
      (await rejection(() => Pdf.Signature.verify(bytes, "00", [Number.NaN, 1, 2, 3]))).message
    ).toMatch(/byteRange/);
  });
});

describe("constants are frozen", () => {
  it("PageSizes cannot be changed for every later export", () => {
    expect(Object.isFrozen(Pdf.PageSizes)).toBe(true);
    expect(Object.isFrozen(Pdf.PageSizes.A4)).toBe(true);
  });
});

describe("PdfContentStream names", () => {
  it("refuses a resource name that would write operators", () => {
    const stream = new Pdf.Builder().addPage().getContentStream();
    // `setFont("F1) Tj (", 12)` wrote a text-showing operator into the stream.
    expect(() => stream.setFont("F1) Tj (", 12)).toThrow(PdfArgumentError);
    expect(() => stream.doXObject("Im1 Do Q")).toThrow(PdfArgumentError);
    expect(() => stream.setGraphicsState("GS1 gs q")).toThrow(PdfArgumentError);
    expect(() => stream.setLineCap(7 as never)).toThrow(PdfArgumentError);
    expect(stream.setFont("F1", 12).toString()).toBe("/F1 12 Tf");
  });
});

describe("EditorPage draws SVG like PageBuilder", () => {
  it("has drawSvg", async () => {
    const editor = Pdf.Editor.read(await onePage());
    editor.getPage(0).drawSvg({
      svg: '<svg width="10" height="10"><rect width="10" height="10" fill="red"/></svg>',
      x: 0,
      y: 0
    });
    expect(latin1(await editor.toBuffer())).toContain("1 0 0 rg");
  });
});
