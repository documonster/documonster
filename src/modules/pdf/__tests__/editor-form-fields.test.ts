import { describe, expect, it } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import { Pdf, PdfIncrementalUpdateError, PdfXfaFormError } from "../index";
import { buildRawPdf, twoPageFormBodies } from "./raw-pdf.helpers";

async function makeForm(): Promise<Uint8Array> {
  const builder = new Pdf.Builder();
  const page = builder.addPage();
  page.drawText("Application", { x: 60, y: 740, fontSize: 18 });
  page.addFormField({ type: "text", name: "fullName", rect: [60, 690, 300, 712] });
  page.addFormField({ type: "checkbox", name: "agree", rect: [60, 650, 76, 666] });
  return builder.toBuffer();
}

async function fieldValues(bytes: Uint8Array): Promise<Record<string, string | undefined>> {
  const result = await Pdf.extract(bytes, { extractFormFields: true, extractText: false });
  return Object.fromEntries(result.formFields.map(field => [field.name, field.value]));
}

describe("PdfEditor — form values survive a full rewrite", () => {
  it("a full toBuffer() keeps the values set with setFormFields", async () => {
    const editor = Pdf.Editor.read(await makeForm());
    editor.setFormFields({ fullName: "Jane Doe", agree: "Yes" });
    const out = await editor.toBuffer();
    expect(await fieldValues(out)).toEqual({ fullName: "Jane Doe", agree: "Yes" });
  });

  it("full and incremental output report the same values", async () => {
    const form = await makeForm();
    const full = Pdf.Editor.read(form).setFormField("fullName", "Ada");
    const incremental = Pdf.Editor.read(form).setFormField("fullName", "Ada");
    expect(await fieldValues(await full.toBuffer())).toEqual(
      await fieldValues(await incremental.toBuffer({ incremental: true }))
    );
  });

  it("the AcroForm lists the widgets that are on the page, not copies of them", async () => {
    const editor = Pdf.Editor.read(await makeForm());
    editor.setFormField("fullName", "Grace");
    const text = new TextDecoder("latin1").decode(await editor.toBuffer());
    const fields = /\/Fields\s*\[([^\]]*)\]/.exec(text)![1].match(/\d+ 0 R/g)!;
    const annots = /\/Annots\s*\[([^\]]*)\]/.exec(text)![1].match(/\d+ 0 R/g)!;
    expect([...fields].sort()).toEqual([...annots].sort());
  });

  it("a full rewrite with no edits keeps the form", async () => {
    const out = await Pdf.Editor.read(await makeForm()).toBuffer();
    expect(Object.keys(await fieldValues(out)).sort()).toEqual(["agree", "fullName"]);
  });

  it("a filled form survives a second full rewrite and a page removal", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().addFormField({ type: "text", name: "a", rect: [60, 690, 300, 712] });
    builder.addPage().addFormField({ type: "text", name: "b", rect: [60, 690, 300, 712] });
    const filled = await Pdf.Editor.read(await builder.toBuffer())
      .setFormFields({ a: "first", b: "second" })
      .toBuffer();
    const again = await Pdf.Editor.read(filled).toBuffer();
    expect(await fieldValues(again)).toEqual({ a: "first", b: "second" });

    const pruned = Pdf.Editor.read(filled);
    pruned.removePage(0);
    const out = await pruned.toBuffer();
    expect((await Pdf.extract(out)).pages).toHaveLength(1);
    expect((await fieldValues(out)).b).toBe("second");
  });
});

describe("PdfEditor — removing a page removes its fields", () => {
  const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);

  it("a builder form: the deleted page's field is gone, the rest keep their values", async () => {
    const builder = new Pdf.Builder();
    builder.addPage().addFormField({ type: "text", name: "a", rect: [60, 690, 300, 712] });
    builder.addPage().addFormField({ type: "text", name: "b", rect: [60, 690, 300, 712] });
    builder.addPage().addFormField({ type: "text", name: "c", rect: [60, 690, 300, 712] });
    const filled = await Pdf.Editor.read(await builder.toBuffer())
      .setFormFields({ a: "first", b: "second", c: "third" })
      .toBuffer();

    const editor = Pdf.Editor.read(filled);
    editor.removePage(1);
    const out = await editor.toBuffer();
    expect(await fieldValues(out)).toEqual({ a: "first", c: "third" });

    // The result is a working form: it re-opens and fills, both ways.
    const refilled = Pdf.Editor.read(out).setFormField("c", "again");
    expect(await fieldValues(await refilled.toBuffer())).toEqual({ a: "first", c: "again" });
    const appended = Pdf.Editor.read(out).setFormField("a", "later");
    expect(await fieldValues(await appended.toBuffer({ incremental: true }))).toEqual({
      a: "later",
      c: "third"
    });
  });

  it("prunes a parent's /Kids, drops an emptied field, and keeps a widget-less one", async () => {
    const form = buildRawPdf(twoPageFormBodies());
    expect(Object.keys(await fieldValues(form)).sort()).toEqual(
      ["group", "hidden", "name", "solo"].sort()
    );

    const editor = Pdf.Editor.read(form);
    editor.removePage(1);
    editor.setFormFields({ name: "Ada", group: "shared" });
    const out = await editor.toBuffer();

    const values = await fieldValues(out);
    expect(Object.keys(values).sort()).toEqual(["group", "hidden", "name"]);
    expect(values.name).toBe("Ada");
    expect(values.hidden).toBe("kept");

    // The surviving parent lists exactly one kid, and it is the widget on the kept page.
    const text = latin1(out);
    const parent =
      /<<[^<>]*\/T \(group\)[^<>]*>>|<<[^<>]*\/T <[0-9a-fA-F]+>[^<>]*\/Kids[^<>]*>>/.exec(text);
    expect(parent).not.toBeNull();
    const kids = /\/Kids \[([^\]]*)\]/.exec(parent![0])![1].match(/\d+ 0 R/g)!;
    expect(kids).toHaveLength(1);
    const annots = /\/Annots \[([^\]]*)\]/.exec(text)![1].match(/\d+ 0 R/g)!;
    expect(annots).toContain(kids[0]);
  });

  it("removing nothing leaves the field tree whole", async () => {
    const out = await Pdf.Editor.read(buildRawPdf(twoPageFormBodies())).toBuffer();
    expect(Object.keys(await fieldValues(out)).sort()).toEqual(["group", "hidden", "name", "solo"]);
  });
});

describe("PdfEditor — a form whose only cross-reference data is a stream", () => {
  it("is filled by a full rewrite and by an incremental update", async () => {
    const form = buildRawPdf(twoPageFormBodies(), { xrefStream: true });
    expect(new TextDecoder("latin1").decode(form)).not.toContain("trailer");
    expect((await fieldValues(form)).hidden).toBe("kept");

    const full = await Pdf.Editor.read(form).setFormFields({ name: "Ada", solo: "S" }).toBuffer();
    expect(await fieldValues(full)).toMatchObject({ name: "Ada", solo: "S", hidden: "kept" });

    const incremental = await Pdf.Editor.read(form)
      .setFormFields({ name: "Ada", solo: "S" })
      .toBuffer({ incremental: true });
    expect(incremental.subarray(0, form.length)).toEqual(form);
    expect(await fieldValues(incremental)).toMatchObject({
      name: "Ada",
      solo: "S",
      hidden: "kept"
    });
    // Still an xref-stream file after the update.
    expect(new TextDecoder("latin1").decode(incremental.subarray(form.length))).toContain(
      "/Type /XRef"
    );
  });
});

describe("PdfEditor — pruning deep and irregular field trees", () => {
  const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
  const page = (annots: string, contents: number): string =>
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents ${contents} 0 R /Annots [${annots}] >>`;
  const widget = (extra: string, pageObj: number, y: number): string =>
    `<< /Type /Annot /Subtype /Widget /Rect [60 ${y} 300 ${y + 22}] /P ${pageObj} 0 R ${extra} >>`;
  const content = "<< /Length 0 >>\nstream\n\nendstream";
  const pageObjects = (text: string): number =>
    (text.match(/\/Type\s*\/Page(?![s\w])/g) ?? []).length;
  const widgets = (text: string): number => (text.match(/\/Subtype\s*\/Widget/g) ?? []).length;

  async function removeSecondPage(bodies: string[]): Promise<string> {
    const editor = Pdf.Editor.read(buildRawPdf(bodies));
    editor.removePage(1);
    return latin1(await editor.toBuffer());
  }

  it("a three-level tree does not resurrect the deleted widget or its page", async () => {
    const out = await removeSecondPage([
      /* 1 */ "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >>",
      /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      /* 3 */ page("7 0 R", 9),
      /* 4 */ page("8 0 R 11 0 R", 10),
      /* 5 */ "<< /T (form) /Kids [6 0 R 11 0 R] >>",
      /* 6 */ "<< /T (sec) /FT /Tx /Parent 5 0 R /Kids [7 0 R 8 0 R] >>",
      /* 7 */ widget("/Parent 6 0 R /T (a) /V (A)", 3, 650),
      /* 8 */ widget("/Parent 6 0 R /T (b) /V (B)", 4, 650),
      /* 9 */ content,
      /* 10 */ content,
      /* 11 */ widget("/Parent 5 0 R /FT /Tx /T (c) /V (C)", 4, 600)
    ]);
    expect(pageObjects(out)).toBe(1);
    expect(widgets(out)).toBe(1);
    expect(Object.keys(await fieldValues(latin1Bytes(out)))).toEqual(["form.sec.a"]);
  });

  it("a radio group across pages keeps only the kept page's button", async () => {
    const out = await removeSecondPage([
      /* 1 */ "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >>",
      /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      /* 3 */ page("6 0 R", 9),
      /* 4 */ page("7 0 R 8 0 R", 10),
      /* 5 */ "<< /T (choice) /FT /Btn /Ff 49152 /V /One /Kids [6 0 R 7 0 R 8 0 R] >>",
      /* 6 */ widget("/Parent 5 0 R /AS /One", 3, 650),
      /* 7 */ widget("/Parent 5 0 R /AS /Off", 4, 650),
      /* 8 */ widget("/Parent 5 0 R /AS /Off", 4, 600),
      /* 9 */ content,
      /* 10 */ content
    ]);
    expect(pageObjects(out)).toBe(1);
    expect(widgets(out)).toBe(1);
  });

  it("a widget listed on both pages survives, once, and its /P cannot pull the deleted page in", async () => {
    const out = await removeSecondPage([
      /* 1 */ "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R 6 0 R] >> >>",
      /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      /* 3 */ page("5 0 R", 7),
      /* 4 */ page("5 0 R 6 0 R", 8),
      /* 5 */ widget("/FT /Tx /T (both) /V (x)", 4, 650),
      /* 6 */ widget("/FT /Tx /T (gone)", 4, 600),
      /* 7 */ content,
      /* 8 */ content
    ]);
    expect(pageObjects(out)).toBe(1);
    expect(widgets(out)).toBe(1);
    expect(await fieldValues(latin1Bytes(out))).toEqual({ both: "x" });
  });

  it("a /Kids cycle terminates and still prunes", async () => {
    const out = await removeSecondPage([
      /* 1 */ "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >>",
      /* 2 */ "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      /* 3 */ page("7 0 R", 9),
      /* 4 */ page("8 0 R", 10),
      /* 5 */ "<< /T (loop) /Kids [6 0 R] >>",
      /* 6 */ "<< /T (inner) /Parent 5 0 R /Kids [5 0 R 7 0 R 8 0 R] >>",
      /* 7 */ widget("/Parent 6 0 R /FT /Tx /T (a)", 3, 650),
      /* 8 */ widget("/Parent 6 0 R /FT /Tx /T (b)", 4, 650),
      /* 9 */ content,
      /* 10 */ content
    ]);
    expect(pageObjects(out)).toBe(1);
    expect(widgets(out)).toBe(1);
  });

  it("a deep /Kids chain is pruned without exhausting the stack", async () => {
    const depth = 5000;
    // Objects: 1 catalog, 2 pages, 3/4 pages, 5/6 contents, 7/8 widgets, 9.. chain nodes.
    const first = 9;
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [9 0 R] >> >>",
      "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
      page("7 0 R", 5),
      page("8 0 R", 6),
      content,
      content,
      widget(`/Parent ${first + depth - 1} 0 R /FT /Tx /T (a)`, 3, 650),
      widget(`/Parent ${first + depth - 1} 0 R /FT /Tx /T (b)`, 4, 650)
    ];
    for (let i = 0; i < depth; i++) {
      const num = first + i;
      const parent = i === 0 ? "" : `/Parent ${num - 1} 0 R`;
      const kids = i === depth - 1 ? "[7 0 R 8 0 R]" : `[${num + 1} 0 R]`;
      bodies.push(`<< /T (n${i}) ${parent} /Kids ${kids} >>`);
    }
    const out = await removeSecondPage(bodies);
    expect(pageObjects(out)).toBe(1);
    expect(widgets(out)).toBe(1);
  });
});

function latin1Bytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) {
    bytes[i] = text.charCodeAt(i) & 0xff;
  }
  return bytes;
}

describe('PdfEditor — toBuffer({ incremental: "required" })', () => {
  it("appends when it can, exactly like incremental: true", async () => {
    const form = await makeForm();
    const out = await Pdf.Editor.read(form)
      .setFormField("fullName", "Ada")
      .toBuffer({ incremental: "required" });
    expect(out.subarray(0, form.length)).toEqual(form);
    expect((await fieldValues(out)).fullName).toBe("Ada");
  });

  it("refuses a structural edit instead of silently rewriting", async () => {
    const editor = Pdf.Editor.read(await makeForm());
    editor.addPage();
    await expect(editor.toBuffer({ incremental: "required" })).rejects.toBeInstanceOf(
      PdfIncrementalUpdateError
    );
    // `true` keeps falling back.
    expect((await editor.toBuffer({ incremental: true })).length).toBeGreaterThan(0);
  });
});

describe("PdfEditor — the AcroForm of a filled form", () => {
  const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
  // Catalog's AcroForm inline (`inline`) or as object 6.
  const form = (inline: boolean): Uint8Array => {
    const acroForm = "<< /Fields [4 0 R] /NeedAppearances false /XFA (xfa-packet) >>";
    return buildRawPdf([
      inline
        ? `<< /Type /Catalog /Pages 2 0 R /AcroForm ${acroForm} >>`
        : "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 5 0 R /Annots [4 0 R] >>",
      "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [60 650 300 672] /P 3 0 R >>",
      "<< /Length 0 >>\nstream\n\nendstream",
      acroForm
    ]);
  };
  // The AcroForm a reader sees: the last revision's.
  const lastAcroForm = (bytes: Uint8Array): string => {
    const text = latin1(bytes);
    const at = text.lastIndexOf("/Fields [");
    const open = text.lastIndexOf("<<", at);
    return text.slice(open, text.indexOf("endobj", at));
  };

  for (const inline of [true, false]) {
    for (const incremental of [false, true]) {
      it(`${inline ? "inline" : "indirect"} AcroForm, ${incremental ? "incremental" : "full"}: NeedAppearances true, XFA dropped`, async () => {
        const out = await Pdf.Editor.read(form(inline))
          .setFormField("name", "Ada")
          .toBuffer({ incremental });
        const acroForm = lastAcroForm(out);
        expect(acroForm).toContain("/NeedAppearances true");
        expect(acroForm).not.toContain("/XFA");
        expect((await fieldValues(out)).name).toBe("Ada");
      });
    }
  }

  it("a rewrite with no edited values keeps the XFA and the flag as they were", async () => {
    const out = await Pdf.Editor.read(form(true)).toBuffer();
    const acroForm = lastAcroForm(out);
    expect(acroForm).toContain("/NeedAppearances false");
    expect(acroForm).toContain("/XFA");
  });
});

describe("PdfEditor — XFA forms", () => {
  const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
  const page =
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 5 0 R /Annots [4 0 R] >>";
  const widget =
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (old) /Rect [60 650 300 672] /P 3 0 R >>";
  const content = "<< /Length 0 >>\nstream\n\nendstream";
  const hybrid = (): Uint8Array =>
    buildRawPdf([
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] /XFA (xfa-packet) >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      page,
      widget,
      content
    ]);
  const dynamic = (variant: "needsRendering" | "noFields"): Uint8Array =>
    buildRawPdf([
      variant === "needsRendering"
        ? "<< /Type /Catalog /Pages 2 0 R /NeedsRendering true /AcroForm << /Fields [4 0 R] /XFA (xfa-packet) >> >>"
        : "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [] /XFA (xfa-packet) >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      page,
      widget,
      content
    ]);

  for (const incremental of [false, true]) {
    it(`hybrid form, ${incremental ? "incremental" : "full"}: XFA removed, values visible`, async () => {
      const out = await Pdf.Editor.read(hybrid())
        .setFormField("name", "Ada")
        .toBuffer({ incremental });
      const text = latin1(out);
      const acroForm = text.slice(text.lastIndexOf("/Fields ["));
      expect(acroForm.slice(0, acroForm.indexOf("endobj"))).not.toContain("/XFA");
      expect((await fieldValues(out)).name).toBe("Ada");
    });
  }

  for (const variant of ["needsRendering", "noFields"] as const) {
    for (const incremental of [false, true]) {
      it(`dynamic XFA form (${variant}), ${incremental ? "incremental" : "full"}: refused, nothing written`, async () => {
        const editor = Pdf.Editor.read(dynamic(variant)).setFormField("name", "Ada");
        await expect(editor.toBuffer({ incremental })).rejects.toBeInstanceOf(PdfXfaFormError);
      });
    }
  }

  it("a dynamic XFA form saved without edits keeps its XFA and /NeedsRendering", async () => {
    for (const incremental of [false, true]) {
      const original = dynamic("needsRendering");
      const out = await Pdf.Editor.read(original).toBuffer({ incremental });
      const text = latin1(out);
      expect(text).toContain("/NeedsRendering true");
      expect(text).toMatch(/\/XFA/);
    }
  });

  it("a dynamic form without /NeedsRendering is recognised by its config packet", async () => {
    const config =
      "<config><acrobat><acrobat7><dynamicRender>required</dynamicRender></acrobat7></acrobat></config>";
    const bytes = buildRawPdf([
      "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] /XFA [(config) 6 0 R] >> >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      page,
      widget,
      content,
      `<< /Length ${config.length} >>\nstream\n${config}\nendstream`
    ]);
    const editor = Pdf.Editor.read(bytes).setFormField("name", "Ada");
    await expect(editor.toBuffer()).rejects.toBeInstanceOf(PdfXfaFormError);
  });

  it("a hybrid form saved without edits keeps its XFA", async () => {
    const text = latin1(await Pdf.Editor.read(hybrid()).toBuffer());
    expect(text).toMatch(/\/XFA/);
  });
});

// The deprecated save()/saveIncremental() forwards keep the behaviour they had before
// toBuffer() existed: no XFA refusal, XFA kept, and an incremental update appends only
// the edited objects. sign() is unchanged too — it predates toBuffer().
describe("PdfEditor — deprecated save()/saveIncremental() keep their behaviour", () => {
  const latin1 = (bytes: Uint8Array): string => new TextDecoder("latin1").decode(bytes);
  const widget =
    "<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (old) /Rect [60 650 300 672] /P 3 0 R >>";
  const build = (catalog: string): Uint8Array =>
    buildRawPdf([
      catalog,
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << >> /Contents 5 0 R /Annots [4 0 R] >>",
      widget,
      "<< /Length 0 >>\nstream\n\nendstream"
    ]);
  const dynamic = (): Uint8Array =>
    build(
      "<< /Type /Catalog /Pages 2 0 R /NeedsRendering true /AcroForm << /Fields [4 0 R] /XFA (xfa-packet) >> >>"
    );
  const hybrid = (): Uint8Array =>
    build("<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] /XFA (xfa-packet) >> >>");

  it("save() on a dynamic XFA form writes instead of throwing, keeping the XFA", async () => {
    const editor = Pdf.Editor.read(dynamic()).setFormField("name", "Ada");
    const text = latin1(await editor.save());
    expect(text).toContain("/XFA");
    expect(text).toContain("/NeedAppearances true");
  });

  it("save() on a hybrid form keeps the XFA", async () => {
    const editor = Pdf.Editor.read(hybrid()).setFormField("name", "Ada");
    const out = await editor.save();
    expect(latin1(out)).toContain("/XFA");
    expect((await fieldValues(out)).name).toBe("Ada");
  });

  it("saveIncremental() appends only the widget — no AcroForm or catalog revision, no throw", async () => {
    for (const original of [hybrid(), dynamic()]) {
      const editor = Pdf.Editor.read(original).setFormField("name", "Ada");
      const out = await editor.saveIncremental();
      const appended = latin1(out.subarray(original.length));
      expect(appended).not.toMatch(/\b1 0 obj/);
      expect(appended).not.toContain("/XFA");
      expect(appended).toMatch(/\b4 0 obj/);
    }
  });

  it("sign() on a dynamic XFA form with edited values does not throw", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate("XfaSign");
    const editor = Pdf.Editor.read(dynamic()).setFormField("name", "Ada");
    const out = await editor.sign({ certificate, privateKey });
    expect(latin1(out)).toContain("/XFA");
  });

  it("a concurrent toBuffer() keeps toBuffer()'s rules while save() is in flight", async () => {
    for (const legacy of ["save", "saveIncremental"] as const) {
      const editor = Pdf.Editor.read(dynamic()).setFormField("name", "Ada");
      const pending = editor[legacy]();
      await expect(editor.toBuffer()).rejects.toBeInstanceOf(PdfXfaFormError);
      expect(latin1(await pending)).toContain("Ada");
      // And the other way round: save() started second still writes.
      const refused = editor.toBuffer();
      await expect(editor[legacy]()).resolves.toBeInstanceOf(Uint8Array);
      await expect(refused).rejects.toBeInstanceOf(PdfXfaFormError);
    }
  });
});
