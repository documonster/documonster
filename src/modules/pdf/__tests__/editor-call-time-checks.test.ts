/**
 * Refusals happen at the call that introduces the problem: a form-field name already in use
 * (builder, editor, copyPagesFrom), and an edit of a page that has been removed.
 */
import { describe, expect, it } from "vitest";

import { Pdf, PdfArgumentError } from "../index";

const field = (name: string, y = 10) =>
  ({ type: "text", name, rect: [10, y, 100, y + 20] }) as const;

async function withFields(...names: string[]): Promise<Uint8Array> {
  const b = new Pdf.Builder();
  const page = b.addPage();
  names.forEach((name, i) => page.addFormField(field(name, 10 + i * 30)));
  return b.toBuffer();
}

async function blank(pages = 1): Promise<Uint8Array> {
  const b = new Pdf.Builder();
  for (let i = 0; i < pages; i++) {
    b.addPage();
  }
  return b.toBuffer();
}

function refusal(fn: () => unknown): PdfArgumentError {
  try {
    fn();
  } catch (err) {
    expect(err).toBeInstanceOf(PdfArgumentError);
    return err as PdfArgumentError;
  }
  throw new Error("expected a PdfArgumentError");
}

describe("a duplicate form-field name is refused by the call that adds it", () => {
  it("Builder#addFormField, against a field on another page", () => {
    const b = new Pdf.Builder();
    b.addPage().addFormField(field("a"));
    const err = refusal(() => b.addPage().addFormField(field("a")));
    expect(err.field).toBe("options.name");
    expect(err.message).toMatch(/^Pdf\.PageBuilder#addFormField:/);
  });

  it("EditorPage#addFormField, against the document's own fields", async () => {
    const editor = Pdf.Editor.read(await withFields("existing"));
    const err = refusal(() => editor.getPage(0).addFormField(field("existing", 50)));
    expect(err.message).toMatch(/^Pdf\.EditorPage#addFormField: the document already has/);
  });

  it("EditorPage#addFormField and an added page, against fields already added", async () => {
    const editor = Pdf.Editor.read(await blank(2));
    editor.getPage(0).addFormField(field("added"));
    expect(() => editor.getPage(1).addFormField(field("added"))).toThrow(
      /^Pdf\.EditorPage#addFormField:/
    );
    expect(() => editor.addPage().addFormField(field("added"))).toThrow(
      /^Pdf\.PageBuilder#addFormField:/
    );
  });

  it("a field on a removed page no longer holds its name", async () => {
    const editor = Pdf.Editor.read(await blank(2));
    editor.getPage(0).addFormField(field("moved"));
    editor.removePage(0);
    editor.getPage(1).addFormField(field("moved"));
    const out = await Pdf.extract(await editor.toBuffer());
    expect(out.formFields.map(f => f.name)).toEqual(["moved"]);
  });

  it("copyPagesFrom, against the document's fields — nothing is copied", async () => {
    const editor = Pdf.Editor.read(await withFields("same"));
    const source = await withFields("same");
    const err = refusal(() => editor.copyPagesFrom(source));
    expect(err.message).toMatch(/^Pdf\.Editor#copyPagesFrom:/);
    const out = await Pdf.extract(await editor.toBuffer());
    expect(out.pages).toHaveLength(1);
  });

  it("copyPagesFrom, against an earlier copy, an added field, and itself", async () => {
    const source = await withFields("guest");
    const editor = Pdf.Editor.read(await blank());
    editor.copyPagesFrom(source);
    expect(() => editor.copyPagesFrom(source)).toThrow(/^Pdf\.Editor#copyPagesFrom:/);
    expect(() => editor.getPage(0).addFormField(field("guest"))).toThrow(
      /^Pdf\.EditorPage#addFormField:/
    );

    const other = Pdf.Editor.read(await blank());
    other.getPage(0).addFormField(field("guest"));
    expect(() => other.copyPagesFrom(source)).toThrow(/^Pdf\.Editor#copyPagesFrom:/);

    const twice = Pdf.Editor.read(await blank());
    expect(() => twice.copyPagesFrom(source, { pageIndices: [0, 0] })).toThrow(
      /^Pdf\.Editor#copyPagesFrom:/
    );
    expect((await Pdf.extract(await twice.toBuffer())).pages).toHaveLength(1);
  });
});

describe("an edit of a removed page is refused (X8)", () => {
  it("a held page handle refuses drawing once its page is removed", async () => {
    const editor = Pdf.Editor.read(await blank(2));
    const page = editor.getPage(0);
    editor.removePage(0);
    expect(refusal(() => page.drawText("x", { x: 1, y: 1 })).message).toMatch(/removed/);
    expect(() => page.drawRect({ x: 0, y: 0, width: 5, height: 5 })).toThrow(PdfArgumentError);
    expect(() => page.addAnnotation({ type: "Text", rect: [0, 0, 5, 5] })).toThrow(
      PdfArgumentError
    );
    expect(() => page.addFormField(field("f"))).toThrow(PdfArgumentError);
    expect(() => page.getContentStream()).toThrow(PdfArgumentError);
    // Measuring changes nothing and is still allowed.
    expect(page.measureText("x", { fontSize: 10 })).toBeGreaterThan(0);
  });

  it("getPage and rotatePage refuse a removed page", async () => {
    const editor = Pdf.Editor.read(await blank(2));
    editor.removePage(0);
    expect(refusal(() => editor.getPage(0)).field).toBe("index");
    expect(refusal(() => editor.rotatePage(0, 90)).field).toBe("index");
    expect(editor.isPageRemoved(0)).toBe(true);
    editor.getPage(1).drawText("kept", { x: 10, y: 10 });
    expect((await Pdf.extract(await editor.toBuffer())).pages).toHaveLength(1);
  });
});

describe("the PDF readers accept an ArrayBuffer (X5)", () => {
  async function arrayBuffer(): Promise<ArrayBuffer> {
    const b = new Pdf.Builder();
    b.addPage().drawText("hello", { x: 10, y: 10 });
    const bytes = await b.toBuffer();
    // A buffer holding exactly the file, as `fetch().arrayBuffer()` gives one.
    return bytes.slice().buffer as ArrayBuffer;
  }

  it("Pdf.extract, Pdf.Editor.read, copyPagesFrom and Pdf.verifySignatures", async () => {
    const buffer = await arrayBuffer();
    expect((await Pdf.extract(buffer)).pages[0].text).toContain("hello");
    const editor = Pdf.Editor.read(buffer);
    expect(editor.pageCount).toBe(1);
    editor.copyPagesFrom(buffer);
    expect((await Pdf.extract(await editor.toBuffer())).pages).toHaveLength(2);
    expect(await Pdf.verifySignatures(buffer)).toEqual([]);
  });

  it("still refuses what is not bytes", async () => {
    await expect(Pdf.extract("x" as never)).rejects.toThrow(PdfArgumentError);
    expect(() => Pdf.Editor.read(42 as never)).toThrow(PdfArgumentError);
  });
});

describe('copyPagesFrom formFields: "drop" keeps the old capability', () => {
  async function form(name: string): Promise<Uint8Array> {
    const b = new Pdf.Builder();
    const page = b.addPage();
    page.drawText(`page of ${name}`, { x: 10, y: 200 });
    page.addFormField(field(name));
    page.addAnnotation({ type: "Text", rect: [10, 60, 30, 80], contents: `note ${name}` });
    return b.toBuffer();
  }

  it("the clash error names the option", async () => {
    const editor = Pdf.Editor.read(await form("same"));
    const source = await form("same");
    expect(refusal(() => editor.copyPagesFrom(source)).message).toMatch(/formFields: "drop"/);
  });

  it("copies a clashing form's pages without its fields, keeping its other annotations", async () => {
    const source = await form("same");
    const editor = Pdf.Editor.read(source);
    editor.copyPagesFrom(source, { formFields: "drop" });
    editor.copyPagesFrom(source, { formFields: "drop", pageIndices: [0, 0] });
    const out = await Pdf.extract(await editor.toBuffer());
    expect(out.pages).toHaveLength(4);
    expect(out.pages[1].text).toContain("page of same");
    expect(out.formFields.map(f => f.name)).toEqual(["same"]);
    for (const page of out.pages.slice(1)) {
      expect(page.annotations.map(a => a.subtype)).toEqual(["Text"]);
      expect(page.annotations[0].contents).toBe("note same");
    }
  });

  it('dropped names hold nothing; "copy" is the default and is checked', async () => {
    const editor = Pdf.Editor.read(await blank());
    editor.copyPagesFrom(await form("guest"), { formFields: "drop" });
    editor.getPage(0).addFormField(field("guest"));
    expect((await Pdf.extract(await editor.toBuffer())).formFields.map(f => f.name)).toEqual([
      "guest"
    ]);
    expect(
      refusal(() => editor.copyPagesFrom(new Uint8Array(1), { formFields: "x" as never })).field
    ).toBe("options.formFields");
  });
});
