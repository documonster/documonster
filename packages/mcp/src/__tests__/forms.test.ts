/**
 * `form_fill` and `pdf_edit` tests.
 *
 * The single most important assertion in this file is that a filled PDF still
 * contains the values after being written and re-read. A full `editor.toBuffer()` once
 * silently discarded them (fixed in the editor, and pinned there by
 * `editor-form-fields.test.ts`); the tool saves incrementally so the original bytes
 * survive, and it re-reads its own output either way.
 */

import { mkdtemp, readFile, realpath, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { Pdf } from "documonster/pdf";
import { Build, Document, Query } from "documonster/word";
import { WordMarkdown } from "documonster/word/markdown";
import { vi } from "vitest";

import { resolveConfig, type ServerConfig } from "../config.js";
import { McpToolError } from "../errors.js";
import { saveKind } from "../tools/document.js";
import { formFillTool } from "../tools/form-fill.js";
import { pdfEditTool } from "../tools/pdf-edit.js";
import { signedXrefStreamPdf, xrefStreamForm, xrefStreamPages } from "./helpers/xref-stream-pdf.js";

interface Fixture {
  readonly config: ServerConfig;
  readonly root: string;
}

async function fixture(args: readonly string[] = []): Promise<Fixture> {
  const root = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-forms-")));
  const base = resolveConfig(args, { cwd: root });
  // Business-behaviour tests use the explicit compatibility layout; dedicated
  // sandbox regressions below cover the secure dual-root default.
  return {
    config: { ...base, outputRoot: root, allowInPlace: true },
    root
  };
}

async function run(
  tool: typeof formFillTool,
  fx: Fixture,
  args: Record<string, unknown>
): Promise<string> {
  const result = await tool.handler(args, { config: fx.config });
  const text = result.content
    .filter((block): block is { type: "text"; text: string } => block.type === "text")
    .map(block => block.text)
    .join("\n");
  if (result.isError === true) {
    throw new Error(text);
  }
  return text;
}

/** A Word form with a text field, a second text field and a checkbox. */
async function makeWordForm(fx: Fixture, name = "form.docx"): Promise<string> {
  const doc = Document.create();
  Document.addHeading(doc, "Application", 1);
  Document.addContent(doc, {
    type: "paragraph",
    children: [Build.text("Name: "), Build.formTextField({ name: "applicantName", default: "" })]
  });
  Document.addContent(doc, {
    type: "paragraph",
    children: [Build.text("Company: "), Build.formTextField({ name: "company", default: "" })]
  });
  Document.addContent(doc, {
    type: "paragraph",
    children: [
      Build.text("Agree: "),
      Build.formCheckboxField({ name: "agreeTerms", default: false })
    ]
  });
  await Document.writeFile(doc, path.join(fx.root, name));
  return name;
}

/** A PDF with an AcroForm: one text field and one checkbox. */
async function makePdfForm(fx: Fixture, name = "form.pdf"): Promise<string> {
  const builder = new Pdf.Builder();
  const page = builder.addPage();
  page.drawText("Application", { x: 60, y: 740, fontSize: 18 });
  page.addFormField({ type: "text", name: "fullName", rect: [60, 690, 300, 712] });
  page.addFormField({ type: "checkbox", name: "agree", rect: [60, 650, 76, 666] });
  await writeFile(path.join(fx.root, name), await builder.toBuffer());
  return name;
}

/** A one-page form whose only cross-reference data is a stream — see `helpers/xref-stream-pdf`. */
async function makeXrefStreamPdfForm(fx: Fixture, name = "xref-stream.pdf"): Promise<string> {
  await writeFile(path.join(fx.root, name), xrefStreamForm());
  return name;
}

/** A multi-page PDF for the structural operations. */
async function makeMultiPagePdf(fx: Fixture, name = "multi.pdf"): Promise<string> {
  const markdown = Array.from(
    { length: 4 },
    (_, section) =>
      `# Section ${section + 1}\n\n` +
      Array.from(
        { length: 30 },
        (_, index) =>
          `Paragraph ${index + 1} of section ${section + 1} with enough text to fill a line.`
      ).join("\n\n")
  ).join("\n\n");
  await writeFile(
    path.join(fx.root, name),
    await Pdf.fromDocx(await WordMarkdown.toDocx(markdown))
  );
  return name;
}

async function pdfFields(fx: Fixture, name: string): Promise<Record<string, unknown>> {
  const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, name))), {
    extractFormFields: true,
    extractText: false
  });
  return Object.fromEntries((result.formFields ?? []).map(field => [field.name, field.value]));
}

async function pageCount(fx: Fixture, name: string): Promise<number> {
  const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, name))), {
    extractText: false
  });
  return result.metadata?.pageCount ?? 0;
}

describe("form_fill — Word", () => {
  it("lists fields with their types and a values shape", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    const text = await run(formFillTool, fx, { path: file });

    expect(text).toContain("3 form field(s)");
    expect(text).toContain("`applicantName`");
    expect(text).toContain("checkBox");
    expect(text).toContain("## Values shape");
  });

  it("reads its values from a JSON file with valuesFrom", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await writeFile(
      path.join(fx.root, "values.json"),
      JSON.stringify({ applicantName: "From File", company: "Acme", agreeTerms: true })
    );
    await run(formFillTool, fx, { path: file, valuesFrom: "values.json" });
    const fields = Query.listFormFields(await Document.readFile(path.join(fx.root, file)));
    expect(Object.fromEntries(fields.map(f => [f.name, f.value]))).toEqual({
      applicantName: "From File",
      company: "Acme",
      agreeTerms: true
    });
  });

  it("fills text and checkbox fields", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await run(formFillTool, fx, {
      path: file,
      values: { applicantName: "Jane Doe", company: "Acme Pty Ltd", agreeTerms: true }
    });

    const fields = Query.listFormFields(await Document.readFile(path.join(fx.root, file)));
    expect(Object.fromEntries(fields.map(f => [f.name, f.value]))).toEqual({
      applicantName: "Jane Doe",
      company: "Acme Pty Ltd",
      agreeTerms: true
    });
  });

  it("fills only the named fields", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await run(formFillTool, fx, { path: file, values: { applicantName: "Only me" } });

    const fields = Query.listFormFields(await Document.readFile(path.join(fx.root, file)));
    expect(fields.find(f => f.name === "applicantName")?.value).toBe("Only me");
    expect(fields.find(f => f.name === "company")?.value).toBe("");
  });

  it("names the real fields when given an unknown one, and writes nothing", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    const before = await readFile(path.join(fx.root, file));

    await expect(
      formFillTool.handler({ path: file, values: { nope: "x" } }, { config: fx.config })
    ).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof McpToolError &&
        error.message.includes('"nope"') &&
        (error.hint ?? "").includes("applicantName") &&
        (error.hint ?? "").includes("Nothing was written"),
      "expected the error to list the real field names"
    );

    expect(await readFile(path.join(fx.root, file))).toEqual(before);
  });

  it("writes to `out` and leaves the original blank", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await run(formFillTool, fx, {
      path: file,
      values: { applicantName: "Jane" },
      out: "out/filled.docx"
    });

    const original = Query.listFormFields(await Document.readFile(path.join(fx.root, file)));
    expect(original.find(f => f.name === "applicantName")?.value).toBe("");
    const filled = Query.listFormFields(
      await Document.readFile(path.join(fx.root, "out/filled.docx"))
    );
    expect(filled.find(f => f.name === "applicantName")?.value).toBe("Jane");
  });

  it("backs up before filling in place", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await run(formFillTool, fx, { path: file, values: { applicantName: "Jane" } });
    await expect(stat(path.join(fx.root, `${file}.bak`))).resolves.toBeDefined();
  });

  it("redirects a document with no form fields", async () => {
    const fx = await fixture();
    const doc = Document.create();
    Document.addParagraph(doc, "No fields here.");
    await Document.writeFile(doc, path.join(fx.root, "plain.docx"));

    const text = await run(formFillTool, fx, { path: "plain.docx" });
    expect(text).toContain("no legacy form fields");
    expect(text).toContain("template_inspect");
  });
});

describe("form_fill — PDF", () => {
  it("lists AcroForm fields", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    const text = await run(formFillTool, fx, { path: file });

    expect(text).toContain("2 form field(s)");
    expect(text).toContain("`fullName`");
    expect(text).toContain("checkbox");
  });

  it("fills fields so the values survive being written and re-read", async () => {
    // The assertion that matters: the values are in the file that was written.
    const fx = await fixture();
    const file = await makePdfForm(fx);

    const report = await run(formFillTool, fx, {
      path: file,
      values: { fullName: "Jane Doe", agree: true }
    });

    expect(report).toContain("Verified by re-reading");
    expect(await pdfFields(fx, file)).toEqual({ fullName: "Jane Doe", agree: "Yes" });
  });

  it("appends the values as an incremental update, leaving the original bytes intact", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    const before = new Uint8Array(await readFile(path.join(fx.root, file)));

    const report = await run(formFillTool, fx, { path: file, values: { fullName: "Ada" } });

    expect(report).toContain("incremental update");
    const after = new Uint8Array(await readFile(path.join(fx.root, file)));
    expect(after.length).toBeGreaterThan(before.length);
    expect(after.subarray(0, before.length)).toEqual(before);
  });

  it("reports a save failure for what it is, not as a signature problem", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    const spy = vi
      .spyOn(Pdf.Editor.prototype, "toBuffer")
      .mockRejectedValueOnce(new Error("disk on fire"));
    try {
      const error = await formFillTool
        .handler({ path: file, values: { fullName: "Ada" } }, { config: fx.config })
        .catch((cause: unknown) => cause);
      expect(String((error as Error).message)).toContain("disk on fire");
      expect(JSON.stringify(error)).not.toMatch(/signature/i);
      expect(String((error as McpToolError).hint ?? "")).not.toMatch(/signature/i);
    } finally {
      spy.mockRestore();
    }
  });

  it("fills a form whose cross-reference data is a stream", async () => {
    const fx = await fixture();
    const file = await makeXrefStreamPdfForm(fx);
    const before = new Uint8Array(await readFile(path.join(fx.root, file)));
    expect(Buffer.from(before).toString("latin1")).not.toContain("trailer");

    const report = await run(formFillTool, fx, { path: file, values: { fullName: "Grace" } });

    expect(report).toContain("Verified by re-reading");
    expect(report).toContain("incremental update");
    expect(await pdfFields(fx, file)).toEqual({ fullName: "Grace" });
    const after = new Uint8Array(await readFile(path.join(fx.root, file)));
    expect(after.subarray(0, before.length)).toEqual(before);
  });

  it("maps a boolean to a checkbox's on and off states", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    await run(formFillTool, fx, { path: file, values: { agree: false } });
    expect((await pdfFields(fx, file)).agree).toBe("Off");
  });

  it("writes to `out` and leaves the original empty", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    await run(formFillTool, fx, { path: file, values: { fullName: "Jane" }, out: "out/f.pdf" });

    expect((await pdfFields(fx, file)).fullName).toBe("");
    expect((await pdfFields(fx, "out/f.pdf")).fullName).toBe("Jane");
  });

  it("rejects an unknown field without writing", async () => {
    const fx = await fixture();
    const file = await makePdfForm(fx);
    await expect(
      formFillTool.handler({ path: file, values: { nope: "x" } }, { config: fx.config })
    ).rejects.toThrow(/no such field/);
  });

  it("says plainly when a PDF has no form at all", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx, "plain.pdf");
    const text = await run(formFillTool, fx, { path: file });
    expect(text).toContain("no AcroForm fields");
    expect(text).toContain("no OCR");
  });
});

describe("form_fill — boundaries", () => {
  it("rejects a format it cannot fill", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "b.xlsx"), "x", "utf8");
    await expect(formFillTool.handler({ path: "b.xlsx" }, { config: fx.config })).rejects.toSatisfy(
      (error: unknown) =>
        error instanceof McpToolError && (error.hint ?? "").includes("sheet_edit"),
      "expected a redirect to sheet_edit"
    );
  });

  it("refuses an `out` in a different format", async () => {
    const fx = await fixture();
    const file = await makeWordForm(fx);
    await expect(
      formFillTool.handler(
        { path: file, values: { applicantName: "x" }, out: "out/f.pdf" },
        { config: fx.config }
      )
    ).rejects.toThrow(/same format/);
  });

  it("is withheld under --readonly for filling but not for listing", async () => {
    const fx = await fixture(["--readonly"]);
    const file = await makeWordForm(fx);
    // Listing still works…
    expect(await run(formFillTool, fx, { path: file })).toContain("form field(s)");
    // …but filling does not.
    await expect(
      formFillTool.handler({ path: file, values: { applicantName: "x" } }, { config: fx.config })
    ).rejects.toSatisfy(
      (error: unknown) => error instanceof McpToolError && error.code === "readonly",
      "expected a readonly error"
    );
  });

  it("cannot escape the sandbox root", async () => {
    const fx = await fixture();
    await expect(
      formFillTool.handler({ path: "../../etc/hosts" }, { config: fx.config })
    ).rejects.toThrow(/outside (?:the server root|--output-root)/);
  });
});

describe("pdf_edit — overlays", () => {
  it("adds a watermark over the original content", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "watermark", text: "CONFIDENTIAL" }]
    });

    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, file))), {
      extractText: true
    });
    expect(result.pages[0]?.text).toContain("CONFIDENTIAL");
    // The original content must still be there — overlays never rewrite it.
    expect(result.pages[0]?.text).toContain("Section 1");
  });

  it("watermarks only the named pages", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "watermark", text: "DRAFT", pages: [1] }]
    });

    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, file))), {
      extractText: true
    });
    expect(result.pages[0]?.text).toContain("DRAFT");
    expect(result.pages[1]?.text).not.toContain("DRAFT");
  });

  it("adds page numbers using the template", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const total = await pageCount(fx, file);
    await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "page_numbers", format: "{page} / {total}" }]
    });

    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, file))), {
      extractText: true
    });
    expect(result.pages[0]?.text).toContain(`1 / ${total}`);
    expect(result.pages[1]?.text).toContain(`2 / ${total}`);
  });

  it("stamps text at a fixed position", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "stamp", text: "APPROVED", x: 400, y: 60, pages: [1] }]
    });

    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, file))), {
      extractText: true
    });
    expect(result.pages[0]?.text).toContain("APPROVED");
  });

  it("rejects a bad colour", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await expect(
      pdfEditTool.handler(
        { path: file, ops: [{ op: "watermark", text: "X", color: "red" }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/not a hex colour/);
  });
});

describe("pdf_edit — structure", () => {
  it("deletes pages", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const before = await pageCount(fx, file);

    await run(pdfEditTool, fx, { path: file, ops: [{ op: "delete_pages", pages: [2] }] });
    expect(await pageCount(fx, file)).toBe(before - 1);
  });

  it("keeps only the named pages", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await run(pdfEditTool, fx, { path: file, ops: [{ op: "keep_pages", pages: "1-2" }] });
    expect(await pageCount(fx, file)).toBe(2);
  });

  it("refuses to delete every page", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const total = await pageCount(fx, file);
    await expect(
      pdfEditTool.handler(
        { path: file, ops: [{ op: "delete_pages", pages: `1-${total}` }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/every page/);
  });

  it("rotates pages", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const report = await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "rotate", degrees: 90, pages: [1] }]
    });
    expect(report).toContain("set 1 page(s) to 90° rotation");
  });

  it("appends another PDF", async () => {
    const fx = await fixture();
    const first = await makeMultiPagePdf(fx, "a.pdf");
    const second = await makeMultiPagePdf(fx, "b.pdf");
    const before = await pageCount(fx, first);

    await run(pdfEditTool, fx, {
      path: first,
      ops: [{ op: "append", path: second, pages: [1] }]
    });
    expect(await pageCount(fx, first)).toBe(before + 1);
  });

  it("removes a page appended earlier in the same call", async () => {
    const fx = await fixture();
    const first = await makeMultiPagePdf(fx, "a.pdf");
    const second = await makeMultiPagePdf(fx, "b.pdf");
    const before = await pageCount(fx, first);

    const report = await run(pdfEditTool, fx, {
      path: first,
      ops: [
        { op: "append", path: second, pages: [1, 2] },
        { op: "delete_pages", pages: [before + 1] }
      ]
    });
    expect(report).toContain(`deleted page(s) ${before + 1}`);
    expect(report).toContain(`pages: ${before} → ${before + 1}`);
    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, first))), {
      extractText: true
    });
    expect(result.pages).toHaveLength(before + 1);
    expect(result.pages[before]?.text).toContain("Section 2");
  });

  it("rotates a page appended earlier in the same call", async () => {
    const fx = await fixture();
    const first = await makeMultiPagePdf(fx, "a.pdf");
    const second = await makeMultiPagePdf(fx, "b.pdf");
    const before = await pageCount(fx, first);

    await run(pdfEditTool, fx, {
      path: first,
      ops: [
        { op: "append", path: second, pages: [1] },
        { op: "rotate", degrees: 90, pages: [before + 1] }
      ]
    });
    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, first))));
    expect(result.pages).toHaveLength(before + 1);
    expect(result.pages[before]?.rotation).toBe(90);
    expect(result.pages[0]?.rotation).toBe(0);
  });

  it("stamps text on a page appended earlier in the same call", async () => {
    const fx = await fixture();
    const first = await makeMultiPagePdf(fx, "a.pdf");
    const second = await makeMultiPagePdf(fx, "b.pdf");
    const before = await pageCount(fx, first);

    await run(pdfEditTool, fx, {
      path: first,
      ops: [
        { op: "append", path: second, pages: [1] },
        { op: "stamp", text: "APPENDIX", x: 400, y: 60, pages: [before + 1] },
        { op: "page_numbers" }
      ]
    });
    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, first))), {
      extractText: true
    });
    expect(result.pages[before]?.text).toContain("APPENDIX");
    expect(result.pages[before]?.text).toContain("Section 1");
    expect(result.pages[before]?.text).toContain(`Page ${before + 1} of ${before + 1}`);
    expect(result.pages[0]?.text).not.toContain("APPENDIX");
  });

  it("appends a form with its fields, and refuses a field name the PDF already has", async () => {
    const fx = await fixture();
    const host = await makeMultiPagePdf(fx, "host.pdf");
    const form = await makePdfForm(fx, "form.pdf");
    const before = await pageCount(fx, host);

    await run(pdfEditTool, fx, { path: host, ops: [{ op: "append", path: form }] });
    const merged = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, host))));
    expect(merged.pages).toHaveLength(before + 1);
    expect(merged.formFields.map(field => field.name).sort()).toEqual(["agree", "fullName"]);

    // The same form again clashes, by name, and the error says how to proceed.
    await expect(
      pdfEditTool.handler(
        { path: host, ops: [{ op: "append", path: form }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/formFields: "drop"/);
    expect(await pageCount(fx, host)).toBe(before + 1);
  });

  it('appends the same form twice with formFields: "drop"', async () => {
    const fx = await fixture();
    const form = await makePdfForm(fx, "form.pdf");
    const report = await run(pdfEditTool, fx, {
      path: form,
      ops: [{ op: "append", path: form, formFields: "drop" }]
    });
    expect(report).toContain("without their form fields");
    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, form))));
    expect(result.pages).toHaveLength(2);
    expect(result.pages[1]?.text).toContain("Application");
    // Only the original's fields: the appended copy brought none.
    expect(result.formFields.map(field => field.name).sort()).toEqual(["agree", "fullName"]);
    expect(result.pages[1]?.annotations ?? []).toHaveLength(0);
  });

  it("applies several ops in order", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const report = await run(pdfEditTool, fx, {
      path: file,
      ops: [
        { op: "keep_pages", pages: "1-2" },
        { op: "watermark", text: "DRAFT" },
        { op: "page_numbers" }
      ]
    });

    expect(report).toContain("3 operation(s) applied");
    expect(await pageCount(fx, file)).toBe(2);
    const result = await Pdf.extract(new Uint8Array(await readFile(path.join(fx.root, file))), {
      extractText: true
    });
    // Page numbering must see the post-deletion count, not the original.
    expect(result.pages[0]?.text).toContain("Page 1 of 2");
  });

  it("rejects a page that does not exist", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await expect(
      pdfEditTool.handler(
        { path: file, ops: [{ op: "watermark", text: "X", pages: [99] }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/do not exist/);
  });
});

describe("pdf_edit — cross-reference-stream files", () => {
  async function place(fx: Fixture, name: string, bytes: Uint8Array): Promise<Uint8Array> {
    await writeFile(path.join(fx.root, name), bytes);
    expect(Buffer.from(bytes).toString("latin1")).not.toContain("trailer");
    return bytes;
  }
  const readBack = async (fx: Fixture, name: string): Promise<Uint8Array> =>
    new Uint8Array(await readFile(path.join(fx.root, name)));

  it("unsigned, overlay: reports and performs an incremental update", async () => {
    const fx = await fixture();
    const before = await place(fx, "x.pdf", xrefStreamPages(2));
    const report = await run(pdfEditTool, fx, {
      path: "x.pdf",
      ops: [{ op: "stamp", text: "APPROVED", x: 72, y: 72 }]
    });
    expect(report).toContain("saved as an incremental update");
    expect(report).not.toContain("rebuilt");
    const after = await readBack(fx, "x.pdf");
    expect(after.subarray(0, before.length)).toEqual(before);
  });

  it("unsigned, structural: reports a rebuild, and makes no signature claim it cannot back", async () => {
    const fx = await fixture();
    const before = await place(fx, "x.pdf", xrefStreamPages(2));
    const report = await run(pdfEditTool, fx, {
      path: "x.pdf",
      ops: [{ op: "delete_pages", pages: [2] }]
    });
    expect(report).toContain("structural change, so the file was rebuilt");
    expect(report).not.toContain("signature is now invalid");
    const after = await readBack(fx, "x.pdf");
    expect(after.subarray(0, before.length)).not.toEqual(before);
    expect(await pageCount(fx, "x.pdf")).toBe(1);
  });

  it("signed, overlay: appends, and the signature still verifies", async () => {
    const fx = await fixture();
    const before = await place(fx, "s.pdf", await signedXrefStreamPdf());
    const [original] = await Pdf.verifySignatures(before);
    expect(original?.intact).toBe(true);
    expect(original?.coversWholeFile).toBe(true);

    const report = await run(pdfEditTool, fx, {
      path: "s.pdf",
      ops: [{ op: "watermark", text: "COPY" }]
    });
    expect(report).toContain("original signature bytes are preserved");

    const after = await readBack(fx, "s.pdf");
    const reports = await Pdf.verifySignatures(after);
    expect(reports).toHaveLength(1);
    // What the report claims is what verification finds: the signed revision is
    // intact, and the watermark is an unsigned update after it.
    expect(reports[0]?.intact).toBe(true);
    expect(reports[0]?.byteRangeValid).toBe(true);
    expect(reports[0]?.coversWholeFile).toBe(false);
  });

  it("signed, structural: refuses, writes nothing, leaves no backup", async () => {
    const fx = await fixture();
    const before = await place(fx, "s.pdf", await signedXrefStreamPdf());
    await expect(
      pdfEditTool.handler(
        { path: "s.pdf", ops: [{ op: "delete_pages", pages: [2] }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/digitally signed/);
    const after = await readBack(fx, "s.pdf");
    expect(after).toEqual(before);
    expect((await Pdf.verifySignatures(after))[0]?.intact).toBe(true);
    await expect(stat(path.join(fx.root, "s.pdf.bak"))).rejects.toThrow();
  });
});

describe("pdf_edit — safety", () => {
  it("dryRun writes nothing", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const total = await pageCount(fx, file);
    const before = await readFile(path.join(fx.root, file));

    const report = await run(pdfEditTool, fx, {
      path: file,
      ops: [{ op: "delete_pages", pages: [1] }],
      dryRun: true
    });

    expect(report).toContain("Dry run");
    expect(report).toContain(`pages: ${total} → ${total - 1}`);
    expect(await readFile(path.join(fx.root, file))).toEqual(before);
    await expect(stat(path.join(fx.root, `${file}.bak`))).rejects.toThrow();
  });

  it("backs up an in-place edit and can write elsewhere instead", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    const before = await pageCount(fx, file);

    await run(pdfEditTool, fx, { path: file, ops: [{ op: "delete_pages", pages: [1] }] });
    await expect(stat(path.join(fx.root, `${file}.bak`))).resolves.toBeDefined();

    const other = await makeMultiPagePdf(fx, "c.pdf");
    await run(pdfEditTool, fx, {
      path: other,
      ops: [{ op: "delete_pages", pages: [1] }],
      out: "out/trimmed.pdf"
    });
    expect(await pageCount(fx, other)).toBe(before);
    expect(await pageCount(fx, "out/trimmed.pdf")).toBe(before - 1);
  });

  it("rejects a non-PDF", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "fake.pdf"), "not a pdf", "utf8");
    await expect(
      pdfEditTool.handler(
        { path: "fake.pdf", ops: [{ op: "page_numbers" }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/could not read fake\.pdf as a PDF/);
  });

  it("is withheld under --readonly", async () => {
    const fx = await fixture(["--readonly"]);
    await expect(
      pdfEditTool.handler({ path: "a.pdf", ops: [{ op: "page_numbers" }] }, { config: fx.config })
    ).rejects.toSatisfy(
      (error: unknown) => error instanceof McpToolError && error.code === "readonly",
      "expected a readonly error"
    );
  });

  it("cannot read or write outside the sandbox root", async () => {
    const fx = await fixture();
    const file = await makeMultiPagePdf(fx);
    await expect(
      pdfEditTool.handler(
        { path: file, ops: [{ op: "append", path: "../../etc/hosts" }] },
        { config: fx.config }
      )
    ).rejects.toThrow(/outside (?:the server root|--output-root)/);
    await expect(
      pdfEditTool.handler(
        { path: file, ops: [{ op: "page_numbers" }], out: "../escape.pdf" },
        { config: fx.config }
      )
    ).rejects.toThrow(/outside (?:the server root|--output-root)/);
  });
});

describe("saveKind", () => {
  const original = Uint8Array.from([1, 2, 3]);
  it("tells an unchanged file from an appended and a rewritten one", () => {
    expect(saveKind(original, Uint8Array.from([1, 2, 3]))).toBe("unchanged");
    expect(saveKind(original, Uint8Array.from([1, 2, 3, 4]))).toBe("appended");
    expect(saveKind(original, Uint8Array.from([1, 9, 3, 4]))).toBe("rewritten");
    expect(saveKind(original, Uint8Array.from([1, 2]))).toBe("rewritten");
  });
});
