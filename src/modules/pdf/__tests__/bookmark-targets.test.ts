/**
 * A bookmark names a page by index, and a page can be added after the bookmark that points
 * at it — so the target is checked when the document is written. One that names no page was
 * silently left out of the outline; it is a `PdfArgumentError` naming the bookmark.
 */
import { Pdf, PdfArgumentError } from "@pdf/index";
import { describe, expect, it } from "vitest";

describe("bookmark targets", () => {
  it("refuses a bookmark to a page the document does not have, naming it", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.addBookmark("Intro", 0).addBookmark("Appendix", 3);
    const err = await doc.toBuffer().catch((e: unknown) => e);
    expect(err).toBeInstanceOf(PdfArgumentError);
    expect((err as PdfArgumentError).field).toBe("pageIndex");
    expect((err as PdfArgumentError).message).toContain('"Appendix"');
    expect((err as PdfArgumentError).message).toContain("3");
  });

  it("checks nested bookmarks too", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.addBookmark("Part", 0).addBookmark("Chapter", 1, 0);
    await expect(doc.toBuffer()).rejects.toThrow(/"Chapter"/);
  });

  it("accepts a bookmark added before the page it points at", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.addBookmark("Later", 1);
    doc.addPage();
    const text = new TextDecoder("latin1").decode(await doc.toBuffer());
    expect(text).toContain("/Title (Later)");
  });

  it("accepts the document again once the missing page exists", async () => {
    const doc = new Pdf.Builder();
    doc.addPage();
    doc.addBookmark("Second", 1);
    await expect(doc.toBuffer()).rejects.toBeInstanceOf(PdfArgumentError);
    doc.addPage();
    expect((await doc.toBuffer()).length).toBeGreaterThan(0);
  });
});
