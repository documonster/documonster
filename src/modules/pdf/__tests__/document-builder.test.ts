import { describe, it, expect } from "vitest";
import { PdfDocumentBuilder, PdfEditor, readPdf, pdf } from "../index";

// =============================================================================
// PdfDocumentBuilder — Free Text & Vector Drawing
// =============================================================================

describe("PdfDocumentBuilder", () => {
  it("should create a blank PDF with one page", async () => {
    const doc = new PdfDocumentBuilder();
    doc.addPage();

    const bytes = await doc.build();
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes.length).toBeGreaterThan(0);

    // Should be a valid PDF
    const header = new TextDecoder().decode(bytes.subarray(0, 8));
    expect(header).toBe("%PDF-2.0");

    // Should be readable
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should create a PDF with custom page size", async () => {
    const doc = new PdfDocumentBuilder();
    doc.addPage({ width: 612, height: 792 }); // US Letter

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages[0].width).toBe(612);
    expect(result.pages[0].height).toBe(792);
  });

  it("should create a PDF with multiple pages", async () => {
    const doc = new PdfDocumentBuilder();
    doc.addPage();
    doc.addPage();
    doc.addPage();

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(3);
  });

  it("should draw text at specific coordinates", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawText("Hello World", { x: 72, y: 750, fontSize: 24 });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.text).toContain("Hello World");
  });

  it("should draw text with different fonts", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawText("Normal text", { x: 72, y: 750 });
    page.drawText("Bold text", { x: 72, y: 720, bold: true });
    page.drawText("Italic text", { x: 72, y: 690, italic: true });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.text).toContain("Normal text");
    expect(result.text).toContain("Bold text");
    expect(result.text).toContain("Italic text");
  });

  it("should draw text with word-wrap", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    const longText =
      "This is a long piece of text that should be wrapped to fit within the specified maximum width";
    page.drawText(longText, { x: 72, y: 750, fontSize: 12, maxWidth: 200 });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    // Text should be present (may be split across lines)
    expect(result.text).toContain("This is a long");
  });

  it("should measure text width", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    const width = page.measureText("Hello", { fontSize: 12 });
    expect(width).toBeGreaterThan(0);
    expect(width).toBeLessThan(100); // Reasonable for "Hello" at 12pt

    // Wider text should have larger width
    const widerWidth = page.measureText("Hello World", { fontSize: 12 });
    expect(widerWidth).toBeGreaterThan(width);
  });

  it("should draw a rectangle", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawRect({
      x: 72,
      y: 700,
      width: 200,
      height: 50,
      fill: { r: 1, g: 0, b: 0 },
      stroke: { r: 0, g: 0, b: 0 },
      lineWidth: 2
    });

    const bytes = await doc.build();
    expect(bytes.length).toBeGreaterThan(0);

    // Verify it's a valid PDF
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a rounded rectangle", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawRect({
      x: 72,
      y: 700,
      width: 200,
      height: 50,
      fill: { r: 0, g: 0.5, b: 1 },
      borderRadius: 10
    });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a circle", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawCircle({
      cx: 300,
      cy: 400,
      r: 50,
      fill: { r: 0, g: 1, b: 0 },
      stroke: { r: 0, g: 0, b: 0 }
    });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw an ellipse", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawEllipse({
      cx: 300,
      cy: 400,
      rx: 100,
      ry: 50,
      fill: { r: 1, g: 1, b: 0 }
    });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a line", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawLine({
      x1: 72,
      y1: 750,
      x2: 500,
      y2: 750,
      color: { r: 0, g: 0, b: 0 },
      lineWidth: 2
    });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a dashed line", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawLine({
      x1: 72,
      y1: 700,
      x2: 500,
      y2: 700,
      color: { r: 1, g: 0, b: 0 },
      lineWidth: 1,
      dashPattern: [5, 3]
    });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a complex path", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    // Draw a triangle
    page.drawPath(
      [
        { op: "move", x: 300, y: 600 },
        { op: "line", x: 250, y: 500 },
        { op: "line", x: 350, y: 500 },
        { op: "close" }
      ],
      {
        fill: { r: 0, g: 0, b: 1 },
        stroke: { r: 0, g: 0, b: 0 },
        lineWidth: 2
      }
    );

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should draw a path with Bezier curves", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    page.drawPath(
      [
        { op: "move", x: 100, y: 400 },
        { op: "curve", x1: 200, y1: 500, x2: 300, y2: 300, x3: 400, y3: 400 }
      ],
      { stroke: { r: 1, g: 0, b: 0 }, lineWidth: 2 }
    );

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });

  it("should set document metadata", async () => {
    const doc = new PdfDocumentBuilder();
    doc.setMetadata({ title: "Test Doc", author: "Test Author" });
    doc.addPage();

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.metadata.title).toBe("Test Doc");
    expect(result.metadata.author).toBe("Test Author");
  });

  it("should support encryption", async () => {
    const doc = new PdfDocumentBuilder();
    doc.setEncryption({
      ownerPassword: "owner123",
      userPassword: "user123"
    });
    const page = doc.addPage();
    page.drawText("Secret content", { x: 72, y: 750 });

    const bytes = await doc.build();

    // Should fail without password
    await expect(readPdf(bytes)).rejects.toThrow();

    // Should succeed with password
    const result = await readPdf(bytes, { password: "user123" });
    expect(result.text).toContain("Secret content");
  });

  it("should combine text and shapes on the same page", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    // Draw a background rect
    page.drawRect({ x: 50, y: 700, width: 500, height: 80, fill: { r: 0.9, g: 0.9, b: 1 } });

    // Draw text on top
    page.drawText("Title", { x: 72, y: 740, fontSize: 24, bold: true });
    page.drawText("Subtitle", { x: 72, y: 710, fontSize: 14 });

    // Draw a divider line
    page.drawLine({ x1: 50, y1: 695, x2: 550, y2: 695, color: { r: 0, g: 0, b: 0.5 } });

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.text).toContain("Title");
    expect(result.text).toContain("Subtitle");
  });

  it("should expose raw content stream", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    // Use the raw content stream for custom operations
    const stream = page.getContentStream();
    stream.save();
    stream.setFillColor({ r: 1, g: 0, b: 0 });
    stream.circle(300, 400, 50);
    stream.fill();
    stream.restore();

    const bytes = await doc.build();
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });
});

// =============================================================================
// PdfEditor — Modify Existing PDF
// =============================================================================

describe("PdfEditor", () => {
  it("should load an existing PDF", async () => {
    const pdfBytes = await pdf([["Hello", "World"]]);
    const editor = PdfEditor.load(pdfBytes);

    expect(editor.pageCount).toBeGreaterThan(0);
  });

  it("should overlay text on an existing page", async () => {
    const pdfBytes = await pdf([["Original"]]);
    const editor = PdfEditor.load(pdfBytes);

    editor.getPage(0).drawText("Overlay", {
      x: 200,
      y: 400,
      fontSize: 24,
      color: { r: 1, g: 0, b: 0 }
    });

    const result = await editor.save();
    const readResult = await readPdf(result);
    expect(readResult.text).toContain("Overlay");
  });

  it("should overlay shapes on an existing page", async () => {
    const pdfBytes = await pdf([["Data"]]);
    const editor = PdfEditor.load(pdfBytes);

    editor.getPage(0).drawRect({
      x: 72,
      y: 300,
      width: 200,
      height: 100,
      fill: { r: 1, g: 1, b: 0 },
      stroke: { r: 0, g: 0, b: 0 }
    });

    const result = await editor.save();
    expect(result).toBeInstanceOf(Uint8Array);
    expect(result.length).toBeGreaterThan(0);
  });

  it("should add new pages to an existing PDF", async () => {
    const pdfBytes = await pdf([["Page 1"]]);
    const editor = PdfEditor.load(pdfBytes);

    const newPage = editor.addPage();
    newPage.drawText("Page 2", { x: 72, y: 750, fontSize: 20 });

    const result = await editor.save();
    const readResult = await readPdf(result);
    expect(readResult.pages.length).toBe(2);
    expect(readResult.pages[1].text).toContain("Page 2");
  });

  it("should throw for invalid page index", async () => {
    const pdfBytes = await pdf([["Test"]]);
    const editor = PdfEditor.load(pdfBytes);

    expect(() => editor.getPage(-1)).toThrow();
    expect(() => editor.getPage(100)).toThrow();
  });

  it("should get page dimensions", async () => {
    const pdfBytes = await pdf([["Test"]]);
    const editor = PdfEditor.load(pdfBytes);

    const page = editor.getPage(0);
    expect(page.width).toBeGreaterThan(0);
    expect(page.height).toBeGreaterThan(0);
  });

  it("should read form fields", async () => {
    // Create a PDF with a form field (hand-crafted)
    const src = [
      "%PDF-1.4",
      "1 0 obj << /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >> endobj",
      "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj",
      "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R >> endobj",
      "4 0 obj << /Length 0 >> stream",
      "endstream endobj",
      "5 0 obj << /Type /Annot /Subtype /Widget /FT /Tx /T (name) /V (old_value) /Rect [72 700 200 720] >> endobj",
      "xref",
      "0 6",
      "0000000000 65535 f ",
      "0000000009 00000 n ",
      "0000000100 00000 n ",
      "0000000157 00000 n ",
      "0000000252 00000 n ",
      "0000000302 00000 n ",
      "trailer << /Size 6 /Root 1 0 R >>",
      "startxref",
      "440",
      "%%EOF"
    ].join("\n");

    const pdfBytes = new TextEncoder().encode(src);
    const editor = PdfEditor.load(pdfBytes);

    const fields = editor.getFormFields();
    expect(fields.length).toBe(1);
    expect(fields[0].name).toBe("name");
    expect(fields[0].value).toBe("old_value");
  });

  it("should copy pages from another PDF", async () => {
    const sourcePdf = await pdf([["Source Page 1"], ["Source Row 2"]]);
    const targetPdf = await pdf([["Target Page 1"]]);

    const editor = PdfEditor.load(targetPdf);
    editor.copyPagesFrom(sourcePdf);

    const result = await editor.save();
    const readResult = await readPdf(result);

    // Should have 2 pages: original + copied
    expect(readResult.pages.length).toBe(2);
  });

  it("should copy specific pages from another PDF", async () => {
    const doc = new PdfDocumentBuilder();
    const p1 = doc.addPage();
    p1.drawText("Page A", { x: 72, y: 750 });
    const p2 = doc.addPage();
    p2.drawText("Page B", { x: 72, y: 750 });
    const p3 = doc.addPage();
    p3.drawText("Page C", { x: 72, y: 750 });
    const sourcePdf = await doc.build();

    const targetPdf = await pdf([["Target"]]);
    const editor = PdfEditor.load(targetPdf);
    editor.copyPagesFrom(sourcePdf, [0, 2]); // Copy pages A and C

    const result = await editor.save();
    const readResult = await readPdf(result);
    expect(readResult.pages.length).toBe(3); // 1 original + 2 copied
  });

  it("should preserve original page content when overlaying", async () => {
    const pdfBytes = await pdf([["Original Content"]]);
    const editor = PdfEditor.load(pdfBytes);

    editor.getPage(0).drawText("Overlay Text", {
      x: 200,
      y: 400,
      fontSize: 20,
      color: { r: 1, g: 0, b: 0 }
    });

    const result = await editor.save();
    const readResult = await readPdf(result);
    // Both original and overlay text must be present
    expect(readResult.text).toContain("Original Content");
    expect(readResult.text).toContain("Overlay Text");
  });

  it("should preserve page Rotate property after save", async () => {
    // Hand-craft a PDF with /Rotate 90
    const src = [
      "%PDF-1.4",
      "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj",
      "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj",
      "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Rotate 90 /Contents 4 0 R /Resources << >> >> endobj",
      "4 0 obj << /Length 0 >> stream",
      "endstream endobj",
      "xref",
      "0 5",
      "0000000000 65535 f ",
      "0000000009 00000 n ",
      "0000000058 00000 n ",
      "0000000115 00000 n ",
      "0000000252 00000 n ",
      "trailer << /Size 5 /Root 1 0 R >>",
      "startxref",
      "302",
      "%%EOF"
    ].join("\n");
    const pdfBytes = new TextEncoder().encode(src);

    const editor = PdfEditor.load(pdfBytes);
    const result = await editor.save();

    // The saved PDF should contain /Rotate 90 in the page dict
    const pdfStr = new TextDecoder().decode(result);
    expect(pdfStr).toContain("/Rotate 90");
  });

  it("should preserve original metadata after save", async () => {
    const doc = new PdfDocumentBuilder();
    doc.setMetadata({ title: "My Title", author: "My Author" });
    doc.addPage().drawText("Hello", { x: 72, y: 750 });
    const pdfBytes = await doc.build();

    const editor = PdfEditor.load(pdfBytes);
    editor.getPage(0).drawText("Extra", { x: 72, y: 700 });
    const result = await editor.save();

    const readResult = await readPdf(result);
    expect(readResult.metadata.title).toBe("My Title");
    expect(readResult.metadata.author).toBe("My Author");
  });

  it("should set and save form field values", async () => {
    // Hand-craft a PDF with a text form field
    const src = [
      "%PDF-1.4",
      "1 0 obj << /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [5 0 R] >> >> endobj",
      "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj",
      "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Annots [5 0 R] >> endobj",
      "4 0 obj << /Length 0 >> stream",
      "endstream endobj",
      "5 0 obj << /Type /Annot /Subtype /Widget /FT /Tx /T (username) /V (old) /Rect [72 700 200 720] >> endobj",
      "xref",
      "0 6",
      "0000000000 65535 f ",
      "0000000009 00000 n ",
      "0000000100 00000 n ",
      "0000000157 00000 n ",
      "0000000280 00000 n ",
      "0000000330 00000 n ",
      "trailer << /Size 6 /Root 1 0 R >>",
      "startxref",
      "460",
      "%%EOF"
    ].join("\n");
    const pdfBytes = new TextEncoder().encode(src);

    const editor = PdfEditor.load(pdfBytes);
    editor.setFormField("username", "new_value");
    const result = await editor.save();

    // The saved PDF should contain the new value
    const pdfStr = new TextDecoder().decode(result);
    expect(pdfStr).toContain("new_value");
  });

  it("should preserve Rotate on copied pages", async () => {
    // Source PDF with /Rotate 90
    const src = [
      "%PDF-1.4",
      "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj",
      "2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj",
      "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Rotate 90 /Contents 4 0 R /Resources << >> >> endobj",
      "4 0 obj << /Length 0 >> stream",
      "endstream endobj",
      "xref",
      "0 5",
      "0000000000 65535 f ",
      "0000000009 00000 n ",
      "0000000058 00000 n ",
      "0000000115 00000 n ",
      "0000000252 00000 n ",
      "trailer << /Size 5 /Root 1 0 R >>",
      "startxref",
      "302",
      "%%EOF"
    ].join("\n");
    const sourcePdf = new TextEncoder().encode(src);

    const targetPdf = await pdf([["Target"]]);
    const editor = PdfEditor.load(targetPdf);
    editor.copyPagesFrom(sourcePdf);

    const result = await editor.save();
    const pdfStr = new TextDecoder().decode(result);
    // The copied page should still have /Rotate 90
    expect(pdfStr).toContain("/Rotate 90");
  });

  it("should produce valid PDF with drawImage in builder", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    // Minimal valid JPEG: SOI + SOF0 + EOI
    // This is a 1x1 pixel JPEG
    const jpegBytes = new Uint8Array([
      0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00,
      0x01, 0x00, 0x01, 0x00, 0x00, 0xff, 0xc0, 0x00, 0x0b, 0x08, 0x00, 0x01, 0x00, 0x01, 0x01,
      0x01, 0x11, 0x00, 0xff, 0xc4, 0x00, 0x14, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xff, 0xda, 0x00, 0x08, 0x01,
      0x01, 0x00, 0x00, 0x3f, 0x00, 0x7b, 0x40, 0xff, 0xd9
    ]);

    page.drawImage({
      data: jpegBytes,
      format: "jpeg",
      x: 72,
      y: 600,
      width: 200,
      height: 150
    });

    const bytes = await doc.build();
    expect(bytes).toBeInstanceOf(Uint8Array);

    // Should be readable
    const result = await readPdf(bytes);
    expect(result.pages.length).toBe(1);
  });
});

// =============================================================================
// Content Stream — Bezier / Circle / Ellipse
// =============================================================================

describe("PdfContentStream — Vector Drawing", () => {
  it("should produce correct Bezier curve operator", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    const stream = page.getContentStream();
    stream.moveTo(100, 200);
    stream.curveTo(150, 300, 250, 300, 300, 200);
    stream.stroke();

    const bytes = await doc.build();
    expect(bytes.length).toBeGreaterThan(0);
  });

  it("should produce a valid circle path", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    const stream = page.getContentStream();
    stream.circle(200, 400, 80);
    stream.stroke();

    const str = stream.toString();
    // Circle uses 4 Bezier curves: should contain 'm' and 4 'c' operators
    expect(str).toContain("m");
    expect((str.match(/ c$/gm) || []).length).toBe(4);

    const bytes = await doc.build();
    expect(bytes.length).toBeGreaterThan(0);
  });

  it("should produce a valid rounded rectangle path", async () => {
    const doc = new PdfDocumentBuilder();
    const page = doc.addPage();

    const stream = page.getContentStream();
    stream.roundedRect(100, 300, 200, 100, 15);
    stream.stroke();

    const str = stream.toString();
    // Rounded rect: 4 curves + 4 lines + 1 move
    expect(str).toContain("m");
    expect(str).toContain(" l");
    expect(str).toContain(" c");

    const bytes = await doc.build();
    expect(bytes.length).toBeGreaterThan(0);
  });
});
