import { extractAll } from "@archive/unzip/extract";
/**
 * Tests for the body-level `<w14:checkbox>` SDT writer.
 *
 * The writer is reachable from the public API via the `checkBox()` builder
 * exposed at `excelts/word`. ECMA-376 Part 4 §17.5.2.41 says
 * `w14:checkedState/@w14:val` must be a hexadecimal Unicode code point —
 * but the public TS type only declares it as `string`, so callers will
 * naturally try to pass a literal glyph (e.g. `"✓"`). Earlier the writer
 * blindly fed the value into `parseInt(_, 16)`, throwing
 * `RangeError: Invalid code point NaN` for any non-hex input.
 *
 * The writer now normalises both forms:
 *   - 1-6 hex digits → use as the canonical code point
 *   - any other string → take the first Unicode code point as the glyph
 *
 * These tests lock that down so we don't regress.
 */
import { describe, it, expect } from "vitest";

import { Document, checkBox, toBuffer } from "../index";

async function getDocumentXml(buf: Uint8Array): Promise<string> {
  const entries = await extractAll(buf);
  const part = entries.get("word/document.xml");
  if (!part) {
    throw new Error("document.xml missing");
  }
  return new TextDecoder().decode(part.data);
}

describe("checkBox builder + w14:checkbox writer", () => {
  it("default checked state renders ☒ and uses the canonical hex 2612", async () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(d, checkBox({ checked: true }));

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    expect(xml).toContain('<w14:checked w14:val="1"/>');
    expect(xml).toContain('w14:val="2612"');
    // Visible glyph: ☒ U+2612 BALLOT BOX WITH X
    expect(xml).toContain("\u2612");
  });

  it("default unchecked state renders ☐", async () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(d, checkBox({ checked: false }));

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    expect(xml).toContain('<w14:checked w14:val="0"/>');
    expect(xml).toContain('w14:val="2610"');
    expect(xml).toContain("\u2610");
  });

  it("wraps a body-level checkbox SDT in a paragraph (valid CT_SdtContentBlock)", async () => {
    // A checkbox SDT's sdtContent holds a run, which is illegal directly at
    // block level (CT_SdtContentBlock forbids bare runs). The writer must wrap
    // the whole SDT in a <w:p> so it becomes a valid run-level SDT inside a
    // paragraph — otherwise Word refuses to open the file.
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(d, checkBox({ checked: true }));

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    // The SDT must be wrapped: <w:p><w:sdt>…</w:sdt></w:p>, never <w:body><w:sdt>.
    expect(xml).toMatch(/<w:p><w:sdt><w:sdtPr><w14:checkbox/);
    expect(xml).not.toMatch(/<w:body><w:sdt><w:sdtPr><w14:checkbox/);
  });

  it("accepts a literal glyph string and normalises to hex (✓ → 2713)", async () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(
      d,
      checkBox({
        checked: true,
        checkedState: { value: "\u2713", font: "Arial" }, // ✓
        uncheckedState: { value: "\u2717", font: "Arial" } // ✗
      })
    );

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    // Hex normalisation in the attribute …
    expect(xml).toContain('w14:val="2713"');
    expect(xml).toContain('w14:val="2717"');
    // … and the actual glyph in the displayed run
    expect(xml).toContain("\u2713");
  });

  it("accepts the canonical hex form and uses it verbatim", async () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(
      d,
      checkBox({
        checked: true,
        checkedState: { value: "2611", font: "MS Gothic" } // ☑
      })
    );

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    expect(xml).toContain('w14:val="2611"');
    // Visible glyph for U+2611
    expect(xml).toContain("\u2611");
  });

  it("does not throw on emoji glyphs that take more than one UTF-16 code unit", async () => {
    const d = Document.create();
    Document.useDefaultStyles(d);
    Document.addContent(
      d,
      checkBox({
        checked: true,
        checkedState: { value: "\u{1F4AF}", font: "Segoe UI Emoji" } // 💯
      })
    );

    const buf = await toBuffer(Document.build(d));
    const xml = await getDocumentXml(buf);

    // Hex form of U+1F4AF is 1F4AF
    expect(xml).toContain('w14:val="1F4AF"');
    expect(xml).toContain("\u{1F4AF}");
  });
});
