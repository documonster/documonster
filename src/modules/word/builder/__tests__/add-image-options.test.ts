/**
 * `Document.addImage` / `addFloatingImage` take the format and size in an options object; the
 * positional `(data, mediaType, width, height, options?)` form is a deprecated forward that must
 * still produce the same document.
 */
import { describe, it, expect } from "vitest";

import { Document, DocxArgumentError, PageSizes, Units } from "../../index";
import type { DocxDocument } from "../../index";

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

function lastImage(doc: DocxDocument): unknown {
  const block = doc.body[doc.body.length - 1] as { children?: { content?: unknown[] }[] };
  return block.children?.[0]?.content?.[0] ?? block;
}

describe("Document.addImage options form", () => {
  it("produces the same picture as the deprecated positional form", () => {
    const a = Document.create();
    const b = Document.create();
    const width = Units.cmToEmu(5);
    const height = Units.cmToEmu(3);
    const ra = Document.addImage(a, PNG, { type: "png", width, height, altText: "logo" });
    const rb = Document.addImage(b, PNG, "png", width, height, { altText: "logo" });
    expect(ra).toEqual(rb);
    expect(lastImage(a)).toEqual(lastImage(b));
    expect(a.images).toEqual(b.images);
  });

  it("names the options field when the type or size is wrong", () => {
    const doc = Document.create();
    expect(() =>
      Document.addImage(doc, PNG, { type: "image/png" as never, width: 1, height: 1 })
    ).toThrow(/options\.type|invalid image media type/);
    try {
      Document.addImage(doc, PNG, { type: "png", width: -1, height: 1 });
      expect.unreachable();
    } catch (error) {
      expect((error as DocxArgumentError).field).toBe("options.width");
    }
  });

  it("refuses positional arguments after the options object", () => {
    const doc = Document.create();
    const call = Document.addImage as (...args: unknown[]) => unknown;
    expect(() => call(doc, PNG, { type: "png", width: 1, height: 1 }, 5)).toThrow(
      DocxArgumentError
    );
    expect(doc.body).toHaveLength(0);
  });
});

describe("Document.addFloatingImage options form", () => {
  it("produces the same anchor as the deprecated positional form", () => {
    const a = Document.create();
    const b = Document.create();
    const size = Units.cmToEmu(4);
    const ra = Document.addFloatingImage(a, PNG, {
      type: "png",
      width: size,
      height: size,
      wrap: { style: "square" },
      behindDoc: true
    });
    const rb = Document.addFloatingImage(b, PNG, "png", size, size, {
      wrap: { style: "square" },
      behindDoc: true
    });
    expect(ra).toBe(rb);
    expect(a.body).toEqual(b.body);
  });

  it("names the options field for a bad anchor distance", () => {
    const doc = Document.create();
    try {
      Document.addFloatingImage(doc, PNG, { type: "png", width: 1, height: 1, distT: Number.NaN });
      expect.unreachable();
    } catch (error) {
      expect((error as DocxArgumentError).field).toBe("options.distT");
    }
  });
});

describe("PageSizes", () => {
  it("holds the paper sizes in twips, frozen", () => {
    expect(PageSizes.A4).toEqual({ width: 11906, height: 16838 });
    expect(PageSizes.LETTER).toEqual({ width: 12240, height: 15840 });
    expect(Object.isFrozen(PageSizes)).toBe(true);
    expect(Object.isFrozen(PageSizes.A4)).toBe(true);
    expect(PageSizes.A4.width).toBe(Units.mmToTwips(210));
  });

  it("is accepted as a section's page size", () => {
    const doc = Document.create();
    Document.setSectionProperties(doc, { pageSize: PageSizes.A4 });
    expect(doc.sectionProperties?.pageSize).toEqual(PageSizes.A4);
  });
});
