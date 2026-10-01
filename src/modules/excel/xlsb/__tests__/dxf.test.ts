import type { Style } from "@excel/types";
import { encodeDxf, readDxf } from "@excel/xlsb/dxf";
import { describe, expect, it } from "vitest";

/**
 * `BrtDXF` in both directions, from one property table.
 *
 * The writer used to express a subset and the reader a smaller one — an underline was written and not read, a
 * pattern was always written solid, a palette colour became "not set", and alignment, number format and
 * protection were written by neither — all without a word in the fidelity report.
 */
const FORMATS = new Map<number, string>([[164, "0.0%"]]);
const formatId = (code: string): number =>
  [...FORMATS].find(([, value]) => value === code)?.[0] ?? 0;
const formatById = (id: number): string | undefined => FORMATS.get(id);

function roundTrip(style: Partial<Style>): {
  style: Partial<Style> | undefined;
  dropped: string[];
} {
  const encoded = encodeDxf(style, formatId);
  expect(encoded.payload).toBeDefined();
  const decoded = readDxf(encoded.payload!, formatById);
  expect(decoded.unread).toEqual([]);
  return { style: decoded.style, dropped: [...encoded.dropped] };
}

describe("every property the model can hold survives", () => {
  it("font", () => {
    const font = {
      // A tint that 16 bits hold exactly: `nTintShade` is the tint × 32767.
      color: { theme: 4, tint: 16384 / 32767 },
      name: "Arial",
      bold: true,
      underline: "double" as const,
      vertAlign: "superscript" as const,
      italic: true,
      strike: true,
      outline: true,
      shadow: true,
      condense: true,
      extend: true,
      charset: 134,
      family: 2,
      size: 14,
      scheme: "minor" as const
    };
    expect(roundTrip({ font })).toEqual({ style: { font }, dropped: [] });
  });

  it("an explicit false or none, which in a differential format removes the property", () => {
    const font = {
      bold: false,
      italic: false,
      strike: false,
      underline: false as const,
      outline: false,
      shadow: false
    };
    expect(roundTrip({ font }).style).toEqual({ font });
  });

  it("alignment, number format and protection — none of which was written before", () => {
    const style: Partial<Style> = {
      alignment: {
        horizontal: "center",
        vertical: "bottom",
        textRotation: -45,
        indent: 3,
        readingOrder: "rtl",
        wrapText: true,
        shrinkToFit: false
      },
      numFmt: "0.0%",
      protection: { locked: false, hidden: true }
    };
    expect(roundTrip(style)).toEqual({ style, dropped: [] });
    expect(roundTrip({ alignment: { textRotation: "vertical" } }).style).toEqual({
      alignment: { textRotation: "vertical" }
    });
  });

  it("the resets a differential format states: general, context, baseline, and a relative indent", () => {
    // Each is Excel's default, so a cell never needs to say it — but in a differential format a stated default
    // *resets* the property where omitting it changes nothing.
    const style: Partial<Style> = {
      alignment: {
        horizontal: "general",
        readingOrder: "context",
        relativeIndent: -2
      },
      font: { vertAlign: "baseline" }
    };
    expect(roundTrip(style)).toEqual({ style, dropped: [] });
    const distributed: Partial<Style> = {
      alignment: { horizontal: "distributed", justifyLastLine: true }
    };
    expect(roundTrip(distributed)).toEqual({ style: distributed, dropped: [] });
  });

  it("refuses justifyLastLine on text that is not distributed, as the record requires", () => {
    const encoded = encodeDxf(
      { alignment: { horizontal: "left", justifyLastLine: true } },
      formatId
    );
    expect(encoded.dropped).toEqual(["alignment.justifyLastLine"]);
    expect(readDxf(encoded.payload!, formatById).style).toEqual({
      alignment: { horizontal: "left" }
    });
  });

  it("a built-in number format by its reserved id", () => {
    const encoded = encodeDxf({ numFmt: "0.00" }, () => 2);
    expect(readDxf(encoded.payload!, id => (id === 2 ? "0.00" : undefined)).style).toEqual({
      numFmt: "0.00"
    });
  });

  it("fill: the pattern it names, not always solid; and a missing pattern stays missing", () => {
    for (const fill of [
      { type: "pattern" as const, pattern: "lightGray" as const, fgColor: { indexed: 10 } },
      { type: "pattern" as const, pattern: "none" as const },
      // Excel's own preset highlight: colours, no pattern — "keep the pattern, paint this solid".
      { type: "pattern" as const, bgColor: { argb: "FFFFC7CE" } }
    ]) {
      expect(roundTrip({ fill }).style).toEqual({ fill });
    }
  });

  it("gradient fills, linear and rectangular", () => {
    const stops = [
      { position: 0, color: { argb: "FFFFFFFF" } },
      { position: 1, color: { theme: 4 } }
    ];
    const angle = { type: "gradient" as const, gradient: "angle" as const, degree: 90, stops };
    const path = {
      type: "gradient" as const,
      gradient: "path" as const,
      center: { left: 0.25, top: 0.5, right: 0.75 },
      stops
    };
    expect(roundTrip({ fill: angle }).style).toEqual({ fill: angle });
    expect(roundTrip({ fill: path }).style).toEqual({ fill: path });
  });

  it("borders, the diagonal's direction, and the inner edges of a range", () => {
    const border = {
      top: { style: "thin" as const, color: { argb: "FFFF0000" } },
      bottom: { style: "double" as const },
      left: { style: "dashed" as const, color: { indexed: 8 } },
      right: { style: "hair" as const },
      diagonal: { style: "thin" as const, up: true, down: false },
      vertical: { style: "thin" as const },
      horizontal: { style: "medium" as const }
    };
    expect(roundTrip({ border })).toEqual({ style: { border }, dropped: [] });
  });
});

describe("what cannot be expressed is reported, not silently lost", () => {
  it("on write", () => {
    const encoded = encodeDxf(
      {
        font: { name: "x".repeat(33), bold: true },
        alignment: { indent: 20 },
        styleName: "Heading 1"
      },
      formatId
    );
    // The rest is still written — the bold survives the over-long name.
    expect(readDxf(encoded.payload!, formatById).style).toEqual({ font: { bold: true } });
    expect(encoded.dropped).toEqual(["font.name", "alignment.indent", "styleName"]);
  });

  it("on read", () => {
    // `BrtDXF` with two properties the model has no field for: fMergeCell (0x17) and an undefined type (0x27).
    const payload = new Uint8Array([
      0x00, 0x80, 0x00, 0x00, 0x02, 0x00, 0x17, 0x00, 0x05, 0x00, 0x01, 0x27, 0x00, 0x05, 0x00, 0x00
    ]);
    expect(readDxf(payload, formatById)).toEqual({
      style: undefined,
      unread: ["merge cell", "property 0x27"]
    });
  });

  it("a style with nothing writable gets no payload", () => {
    expect(encodeDxf({}, formatId)).toEqual({ payload: undefined, dropped: [] });
  });
});

describe("a malformed record costs the bad property, not the record", () => {
  /** A `BrtDXF` from raw `[type, cb, ...data]` properties, with `cprops` stated separately. */
  function record(cprops: number, ...properties: number[][]): Uint8Array {
    const bytes = [0x00, 0x80, 0x00, 0x00, cprops & 0xff, cprops >> 8];
    for (const [type, cb, ...data] of properties) {
      bytes.push(type! & 0xff, type! >> 8, cb! & 0xff, cb! >> 8, ...data);
    }
    return new Uint8Array(bytes);
  }
  const italic = [0x1c, 5, 1];

  it("skips a property too short for its type, and reads the next one", () => {
    // A gradient needs 44 bytes of data; this one has 2. Reading it anyway walked into the italic after it.
    expect(readDxf(record(2, [0x03, 6, 0, 0], italic), formatById)).toEqual({
      style: { font: { italic: true } },
      unread: ["truncated property 0x03"]
    });
  });

  it("does not read a font name past the end of its own property", () => {
    // Claims 4 characters and holds 1; the italic that follows is not part of the name.
    expect(readDxf(record(2, [0x18, 8, 4, 0, 0x41, 0], italic), formatById)).toEqual({
      style: { font: { italic: true } },
      unread: ["truncated property 0x18"]
    });
  });

  it("reports a property count that disagrees with the properties present", () => {
    expect(readDxf(record(3, italic), formatById)).toEqual({
      style: { font: { italic: true } },
      unread: ["property count: 3 declared, 1 present"]
    });
  });
});
