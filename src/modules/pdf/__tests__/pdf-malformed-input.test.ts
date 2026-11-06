import { zlibSync } from "@archive/compression/compress";
import { resolveFont } from "@pdf/reader/font-decoder";
/**
 * Corrupted input must end in a result or a typed `PdfError` — never a hang, a stack overflow
 * or a `TypeError`. Each block is one root cause found by fuzzing. The cases that need a file
 * from `data/` are in `pdf-malformed-input.node.test.ts`.
 */
import { PdfDocument } from "@pdf/reader/pdf-document";
import { parseObject } from "@pdf/reader/pdf-parser";
import type { PdfObject } from "@pdf/reader/pdf-parser";
import { PdfTokenizer, TokenType } from "@pdf/reader/pdf-tokenizer";
import { decodeStreamFilters } from "@pdf/reader/stream-filters";
import { describe, expect, it, vi } from "vitest";

import { Pdf, PdfError } from "../index";
import { buildRawPdf } from "./raw-pdf.helpers";

const enc = new TextEncoder();

/** Run a read and report how it ended: `"ok"` or the error's class name. */
async function outcome(read: () => unknown): Promise<string> {
  try {
    await read();
    return "ok";
  } catch (error) {
    expect(error).toBeInstanceOf(PdfError);
    return (error as Error).constructor.name;
  }
}

describe("tokenizer always makes progress", () => {
  it.each([")", "{", "}"])("consumes a stray %s as a one-character keyword", delimiter => {
    const tokenizer = new PdfTokenizer(enc.encode(`${delimiter} /Next`));
    const token = tokenizer.next();
    expect(token.type).toBe(TokenType.Keyword);
    expect(token.strValue).toBe(delimiter);
    expect(tokenizer.pos).toBe(1);
    expect(tokenizer.next()).toMatchObject({ type: TokenType.Name, strValue: "Next" });
  });

  it("parses past a stray delimiter where a dictionary key belongs", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("<< ) /A 1 } /B 2 >>")))).toEqual(
      new Map([
        ["A", 1],
        ["B", 2]
      ])
    );
  });

  it("parses past a stray delimiter inside an array", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("[1 } 2]")))).toEqual([1, "}", 2]);
  });

  it("reads a keyword and a name longer than the engine's argument limit", () => {
    const long = "k".repeat(300_000);
    expect(new PdfTokenizer(enc.encode(long)).next().strValue).toHaveLength(300_000);
    expect(new PdfTokenizer(enc.encode("/" + long)).next().strValue).toHaveLength(300_000);
  });
});

describe("editor clones a reference cycle that never passes through a dictionary", () => {
  // A dictionary or stream behind a reference was cloned once and cached, which is what
  // ends a cycle. An array or a reference behind a reference was inlined instead — so a
  // cycle made only of those recursed until the stack overflowed.
  const page = (resources: string): string[] => [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /X ${resources} >> >>`
  ];
  const host = async (): Promise<Uint8Array> => {
    const builder = new Pdf.Builder();
    builder.addPage().drawText("host", { x: 10, y: 10 });
    return builder.toBuffer();
  };

  it.each([
    ["an array that contains itself", [...page("4 0 R"), "[4 0 R]"]],
    ["two references to each other", [...page("4 0 R"), "5 0 R", "4 0 R"]],
    ["arrays that contain each other", [...page("4 0 R"), "[5 0 R 1]", "[4 0 R 2]"]]
  ])("%s", async (_, bodies) => {
    const bytes = buildRawPdf(bodies);
    for (const write of [
      () => Pdf.Editor.read(bytes).toBuffer(),
      () => Pdf.Editor.read(bytes).splitPages(),
      async () =>
        Pdf.Editor.read(await host())
          .copyPagesFrom(bytes)
          .toBuffer()
    ]) {
      expect(await outcome(write)).toBe("ok");
    }
    const out = await Pdf.Editor.read(bytes).toBuffer();
    expect((await Pdf.extract(out)).pages).toHaveLength(1);
  });

  it("a long chain of indirect arrays is not flattened into one nested array", async () => {
    // Inlined, the chain came out 20,000 arrays deep: past the stack on a small one, and past
    // the reader's own nesting limit on any — the editor wrote a file it could not read back.
    const chain = Array.from({ length: 20_000 }, (_, i) => `[${i + 5} 0 R]`);
    const bytes = buildRawPdf([...page("4 0 R"), ...chain, "1"]);
    const out = await Pdf.Editor.read(bytes).toBuffer();
    expect((await Pdf.extract(out)).pages).toHaveLength(1);
  });

  it("keeps an indirect array's contents", async () => {
    const bytes = buildRawPdf([...page("4 0 R"), "[1 2 3]"]);
    const out = new TextDecoder("latin1").decode(await Pdf.Editor.read(bytes).toBuffer());
    expect(out).toMatch(/\[1 2 3\]/);
  });
});

describe("cross-reference table with an overstated count", () => {
  const latin1 = (bytes: Uint8Array): string =>
    Array.from(bytes, byte => String.fromCharCode(byte)).join("");
  const withCount = (count: string): Uint8Array => {
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>"
    ];
    const text = latin1(buildRawPdf(bodies)).replace("xref\n0 4\n", `xref\n0 ${count}\n`);
    return Uint8Array.from(text, ch => ch.charCodeAt(0));
  };

  // The count bounded the loop and nothing else did: past the last entry it kept reading
  // empty lines at the end of the data, four billion times for this one.
  it.each(["4294967296", "999999999", "40"])(
    "%s entries claimed, 4 present: the table ends at its last entry",
    async count => {
      const bytes = withCount(count);
      expect((await Pdf.extract(bytes)).pages).toHaveLength(1);
      expect(await outcome(() => Pdf.Editor.read(bytes).toBuffer())).toBe("ok");
    }
  );

  it("a subsection that follows an overstated one is still read", async () => {
    const bodies = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>"
    ];
    const text = latin1(buildRawPdf(bodies));
    // Split the table into "0 3" (claimed as 9) and "3 1".
    const split = text.replace(/xref\n0 4\n((?:.{20}){3})(.{20})/s, "xref\n0 9\n$1" + "3 1\n$2");
    expect(split).not.toBe(text);
    const bytes = Uint8Array.from(split, ch => ch.charCodeAt(0));
    expect((await Pdf.extract(bytes)).pages).toHaveLength(1);
  });
});

describe("splitting a file whose cross-reference table had to be rebuilt", () => {
  it("rebuilds it once, not once per page", async () => {
    const builder = new Pdf.Builder();
    for (let i = 0; i < 6; i++) {
      builder.addPage().drawText(`Page ${i + 1}`, { x: 72, y: 700, fontSize: 12 });
    }
    const bytes = await builder.toBuffer();
    // Point startxref at nothing, so every read has to scan the file for its objects.
    const text = Array.from(bytes, byte => String.fromCharCode(byte)).join("");
    const damaged = text.replace(/startxref\s+\d+/, "startxref\n1");
    expect(damaged).not.toBe(text);
    const input = Uint8Array.from(damaged, ch => ch.charCodeAt(0));

    const rebuild = vi.spyOn(
      PdfDocument.prototype as unknown as { reconstructXref: () => unknown },
      "reconstructXref"
    );
    try {
      const parts = await Pdf.Editor.read(input).splitPages();
      expect(parts).toHaveLength(6);
      expect(rebuild).toHaveBeenCalledTimes(1);
      const texts = await Promise.all(
        parts.map(async part => (await Pdf.extract(part)).pages.map(page => page.text.trim()))
      );
      expect(texts).toEqual([1, 2, 3, 4, 5, 6].map(n => [`Page ${n}`]));
    } finally {
      rebuild.mockRestore();
    }
  });
});

describe("an object that fails to parse", () => {
  it("is parsed once, not on every lookup", () => {
    // Object 4 never closes its array, so parsing it runs to the end of the file and fails.
    const bytes = buildRawPdf([
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Annots 4 0 R >>",
      "[ 1 2 3"
    ]);
    const doc = new PdfDocument(bytes);
    const parseAt = vi.spyOn(
      doc as unknown as { parseObjectAt: (...args: unknown[]) => unknown },
      "parseObjectAt"
    );
    const first = doc.resolve(4);
    for (let i = 0; i < 5; i++) {
      expect(doc.resolve(4)).toEqual(first);
    }
    expect(parseAt).toHaveBeenCalledTimes(1);
  });
});

describe("an array or dictionary that is never closed", () => {
  it("ends at its object's endobj instead of swallowing the objects after it", async () => {
    // Each later `N 0 obj` was parsed as an object nested inside the open array, so 600
    // objects after it nested 600 deep and the whole file failed with a limit error.
    const later = Array.from({ length: 600 }, (_, i) => `<< /Index ${i} >>`);
    const bytes = buildRawPdf([
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Annots 4 0 R >>",
      "[ 1 2",
      "<< /Open true /Dangling",
      ...later
    ]);
    expect((await Pdf.extract(bytes)).pages).toHaveLength(1);
    const doc = new PdfDocument(bytes);
    expect(doc.resolve(4)).toEqual([1, 2]);
    expect(doc.resolve(5)).toEqual(new Map([["Open", true]]));
    expect(doc.resolve(6)).toEqual(new Map([["Index", 0]]));
    expect(await outcome(() => Pdf.Editor.read(bytes).toBuffer())).toBe("ok");
  });

  it("a number pair followed by obj inside an array is not an object definition", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("[1 0 obj 2]")))).toEqual([1, 0, "obj", 2]);
  });
});

describe("a dictionary key with no value", () => {
  it("does not take the closing >> as its value", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("[<< /A 1 /B >> 7]")))).toEqual([
      new Map([["A", 1]]),
      7
    ]);
  });
});

describe("a CID font /W range", () => {
  it("is clamped to the CIDs that exist", () => {
    const doc = new PdfDocument(
      buildRawPdf(["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [] /Count 0 >>"])
    );
    const cidFont = new Map<string, PdfObject>([
      ["Type", "Font"],
      ["Subtype", "CIDFontType2"],
      ["W", [1, [500, 600], 3, 4294967296, 700]]
    ]);
    const font = resolveFont(
      new Map<string, PdfObject>([
        ["Type", "Font"],
        ["Subtype", "Type0"],
        ["Encoding", "Identity-H"],
        ["DescendantFonts", [cidFont]]
      ]),
      doc
    );
    expect(font.widths.get(2)).toBe(600);
    expect(font.widths.get(0xffff)).toBe(700);
    expect(font.widths.size).toBe(0xffff);
  });
});

describe("predictor parameters that describe no row", () => {
  const flate = (data: Uint8Array, parms: Array<[string, PdfObject]>): Map<string, PdfObject> =>
    new Map<string, PdfObject>([
      ["Filter", "FlateDecode"],
      ["DecodeParms", new Map(parms)]
    ]);
  const data = Uint8Array.from({ length: 64 }, (_, i) => i);

  it("TIFF predictor with /Columns 0 takes the default instead of looping forever", () => {
    const out = decodeStreamFilters(
      zlibSync(data),
      flate(data, [
        ["Predictor", 2],
        ["Columns", 0]
      ])
    );
    expect(out).toEqual(data);
  });

  it("PNG predictor with a negative /Columns does not size its output from it", () => {
    const rows = Uint8Array.from({ length: 64 }, (_, i) => (i % 2 === 0 ? 0 : i));
    const out = decodeStreamFilters(
      zlibSync(rows),
      flate(rows, [
        ["Predictor", 12],
        ["Columns", -100_000_000]
      ])
    );
    expect(out).toEqual(Uint8Array.from({ length: 32 }, (_, i) => 2 * i + 1));
  });
});

describe("a cross-reference stream whose entries take no bytes", () => {
  it("is rejected and the file is read by scanning instead", async () => {
    const bytes = buildRawPdf(
      [
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] >>"
      ],
      { xrefStream: true }
    );
    const text = Array.from(bytes, byte => String.fromCharCode(byte)).join("");
    // Only the xref stream's own dictionary changes, so no offset moves.
    const zero = text.replace("/W [1 4 2]", "/W [0 0 0] /Index [0 999999999]");
    expect(zero).not.toBe(text);
    const damaged = Uint8Array.from(zero, ch => ch.charCodeAt(0));
    const doc = new PdfDocument(damaged);
    expect(doc.xrefRecovered).toBe(true);
    expect((await Pdf.extract(damaged)).pages).toHaveLength(1);
  });
});

describe("damage a reader tolerates", () => {
  // Every case here was read by the reader before the malformed-input hardening, and must
  // still be read after it: a guard against hostile input may not cost a damaged file.
  const latin1 = (bytes: Uint8Array): string =>
    Array.from(bytes, byte => String.fromCharCode(byte)).join("");
  const bytesOf = (text: string): Uint8Array => Uint8Array.from(text, ch => ch.charCodeAt(0));
  const content = "BT /F1 12 Tf 72 700 Td (Hello damage) Tj ET";
  const textPdf = (catalog = "<< /Type /Catalog /Pages 2 0 R >>"): string =>
    latin1(
      buildRawPdf([
        catalog,
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        "<< /Title (Kept) >>"
      ])
    ).replace("/Root 1 0 R", "/Root 1 0 R /Info 6 0 R");
  const expectReadable = async (text: string): Promise<void> => {
    const bytes = bytesOf(text);
    const result = await Pdf.extract(bytes);
    expect(result.pages.map(page => page.text.trim())).toEqual(["Hello damage"]);
    expect(result.metadata?.title).toBe("Kept");
    const out = await Pdf.Editor.read(bytes).toBuffer();
    expect((await Pdf.extract(out)).pages[0].text.trim()).toBe("Hello damage");
  };
  /**
   * A table that parses but has a wrong entry still takes an incremental update — chained
   * onto it, as for a signed document it has to be — and the update reads back.
   */
  const expectUpdatable = async (bytes: Uint8Array): Promise<void> => {
    expect(new PdfDocument(bytes).xrefRecovered).toBe(false);
    const editor = Pdf.Editor.read(bytes);
    editor.getPage(0).drawText("stamp", { x: 72, y: 72, fontSize: 9 });
    const out = await editor.toBuffer({ incremental: "required" });
    expect(latin1(out.subarray(0, bytes.length))).toBe(latin1(bytes));
    const text = (await Pdf.extract(out)).pages[0].text;
    expect(text).toContain("Hello damage");
    expect(text).toContain("stamp");
  };
  const xrefLine = (text: string, objNum: number): string => {
    const table = text.slice(text.lastIndexOf("xref\n0 "));
    return table.split("\n")[2 + objNum] + "\n";
  };

  it("a stray endobj inside the catalog dictionary", async () => {
    const text = textPdf("<< /Type /Catalog endobj /Pages 2 0 R >>");
    await expectReadable(text);
  });

  it("a stray endobj at the start of the trailer dictionary", async () => {
    const text = textPdf().replace("trailer\n<< ", "trailer\n<<endobj ");
    await expectReadable(text);
  });

  it("an xref whose first two entries were merged into one over-long line", async () => {
    const text = textPdf();
    const merged = text.replace(
      xrefLine(text, 0) + xrefLine(text, 1),
      `0${"9".repeat(40)}${xrefLine(text, 2).slice(3)}`
    );
    expect(merged).not.toBe(text);
    await expectReadable(merged);
  });

  it("an xref entry that points the catalog into another object", async () => {
    const text = textPdf();
    const line = xrefLine(text, 1);
    const wrong = String(Number(line.slice(0, 10)) + 12).padStart(10, "0") + line.slice(10);
    await expectReadable(text.replace(line, wrong));
  });

  it("an xref entry that points a page at another object is relocated by its header", async () => {
    const text = textPdf();
    const page = xrefLine(text, 3);
    const damaged = bytesOf(text.replace(page, xrefLine(text, 4)));
    expect((await Pdf.extract(damaged)).pages.map(p => p.text.trim())).toEqual(["Hello damage"]);
    await expectUpdatable(damaged);
  });

  it("a garbled xref entry is skipped, not the end of its subsection", async () => {
    // Ending the subsection there lost every entry after it, the catalog's included.
    const text = textPdf();
    const garbled = text.replace(xrefLine(text, 0), "00000Q00000 65535 f \n");
    expect(garbled).not.toBe(text);
    await expectReadable(garbled);
    await expectUpdatable(bytesOf(garbled));
  });

  it("the object a garbled entry listed is found by its header", async () => {
    const text = textPdf();
    const garbled = text.replace(xrefLine(text, 2), "0000000\x1015\x0000000 n \n");
    expect(garbled).not.toBe(text);
    await expectReadable(garbled);
    await expectUpdatable(bytesOf(garbled));
  });

  it("an entry whose header was damaged is still read from where the table says", async () => {
    // Nothing else in the file is object 1, so the entry is the best information there is.
    const text = textPdf();
    const at = text.indexOf("1 0 obj");
    await expectReadable(text.slice(0, at) + "0 0 obj" + text.slice(at + 7));
  });

  it("an endobj that does not end the object is read as before", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("[1 endobj 2]")))).toEqual([1, "endobj", 2]);
    expect(parseObject(new PdfTokenizer(enc.encode("<< /A endobj /B 2 >>")))).toEqual(
      new Map<string, PdfObject>([
        ["A", "endobj"],
        ["B", 2]
      ])
    );
    expect(parseObject(new PdfTokenizer(enc.encode("<< endobj /B 2 >>")))).toEqual(
      new Map([["B", 2]])
    );
  });

  it("an endobj followed by the next object, or by the end of the file, ends the object", () => {
    expect(parseObject(new PdfTokenizer(enc.encode("[1 2\nendobj\n5 0 obj\n3")))).toEqual([1, 2]);
    expect(parseObject(new PdfTokenizer(enc.encode("<< /A 1 endobj\nxref\n")))).toEqual(
      new Map([["A", 1]])
    );
    expect(parseObject(new PdfTokenizer(enc.encode("<< /A 1 endobj %%EOF\n")))).toEqual(
      new Map([["A", 1]])
    );
  });
});
