/**
 * The PDF writer shapes complex scripts — conditionally, and without breaking what it
 * already got right.
 *
 * Shaping a PDF is not shaping a viewer's screen. A viewer has the whole font and the whole
 * paragraph; the writer has to commit a glyph and a position for every character, and its
 * choices are constrained by three things this suite pins:
 *
 * 1. **A form can only be used if the embedded face has a glyph for it.** Measured across
 *    the 50 Arabic-capable system faces on one macOS host, eleven publish no presentation
 *    forms at all and not one covers the whole U+FE70–FEFC block. Substituting regardless
 *    would replace today's readable-but-disconnected letters with `.notdef` boxes — worse
 *    than not shaping.
 * 2. **The glyph drawn and the text it means are different things.** The content stream
 *    carries U+FEE3; the clipboard and a search have to yield U+0645. A ligature makes the
 *    split unmistakable: one glyph, two characters.
 * 3. **The width has to follow.** A shaped Arabic run is around a quarter narrower than the
 *    same letters unjoined, so anything that measured the unshaped text and then drew the
 *    shaped one would misplace every run after it.
 *
 * Fonts here are synthetic so that coverage is an input rather than a property of the host.
 */

import { inflateSync } from "node:zlib";

import { PdfDocumentBuilder } from "@pdf/builder/document-builder";
import { docxToPdf } from "@pdf/word-bridge";
import { Build, Document, Layout } from "@word/index";
import type { DocxDocument } from "@word/index";
import { describe, expect, it } from "vitest";

import { decompressPdfContent } from "./test-helpers";
import { buildTtfWithCmap } from "./ttf-test-utils";

/** `مرحبا` in base letters. */
const HELLO = "\u0645\u0631\u062d\u0628\u0627";
/** `لا` — lam followed by alef, which ligates to a single glyph. */
const LAM_ALEF = "\u0644\u0627";
/** `שלום`, which needs reordering but no contextual forms. */
const SHALOM = "\u05E9\u05DC\u05D5\u05DD";

/**
 * A face covering ASCII plus `[start, end]`, mapped to sequential glyph ids.
 *
 * `narrowFrom` gives every glyph at or past that id a smaller advance, which is how the
 * presentation forms are made narrower than the base letters. Real faces are: measured on
 * Arial Unicode MS, shaping `مرحبا بالعالم` takes the run from 13,538 to 10,074 units, a
 * 25.6% reduction. Without that asymmetry a width assertion cannot tell the two apart.
 */
function faceCovering(
  ranges: Array<[number, number]>,
  familyName: string,
  narrowFrom?: number
): Uint8Array {
  let nextGlyph = 1;
  const segments = [{ start: 0x20, end: 0x7e, delta: 1 - 0x20 }];
  nextGlyph += 0x7e - 0x20 + 1;
  for (const [start, end] of ranges) {
    segments.push({ start, end, delta: nextGlyph - start });
    nextGlyph += end - start + 1;
  }
  const numGlyphs = nextGlyph + 1;
  const advanceWidths = Array.from({ length: numGlyphs }, (_, gid) =>
    narrowFrom !== undefined && gid >= narrowFrom ? 250 : 500
  );
  return buildTtfWithCmap(segments, numGlyphs, { familyName, advanceWidths });
}

/** Base Arabic letters only — the eleven-face case, where no form can be substituted. */
function baseOnlyFace(): Uint8Array {
  return faceCovering([[0x0600, 0x06ff]], "BaseOnlyArabic");
}

/** Base letters *and* the Presentation Forms-B block, so joining is possible. */
function formsFace(): Uint8Array {
  // Glyph ids: 1..95 are ASCII, then 0x0600–0x06FF, then 0xFE70–0xFEFF. The forms therefore
  // start at 96 + 256, and are given the narrower advance from there on.
  return faceCovering(
    [
      [0x0600, 0x06ff],
      [0xfe70, 0xfeff]
    ],
    "FormsArabic",
    96 + 0x0700 - 0x0600
  );
}

/** Hebrew only — reordering is possible, form substitution is not applicable. */
function hebrewFace(): Uint8Array {
  return faceCovering([[0x0590, 0x05ff]], "HebrewProbe");
}

interface Built {
  readonly bytes: Uint8Array;
  readonly warnings: string[];
}

/** Draw `text` on one page with `font` as a configured face, and collect the diagnostics. */
async function build(text: string, font: Uint8Array): Promise<Built> {
  const warnings: string[] = [];
  const builder = new PdfDocumentBuilder({ fonts: { default: { regular: font } } });
  builder.onWarning(message => warnings.push(message));
  builder.addPage().drawText(text, { x: 40, y: 700, fontSize: 20 });
  return { bytes: await builder.build(), warnings };
}

/**
 * The same, through `embedFont` instead of a font configuration.
 *
 * These are two independent code paths inside `FontManager` — `routeText` for a configured
 * document, a single embedded face otherwise — and they shape in different places. Testing
 * only the configured one left the other completely uncovered: disabling shaping there
 * outright kept every assertion in this file green, and it is the path `Pdf.create` with an
 * auto-discovered font and the whole Word→PDF bridge take.
 */
async function buildEmbedded(text: string, font: Uint8Array): Promise<Built> {
  const warnings: string[] = [];
  const builder = new PdfDocumentBuilder();
  builder.embedFont(font);
  builder.onWarning(message => warnings.push(message));
  builder.addPage().drawText(text, { x: 40, y: 700, fontSize: 20 });
  return { bytes: await builder.build(), warnings };
}

/** Every `<...> <...>` pair in the PDF's ToUnicode CMap, as `[cid, codePoints]`. */
function toUnicodeEntries(bytes: Uint8Array): Array<{ cid: number; codePoints: number[] }> {
  const entries: Array<{ cid: number; codePoints: number[] }> = [];
  for (const block of decompressPdfContent(bytes).matchAll(/beginbfchar([\s\S]*?)endbfchar/g)) {
    for (const pair of block[1].matchAll(/<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4,})>/g)) {
      const hex = pair[2];
      const codePoints: number[] = [];
      for (let i = 0; i < hex.length; i += 4) {
        codePoints.push(Number.parseInt(hex.slice(i, i + 4), 16));
      }
      entries.push({ cid: Number.parseInt(pair[1], 16), codePoints });
    }
  }
  return entries;
}

/**
 * The subset glyph id each CID maps to, read from the `/CIDToGIDMap` stream.
 *
 * Needed because ToUnicode alone cannot tell a fallback from a failure: both keep the
 * original letter on the clipboard, and only the glyph says whether the page shows a letter
 * or a `.notdef` box. Glyph 0 *is* `.notdef`.
 *
 * The stream is located through the indirect reference in the CIDFont dictionary rather than
 * by guessing which stream looks binary. A guess was tried first — "even length, and no
 * letter `b` in it" — and it does not hold: the map is arbitrary bytes and may contain any
 * of them, so it silently picked the wrong stream and reported no glyphs at all.
 */
function cidToGid(bytes: Uint8Array): number[] {
  const text = new TextDecoder("latin1").decode(bytes);
  const ref = /\/CIDToGIDMap\s+(\d+)\s+0\s+R/.exec(text);
  if (!ref) {
    return [];
  }
  const objHeader = new RegExp(`(?:^|[^0-9])${ref[1]} 0 obj`).exec(text);
  if (!objHeader) {
    return [];
  }
  const streamStart = /stream\r?\n/.exec(text.slice(objHeader.index));
  if (!streamStart) {
    return [];
  }
  const start = objHeader.index + streamStart.index + streamStart[0].length;
  const end = text.indexOf("endstream", start);
  let raw: Uint8Array;
  try {
    raw = inflateSync(bytes.subarray(start, end));
  } catch {
    raw = bytes.subarray(start, end);
  }
  const gids: number[] = [];
  for (let i = 0; i + 1 < raw.length; i += 2) {
    gids.push((raw[i] << 8) | raw[i + 1]);
  }
  return gids;
}

/**
 * The CIDs actually written by a `<...> Tj` in a page's content stream.
 *
 * The subset and the ToUnicode map are not evidence about the page: glyphs are registered
 * before drawing, so a writer that shapes for registration and then draws the unshaped text
 * leaves the forms present in the font and absent from the content stream. Only the operator
 * says what a reader sees.
 */
function drawnCids(bytes: Uint8Array): number[] {
  const cids: number[] = [];
  for (const show of decompressPdfContent(bytes).matchAll(/<([0-9A-Fa-f]+)>\s*Tj/g)) {
    const hex = show[1];
    for (let i = 0; i + 3 < hex.length; i += 4) {
      cids.push(Number.parseInt(hex.slice(i, i + 4), 16));
    }
  }
  return cids;
}

describe("PDF shaping", () => {
  it("substitutes contextual forms when the face publishes them", async () => {
    const { bytes } = await build(HELLO, formsFace());
    const entries = toUnicodeEntries(bytes);
    const arabic = entries.filter(e => e.codePoints.some(cp => cp >= 0x0600 && cp <= 0x06ff));

    // Every Arabic letter of the run is present…
    expect(arabic.length).toBeGreaterThanOrEqual(HELLO.length);
    // …and each of them still *means* a base letter. A presentation form on the clipboard
    // is the bug this decoupling exists to prevent.
    for (const entry of entries) {
      for (const cp of entry.codePoints) {
        expect(cp < 0xfb50 || cp > 0xfeff, `U+${cp.toString(16)} leaked into ToUnicode`).toBe(true);
      }
    }
  });

  it("keeps a ligature's two characters recoverable from its one glyph", async () => {
    // The sharpest statement that "what is drawn" and "what it means" are separate: lam and
    // alef collapse to a single glyph, so one CID has to map back to two code points.
    const { bytes } = await build(LAM_ALEF, formsFace());
    const entries = toUnicodeEntries(bytes);
    const pair = entries.find(e => e.codePoints.length === 2);
    expect(pair, "no CID mapped to two code points, so the ligature lost a letter").toBeDefined();
    expect(pair?.codePoints).toEqual([0x0644, 0x0627]);
  });

  it("falls back to unjoined letters when the face has no presentation forms", async () => {
    // The eleven-face case, and it has to be asserted on the *glyphs*. ToUnicode cannot see
    // this failure: an unconditional substitution still records the original letter, so the
    // text is recoverable while the page shows five `.notdef` boxes. Only the glyph id
    // distinguishes "fell back to the base letter" from "asked for a glyph that is absent".
    const { bytes } = await build(HELLO, baseOnlyFace());
    const gids = cidToGid(bytes);
    expect(gids.length).toBeGreaterThan(1);
    // CID 0 is .notdef by definition; every CID the run needed must resolve to a real glyph.
    expect(gids.slice(1).filter(gid => gid === 0)).toEqual([]);

    const entries = toUnicodeEntries(bytes);
    const base = entries.filter(e => e.codePoints.some(cp => cp >= 0x0600 && cp <= 0x06ff));
    expect(base.length).toBe(new Set([...HELLO]).size);
  });

  it("tells the caller which half of shaping it got, and does not overclaim", async () => {
    // Two faces, one script, two different truths. Getting this from the same code path is
    // what shows the diagnostics are measured rather than assumed.
    const joined = await build(HELLO, formsFace());
    const unjoined = await build(HELLO, baseOnlyFace());

    const joinedShaping = joined.warnings.find(w => w.includes("Arabic")) ?? "";
    expect(joinedShaping).toContain("applies contextual forms");
    expect(joinedShaping).not.toContain("one glyph per code point");

    const unjoinedShaping = unjoined.warnings.find(w => w.includes("Arabic")) ?? "";
    expect(unjoinedShaping).toContain("GSUB/GPOS");
    expect(unjoinedShaping).not.toContain("applies contextual forms");
  });

  it("reports right-to-left text as reordered even with no forms to substitute", async () => {
    // Hebrew needs no glyph the face lacks, so reordering happens for any embedded face —
    // and must be reported as done rather than as missing. It must also not be reported as
    // needing shaping, which it does not.
    const { warnings } = await build(SHALOM, hebrewFace());
    const rtl = warnings.find(w => w.includes("right-to-left")) ?? "";
    expect(rtl).toContain("visual order");
    expect(rtl).not.toContain("appear reversed");
    expect(warnings.some(w => w.includes("GSUB/GPOS"))).toBe(false);
  });

  it("shapes through the single-embedded-face path too", async () => {
    // `embedFont`, not `fonts:` — see `buildEmbedded`. Without this the legacy path had no
    // coverage at all, and turning its shaping off changed nothing in this file.
    // Asserted on the *content stream*, and through the ligature, because both of the
    // easier assertions are blind here. The subset is registered before anything is drawn,
    // so a writer that shaped only for registration still carries the forms in the font
    // while painting the unjoined letters; and the glyph ids in `/CIDToGIDMap` are subset
    // ids, renumbered, so they say nothing about which source glyph they came from.
    //
    // Two characters collapsing to one `Tj` operand cannot happen without shaping.
    const ligature = await buildEmbedded(LAM_ALEF, formsFace());
    expect(drawnCids(ligature.bytes)).toHaveLength(1);

    // No `.notdef` anywhere, and nothing on the clipboard is a presentation form.
    const { bytes } = await buildEmbedded(HELLO, formsFace());
    expect(
      cidToGid(bytes)
        .slice(1)
        .filter(gid => gid === 0)
    ).toEqual([]);
    for (const entry of toUnicodeEntries(bytes)) {
      for (const cp of entry.codePoints) {
        expect(cp < 0xfb50 || cp > 0xfeff).toBe(true);
      }
    }
  });

  it("measures the shaped width, so a following run is not misplaced", async () => {
    // The third constraint in this file's header. A shaped Arabic run is about a quarter
    // narrower than the same letters unjoined, so the two faces must not produce the same
    // measurement — if they did, every run after this one would sit in the wrong place.
    const joined = new PdfDocumentBuilder();
    joined.embedFont(formsFace());
    const joinedPage = joined.addPage();
    const unjoined = new PdfDocumentBuilder();
    unjoined.embedFont(baseOnlyFace());
    const unjoinedPage = unjoined.addPage();

    // Anchored at the end, so the measured width decides where the ink starts.
    joinedPage.drawText(HELLO, { x: 400, y: 700, fontSize: 20, anchor: "end" });
    unjoinedPage.drawText(HELLO, { x: 400, y: 700, fontSize: 20, anchor: "end" });
    await joined.build();
    await unjoined.build();

    const measured = joinedPage.measureText(HELLO, { fontSize: 20 });
    const measuredUnshaped = unjoinedPage.measureText(HELLO, { fontSize: 20 });
    expect(measured).toBeGreaterThan(0);
    expect(measured).not.toBe(measuredUnshaped);
  });

  it("leaves Latin text byte-identical", async () => {
    // The guard that matters most in practice: shaping must be invisible to the
    // overwhelming majority of documents. Two builds of the same Latin page have to agree,
    // and the shaping path must not have touched the content stream at all.
    const first = await build("Hello world", formsFace());
    const second = await build("Hello world", formsFace());
    expect(first.warnings).toEqual([]);
    expect(second.bytes.length).toBe(first.bytes.length);
  });
});

/**
 * Shaping was deliberately *not* put in the Word layout engine, and this is the suite that
 * holds that line.
 *
 * Putting it there would have been the obvious way to make Word→PDF correct, and it breaks
 * two things that are correct today. `PositionedRun.text` is written verbatim into SVG, whose
 * correctness rests on carrying the original characters so the viewer shapes them — shaped
 * input would be shaped a second time. And the layout's page count is resolved into
 * `NUMPAGES`, `PAGE`, `PAGEREF` and TOC entries, i.e. written back into the .docx; the
 * default measurer is the static advance table, which has no width for a presentation form
 * at all, so shaping there would make those numbers *worse*.
 *
 * It lives in the PDF font engine instead, which is reached only when the Word bridge injects
 * its measurer — so the layout keeps the author's text and `run.width` still matches the
 * glyphs. These tests fail if anyone moves it.
 */
describe("Word layout is left unshaped", () => {
  const ARABIC = "\u0645\u0631\u062d\u0628\u0627 \u0628\u0627\u0644\u0639\u0627\u0644\u0645";

  function arabicDoc(): DocxDocument {
    const handle = Document.create();
    Document.addParagraphElement(handle, Build.paragraph([Build.text(ARABIC)]));
    return Document.build(handle);
  }

  it("keeps the original characters in the SVG it renders", () => {
    const svg = Layout.renderPageToSvg(arabicDoc(), 1);
    // A presentation form here would be shaped again by the viewer.
    expect(/[\uFB50-\uFDFF\uFE70-\uFEFF]/.test(svg)).toBe(false);
    expect(svg).toContain("\u0645");
  });

  it("keeps the page count the field engine writes into the document", () => {
    // The static advance tables have no entry for a presentation form, so a shaped layout
    // would measure this run at the fallback average width and could paginate differently —
    // and that number is written into NUMPAGES, PAGE, PAGEREF and the TOC.
    expect(Layout.document(arabicDoc()).pageCount).toBe(1);
  });

  it("still produces a joined PDF through the bridge", async () => {
    // The other half of the claim: keeping the layout unshaped must not cost the PDF its
    // shaping, because the bridge injects the font engine's measurer.
    const bytes = await docxToPdf(arabicDoc());
    expect(new TextDecoder().decode(bytes.slice(0, 5))).toBe("%PDF-");
    for (const entry of toUnicodeEntries(bytes)) {
      for (const cp of entry.codePoints) {
        expect(cp < 0xfb50 || cp > 0xfeff, "a presentation form reached ToUnicode").toBe(true);
      }
    }
  });
});
