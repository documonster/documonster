/**
 * Which text this library can and cannot lay out.
 *
 * The distinction matters because the failure is invisible. A missing glyph leaves a
 * hole; text that needs shaping is drawn in full, in the wrong shapes and the wrong
 * order, and looks fine to anyone who cannot read the script. Arabic came out as
 * disconnected isolated letters and Devanagari with its vowel signs on the wrong side
 * of the consonant, from both the PDF writer and the rasteriser, with nothing said.
 *
 * These tests pin the two halves of the answer: what is genuinely fine (which must
 * stay quiet, or the warning is noise) and what is not (which must be reported).
 */

import {
  TextFeatureTally,
  isSimpleText,
  isWellMeasuredText,
  textFeaturesOf
} from "@utils/complex-text";
import { describe, expect, it } from "vitest";

describe("text that needs no shaping", () => {
  const fine: [string, string][] = [
    ["English", "The quick brown fox"],
    ["French", "Café à l'œuvre — ça va"],
    ["German", "Größe, Übung, Straße"],
    ["Polish", "Zażółć gęślą jaźń"],
    ["Czech", "Příšerně žluťoučký kůň"],
    ["Hungarian", "Árvíztűrő tükörfúrógép"],
    ["Turkish", "Iğdır ışıklı şoför"],
    ["Vietnamese", "Tiếng Việt nghiêng"],
    ["Romanian", "Șoseaua întâi"],
    ["Icelandic", "Þetta er þröngt"],
    ["Russian", "Съешь ещё этих булок"],
    ["Greek", "Ξεσκεπάζω την ψυχή"],
    ["Ukrainian", "Їжак ґедзь щавель"],
    ["Simplified Chinese", "数据校验写入数据库"],
    ["Japanese", "日本語のテスト・カタカナ"],
    ["Korean", "한국어 테스트"],
    ["digits and punctuation", "1,234.56 (a) — «b» [c]"],
    ["mixed Latin and CJK", "Mixed 混合 ABC 123"],
    // Decomposed diacritics: combining marks are `Script=Inherited`, and a Latin base
    // with a combining acute needs no shaping — the mark has its own advance.
    ["French decomposed", "Cafe\u0301 a\u0300"],
    ["Vietnamese decomposed", "Tie\u0302\u0301ng Vie\u0323\u0302t"]
  ];

  it.each(fine)("says %s is simple", (_label, text) => {
    expect(isSimpleText(text)).toBe(true);
    expect(textFeaturesOf(text)).toEqual({ scripts: [], rtl: false });
  });

  it("stays quiet for an empty string", () => {
    expect(isSimpleText("")).toBe(true);
    const tally = new TextFeatureTally();
    tally.note("");
    expect(tally.isEmpty()).toBe(true);
  });
});

describe("text that needs shaping", () => {
  it("reports Arabic as needing both shaping and reordering", () => {
    // Arabic letters change shape by position *and* run right to left, so it is the
    // one script that trips both checks.
    const features = textFeaturesOf("مرحبا بالعالم");
    expect(features.scripts).toEqual(["Arabic"]);
    expect(features.rtl).toBe(true);
  });

  it("reports Hebrew as needing reordering but not shaping", () => {
    // The case that rules out treating "right to left" and "needs shaping" as one
    // question: Hebrew letters do not join, they are merely written the other way.
    const features = textFeaturesOf("שלום עולם");
    expect(features.scripts).toEqual([]);
    expect(features.rtl).toBe(true);
  });

  it.each([
    ["Devanagari", "नमस्ते दुनिया"],
    ["Bengali", "নমস্কার বিশ্ব"],
    ["Tamil", "வணக்கம் உலகம்"],
    ["Telugu", "నమస్కారం ప్రపంచం"],
    ["Gujarati", "નમસ્તે વિશ્વ"],
    ["Kannada", "ನಮಸ್ಕಾರ"],
    ["Malayalam", "നമസ്കാരം"],
    ["Gurmukhi", "ਸਤ ਸ੍ਰੀ ਅਕਾਲ"],
    ["Oriya", "ନମସ୍କାର"],
    ["Sinhala", "ආයුබෝවන්"],
    ["Thai", "สวัสดีชาวโลก"],
    ["Lao", "ສະບາຍດີ"],
    ["Khmer", "សួស្តី"],
    ["Myanmar", "မင်္ဂလာပါ"],
    ["Tibetan", "བཀྲ་ཤིས་བདེ་ལེགས"]
  ])("reports %s as needing shaping", (script, text) => {
    const features = textFeaturesOf(text);
    expect(features.scripts).toContain(script);
    expect(features.rtl).toBe(false);
    expect(isSimpleText(text)).toBe(false);
  });

  it("finds a complex script inside otherwise-Latin text", () => {
    // A single Arabic word in an English sentence still makes the line unrenderable,
    // so the scan cannot stop at "mostly Latin".
    const features = textFeaturesOf("The word مرحبا means hello");
    expect(features.scripts).toEqual(["Arabic"]);
    expect(features.rtl).toBe(true);
  });

  it("treats an explicit bidi control as right-to-left text", () => {
    expect(textFeaturesOf("abc\u202Bdef\u202C").rtl).toBe(true);
  });
});

describe("accumulating across many runs", () => {
  it("records each script once, however often it appears", () => {
    const tally = new TextFeatureTally();
    for (let i = 0; i < 100; i++) {
      tally.note("नमस्ते");
      tally.note("مرحبا");
    }
    expect(tally.shapingScripts()).toEqual(["Arabic", "Devanagari"]);
    expect(tally.hasRtl()).toBe(true);
    // One warning per feature, not per run: a 100k-cell sheet must not produce 100k.
    expect(tally.warnings("This renderer")).toHaveLength(2);
  });

  it("stays empty for a document of simple text", () => {
    const tally = new TextFeatureTally();
    for (const text of ["Hello", "数据校验", "Привет", "Café"]) {
      tally.note(text);
    }
    expect(tally.isEmpty()).toBe(true);
    expect(tally.warnings("This renderer")).toEqual([]);
  });

  it("names the renderer and points at a backend that gets it right", () => {
    const tally = new TextFeatureTally();
    tally.note("مرحبا");
    const [shaping, bidi] = tally.warnings("This rasteriser");

    expect(shaping).toContain("This rasteriser");
    expect(shaping).toContain("Arabic");
    // The message has to be actionable. Supplying a font does not help here, unlike
    // an uncovered code point, so it must say what does.
    expect(shaping).toContain("SVG");
    expect(shaping).toContain("DOCX");
    expect(bidi).toContain("Bidi");
    expect(bidi).toContain("SVG");
  });
});

describe("whether the static advance tables describe the text", () => {
  // `isWellMeasuredText` exists because the rasteriser has two separate decisions to make
  // and only one of them is "does this need shaping". Using one predicate for both broke a
  // line in each direction, so the difference between them is pinned here rather than left
  // to the caller to rediscover.

  it("agrees with isSimpleText on every script", () => {
    for (const text of ["Hello", "混合 ABC", "Привет", "Ελληνικά", "123 —"]) {
      expect(isWellMeasuredText(text)).toBe(true);
      expect(isSimpleText(text)).toBe(true);
    }
    for (const text of ["مرحبا", "שלום", "வணக்கம்", "नमस्ते", "สวัสดี"]) {
      expect(isWellMeasuredText(text)).toBe(false);
      expect(isSimpleText(text)).toBe(false);
    }
  });

  it("differs from isSimpleText on an explicit bidi control, which is the whole point", () => {
    // The controls are `Script=Common`, so the measurer covers them and charges them
    // nothing — the text is measured correctly. They do request reordering, so shaping
    // must still run. A renderer that reads one answer for both questions either squeezes
    // a Tamil line into an overlap or shifts a Latin one by a rounding step.
    for (const control of ["\u2066", "\u202B", "\u200F"]) {
      const text = `AB${control}CD`;
      expect(isWellMeasuredText(text)).toBe(true);
      expect(isSimpleText(text)).toBe(false);
    }
  });

  it("says a complex script is badly measured even when shaping would be a no-op", () => {
    // Tamil base consonants with no vowel sign: the shaper returns them unchanged, so a
    // renderer deciding the layout mode from what shaping *did* puts this on the measured
    // width — which reports roughly half the real advance and overlaps the glyphs.
    expect(isWellMeasuredText("வணகம")).toBe(false);
  });
});

describe("warnings adapt to what the renderer actually does", () => {
  // The advice is the payload, not the prose. A backend that shapes must not be described
  // with the message written for one that does not: it would tell a caller whose Arabic now
  // renders correctly that it "will render incorrectly" and send them to a different
  // backend to fix a problem they no longer have.

  it("does not claim incorrect output from a renderer that shapes", () => {
    const tally = new TextFeatureTally();
    tally.note("مرحبا بالعالم");

    const [shaping, bidi] = tally.warnings("This rasteriser", {
      contextualForms: true,
      visualOrder: true
    });
    expect(shaping).not.toContain("render incorrectly");
    expect(shaping).not.toContain("one glyph per code point");
    // It names what is genuinely still missing, so the warning stays actionable.
    expect(shaping).toContain("GSUB");
    expect(shaping).toContain("GPOS");
    expect(bidi).toContain("visual order");
    expect(bidi).not.toContain("appear reversed");
  });

  it("still gives the full warning to a renderer that does not shape", () => {
    const tally = new TextFeatureTally();
    tally.note("مرحبا بالعالم");

    const [shaping, bidi] = tally.warnings("The PDF writer");
    expect(shaping).toContain("render incorrectly");
    expect(bidi).toContain("appear reversed");
  });
});
