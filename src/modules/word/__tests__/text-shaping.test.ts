import { describe, it, expect } from "vitest";

import { shapeText, detectScript, detectDirection } from "../font/text-shaping";

describe("text-shaping", () => {
  // ===========================================================================
  // detectScript
  // ===========================================================================
  describe("detectScript", () => {
    it("should detect Latin script", () => {
      expect(detectScript("Hello World")).toBe("latin");
      expect(detectScript("café")).toBe("latin");
      expect(detectScript("naïve résumé")).toBe("latin");
    });

    it("should detect Arabic script", () => {
      expect(detectScript("مرحبا")).toBe("arabic");
      expect(detectScript("بسم الله")).toBe("arabic");
    });

    it("should detect Hebrew script", () => {
      expect(detectScript("שלום")).toBe("hebrew");
      expect(detectScript("עברית")).toBe("hebrew");
    });

    it("should detect CJK script", () => {
      expect(detectScript("你好世界")).toBe("cjk");
      expect(detectScript("日本語テスト")).toBe("cjk");
      expect(detectScript("漢字")).toBe("cjk");
    });

    it("should detect Hangul script", () => {
      expect(detectScript("안녕하세요")).toBe("hangul");
      expect(detectScript("한글")).toBe("hangul");
    });

    it("should detect Devanagari script", () => {
      expect(detectScript("नमस्ते")).toBe("devanagari");
      expect(detectScript("हिन्दी")).toBe("devanagari");
    });

    it("should detect Thai script", () => {
      expect(detectScript("สวัสดี")).toBe("thai");
      expect(detectScript("ภาษาไทย")).toBe("thai");
    });

    it("should detect dominant script in mixed text", () => {
      // Latin dominates
      expect(detectScript("Hello مرحبا World Test")).toBe("latin");
      // Arabic dominates
      expect(detectScript("مرحبا بكم في العالم Hi")).toBe("arabic");
    });

    it("should default to latin for empty-like input with no strong characters", () => {
      // Only spaces/punctuation → classifyChar returns "other", so maxScript stays "latin"
      expect(detectScript("   ")).toBe("latin");
    });
  });

  // ===========================================================================
  // detectDirection
  // ===========================================================================
  describe("detectDirection", () => {
    it("should return ltr for Latin text", () => {
      expect(detectDirection("Hello World")).toBe("ltr");
    });

    it("should return ltr for CJK text", () => {
      expect(detectDirection("你好")).toBe("ltr");
    });

    it("should return ltr for Hangul text", () => {
      expect(detectDirection("안녕")).toBe("ltr");
    });

    it("should return rtl for Arabic text", () => {
      expect(detectDirection("مرحبا")).toBe("rtl");
    });

    it("should return rtl for Hebrew text", () => {
      expect(detectDirection("שלום")).toBe("rtl");
    });

    it("should detect direction based on first strong character", () => {
      // Starts with Arabic (RTL strong)
      expect(detectDirection("بHello")).toBe("rtl");
      // Starts with Latin (LTR strong)
      expect(detectDirection("Helloب")).toBe("ltr");
    });

    it("should return ltr as default when no strong characters found", () => {
      expect(detectDirection("123")).toBe("ltr");
      expect(detectDirection("   ")).toBe("ltr");
      expect(detectDirection("")).toBe("ltr");
    });
  });

  // ===========================================================================
  // shapeText — Latin text (simple passthrough)
  // ===========================================================================
  describe("shapeText — Latin text", () => {
    it("should pass through Latin text unchanged", () => {
      const result = shapeText("Hello");
      expect(result).toHaveLength(5);
      expect(result.map(c => c.visual).join("")).toBe("Hello");
      expect(result.every(c => c.script === "latin")).toBe(true);
      expect(result.every(c => c.direction === "ltr")).toBe(true);
      expect(result.every(c => c.advanceMultiplier === 1.0)).toBe(true);
    });

    it("should preserve chars matching visual for Latin", () => {
      const result = shapeText("ABC");
      for (const cluster of result) {
        expect(cluster.chars).toBe(cluster.visual);
      }
    });
  });

  // ===========================================================================
  // shapeText — Arabic joining
  // ===========================================================================
  describe("shapeText — Arabic joining", () => {
    // Ba (ب U+0628) — Dual-joining
    // Sin (س U+0633) — Dual-joining
    // Meem (م U+0645) — Dual-joining
    // Alef (ا U+0627) — Right-joining only

    it("should apply initial form for first dual-joining letter followed by another", () => {
      // بس — Ba should be in Initial form (connects to right in visual, which is next in logical)
      // Ba initial: 0xFE91
      const result = shapeText("\u0628\u0633");
      const ba = result.find(c => c.chars === "\u0628");
      expect(ba).toBeDefined();
      expect(ba!.visual).toBe(String.fromCodePoint(0xfe91)); // Ba Initial
    });

    it("should apply medial form for dual-joining letter between two others", () => {
      // بسم — Sin should be in Medial form (preceded by Ba, followed by Meem)
      // Sin medial: 0xFEB4
      const result = shapeText("\u0628\u0633\u0645");
      const sin = result.find(c => c.chars === "\u0633");
      expect(sin).toBeDefined();
      expect(sin!.visual).toBe(String.fromCodePoint(0xfeb4)); // Sin Medial
    });

    it("should apply final form for last dual-joining letter preceded by another", () => {
      // بسم — Meem should be in Final form (preceded by Sin, nothing follows)
      // Meem final: 0xFEE2
      const result = shapeText("\u0628\u0633\u0645");
      const meem = result.find(c => c.chars === "\u0645");
      expect(meem).toBeDefined();
      expect(meem!.visual).toBe(String.fromCodePoint(0xfee2)); // Meem Final
    });

    it("should apply initial form for Ba in بسم", () => {
      // Ba initial: 0xFE91
      const result = shapeText("\u0628\u0633\u0645");
      const ba = result.find(c => c.chars === "\u0628");
      expect(ba).toBeDefined();
      expect(ba!.visual).toBe(String.fromCodePoint(0xfe91)); // Ba Initial
    });

    it("joins a dual-joining letter to a following right-joining one", () => {
      // با — Ba (dual) then Alef (right-joining). **Both** letters join here: a
      // right-joining letter attaches on its right, which in logical order is the
      // character before it. So Ba is initial and Alef is final.
      //
      // This case is why the joining pass was wrong for most Arabic words. The test
      // that used to be here asserted only Alef, and its comment reasoned its way to
      // "So Ba stays isolated" — recording the bug as expected behaviour. Every other
      // case in this file is dual-joining throughout (بس, بسم), so nothing else
      // exercised a dual letter followed by a right-joining one. Alef, dal, reh and waw
      // are all right-joining and appear in almost every Arabic word, so the effect was
      // that `مرحبا` rendered as disconnected letters.
      const result = shapeText("\u0628\u0627");
      const ba = result.find(c => c.chars === "\u0628");
      const alef = result.find(c => c.chars === "\u0627");
      expect(ba!.visual).toBe(String.fromCodePoint(0xfe91)); // Ba Initial
      expect(alef!.visual).toBe(String.fromCodePoint(0xfe8e)); // Alef Final
    });

    it("shapes مرحبا the way the Unicode joining rules require", () => {
      // A whole real word, letter by letter, because the per-letter cases above can all
      // pass while a word made of them does not. meem(D) reh(R) hah(D) beh(D) alef(R):
      // reh and alef are right-joining, so meem is initial and beh is medial.
      const result = shapeText("\u0645\u0631\u062D\u0628\u0627");
      const form = (ch: string) => result.find(c => c.chars === ch)!.visual.codePointAt(0);
      expect(form("\u0645")).toBe(0xfee3); // meem INITIAL
      expect(form("\u0631")).toBe(0xfeae); // reh  FINAL
      expect(form("\u062D")).toBe(0xfea3); // hah  INITIAL
      expect(form("\u0628")).toBe(0xfe92); // beh  MEDIAL
      expect(form("\u0627")).toBe(0xfe8e); // alef FINAL
    });

    it("shapes بالعالم, where a right-joining letter breaks the run twice", () => {
      // beh(D) alef(R) lam(D) ain(D) alef(R) lam(D) meem(D). Each alef ends a join and
      // the letter after it starts a new one, so the word has two initial forms.
      const result = shapeText("\u0628\u0627\u0644\u0639\u0627\u0644\u0645");
      const visual = result.map(c => c.visual.codePointAt(0));
      // Reported in visual (right-to-left) order.
      expect(visual).toEqual([0xfee2, 0xfedf, 0xfe8e, 0xfecc, 0xfedf, 0xfe8e, 0xfe91]);
    });

    it("should apply isolated form for standalone letter", () => {
      // Single Ba — no neighbors → isolated
      // Ba isolated: 0xFE8F
      const result = shapeText("\u0628");
      expect(result).toHaveLength(1);
      expect(result[0].visual).toBe(String.fromCodePoint(0xfe8f)); // Ba Isolated
    });

    it("should reverse Arabic clusters for RTL visual ordering", () => {
      // بسم in logical order → after BiDi should be in reversed visual order (Meem, Sin, Ba)
      const result = shapeText("\u0628\u0633\u0645");
      // The visual order for RTL text should be reversed: last logical char appears first visually
      expect(result[0].chars).toBe("\u0645"); // Meem first in visual
      expect(result[1].chars).toBe("\u0633"); // Sin second
      expect(result[2].chars).toBe("\u0628"); // Ba third (last visually)
    });

    it("should not apply joining when arabicJoining is disabled", () => {
      const result = shapeText("\u0628\u0633\u0645", { arabicJoining: false });
      // Visual should remain as original characters (no presentation forms)
      for (const cluster of result) {
        expect(cluster.visual).toBe(cluster.chars);
      }
    });
  });

  // ===========================================================================
  // shapeText — Mixed LTR/RTL (BiDi reordering)
  // ===========================================================================
  describe("shapeText — mixed LTR/RTL BiDi reordering", () => {
    it("should reorder RTL run within LTR paragraph", () => {
      // "Hiبس End" — Arabic run should appear reversed within LTR context
      // No space between Latin and Arabic to avoid neutral char inheritance issues
      const result = shapeText("Hi\u0628\u0633End");
      // There should be latin clusters, then arabic clusters (reversed), then latin
      const arabicClusters = result.filter(c => c.script === "arabic" && c.chars !== " ");
      expect(arabicClusters.length).toBe(2);

      // Arabic chars should be visually reversed (Sin before Ba)
      expect(arabicClusters[0].chars).toBe("\u0633"); // Sin first visually
      expect(arabicClusters[1].chars).toBe("\u0628"); // Ba second visually
    });

    it("should not reorder when bidiReorder is disabled", () => {
      const result = shapeText("Hi \u0628\u0633", { bidiReorder: false });
      const arabicClusters = result.filter(c => c.script === "arabic");
      // Should remain in logical order (Ba then Sin)
      expect(arabicClusters[0].chars).toBe("\u0628");
      expect(arabicClusters[1].chars).toBe("\u0633");
    });

    it("should handle LTR text within base RTL paragraph", () => {
      // Arabic then Latin in an RTL paragraph
      const result = shapeText("\u0628\u0633 Hi", { direction: "rtl" });
      // In RTL base direction:
      // Arabic run: level 1, LTR run: level 2
      // After BiDi reorder: LTR text reversed then whole reversed again
      // The exact behavior depends on the implementation
      expect(result.length).toBeGreaterThan(0);
      // Latin clusters should still read left-to-right within their group
      const latinClusters = result.filter(c => c.script === "latin");
      expect(latinClusters.map(c => c.chars).join("")).toBe("Hi");
    });
  });

  // ===========================================================================
  // shapeText — RTL paragraph direction option
  // ===========================================================================
  describe("shapeText — options.direction = 'rtl'", () => {
    it("should use RTL as base paragraph direction", () => {
      const result = shapeText("\u0628\u0633\u0645", { direction: "rtl" });
      // All clusters should be Arabic and RTL
      expect(result.every(c => c.direction === "rtl")).toBe(true);
      expect(result.every(c => c.script === "arabic")).toBe(true);
    });

    it("should still apply joining with RTL base direction", () => {
      const result = shapeText("\u0628\u0633\u0645", { direction: "rtl" });
      // Ba should get initial form, Sin medial, Meem final
      // After BiDi reversal, visual order is Meem, Sin, Ba
      const ba = result.find(c => c.chars === "\u0628");
      expect(ba!.visual).toBe(String.fromCodePoint(0xfe91)); // Ba Initial
    });
  });

  // ===========================================================================
  // Edge cases
  // ===========================================================================
  describe("shapeText — edge cases", () => {
    it("should return empty array for empty string", () => {
      const result = shapeText("");
      expect(result).toEqual([]);
    });

    it("should handle single Latin character", () => {
      const result = shapeText("A");
      expect(result).toHaveLength(1);
      expect(result[0].chars).toBe("A");
      expect(result[0].visual).toBe("A");
      expect(result[0].script).toBe("latin");
      expect(result[0].direction).toBe("ltr");
      expect(result[0].advanceMultiplier).toBe(1.0);
    });

    it("should handle single Arabic character", () => {
      const result = shapeText("\u0645"); // Meem alone
      expect(result).toHaveLength(1);
      expect(result[0].chars).toBe("\u0645");
      // Isolated form for Meem: 0xFEE1
      expect(result[0].visual).toBe(String.fromCodePoint(0xfee1));
      expect(result[0].script).toBe("arabic");
      expect(result[0].direction).toBe("rtl");
    });

    it("should handle spaces", () => {
      const result = shapeText("   ");
      expect(result).toHaveLength(3);
      for (const cluster of result) {
        expect(cluster.chars).toBe(" ");
        expect(cluster.visual).toBe(" ");
        expect(cluster.advanceMultiplier).toBe(1.0);
      }
    });

    it("should handle numbers (classified as 'other' script)", () => {
      const result = shapeText("123");
      expect(result).toHaveLength(3);
      // Numbers are in "other" category, may inherit surrounding script or default
      for (const cluster of result) {
        expect(cluster.visual).toBe(cluster.chars);
        expect(cluster.advanceMultiplier).toBe(1.0);
      }
    });

    it("should handle mixed numbers and Latin", () => {
      const result = shapeText("A1B");
      // "A" is latin, "1" is neutral (inherits), "B" is latin
      expect(result).toHaveLength(3);
      expect(result.map(c => c.visual).join("")).toBe("A1B");
    });

    it("should handle all options disabled", () => {
      const result = shapeText("\u0628\u0633\u0645", {
        arabicJoining: false,
        bidiReorder: false
      });
      // No joining and no reordering — original chars in logical order
      expect(result[0].chars).toBe("\u0628");
      expect(result[0].visual).toBe("\u0628");
      expect(result[1].chars).toBe("\u0633");
      expect(result[1].visual).toBe("\u0633");
      expect(result[2].chars).toBe("\u0645");
      expect(result[2].visual).toBe("\u0645");
    });
  });
});

describe("Lam-Alef ligatures", () => {
  // These were completely uncovered, and the ligature was broken in the worst way a text
  // renderer can be: the alef was *lost*. `applyLamAlefLigatures` wrote U+FEFB into the
  // lam's cluster and marked the alef's cluster zero-width, but it ran *before* the
  // contextual-form loop, which then read `codePointAt(0)` of the merged cluster — still
  // lam — and overwrote the ligature with lam's own initial form U+FEDD. The alef had
  // already been dropped, so `لا` rendered as a single ﻝ.

  it("joins lam + alef into one ligature glyph instead of dropping the alef", () => {
    const drawn = shapeText("\u0644\u0627").filter(c => c.advanceMultiplier !== 0);
    expect(drawn).toHaveLength(1);
    // U+FEFB, the isolated lam-alef ligature — not U+FEDD, which is lam alone.
    expect(drawn[0].visual).toBe("\uFEFB");
    // The cluster still carries both original characters, so a consumer writing a
    // ToUnicode map or extracting text recovers "لا" rather than one letter.
    expect(drawn[0].chars).toBe("\u0644\u0627");
  });

  it("uses the final ligature form when the lam is joined from the right", () => {
    // بلا — beh is dual-joining, so the lam is attached on its right and the ligature
    // takes its final form U+FEFC. Getting this from the same code path that had the
    // ordering bug is what shows the fix did not simply hardcode the isolated form.
    const drawn = shapeText("\u0628\u0644\u0627").filter(c => c.advanceMultiplier !== 0);
    expect(drawn.map(c => c.visual).join("")).toContain("\uFEFC");
  });

  it.each([
    ["\u0622", "\uFEF5"], // alef with madda
    ["\u0623", "\uFEF7"], // alef with hamza above
    ["\u0625", "\uFEF9"] // alef with hamza below
  ])("ligates lam with the alef variant %s", (alef, ligature) => {
    const drawn = shapeText(`\u0644${alef}`).filter(c => c.advanceMultiplier !== 0);
    expect(drawn).toHaveLength(1);
    expect(drawn[0].visual).toBe(ligature);
  });

  it("leaves lam alone when what follows is not an alef", () => {
    // لم — no ligature, so the lam must keep its ordinary initial form. Without this the
    // fix could pass by ligating anything that follows a lam.
    const drawn = shapeText("\u0644\u0645").filter(c => c.advanceMultiplier !== 0);
    expect(drawn).toHaveLength(2);
    expect(drawn.map(c => c.visual).join("")).not.toContain("\uFEFB");
  });
});
