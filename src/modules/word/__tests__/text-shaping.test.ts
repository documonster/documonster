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

    it("should apply final form for right-joining letter (Alef) preceded by dual-joining", () => {
      // با — Ba followed by Alef
      // Ba is dual-joining, Alef is right-joining
      // Ba sees next=Alef (R-type, not D or C), so nextJoins=false → Ba gets isolated? No:
      // Actually let's check: Ba precedes Alef. For Ba: prevJoins=false, nextJoins depends on Alef type.
      // Alef is R-type. nextJoins checks if next is D or C → Alef is R → nextJoins=false for Ba.
      // So Ba stays isolated. Alef: prevJoins checks if prev (Ba) is D/C/R → yes (D) → final form.
      // Alef final: 0xFE8E
      const result = shapeText("\u0628\u0627");
      const alef = result.find(c => c.chars === "\u0627");
      expect(alef).toBeDefined();
      expect(alef!.visual).toBe(String.fromCodePoint(0xfe8e)); // Alef Final
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
