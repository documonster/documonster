import { getLayoutFontMetrics } from "@utils/font-data";
import { measureFaceText } from "@word/layout/font-face";
/**
 * Aptos — the face of a new document in current Word — is measured in each of
 * its four styles. Its italic is a separate design, about as wide as the
 * regular overall but not letter for letter, so measuring it upright moved
 * line breaks in italic text.
 */
import { describe, expect, it } from "vitest";

describe("Aptos layout metrics", () => {
  it("has a table for regular, bold, italic and bold italic", () => {
    const faces = [
      getLayoutFontMetrics("Aptos"),
      getLayoutFontMetrics("Aptos", true),
      getLayoutFontMetrics("Aptos", false, true),
      getLayoutFontMetrics("Aptos", true, true)
    ];
    expect(faces.every(Boolean)).toBe(true);
    expect(new Set(faces).size).toBe(4);
  });

  it("measures italic text with the italic advances", () => {
    // Advances from Aptos-Italic.ttf as shipped with Word for Mac 16.114.
    expect(measureFaceText("f", "Aptos", 2048, false, true)).toBe(
      getLayoutFontMetrics("Aptos", false, true)!.advances.find(
        ([s, n]) => s <= 0x66 && s + n > 0x66
      )![2]
    );
    expect(measureFaceText("f", "Aptos", 2048, false, true)).not.toBe(
      measureFaceText("f", "Aptos", 2048)
    );
  });

  it("keeps a face without an italic table upright", () => {
    // Calibri now has italic tables; Courier New (monospaced) does not.
    expect(getLayoutFontMetrics("Courier New", false, true)).toBe(
      getLayoutFontMetrics("Courier New")
    );
    expect(getLayoutFontMetrics("Courier New", true, true)).toBe(
      getLayoutFontMetrics("Courier New", true)
    );
  });
});
