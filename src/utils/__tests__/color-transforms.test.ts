/**
 * The one DrawingML colour-transform implementation (`applyColorTransforms`) and the
 * Office tints expressed through it: SpreadsheetML's cell `tint` (§18.8.19) and Word's
 * `themeTint` / `themeShade` (§17.3.2.41 / .40), which Office applies to the HSL luminance
 * — not as an RGB mix — so "Accent 1, Darker 25%" of `4472C4` is Word's `2F5496`, not `335693`.
 *
 * Expected values are what Office shows (the colour picker's named variants of the default
 * theme) and what LibreOffice's MSO-compatible `oox::drawingml::Color::getColor` computes.
 */
import { describe, expect, it } from "vitest";

import {
  applyColorTransforms,
  applyShade,
  applyTint,
  applyTintRgb01,
  rgb01ToHex,
  spreadsheetTintTransforms,
  wordThemeShadeTransforms,
  wordThemeTintTransforms
} from "../theme-colors";
import type { ColorTransform } from "../theme-colors";

const hex = (h: string) => ({
  r: parseInt(h.slice(0, 2), 16) / 255,
  g: parseInt(h.slice(2, 4), 16) / 255,
  b: parseInt(h.slice(4, 6), 16) / 255
});
const apply = (h: string, transforms: readonly ColorTransform[]) =>
  rgb01ToHex(applyColorTransforms(hex(h), transforms));

/** Within one step per channel: Office rounds through integer HSL. */
function near(actual: string, expected: string) {
  for (let i = 0; i < 6; i += 2) {
    expect(
      Math.abs(parseInt(actual.slice(i, i + 2), 16) - parseInt(expected.slice(i, i + 2), 16)),
      `${actual} vs ${expected}`
    ).toBeLessThanOrEqual(2);
  }
}

describe("applyColorTransforms — DrawingML §20.1.2.3", () => {
  it("lumMod / lumOff act on HSL luminance: Excel's automatic gridline is #D9D9D9", () => {
    expect(
      apply("000000", [
        { kind: "lumMod", val: 15000 },
        { kind: "lumOff", val: 85000 }
      ])
    ).toBe("D9D9D9");
    // Office's named variants of accent1 (4472C4).
    near(apply("4472C4", [{ kind: "lumMod", val: 75000 }]), "2F5597");
    near(
      apply("4472C4", [
        { kind: "lumMod", val: 60000 },
        { kind: "lumOff", val: 40000 }
      ]),
      "8FAADC"
    );
    near(
      apply("000000", [
        { kind: "lumMod", val: 65000 },
        { kind: "lumOff", val: 35000 }
      ]),
      "595959"
    );
  });

  it("tint and shade mix in linear RGB (gamma 2.3), as Office does", () => {
    // 50 % shade of white: half the linear intensity, not half the sRGB value.
    near(apply("FFFFFF", [{ kind: "shade", val: 50000 }]), "BDBDBD");
    near(apply("000000", [{ kind: "tint", val: 50000 }]), "BDBDBD");
    expect(apply("4472C4", [{ kind: "tint", val: 100000 }])).toBe("4472C4");
    expect(apply("4472C4", [{ kind: "shade", val: 0 }])).toBe("000000");
  });

  it("applies the transforms in document order", () => {
    const a = apply("4472C4", [
      { kind: "tint", val: 50000 },
      { kind: "lumMod", val: 50000 }
    ]);
    const b = apply("4472C4", [
      { kind: "lumMod", val: 50000 },
      { kind: "tint", val: 50000 }
    ]);
    expect(a).not.toBe(b);
  });

  it("hue, saturation, complement, inverse, gray and gamma", () => {
    near(apply("FF0000", [{ kind: "hueOff", val: 120 * 60000 }]), "00FF00");
    near(apply("FF0000", [{ kind: "hue", val: 240 * 60000 }]), "0000FF");
    near(apply("FF0000", [{ kind: "comp" }]), "00FFFF");
    near(apply("FF0000", [{ kind: "satMod", val: 0 }]), "808080");
    near(apply("C04040", [{ kind: "satOff", val: -100000 }]), "808080");
    expect(apply("FFFFFF", [{ kind: "inv" }])).toBe("000000");
    // gray: weighted 22 % red, 72 % green, 6 % blue.
    near(apply("FF0000", [{ kind: "gray" }]), "383838");
    expect(apply("808080", [{ kind: "gamma" }, { kind: "invGamma" }])).toBe("808080");
    near(apply("FF0000", [{ kind: "greenOff", val: 100000 }]), "FFFF00");
    near(apply("808080", [{ kind: "redMod", val: 0 }]), "008080");
  });

  it("alpha, alphaMod and alphaOff set the opacity", () => {
    expect(applyColorTransforms(hex("000000"), [{ kind: "alpha", val: 40000 }]).a).toBeCloseTo(0.4);
    expect(
      applyColorTransforms(hex("000000"), [
        { kind: "alpha", val: 40000 },
        { kind: "alphaMod", val: 50000 },
        { kind: "alphaOff", val: 10000 }
      ]).a
    ).toBeCloseTo(0.3);
    expect(applyColorTransforms(hex("000000"), []).a).toBe(1);
  });
});

describe("Office's other tints, through the same engine", () => {
  it("SpreadsheetML cell tint (§18.8.19) is HSL luminance", () => {
    expect(spreadsheetTintTransforms(0.4)).toEqual([
      { kind: "lumMod", val: 60000 },
      { kind: "lumOff", val: 40000 }
    ]);
    expect(spreadsheetTintTransforms(-0.25)).toEqual([{ kind: "lumMod", val: 75000 }]);
    expect(spreadsheetTintTransforms(0)).toEqual([]);
    // Excel's "Blue, Accent 1, Lighter 40%" and "Darker 25%".
    near(rgb01ToHex(applyTintRgb01(hex("4472C4"), 0.3999755851924192)), "8FAADC");
    near(rgb01ToHex(applyTintRgb01(hex("4472C4"), -0.249977111117893)), "2F5597");
  });

  it("Word themeTint / themeShade are HSL luminance: Lighter 40 %, Darker 25 %", () => {
    // themeTint="99" is 60 % colour (Lighter 40 %); themeShade="BF" is 75 % (Darker 25 %).
    near(apply("4472C4", wordThemeTintTransforms(0x99)), "8EAADB");
    near(apply("4472C4", wordThemeShadeTransforms(0xbf)), "2F5496");
  });

  it("applyTint / applyShade keep their meaning (toward white / black) on Office's curve", () => {
    expect(applyTint("000000", 1)).toBe("FFFFFF");
    expect(applyTint("808080", 0)).toBe("808080");
    expect(applyShade("FFFFFF", 0)).toBe("000000");
    expect(applyShade("FFFFFF", 1)).toBe("FFFFFF");
    near(applyTint("4472C4", 0.4), "8FAADC");
    near(applyShade("4472C4", 0.75), "2F5597");
  });
});
