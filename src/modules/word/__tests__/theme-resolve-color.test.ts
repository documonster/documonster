/**
 * `Theme.resolveColor` answers with what its TSDoc promises — six hex digits,
 * or `undefined` when there is no colour to give — for the forms it documents:
 * a `HexColor` (`RRGGBB` or `"auto"`) or a `ColorSpec`. A value in neither form
 * is refused, not handed back as if it were a colour.
 */

import { describe, it, expect } from "vitest";

import { DocxArgumentError } from "../errors";
import { Theme } from "../index";
import type { ColorSpec, DocumentTheme } from "../types";

const theme: DocumentTheme = {
  name: "T",
  colorScheme: {
    name: "T",
    colors: {
      dk1: "1F1F1F",
      lt1: "FFFFFF",
      dk2: "303030",
      lt2: "F2F2F2",
      accent1: "C0392B",
      accent2: "E67E22",
      accent3: "F1C40F",
      accent4: "27AE60",
      accent5: "2980B9",
      accent6: "8E44AD",
      hlink: "0563C1",
      folHlink: "954F72"
    }
  },
  fontScheme: { name: "T", majorFont: "Cambria", minorFont: "Calibri" }
};

function refused(label: string, run: () => unknown, field: string): void {
  let error: unknown;
  try {
    run();
  } catch (e) {
    error = e;
  }
  expect(error, label).toBeInstanceOf(DocxArgumentError);
  expect((error as DocxArgumentError).field, label).toBe(field);
}

describe("Theme.resolveColor", () => {
  it("refuses a value that is not a documented colour form", () => {
    refused("bogus", () => Theme.resolveColor("bogus"), "color");
    refused("#hex", () => Theme.resolveColor("#FF0000"), "color");
    refused("short", () => Theme.resolveColor("F00"), "color");
    refused("number", () => Theme.resolveColor(42 as never), "color");
    refused("null", () => Theme.resolveColor(null as never), "color");
    refused("bad val", () => Theme.resolveColor({ val: "bogus" }), "color.val");
    refused(
      "bad tint",
      () => Theme.resolveColor({ val: "auto", themeColor: "accent1", themeTint: "zz" }, theme),
      "color.themeTint"
    );
    refused(
      "bad shade",
      () => Theme.resolveColor({ val: "auto", themeColor: "accent1", themeShade: "1234" }, theme),
      "color.themeShade"
    );
    refused(
      "bad themeColor",
      () => Theme.resolveColor({ val: "auto", themeColor: 7 } as never),
      "color.themeColor"
    );
    refused(
      "bad theme",
      () => Theme.resolveColor({ val: "auto", themeColor: "accent1" }, 42 as never),
      "theme"
    );
  });

  it("gives six hex digits, or undefined when there is no colour", () => {
    expect(Theme.resolveColor("FF8800")).toBe("FF8800");
    expect(Theme.resolveColor({ val: "00FF00" })).toBe("00FF00");
    expect(Theme.resolveColor(undefined)).toBeUndefined();
    // "auto" is the application's default colour: there is no RGB to give.
    expect(Theme.resolveColor("auto")).toBeUndefined();
    expect(Theme.resolveColor({ val: "auto" })).toBeUndefined();
    expect(Theme.resolveColor({ val: "auto", themeColor: "accent1" })).toBeUndefined();
    expect(Theme.resolveColor({ val: "auto", themeColor: "accent1" }, theme)).toBe("C0392B");
    // A theme colour this theme does not define falls back to `val`.
    expect(Theme.resolveColor({ val: "123456", themeColor: "unknownSlot" }, theme)).toBe("123456");
  });

  it("resolves Word's theme colour names through the default colour mapping", () => {
    const cases: [string, string][] = [
      ["text1", "1F1F1F"],
      ["background1", "FFFFFF"],
      ["text2", "303030"],
      ["background2", "F2F2F2"],
      ["dark1", "1F1F1F"],
      ["hyperlink", "0563C1"]
    ];
    for (const [name, hex] of cases) {
      const spec: ColorSpec = { val: "auto", themeColor: name };
      expect(Theme.resolveColor(spec, theme), name).toBe(hex);
    }
  });

  it("applies themeTint / themeShade to the luminance, as Word does", () => {
    const office = {
      ...theme,
      colorScheme: {
        ...theme.colorScheme,
        colors: { ...theme.colorScheme.colors, accent1: "4472C4" }
      }
    };
    const near = (actual: string | undefined, expected: string) => {
      for (let i = 0; i < 6; i += 2) {
        expect(
          Math.abs(parseInt(actual!.slice(i, i + 2), 16) - parseInt(expected.slice(i, i + 2), 16)),
          `${actual} vs ${expected}`
        ).toBeLessThanOrEqual(2);
      }
    };
    // Word's "Blue, Accent 1, Lighter 40%" (themeTint="99") and "Darker 25%" (themeShade="BF").
    near(
      Theme.resolveColor({ val: "auto", themeColor: "accent1", themeTint: "99" }, office),
      "8EAADB"
    );
    near(
      Theme.resolveColor({ val: "auto", themeColor: "accent1", themeShade: "BF" }, office),
      "2F5496"
    );
    // "Lighter 80%" is themeTint="33": mostly white.
    near(
      Theme.resolveColor({ val: "auto", themeColor: "accent1", themeTint: "33" }, office),
      "DAE3F3"
    );
  });
});
