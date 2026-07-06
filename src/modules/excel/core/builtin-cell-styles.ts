import type { NamedStyle } from "@excel/types";

/**
 * Built-in Excel cell-style presets, keyed by a stable identifier. Each entry
 * carries the OOXML `builtinId` that Excel and accessibility software use to
 * recognise the style's semantic role, plus approximate default-theme
 * formatting so the style is visually meaningful without a theme part.
 *
 * The `name` is the exact OOXML `cellStyle` name (e.g. "Heading 1"); it is what
 * cells reference via `Style.styleName`.
 */
export interface BuiltinCellStyleDef extends NamedStyle {
  name: string;
  builtinId: number;
}

// Default-theme accent/text colours (theme indices) used by Excel's presets.
const THEME_DARK1 = { theme: 1 } as const; // text 1 (near-black)
const HEADING_BLUE = { argb: "FF1F4E79" } as const; // Heading blue (accent-ish)
const TITLE_BLUE = { argb: "FF1F4E79" } as const;
const GOOD_TEXT = { argb: "FF006100" } as const;
const GOOD_FILL = { argb: "FFC6EFCE" } as const;
const BAD_TEXT = { argb: "FF9C0006" } as const;
const BAD_FILL = { argb: "FFFFC7CE" } as const;
const NEUTRAL_TEXT = { argb: "FF9C6500" } as const;
const NEUTRAL_FILL = { argb: "FFFFEB9C" } as const;

function solidFill(argb: string): NamedStyle["fill"] {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

export const BUILTIN_CELL_STYLES = {
  Title: {
    name: "Title",
    builtinId: 15,
    font: { name: "Calibri Light", size: 18, bold: false, color: TITLE_BLUE }
  },
  Heading1: {
    name: "Heading 1",
    builtinId: 16,
    font: { size: 15, bold: true, color: HEADING_BLUE }
  },
  Heading2: {
    name: "Heading 2",
    builtinId: 17,
    font: { size: 13, bold: true, color: HEADING_BLUE }
  },
  Heading3: {
    name: "Heading 3",
    builtinId: 18,
    font: { size: 11, bold: true, color: HEADING_BLUE }
  },
  Heading4: {
    name: "Heading 4",
    builtinId: 19,
    font: { size: 11, bold: true, color: THEME_DARK1 }
  },
  Good: {
    name: "Good",
    builtinId: 26,
    font: { color: GOOD_TEXT },
    fill: solidFill(GOOD_FILL.argb)
  },
  Bad: {
    name: "Bad",
    builtinId: 27,
    font: { color: BAD_TEXT },
    fill: solidFill(BAD_FILL.argb)
  },
  Neutral: {
    name: "Neutral",
    builtinId: 28,
    font: { color: NEUTRAL_TEXT },
    fill: solidFill(NEUTRAL_FILL.argb)
  }
} as const satisfies Record<string, BuiltinCellStyleDef>;

/** Identifier keys for {@link BUILTIN_CELL_STYLES}. */
export type BuiltinCellStyle = keyof typeof BUILTIN_CELL_STYLES;
