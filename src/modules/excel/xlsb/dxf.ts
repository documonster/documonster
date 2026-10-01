/**
 * `BrtDXF` — a differential format, the formatting a conditional-formatting rule (and a table style, and a
 * pivot table's `<formats>`) applies *on top of* a cell's own.
 *
 * ```text
 * BrtDXF          flags (2 bytes, bit 15 fNewBorder), then XFProps
 * XFProps         reserved (2), cprops (2), then cprops × XFProp
 * XFProp          xfPropType (2), cb (2, the size of the whole XFProp), xfPropDataBlob
 * ```
 *
 * MS-XLSB 2.4.380, 2.5.159–2.5.165. Every `xfPropType` the specification defines is listed in {@link PROPERTY},
 * and every one the model's `Style` can express is written *and* read — in both directions, from one table.
 *
 * **Why both directions, and why they report.** This used to write a subset and read a smaller one: the reader
 * had no branch for the underline the writer emitted, so an underlined rule came back plain; alignment, number
 * format and protection were written by neither, silently; a patterned fill was always written solid; and a
 * palette colour became "not set". Each is a format a user sees change after a save with nothing telling them.
 * So the encoder returns what it could *not* express ({@link DxfEncoding.dropped}) and the decoder what it could
 * not read ({@link DxfDecoding.unread}), and the callers put both in the fidelity report.
 *
 * **Absent and explicit are different here, and the model keeps them apart.** A differential format changes only
 * what it states: `bold: false` removes bold where no `bold` leaves it alone, and a fill with colours and no
 * pattern keeps the cell's pattern (Excel paints it solid). So `false`, `0` and a missing pattern are all written
 * and read as themselves rather than folded into "absent", which is what a cell format is free to do.
 */
import type {
  Alignment,
  Border,
  BorderStyle,
  Borders,
  Color,
  Fill,
  Font,
  GradientStop,
  Protection,
  Style
} from "@excel/types";
import { borderStyleName, borderStyleValue } from "@excel/xlsb/border";
import { fillPatternName, fillPatternValue } from "@excel/xlsb/fill";
import { SCHEME, UNDERLINE, VERTICAL_ALIGN, nameOf, underlineValue } from "@excel/xlsb/font";
import { dxfExtensionsOf } from "@excel/xlsx/xform/style/dxf-extensions";
import { BinaryWriter, concatUint8Arrays } from "@utils/binary";

/** `xfPropType` values — MS-XLSB 2.5.159. 0x27 and 0x28 are not defined. */
const PROPERTY = {
  fillPattern: 0x00,
  fillForeground: 0x01,
  fillBackground: 0x02,
  gradient: 0x03,
  gradientStop: 0x04,
  fontColor: 0x05,
  borderTop: 0x06,
  borderBottom: 0x07,
  borderLeft: 0x08,
  borderRight: 0x09,
  borderDiagonal: 0x0a,
  borderVertical: 0x0b,
  borderHorizontal: 0x0c,
  diagonalUp: 0x0d,
  diagonalDown: 0x0e,
  horizontalAlignment: 0x0f,
  verticalAlignment: 0x10,
  textRotation: 0x11,
  indent: 0x12,
  readingOrder: 0x13,
  wrapText: 0x14,
  justifyLastLine: 0x15,
  shrinkToFit: 0x16,
  mergeCell: 0x17,
  fontName: 0x18,
  fontWeight: 0x19,
  underline: 0x1a,
  script: 0x1b,
  italic: 0x1c,
  strike: 0x1d,
  outline: 0x1e,
  shadow: 0x1f,
  condense: 0x20,
  extend: 0x21,
  charset: 0x22,
  fontFamily: 0x23,
  fontSize: 0x24,
  fontScheme: 0x25,
  numberFormatString: 0x26,
  numberFormatId: 0x29,
  relativeIndent: 0x2a,
  locked: 0x2b,
  hidden: 0x2c
} as const;

/** The border edge each `XFPropBorder` type formats, in the model's names. */
const BORDER_EDGE: readonly (readonly [number, keyof Borders])[] = [
  [PROPERTY.borderTop, "top"],
  [PROPERTY.borderBottom, "bottom"],
  [PROPERTY.borderLeft, "left"],
  [PROPERTY.borderRight, "right"],
  [PROPERTY.borderDiagonal, "diagonal"],
  [PROPERTY.borderVertical, "vertical"],
  [PROPERTY.borderHorizontal, "horizontal"]
];

/** `HorizAlign` — MS-XLSB 2.5.74, by value. A stated `general` resets alignment, so it is a value, not an absence. */
const HORIZONTAL: readonly Alignment["horizontal"][] = [
  "general",
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed"
];

/** `VertAlign` — MS-XLSB 2.5.158. Unlike a cell format's table, `bottom` is a stated value here, not a default. */
const VERTICAL: readonly Alignment["vertical"][] = [
  "top",
  "middle",
  "bottom",
  "justify",
  "distributed"
];

/** `ReadingOrder` — MS-XLSB 2.5.114. A stated `context` resets the direction, so it too is a value. */
const READING_ORDER: readonly Alignment["readingOrder"][] = ["context", "ltr", "rtl"];

/** `relativeIndent` meaning "no relative indentation" — MS-XLSB 2.5.159, type 0x2A. */
const NO_RELATIVE_INDENT = 0xff;

/** `Bold` — MS-XLSB 2.5.6. An enumeration of two weights, not a flag. */
const WEIGHT_BOLD = 0x02bc;
const WEIGHT_NORMAL = 0x0190;

/** `XFPropTextRotation` for vertically stacked text — MS-XLSB 2.5.165. */
const ROTATION_VERTICAL = 0xff;

/** `fNewBorder`, bit 15 of `BrtDXF`'s flag word: the inner-border types 0x0B and 0x0C are permitted. */
const DXF_NEW_BORDER = 0x8000;

/** Twips per point, for a font size. */
const TWIPS_PER_POINT = 20;

/**
 * The fixed size of each property's data, in bytes — MS-XLSB 2.5.159 and the structures it names. `fontName` is a
 * minimum (its `LPWideString` count; the characters follow) and `numberFormatString` is left out, since it is
 * never decoded. A property whose `cb` leaves less than this is reported, not read: reading it anyway took bytes
 * from the *next* property, or threw a `RangeError` out of the whole styles part.
 */
const DATA_SIZE: Readonly<Record<number, number>> = {
  [PROPERTY.fillPattern]: 1,
  [PROPERTY.fillForeground]: 8,
  [PROPERTY.fillBackground]: 8,
  [PROPERTY.gradient]: 44,
  [PROPERTY.gradientStop]: 18,
  [PROPERTY.fontColor]: 8,
  [PROPERTY.borderTop]: 10,
  [PROPERTY.borderBottom]: 10,
  [PROPERTY.borderLeft]: 10,
  [PROPERTY.borderRight]: 10,
  [PROPERTY.borderDiagonal]: 10,
  [PROPERTY.borderVertical]: 10,
  [PROPERTY.borderHorizontal]: 10,
  [PROPERTY.diagonalUp]: 1,
  [PROPERTY.diagonalDown]: 1,
  [PROPERTY.horizontalAlignment]: 1,
  [PROPERTY.verticalAlignment]: 1,
  [PROPERTY.textRotation]: 1,
  [PROPERTY.indent]: 2,
  [PROPERTY.readingOrder]: 1,
  [PROPERTY.wrapText]: 1,
  [PROPERTY.justifyLastLine]: 1,
  [PROPERTY.shrinkToFit]: 1,
  [PROPERTY.mergeCell]: 1,
  [PROPERTY.fontName]: 2,
  [PROPERTY.fontWeight]: 2,
  [PROPERTY.underline]: 2,
  [PROPERTY.script]: 2,
  [PROPERTY.italic]: 1,
  [PROPERTY.strike]: 1,
  [PROPERTY.outline]: 1,
  [PROPERTY.shadow]: 1,
  [PROPERTY.condense]: 1,
  [PROPERTY.extend]: 1,
  [PROPERTY.charset]: 1,
  [PROPERTY.fontFamily]: 1,
  [PROPERTY.fontSize]: 4,
  [PROPERTY.fontScheme]: 1,
  [PROPERTY.numberFormatId]: 2,
  [PROPERTY.relativeIndent]: 2,
  [PROPERTY.locked]: 1,
  [PROPERTY.hidden]: 1
};

/** `XFPropColor.xclrType` — MS-XLSB 2.5.161. */
const COLOR_TYPE = {
  automatic: 0x00,
  indexed: 0x01,
  rgb: 0x02,
  theme: 0x03,
  notSet: 0x04
} as const;

/** What {@link encodeDxf} produced. */
export interface DxfEncoding {
  /** The `BrtDXF` payload, or `undefined` when nothing in the style could be written. */
  readonly payload: Uint8Array | undefined;
  /** Dotted paths of properties the style carries and the record could not express. */
  readonly dropped: readonly string[];
}

/** What {@link readDxf} produced. */
export interface DxfDecoding {
  /** The format, or `undefined` when the record held nothing the model can express. */
  readonly style: Partial<Style> | undefined;
  /** Properties the record carried that the model cannot hold, by name. */
  readonly unread: readonly string[];
}

// =============================================================================
// Writing
// =============================================================================

/**
 * Encode a differential format.
 *
 * @param formatId - The `ifmt` a number-format code is written under. The format has to be declared by a `BrtFmt`
 *   in the same styles part (or be built in), so the caller that owns that table decides the number.
 */
export function encodeDxf(style: Partial<Style>, formatId: (code: string) => number): DxfEncoding {
  const props: { type: number; bytes: Uint8Array }[] = [];
  const dropped: string[] = [];
  const add = (type: number, blob: Uint8Array): void => {
    props.push({ type, bytes: xfProp(type, blob) });
  };
  const flag = (type: number, value: boolean | undefined): void => {
    if (value !== undefined) {
      add(type, u8(value ? 1 : 0));
    }
  };
  const unknownKeys = (
    group: string,
    value: object | undefined,
    known: readonly string[]
  ): void => {
    for (const key of Object.keys(value ?? {})) {
      if (!known.includes(key) && (value as Record<string, unknown>)[key] !== undefined) {
        dropped.push(`${group}.${key}`);
      }
    }
  };

  encodeFill(style.fill, add, dropped);
  encodeFont(style.font, add, flag, dropped);
  unknownKeys("font", style.font, FONT_KEYS);
  encodeBorders(style.border, add, flag, dropped);
  unknownKeys("border", style.border, BORDER_KEYS);
  encodeAlignment(style.alignment, add, flag, dropped);
  unknownKeys("alignment", style.alignment, ALIGNMENT_KEYS);

  if (style.numFmt !== undefined) {
    const code = typeof style.numFmt === "string" ? style.numFmt : style.numFmt.formatCode;
    add(PROPERTY.numberFormatId, u16(formatId(code)));
  }

  flag(PROPERTY.locked, style.protection?.locked);
  flag(PROPERTY.hidden, style.protection?.hidden);
  unknownKeys("protection", style.protection, ["locked", "hidden"]);

  // An XLSX `<extLst>` the reader preserved: `BrtDXF` has no extension point to carry it in.
  if (dxfExtensionsOf(style) !== undefined) {
    dropped.push("extLst");
  }

  // Anything at the top level beyond the six facets a differential format has. `numFmtId` is the XLSX
  // writer's own bookkeeping and `styleName` names a cell style, which a differential format cannot apply.
  for (const key of Object.keys(style)) {
    if (!TOP_LEVEL_KEYS.includes(key) && (style as Record<string, unknown>)[key] !== undefined) {
      dropped.push(key);
    }
  }

  // Ascending by type. The specification constrains which types may coexist, not their order, but Excel writes
  // them sorted; the one ordering rule that exists — a gradient (3) followed by its stops (4) — survives a sort
  // because the sort is stable and the stops are pushed after the gradient.
  props.sort((left, right) => left.type - right.type);
  if (props.length === 0) {
    // A differential format that changes nothing is not writable: Excel discards the whole `DXFs` collection when
    // it meets one, and with it every other rule's formatting. The caller gives the style no `dxfId`.
    return { payload: undefined, dropped };
  }
  return {
    payload: concatUint8Arrays([
      // `fNewBorder` is a *capability* — the inner-border types are permitted — not a claim that any is present.
      // Excel sets it on every `BrtDXF` it writes.
      u16(DXF_NEW_BORDER),
      new BinaryWriter().writeUint16(0).writeUint16(props.length).toUint8Array(),
      ...props.map(property => property.bytes)
    ]),
    dropped
  };
}

const TOP_LEVEL_KEYS = ["font", "fill", "border", "alignment", "protection", "numFmt", "numFmtId"];
const FONT_KEYS = [
  "color",
  "name",
  "bold",
  "underline",
  "vertAlign",
  "italic",
  "strike",
  "outline",
  "shadow",
  "condense",
  "extend",
  "charset",
  "family",
  "size",
  "scheme"
];
const BORDER_KEYS = [
  "top",
  "bottom",
  "left",
  "right",
  "diagonal",
  "vertical",
  "horizontal",
  "color"
];
const ALIGNMENT_KEYS = [
  "horizontal",
  "vertical",
  "textRotation",
  "indent",
  "readingOrder",
  "wrapText",
  "shrinkToFit",
  "justifyLastLine",
  "relativeIndent"
];

type Add = (type: number, blob: Uint8Array) => void;
type Flag = (type: number, value: boolean | undefined) => void;

function encodeFill(fill: Fill | undefined, add: Add, dropped: string[]): void {
  if (fill === undefined) {
    return;
  }
  if (fill.type === "pattern") {
    // An absent pattern is written as no pattern property — the differential "keep the cell's pattern", with the
    // colour painted solid — not as `FLSNULL`, which removes the fill.
    if (fill.pattern !== undefined) {
      const value = fillPatternValue(fill.pattern);
      if (value === undefined) {
        dropped.push("fill.pattern");
      } else {
        add(PROPERTY.fillPattern, u8(value));
      }
    }
    if (fill.fgColor !== undefined) {
      add(PROPERTY.fillForeground, xfPropColor(fill.fgColor));
    }
    if (fill.bgColor !== undefined) {
      add(PROPERTY.fillBackground, xfPropColor(fill.bgColor));
    }
    return;
  }
  // `XFPropGradient` — MS-XLSB 2.5.162: a type, then five `Xnum`s. The four rectangle coordinates are the
  // XLSX `left`/`right`/`top`/`bottom`, which the model folds into `center` and leaves `right`/`bottom` out of
  // when they repeat `left`/`top`.
  const center = (
    fill as { center?: { left?: number; top?: number; right?: number; bottom?: number } }
  ).center;
  const path = fill.gradient === "path";
  const left = path ? (center?.left ?? 0) : 0;
  const top = path ? (center?.top ?? 0) : 0;
  add(
    PROPERTY.gradient,
    new BinaryWriter()
      .writeInt32(path ? 1 : 0)
      .writeFloat64(path ? 0 : ((fill as { degree?: number }).degree ?? 0))
      .writeFloat64(left)
      .writeFloat64(path ? (center?.right ?? left) : 0)
      .writeFloat64(top)
      .writeFloat64(path ? (center?.bottom ?? top) : 0)
      .toUint8Array()
  );
  for (const stop of fill.stops ?? []) {
    add(
      PROPERTY.gradientStop,
      concatUint8Arrays([
        new BinaryWriter().writeUint16(0).writeFloat64(stop.position).toUint8Array(),
        xfPropColor(stop.color)
      ])
    );
  }
}

function encodeFont(
  font: Partial<Font> | undefined,
  add: Add,
  flag: Flag,
  dropped: string[]
): void {
  if (font === undefined) {
    return;
  }
  if (font.color !== undefined) {
    add(PROPERTY.fontColor, xfPropColor(font.color));
  }
  if (font.name !== undefined) {
    // An `LPWideString` here "MUST be less than or equal to 32 characters". A longer name cannot be written
    // truncated — that names a different font — so it is not written at all.
    if ([...font.name].length > 32) {
      dropped.push("font.name");
    } else {
      add(PROPERTY.fontName, lpWideString(font.name));
    }
  }
  if (font.bold !== undefined) {
    add(PROPERTY.fontWeight, u16(font.bold ? WEIGHT_BOLD : WEIGHT_NORMAL));
  }
  if (font.underline !== undefined) {
    // `false` and `"none"` are written as `ULSNONE`: in a differential format they remove an underline.
    add(PROPERTY.underline, u16(underlineValue(font.underline)));
  }
  if (font.vertAlign !== undefined) {
    const value =
      font.vertAlign === "baseline"
        ? VERTICAL_ALIGN.none
        : VERTICAL_ALIGN[font.vertAlign as keyof typeof VERTICAL_ALIGN];
    if (value === undefined) {
      dropped.push("font.vertAlign");
    } else {
      add(PROPERTY.script, u16(value));
    }
  }
  flag(PROPERTY.italic, font.italic);
  flag(PROPERTY.strike, font.strike);
  flag(PROPERTY.outline, font.outline);
  flag(PROPERTY.shadow, font.shadow);
  flag(PROPERTY.condense, font.condense);
  flag(PROPERTY.extend, font.extend);
  if (font.charset !== undefined) {
    if (Number.isInteger(font.charset) && font.charset >= 0 && font.charset <= 0xff) {
      add(PROPERTY.charset, u8(font.charset));
    } else {
      dropped.push("font.charset");
    }
  }
  if (font.family !== undefined) {
    // "MUST be greater than or equal to 0 and less than or equal to 5."
    if (Number.isInteger(font.family) && font.family >= 0 && font.family <= 5) {
      add(PROPERTY.fontFamily, u8(font.family));
    } else {
      dropped.push("font.family");
    }
  }
  if (font.size !== undefined) {
    // Twips, and "MUST be greater than or equal to 20 and less than or equal to 8191".
    const twips = Math.round(font.size * TWIPS_PER_POINT);
    if (twips >= 20 && twips <= 8191) {
      add(PROPERTY.fontSize, u32(twips));
    } else {
      dropped.push("font.size");
    }
  }
  if (font.scheme !== undefined) {
    const value = SCHEME[font.scheme];
    if (value === undefined) {
      dropped.push("font.scheme");
    } else {
      add(PROPERTY.fontScheme, u8(value));
    }
  }
}

function encodeBorders(
  border: Partial<Borders> | undefined,
  add: Add,
  flag: Flag,
  dropped: string[]
): void {
  if (border === undefined) {
    return;
  }
  // A border-wide `color` is the XLSX writer's shorthand for "every edge without its own"; applied the same way.
  const fallback = (border as { color?: Partial<Color> }).color;
  for (const [type, edge] of BORDER_EDGE) {
    const value = border[edge] as Partial<Border> | undefined;
    if (value === undefined) {
      continue;
    }
    if (value.style !== undefined && borderStyleValue(value.style) === 0) {
      dropped.push(`border.${edge}.style`);
      continue;
    }
    add(
      type,
      concatUint8Arrays([
        xfPropColor(value.color ?? fallback),
        u16(borderStyleValue(value.style as BorderStyle | undefined))
      ])
    );
  }
  flag(PROPERTY.diagonalUp, border.diagonal?.up);
  flag(PROPERTY.diagonalDown, border.diagonal?.down);
}

function encodeAlignment(
  alignment: Partial<Alignment> | undefined,
  add: Add,
  flag: Flag,
  dropped: string[]
): void {
  if (alignment === undefined) {
    return;
  }
  const enumerated = <T>(
    type: number,
    table: readonly (T | undefined)[],
    value: T | undefined,
    name: string
  ): void => {
    if (value === undefined) {
      return;
    }
    const index = table.indexOf(value);
    if (index < 0) {
      dropped.push(`alignment.${name}`);
    } else {
      add(type, u8(index));
    }
  };
  enumerated(PROPERTY.horizontalAlignment, HORIZONTAL, alignment.horizontal, "horizontal");
  enumerated(PROPERTY.verticalAlignment, VERTICAL, alignment.vertical, "vertical");
  if (alignment.textRotation !== undefined) {
    const rotation = textRotationValue(alignment.textRotation);
    if (rotation === undefined) {
      dropped.push("alignment.textRotation");
    } else {
      add(PROPERTY.textRotation, u8(rotation));
    }
  }
  if (alignment.indent !== undefined) {
    // "MUST be less than or equal to 15."
    if (Number.isInteger(alignment.indent) && alignment.indent >= 0 && alignment.indent <= 15) {
      add(PROPERTY.indent, u16(alignment.indent));
    } else {
      dropped.push("alignment.indent");
    }
  }
  enumerated(PROPERTY.readingOrder, READING_ORDER, alignment.readingOrder, "readingOrder");
  flag(PROPERTY.wrapText, alignment.wrapText);
  flag(PROPERTY.shrinkToFit, alignment.shrinkToFit);
  if (alignment.justifyLastLine !== undefined) {
    // "If this value is 1, then an XFProp with xfPropType equal to 0x000F MUST exist … and MUST equal 0x07" —
    // justify-distributed only means something for distributed text, and the record may not say it otherwise.
    if (alignment.justifyLastLine && alignment.horizontal !== "distributed") {
      dropped.push("alignment.justifyLastLine");
    } else {
      flag(PROPERTY.justifyLastLine, alignment.justifyLastLine);
    }
  }
  if (alignment.relativeIndent !== undefined) {
    // A 2-byte signed integer, -15…15 ("or 255", the absence of one — which the model spells by omission).
    const value = alignment.relativeIndent;
    if (Number.isInteger(value) && value >= -15 && value <= 15) {
      add(PROPERTY.relativeIndent, i16(value));
    } else {
      dropped.push("alignment.relativeIndent");
    }
  }
}

/**
 * `trot` for the model's rotation: 0–90 counter-clockwise as itself, and the model's -1…-90 (clockwise) as
 * 91–180 — the same mapping the XLSX `textRotation` attribute uses.
 */
function textRotationValue(rotation: number | "vertical"): number | undefined {
  if (rotation === "vertical") {
    return ROTATION_VERTICAL;
  }
  if (!Number.isInteger(rotation) || rotation < -90 || rotation > 90) {
    return undefined;
  }
  return rotation >= 0 ? rotation : 90 - rotation;
}

/** One `XFProp`: a type, the size of the *whole* structure, then the blob. */
function xfProp(type: number, blob: Uint8Array): Uint8Array {
  return concatUint8Arrays([
    new BinaryWriter()
      .writeUint16(type)
      .writeUint16(blob.length + 4)
      .toUint8Array(),
    blob
  ]);
}

/**
 * An `XFPropColor` — MS-XLSB 2.5.161. Eight bytes: `fValidRGBA` and `xclrType` sharing a byte, `icv`, the tint,
 * then `LongRGBA`.
 */
function xfPropColor(color: Partial<Color> | undefined): Uint8Array {
  const writer = new BinaryWriter();
  if (color?.theme !== undefined) {
    writer.writeUint8(COLOR_TYPE.theme << 1).writeUint8(color.theme & 0xff);
  } else if (typeof color?.argb === "string") {
    // Bit 0 is `fValidRGBA`: `dwRgba` holds the colour itself.
    writer.writeUint8(0x01 | (COLOR_TYPE.rgb << 1)).writeUint8(0);
  } else if (color?.indexed !== undefined) {
    // A palette colour. It used to be written "not set", so an indexed fill came back with no colour.
    writer.writeUint8(COLOR_TYPE.indexed << 1).writeUint8(color.indexed & 0xff);
  } else {
    writer.writeUint8(COLOR_TYPE.notSet << 1).writeUint8(0);
  }
  // `nTintShade` maps to -1.0…1.0 and MUST NOT be -32768; written unsigned, two's complement done here.
  const tint = Math.max(-32767, Math.min(32767, Math.round((color?.tint ?? 0) * 32767)));
  writer.writeUint16(tint < 0 ? tint + 0x10000 : tint);
  const argb = typeof color?.argb === "string" ? color.argb : "00000000";
  // `LongRGBA` is red, green, blue, alpha — *not* the ARGB order the string spells.
  const bytes = argb.padStart(8, "0").slice(-8);
  const at = (index: number): number => Number.parseInt(bytes.slice(index, index + 2), 16) || 0;
  writer.writeUint8(at(2)).writeUint8(at(4)).writeUint8(at(6)).writeUint8(at(0));
  return writer.toUint8Array();
}

/**
 * An `LPWideString`: a **two-byte** character count, then UTF-16 — MS-XLSB 2.5.92. Writing the count as one byte
 * put the name a byte early, and Excel discarded the whole `DXFs` collection over it.
 */
function lpWideString(value: string): Uint8Array {
  const writer = new BinaryWriter().writeUint16(value.length);
  for (let index = 0; index < value.length; index++) {
    writer.writeUint16(value.charCodeAt(index));
  }
  return writer.toUint8Array();
}

function u8(value: number): Uint8Array {
  return new BinaryWriter().writeUint8(value).toUint8Array();
}

function i16(value: number): Uint8Array {
  return u16(value < 0 ? value + 0x10000 : value);
}

function u16(value: number): Uint8Array {
  return new BinaryWriter().writeUint16(value).toUint8Array();
}

function u32(value: number): Uint8Array {
  return new BinaryWriter().writeUint32(value).toUint8Array();
}

// =============================================================================
// Reading
// =============================================================================

/**
 * Decode a differential format.
 *
 * Walked by `cb` rather than by a fixed field order: the specification constrains which types may coexist, not
 * their sequence, and `cb` covering the whole `XFProp` is what makes the walk possible at all.
 *
 * @param formatById - The format code an `ifmt` names: the styles part's own `BrtFmt`, else the built-in code.
 */
export function readDxf(
  payload: Uint8Array,
  formatById: (id: number) => string | undefined
): DxfDecoding {
  const unread: string[] = [];
  if (payload.length < 6) {
    return { style: undefined, unread };
  }
  const view = new DataView(payload.buffer, payload.byteOffset, payload.length);
  const font: Partial<Font> & Record<string, unknown> = {};
  const fill: Record<string, unknown> = {};
  const stops: GradientStop[] = [];
  const border: Record<string, Record<string, unknown>> = {};
  const alignment: Partial<Alignment> = {};
  const protection: Partial<Protection> = {};
  let numFmt: string | undefined;
  const edge = (name: string): Record<string, unknown> => (border[name] ??= {});

  // Two flag bytes, then `XFProps`: two reserved and the property count, which "MUST match the number of XFProp
  // structures" — so a disagreement is reported, and the walk still reads what is actually there.
  const declared = view.getUint16(4, true);
  let walked = 0;
  let offset = 6;
  while (offset + 4 <= payload.length) {
    const type = view.getUint16(offset, true);
    const size = view.getUint16(offset + 2, true);
    if (size < 4 || offset + size > payload.length) {
      unread.push("truncated property");
      break;
    }
    walked++;
    const at = offset + 4;
    const minimum = DATA_SIZE[type];
    if (minimum !== undefined && size - 4 < minimum) {
      // Too short for its own type. Skipped by its `cb`, which is still trustworthy — it was bounds-checked above
      // — so every property after it is read normally.
      unread.push(`truncated property 0x${type.toString(16).padStart(2, "0")}`);
      offset += size;
      continue;
    }
    const byte = (): number => view.getUint8(at);
    const bool = (): boolean => view.getUint8(at) === 1;
    switch (type) {
      case PROPERTY.fillPattern: {
        const name = fillPatternName(byte());
        if (name === undefined) {
          unread.push("fill pattern");
        } else {
          fill.pattern = name;
        }
        break;
      }
      case PROPERTY.fillForeground:
        fill.fgColor = readColor(view, at);
        break;
      case PROPERTY.fillBackground:
        fill.bgColor = readColor(view, at);
        break;
      case PROPERTY.gradient: {
        const path = view.getInt32(at, true) === 1;
        fill.gradient = path ? "path" : "angle";
        if (path) {
          const left = view.getFloat64(at + 12, true);
          const right = view.getFloat64(at + 20, true);
          const top = view.getFloat64(at + 28, true);
          const bottom = view.getFloat64(at + 36, true);
          fill.center = {
            left,
            top,
            ...(right === left ? {} : { right }),
            ...(bottom === top ? {} : { bottom })
          };
        } else {
          fill.degree = view.getFloat64(at + 4, true);
        }
        break;
      }
      case PROPERTY.gradientStop:
        stops.push({ position: view.getFloat64(at + 2, true), color: readColor(view, at + 10) });
        break;
      case PROPERTY.fontColor:
        font.color = readColor(view, at);
        break;
      case PROPERTY.borderTop:
      case PROPERTY.borderBottom:
      case PROPERTY.borderLeft:
      case PROPERTY.borderRight:
      case PROPERTY.borderDiagonal:
      case PROPERTY.borderVertical:
      case PROPERTY.borderHorizontal: {
        const name = BORDER_EDGE.find(([candidate]) => candidate === type)![1];
        const target = edge(name);
        const color = readColor(view, at);
        const style = borderStyleName(view.getUint16(at + 8, true));
        if (style !== undefined) {
          target.style = style;
        }
        if (Object.keys(color).length > 0) {
          target.color = color;
        }
        break;
      }
      case PROPERTY.diagonalUp:
        edge("diagonal").up = bool();
        break;
      case PROPERTY.diagonalDown:
        edge("diagonal").down = bool();
        break;
      case PROPERTY.horizontalAlignment: {
        const value = HORIZONTAL[byte()];
        if (value === undefined) {
          unread.push("horizontal alignment");
        } else {
          alignment.horizontal = value;
        }
        break;
      }
      case PROPERTY.verticalAlignment: {
        const value = VERTICAL[byte()];
        if (value === undefined) {
          unread.push("vertical alignment");
        } else {
          alignment.vertical = value;
        }
        break;
      }
      case PROPERTY.textRotation: {
        const trot = byte();
        if (trot === ROTATION_VERTICAL) {
          alignment.textRotation = "vertical";
        } else if (trot <= 90) {
          alignment.textRotation = trot;
        } else if (trot <= 180) {
          alignment.textRotation = 90 - trot;
        } else {
          // 254 is "context dependent", which the model has no word for.
          unread.push("text rotation");
        }
        break;
      }
      case PROPERTY.indent:
        alignment.indent = view.getUint16(at, true);
        break;
      case PROPERTY.readingOrder: {
        const value = READING_ORDER[byte()];
        if (value === undefined) {
          unread.push("reading order");
        } else {
          alignment.readingOrder = value;
        }
        break;
      }
      case PROPERTY.wrapText:
        alignment.wrapText = bool();
        break;
      case PROPERTY.shrinkToFit:
        alignment.shrinkToFit = bool();
        break;
      case PROPERTY.justifyLastLine:
        alignment.justifyLastLine = bool();
        break;
      case PROPERTY.mergeCell:
        // Merging is a property of a *range* in the model (`Worksheet.mergeCells`), not of a format — and a
        // differential format cannot merge anything: a conditional format or table style that "merges" is not
        // a thing Excel offers. Excel never writes it here; a producer that does is reported, not obeyed.
        unread.push("merge cell");
        break;
      case PROPERTY.fontName:
        {
          const name = readLpWideString(view, at, offset + size);
          if (name === undefined) {
            unread.push("truncated property 0x18");
          } else {
            font.name = name;
          }
        }
        break;
      case PROPERTY.fontWeight: {
        // Two values only. Anything else is neither bold nor normal, so it is not reported as either.
        const weight = view.getUint16(at, true);
        if (weight === WEIGHT_BOLD || weight === WEIGHT_NORMAL) {
          font.bold = weight === WEIGHT_BOLD;
        } else {
          unread.push("font weight");
        }
        break;
      }
      case PROPERTY.underline: {
        const name = nameOf(UNDERLINE, view.getUint16(at, true));
        if (name === undefined) {
          unread.push("underline");
        } else {
          // `true` for a single underline and `false` for none: what the XLSX reader produces for `<u/>` and
          // `<u val="none"/>`, so the two containers agree about the same format.
          font.underline = name === "single" ? true : name === "none" ? false : name;
        }
        break;
      }
      case PROPERTY.script: {
        const name = nameOf(VERTICAL_ALIGN, view.getUint16(at, true));
        if (name === undefined) {
          unread.push("font script");
        } else {
          // `SSSNONE` is a stated baseline: it resets a super- or subscript.
          font.vertAlign = name === "none" ? "baseline" : name;
        }
        break;
      }
      case PROPERTY.italic:
        font.italic = bool();
        break;
      case PROPERTY.strike:
        font.strike = bool();
        break;
      case PROPERTY.outline:
        font.outline = bool();
        break;
      case PROPERTY.shadow:
        font.shadow = bool();
        break;
      case PROPERTY.condense:
        font.condense = bool();
        break;
      case PROPERTY.extend:
        font.extend = bool();
        break;
      case PROPERTY.charset:
        font.charset = byte();
        break;
      case PROPERTY.fontFamily:
        font.family = byte();
        break;
      case PROPERTY.fontSize:
        font.size = view.getUint32(at, true) / TWIPS_PER_POINT;
        break;
      case PROPERTY.fontScheme: {
        const name = nameOf(SCHEME, byte());
        if (name === undefined) {
          unread.push("font scheme");
        } else {
          font.scheme = name;
        }
        break;
      }
      case PROPERTY.numberFormatId: {
        const code = formatById(view.getUint16(at, true));
        if (code === undefined) {
          unread.push("number format");
        } else {
          numFmt = code;
        }
        break;
      }
      case PROPERTY.numberFormatString:
        // The format code itself, in a structure MS-XLSB defers to MS-XLS for. Excel pairs it with an `ifmt` (0x29),
        // which is read above, so it is only reported when there is no identifier to read instead — see below.
        break;
      case PROPERTY.relativeIndent: {
        const value = view.getInt16(at, true);
        if (value !== NO_RELATIVE_INDENT) {
          alignment.relativeIndent = value;
        }
        break;
      }
      case PROPERTY.locked:
        protection.locked = bool();
        break;
      case PROPERTY.hidden:
        protection.hidden = bool();
        break;
      default:
        unread.push(`property 0x${type.toString(16).padStart(2, "0")}`);
        break;
    }
    offset += size;
  }
  if (walked !== declared) {
    unread.push(`property count: ${declared} declared, ${walked} present`);
  }
  if (numFmt === undefined && hasProperty(view, payload.length, PROPERTY.numberFormatString)) {
    unread.push("number format");
  }

  const style: Partial<Style> = {};
  if (Object.keys(font).length > 0) {
    style.font = font;
  }
  if (fill.gradient !== undefined) {
    style.fill = { type: "gradient", ...fill, stops } as Fill;
  } else if (Object.keys(fill).length > 0) {
    // No pattern stays no pattern: the differential "keep the cell's pattern", as the XLSX reader models it.
    style.fill = { type: "pattern", ...fill } as Fill;
  }
  if (Object.keys(border).length > 0) {
    style.border = border as Partial<Borders>;
  }
  if (Object.keys(alignment).length > 0) {
    style.alignment = alignment;
  }
  if (Object.keys(protection).length > 0) {
    style.protection = protection;
  }
  if (numFmt !== undefined) {
    style.numFmt = numFmt;
  }
  return { style: Object.keys(style).length === 0 ? undefined : style, unread };
}

function hasProperty(view: DataView, length: number, wanted: number): boolean {
  for (let offset = 6; offset + 4 <= length;) {
    const size = view.getUint16(offset + 2, true);
    if (view.getUint16(offset, true) === wanted) {
      return true;
    }
    if (size < 4) {
      return false;
    }
    offset += size;
  }
  return false;
}

/** An `XFPropColor`, in the model's shape. Automatic and "not set" are both the absence of a colour. */
function readColor(view: DataView, offset: number): Partial<Color> {
  const kind = view.getUint8(offset) >> 1;
  const icv = view.getUint8(offset + 1);
  const tint = view.getInt16(offset + 2, true);
  const color: Partial<Color> = {};
  if (kind === COLOR_TYPE.theme) {
    color.theme = icv;
  } else if (kind === COLOR_TYPE.rgb) {
    const hex = (value: number): string => value.toString(16).padStart(2, "0").toUpperCase();
    color.argb =
      hex(view.getUint8(offset + 7)) +
      hex(view.getUint8(offset + 4)) +
      hex(view.getUint8(offset + 5)) +
      hex(view.getUint8(offset + 6));
  } else if (kind === COLOR_TYPE.indexed) {
    color.indexed = icv;
  }
  if (tint !== 0) {
    color.tint = tint / 32767;
  }
  return color;
}

/** An `LPWideString`: a two-byte character count, then UTF-16 code units. */
/**
 * An `LPWideString`: a two-byte character count, then UTF-16 code units — bounded by `end`, the end of the property
 * it sits in, not of the record. `undefined` when the count claims more characters than the property holds: reading
 * on took the name's tail from the next property.
 */
function readLpWideString(view: DataView, offset: number, end: number): string | undefined {
  const length = view.getUint16(offset, true);
  if (offset + 2 + length * 2 > end) {
    return undefined;
  }
  let text = "";
  for (let index = 0; index < length; index++) {
    text += String.fromCharCode(view.getUint16(offset + 2 + index * 2, true));
  }
  return text;
}
