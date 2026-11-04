/**
 * Style values checked where they are set.
 *
 * The serialisers refuse a value that has no XML form (`NaN`, an object where a number goes), but a value of the
 * right shape outside the schema's vocabulary — a border style `"wavy"`, a pattern `"dots!"`, a font size of `-3`, a
 * colour `"red"` — was written as given and Excel repaired the file. Checking the vocabulary here makes the call that
 * passed the value the one that fails, naming the field.
 */
import { ExcelArgumentError } from "@excel/errors";
import type { Alignment, BorderStyle, FillPatterns, Font, Style } from "@excel/types";
import { describeValue as describe } from "@utils/describe-value";

const BORDER_STYLES: Record<BorderStyle, true> = {
  thin: true,
  dotted: true,
  hair: true,
  medium: true,
  double: true,
  thick: true,
  dashed: true,
  dashDot: true,
  dashDotDot: true,
  slantDashDot: true,
  mediumDashed: true,
  mediumDashDotDot: true,
  mediumDashDot: true
};

const FILL_PATTERNS: Record<FillPatterns, true> = {
  none: true,
  solid: true,
  darkVertical: true,
  darkHorizontal: true,
  darkGrid: true,
  darkTrellis: true,
  darkDown: true,
  darkUp: true,
  lightVertical: true,
  lightHorizontal: true,
  lightGrid: true,
  lightTrellis: true,
  lightDown: true,
  lightUp: true,
  darkGray: true,
  mediumGray: true,
  lightGray: true,
  gray125: true,
  gray0625: true
};

const HORIZONTAL: Record<Alignment["horizontal"], true> = {
  general: true,
  left: true,
  center: true,
  right: true,
  fill: true,
  justify: true,
  centerContinuous: true,
  distributed: true
};
const VERTICAL: Record<Alignment["vertical"], true> = {
  top: true,
  middle: true,
  bottom: true,
  distributed: true,
  justify: true
};
const READING_ORDER: Record<Alignment["readingOrder"], true> = {
  rtl: true,
  ltr: true,
  context: true
};
const UNDERLINE: Record<Exclude<Font["underline"], boolean>, true> = {
  none: true,
  single: true,
  double: true,
  singleAccounting: true,
  doubleAccounting: true
};
const VERT_ALIGN: Record<Font["vertAlign"], true> = {
  superscript: true,
  subscript: true,
  baseline: true
};
const SCHEME: Record<Font["scheme"], true> = { minor: true, major: true, none: true };

const ARGB = /^(?:[0-9A-Fa-f]{2})?[0-9A-Fa-f]{6}$/;

class Checker {
  constructor(readonly where: string) {}

  fail(field: string, expected: string, value: unknown): never {
    throw new ExcelArgumentError(
      field,
      `${this.where}: ${field} must be ${expected}, got ${describe(value)}.`
    );
  }

  optional(value: unknown): boolean {
    return value === undefined;
  }

  bool(field: string, value: unknown): void {
    if (!this.optional(value) && typeof value !== "boolean") {
      this.fail(field, "a boolean", value);
    }
  }

  num(field: string, value: unknown, min?: number, max?: number, integer = false): void {
    if (this.optional(value)) {
      return;
    }
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      (integer && !Number.isInteger(value)) ||
      (min !== undefined && value < min) ||
      (max !== undefined && value > max)
    ) {
      const range =
        min !== undefined && max !== undefined
          ? ` from ${min} to ${max}`
          : min !== undefined
            ? ` of at least ${min}`
            : "";
      this.fail(field, `${integer ? "a whole number" : "a number"}${range}`, value);
    }
  }

  str(field: string, value: unknown): void {
    if (!this.optional(value) && typeof value !== "string") {
      this.fail(field, "a string", value);
    }
  }

  oneOf(field: string, value: unknown, allowed: Record<string, true>): void {
    if (!this.optional(value) && (typeof value !== "string" || !Object.hasOwn(allowed, value))) {
      this.fail(field, `one of ${Object.keys(allowed).join(", ")}`, value);
    }
  }

  object(field: string, value: unknown): value is Record<string, unknown> {
    if (this.optional(value) || value === null) {
      return false;
    }
    if (typeof value !== "object" || Array.isArray(value)) {
      this.fail(field, "an object", value);
    }
    return true;
  }

  color(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    if (value.argb !== undefined && (typeof value.argb !== "string" || !ARGB.test(value.argb))) {
      this.fail(
        `${field}.argb`,
        'six or eight hex digits (alpha first), such as "FFFF0000"',
        value.argb
      );
    }
    this.num(`${field}.theme`, value.theme, 0, undefined, true);
    this.num(`${field}.tint`, value.tint, -1, 1);
    this.num(`${field}.indexed`, value.indexed, 0, undefined, true);
  }

  font(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    this.str(`${field}.name`, value.name);
    this.num(`${field}.size`, value.size, 1, 409);
    this.num(`${field}.family`, value.family, 0, 14, true);
    this.oneOf(`${field}.scheme`, value.scheme, SCHEME);
    this.num(`${field}.charset`, value.charset, 0, 255, true);
    this.color(`${field}.color`, value.color);
    for (const flag of ["bold", "italic", "strike", "outline", "condense", "extend", "shadow"]) {
      this.bool(`${field}.${flag}`, value[flag]);
    }
    if (typeof value.underline !== "boolean") {
      this.oneOf(`${field}.underline`, value.underline, UNDERLINE);
    }
    this.oneOf(`${field}.vertAlign`, value.vertAlign, VERT_ALIGN);
  }

  alignment(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    this.oneOf(`${field}.horizontal`, value.horizontal, HORIZONTAL);
    this.oneOf(`${field}.vertical`, value.vertical, VERTICAL);
    this.oneOf(`${field}.readingOrder`, value.readingOrder, READING_ORDER);
    this.bool(`${field}.wrapText`, value.wrapText);
    this.bool(`${field}.shrinkToFit`, value.shrinkToFit);
    this.bool(`${field}.justifyLastLine`, value.justifyLastLine);
    this.num(`${field}.indent`, value.indent, 0, 250, true);
    this.num(`${field}.relativeIndent`, value.relativeIndent, -15, 15, true);
    const rotation = value.textRotation;
    if (
      rotation !== undefined &&
      rotation !== "vertical" &&
      !(
        typeof rotation === "number" &&
        Number.isInteger(rotation) &&
        ((rotation >= -90 && rotation <= 180) || rotation === 255)
      )
    ) {
      this.fail(
        `${field}.textRotation`,
        '"vertical" or a whole number of degrees from -90 to 180',
        rotation
      );
    }
  }

  protection(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    this.bool(`${field}.locked`, value.locked);
    this.bool(`${field}.hidden`, value.hidden);
  }

  border(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    for (const edge of ["top", "left", "bottom", "right", "diagonal", "vertical", "horizontal"]) {
      const line = value[edge];
      if (!this.object(`${field}.${edge}`, line)) {
        continue;
      }
      this.oneOf(`${field}.${edge}.style`, line.style, BORDER_STYLES);
      this.color(`${field}.${edge}.color`, line.color);
      if (edge === "diagonal") {
        this.bool(`${field}.diagonal.up`, line.up);
        this.bool(`${field}.diagonal.down`, line.down);
      }
    }
  }

  fill(field: string, value: unknown): void {
    if (!this.object(field, value)) {
      return;
    }
    if (value.type === "pattern") {
      this.oneOf(`${field}.pattern`, value.pattern, FILL_PATTERNS);
      this.color(`${field}.fgColor`, value.fgColor);
      this.color(`${field}.bgColor`, value.bgColor);
      return;
    }
    if (value.type === "gradient") {
      if (value.gradient !== "angle" && value.gradient !== "path") {
        this.fail(`${field}.gradient`, '"angle" or "path"', value.gradient);
      }
      this.num(`${field}.degree`, value.degree);
      if (value.stops !== undefined) {
        if (!Array.isArray(value.stops)) {
          this.fail(`${field}.stops`, "an array", value.stops);
        }
        value.stops.forEach((stop: unknown, i: number) => {
          if (this.object(`${field}.stops[${i}]`, stop)) {
            this.num(`${field}.stops[${i}].position`, stop.position, 0, 1);
            this.color(`${field}.stops[${i}].color`, stop.color);
          }
        });
      }
      return;
    }
    this.fail(`${field}.type`, '"pattern" or "gradient"', value.type);
  }

  numFmt(field: string, value: unknown): void {
    if (this.optional(value) || typeof value === "string") {
      return;
    }
    // Named for what a number format is — a format code — rather than "an object", which a `42` was told it must be.
    if (typeof value !== "object" || value === null || Array.isArray(value)) {
      this.fail(field, "a format code string or { id, formatCode }", value);
    }
    if (this.object(field, value)) {
      this.num(`${field}.id`, value.id, 0, undefined, true);
      this.str(`${field}.formatCode`, value.formatCode);
      return;
    }
    this.fail(field, "a format code string", value);
  }
}

type Facet = "font" | "alignment" | "protection" | "border" | "fill" | "numFmt";

/** Check one style facet as {@link checkStyle} checks it inside a style. `undefined` clears and is accepted. */
/**
 * Check one facet of a style. `field` is the argument's path from the parameter name (`value` for `Cell.setFont(ws,
 * addr, value)`); it defaults to the facet's own name, as it reads inside a style.
 */
export function checkStyleFacet(
  where: string,
  facet: Facet,
  value: unknown,
  field: string = facet
): void {
  new Checker(where)[facet](field, value);
}

/** Check a colour (`{ argb }`, `{ theme, tint }` or `{ indexed }`) as a style's colours are checked. `undefined` is accepted. */
export function checkColor(where: string, field: string, value: unknown): void {
  new Checker(where).color(field, value);
}

/**
 * Check every facet a (partial) style carries. `field` is the style's path from the parameter name (`style`,
 * `definitions[0].style`); each facet is named beneath it (`style.font.size`).
 */
export function checkStyle(where: string, style: unknown, field = "style"): void {
  const checker = new Checker(where);
  if (!checker.object(field, style)) {
    return;
  }
  const s = style as Partial<Style> & Record<string, unknown>;
  checker.font(`${field}.font`, s.font);
  checker.alignment(`${field}.alignment`, s.alignment);
  checker.protection(`${field}.protection`, s.protection);
  checker.border(`${field}.border`, s.border);
  checker.fill(`${field}.fill`, s.fill);
  checker.numFmt(`${field}.numFmt`, s.numFmt);
  checker.str(`${field}.styleName`, s.styleName);
}
