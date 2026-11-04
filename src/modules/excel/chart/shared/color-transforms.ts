/**
 * The ordered transform list of a chart colour — {@link ChartColor.transforms} — read from the markup, reconciled
 * with the six named fields, and written back.
 *
 * One implementation for every chart colour reader and writer (classic chart parts, ChartEx parts, the colours
 * sidecar, the raw-XML patcher) and for the renderer, so a colour cannot be read one way and drawn another.
 */
import type { ChartColor } from "@excel/chart/model/types";
import { ChartOptionsError } from "@excel/errors";
import type { ColorTransform, ColorTransformKind } from "@utils/theme-colors";

/** Every transform element of ECMA-376 §20.1.2.3, by the name the markup gives it. */
const KINDS = [
  "red",
  "redMod",
  "redOff",
  "green",
  "greenMod",
  "greenOff",
  "blue",
  "blueMod",
  "blueOff",
  "alpha",
  "alphaMod",
  "alphaOff",
  "hue",
  "hueMod",
  "hueOff",
  "sat",
  "satMod",
  "satOff",
  "lum",
  "lumMod",
  "lumOff",
  "shade",
  "tint",
  "gray",
  "comp",
  "inv",
  "gamma",
  "invGamma"
] as const satisfies readonly ColorTransformKind[];

/** A transform element: one of {@link KINDS}, as a whole name (`<a:red …>` is not `<a:redMod …>`). */
const TRANSFORM_RE = new RegExp(
  `<a:(${[...KINDS].sort((a, b) => b.length - a.length).join("|")})(?=[\\s/>])([^>]*)>`,
  "g"
);
const VAL_RE = /\bval\s*=\s*"([^"]*)"/;

/**
 * The transforms of a colour element, in document order — `region` is the element's markup, from its opening tag to
 * its closing one (transforms are its only children). A `val` is rounded to the integer the schema types it as; a
 * flag transform has none.
 */
export function readColorTransforms(region: string): ColorTransform[] {
  const list: ColorTransform[] = [];
  for (const match of region.matchAll(TRANSFORM_RE)) {
    const raw = VAL_RE.exec(match[2]!)?.[1];
    const val = raw === undefined ? Number.NaN : parseFloat(raw);
    list.push(
      Number.isFinite(val)
        ? { kind: match[1] as ColorTransformKind, val: Math.round(val) }
        : { kind: match[1] as ColorTransformKind }
    );
  }
  return list;
}

/** The six transforms a chart colour also names as fields, each as its raw OOXML integer. */
export interface NamedColorTransforms {
  readonly tint?: number;
  readonly shade?: number;
  readonly satMod?: number;
  readonly lumMod?: number;
  readonly lumOff?: number;
  readonly alpha?: number;
}

const NAMED = ["tint", "shade", "satMod", "lumMod", "lumOff", "alpha"] as const;

/**
 * `list` with the named fields applied, as {@link ChartColor.transforms} defines it: a field that is set (a finite
 * number) wins — it replaces the value of the first entry of its kind, or is appended when there is none; a field
 * that is absent leaves the list's entries of its kind as they are.
 */
export function mergeColorTransforms(
  list: readonly ColorTransform[],
  named: NamedColorTransforms
): ColorTransform[] {
  const merged = [...list];
  for (const kind of NAMED) {
    const value = named[kind];
    if (typeof value !== "number" || !Number.isFinite(value)) {
      continue;
    }
    const at = merged.findIndex(t => t.kind === kind);
    if (at === -1) {
      merged.push({ kind, val: Math.round(value) });
    } else {
      merged[at] = { kind, val: Math.round(value) };
    }
  }
  return merged;
}

/**
 * The transforms a {@link ChartColor} is written and drawn with, when it carries a list; `undefined` when it carries
 * only the named fields, which each writer then writes as it always has. `tint` is the model's fraction.
 */
export function chartColorTransformList(color: ChartColor): ColorTransform[] | undefined {
  if (color.transforms === undefined) {
    return undefined;
  }
  return mergeColorTransforms(color.transforms, {
    ...color,
    tint:
      typeof color.tint === "number" && Number.isFinite(color.tint)
        ? color.tint * 100000
        : undefined
  });
}

/**
 * {@link chartColorTransformList} for a colours-sidecar palette entry or variation, whose `tint` is the raw OOXML
 * integer.
 */
export function paletteTransformList(
  entry: NamedColorTransforms & { readonly transforms?: readonly ColorTransform[] }
): ColorTransform[] | undefined {
  return entry.transforms === undefined ? undefined : mergeColorTransforms(entry.transforms, entry);
}

/** The transforms that take no value. */
const FLAGS: ReadonlySet<string> = new Set(["gray", "comp", "inv", "gamma", "invGamma"]);
const KNOWN: ReadonlySet<string> = new Set(KINDS);

/**
 * `list` as the elements it is written as, in order — `[name, val]`, `val` `undefined` for a flag. A kind the markup
 * does not have, or a valued transform without a finite `val`, is refused (`ChartOptionsError` on
 * `transforms[i]`) rather than written as markup Excel refuses: the writers fail loudly on a malformed colour.
 */
export function colorTransformElements(
  list: readonly ColorTransform[]
): [name: string, val: number | undefined][] {
  return list.map((t, i) => {
    if (!KNOWN.has(t?.kind)) {
      throw new ChartOptionsError(
        `transforms[${i}].kind`,
        `transforms[${i}].kind must be a DrawingML colour transform (lumMod, satOff, comp, …), got ${JSON.stringify(t?.kind)}.`
      );
    }
    if (FLAGS.has(t.kind)) {
      return [t.kind, undefined];
    }
    if (typeof t.val !== "number" || !Number.isFinite(t.val)) {
      throw new ChartOptionsError(
        `transforms[${i}].val`,
        `transforms[${i}].val must be a finite number for ${t.kind}, got ${JSON.stringify(t.val)}.`
      );
    }
    return [t.kind, Math.round(t.val)];
  });
}

/** `list` as the markup writes it: `<a:lumMod val="75000"/>`, `<a:comp/>`, in order (see {@link colorTransformElements}). */
export function colorTransformsXml(list: readonly ColorTransform[]): string {
  return colorTransformElements(list)
    .map(([name, val]) => (val === undefined ? `<a:${name}/>` : `<a:${name} val="${val}"/>`))
    .join("");
}
