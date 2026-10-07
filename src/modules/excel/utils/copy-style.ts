/**
 * A style-like bag copied generically by key. The `any` index signature is
 * intentional: this helper performs dynamic, key-driven deep copies
 * (`{ ...obj[key] }`) over arbitrary nested style objects, which `unknown`
 * cannot express (it forbids spreading the indexed value).
 */
interface StyleObject {
  [key: string]: any;
}

/** A border's edge keys, each carrying its own optional `color`. */
const BORDER_EDGES = ["top", "left", "bottom", "right", "diagonal"] as const;

/** The facets of a style that own mutable sub-objects and therefore need copying. */
const COPIED_FACETS = ["font", "alignment", "protection", "border", "fill"] as const;

/**
 * Copy `obj` and each of its `nestKeys` sub-objects one level down.
 *
 * Written as one spread and direct assignments: this runs for every styled cell a streamed sheet yields, and the
 * `reduce` into a second object that was then spread again built three objects per call to say what one does. The
 * result is the same — every nested key already exists on the copy, so assigning it keeps its position.
 */
function oneDepthCopy(obj: StyleObject, nestKeys: readonly string[]): StyleObject {
  const copied: StyleObject = { ...obj };
  for (let i = 0; i < nestKeys.length; i++) {
    const key = nestKeys[i];
    if (obj[key]) {
      copied[key] = { ...obj[key] };
    }
  }
  return copied;
}

const COLOR_KEYS = ["color"] as const;
const FILL_KEYS = ["fgColor", "bgColor", "center"] as const;

/** Whether `obj` has no own enumerable key — without allocating the key array `Object.keys` would. */
function isEmptyObj(obj: StyleObject): boolean {
  for (const key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      return false;
    }
  }
  return true;
}

/**
 * Copy one style facet (`font`, `border`, `fill`, `alignment`, `protection`) deeply
 * enough that the copy shares no mutable sub-object with the original — which is
 * what the "style isolation" invariant needs: mutating `Cell.getStyle(…).border.top`
 * must not reach a sibling cell.
 *
 * **Deep enough, not arbitrarily deep.** Every facet is a fixed OOXML shape, so the
 * nesting is known — a font's `color`, a border's five edges and each edge's `color`,
 * a fill's colours and gradient stops. `structuredClone` was used for this and is a
 * host serialize/deserialize round trip: measured on the four facets of one style,
 * 200k times, it costs 920ms against 47ms here — **19.5×** — and the propagation
 * setters run it once per cell per facet, so it dominated the cost of styling a
 * large sheet.
 *
 * The narrowing is that a value nested deeper than the format allows (a `Date`, a
 * cycle, a caller's own object hung off a facet) is now shared rather than copied.
 * That is already the contract on this module's other path: `copyStyle` — used by
 * `duplicateRow`/`spliceRows` and covered by `style-deep-copy.test.ts` — has always
 * copied styles to exactly this depth. Keeping one definition of "how deep is a
 * style" is the point; two would drift.
 */
function copyStyleFacet<T>(key: string, value: T): T {
  if (!value || typeof value !== "object") {
    return value;
  }
  const facet = value as unknown as StyleObject;
  switch (key) {
    case "font":
      return oneDepthCopy(facet, COLOR_KEYS) as unknown as T;
    case "alignment":
    case "protection":
      return { ...facet } as unknown as T;
    case "border":
      return copyBorder(facet) as unknown as T;
    case "fill":
      return copyFill(facet) as unknown as T;
    default:
      // numFmt / styleName are primitives; anything else is not a known facet.
      return value;
  }
}

function copyBorder(border: StyleObject): StyleObject {
  const copied: StyleObject = { ...border };
  for (const edge of BORDER_EDGES) {
    if (border[edge]) {
      copied[edge] = oneDepthCopy(border[edge], COLOR_KEYS);
    }
  }
  return copied;
}

function copyFill(fill: StyleObject): StyleObject {
  const copied = oneDepthCopy(fill, FILL_KEYS);
  if (fill.stops) {
    copied.stops = fill.stops.map((s: StyleObject) => oneDepthCopy(s, COLOR_KEYS));
  }
  return copied;
}

function copyStyle(style: StyleObject | null | undefined): StyleObject | null | undefined {
  if (!style) {
    return style;
  }
  if (isEmptyObj(style)) {
    return {};
  }

  const copied: StyleObject = { ...style };

  for (const key of COPIED_FACETS) {
    if (style[key]) {
      copied[key] = copyStyleFacet(key, style[key]);
    }
  }

  return copied;
}

export { copyStyle, copyStyleFacet };
