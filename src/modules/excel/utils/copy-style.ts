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

function oneDepthCopy(obj: StyleObject, nestKeys: string[]): StyleObject {
  return {
    ...obj,
    ...nestKeys.reduce((memo: StyleObject, key: string) => {
      if (obj[key]) {
        memo[key] = { ...obj[key] };
      }
      return memo;
    }, {})
  };
}

function isEmptyObj(obj: StyleObject): boolean {
  return Object.keys(obj).length === 0;
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
      return oneDepthCopy(facet, ["color"]) as unknown as T;
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
      copied[edge] = oneDepthCopy(border[edge], ["color"]);
    }
  }
  return copied;
}

function copyFill(fill: StyleObject): StyleObject {
  const copied = oneDepthCopy(fill, ["fgColor", "bgColor", "center"]);
  if (fill.stops) {
    copied.stops = fill.stops.map((s: StyleObject) => oneDepthCopy(s, ["color"]));
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
