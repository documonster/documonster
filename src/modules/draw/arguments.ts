/**
 * Argument checks for the public drawing entry points.
 *
 * The walker and the surfaces trust what they are given: a `NaN` coordinate reached SVG
 * as `0`, a CSS string where a colour struct belongs became `fill="#NaNNaNNaN"`, and a
 * `null` child crashed with a `TypeError` half-way through a render. Each check here
 * throws {@link DrawArgumentError} naming the call and the field, before anything is drawn.
 *
 * Only the public entries run these (`Draw.*`, `Matrix.*`). The producers inside the
 * library build their lists from checked inputs, and re-walking every chart's list on
 * every render would cost them for nothing.
 */

import { DrawArgumentError } from "@draw/errors";
import { checkBackgroundWith } from "@utils/background-check";
import { describeValue as describe } from "@utils/describe-value";

export function fail(where: string, field: string, expected: string, value: unknown): never {
  throw new DrawArgumentError(
    field,
    `${where}: ${field} must be ${expected}, got ${describe(value)}`
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function checkFinite(where: string, field: string, value: unknown): void {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    fail(where, field, "a finite number", value);
  }
}

export function checkOptionalFinite(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkFinite(where, field, value);
  }
}

export function checkNonNegative(where: string, field: string, value: unknown): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    fail(where, field, "a finite number >= 0", value);
  }
}

export function checkOptionalNonNegative(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkNonNegative(where, field, value);
  }
}

export function checkPositive(where: string, field: string, value: unknown): void {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    fail(where, field, "a finite number > 0", value);
  }
}

export function checkOptionalPositive(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkPositive(where, field, value);
  }
}

export function checkString(where: string, field: string, value: unknown): void {
  if (typeof value !== "string") {
    fail(where, field, "a string", value);
  }
}

export function checkOptionalString(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkString(where, field, value);
  }
}

export function checkOptionalBoolean(where: string, field: string, value: unknown): void {
  if (value !== undefined && typeof value !== "boolean") {
    fail(where, field, "a boolean", value);
  }
}

export function checkObject(where: string, field: string, value: unknown): void {
  if (!isRecord(value)) {
    fail(where, field, "an object", value);
  }
}

export function checkOptionalObject(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkObject(where, field, value);
  }
}

function checkOptionalEnum(
  where: string,
  field: string,
  value: unknown,
  allowed: readonly string[]
): void {
  if (value !== undefined && (typeof value !== "string" || !allowed.includes(value))) {
    fail(where, field, `one of ${allowed.map(v => JSON.stringify(v)).join(", ")}`, value);
  }
}

function checkArray(where: string, field: string, value: unknown): asserts value is unknown[] {
  if (!Array.isArray(value)) {
    fail(where, field, "an array", value);
  }
}

/** An `Rgba01`: `r`, `g`, `b` and `a`, each a number from 0 to 1. */
export function checkColor(where: string, field: string, value: unknown): void {
  if (!isRecord(value)) {
    fail(where, field, "a color { r, g, b, a } with channels from 0 to 1", value);
  }
  for (const channel of ["r", "g", "b", "a"] as const) {
    const component = value[channel];
    if (typeof component !== "number" || !(component >= 0 && component <= 1)) {
      fail(where, `${field}.${channel}`, "a number from 0 to 1", component);
    }
  }
}

/** {@link checkBackgroundWith}, refused with `DrawArgumentError`. */
export function checkBackground(where: string, field: string, value: unknown): void {
  checkBackgroundWith(where, field, value, fail);
}

function checkOptionalColor(where: string, field: string, value: unknown): void {
  if (value !== undefined) {
    checkColor(where, field, value);
  }
}

/** A `DrawMatrix`: six finite numbers. */
export function checkMatrix(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const m = value as Record<string, unknown>;
  for (const key of ["a", "b", "c", "d", "e", "f"]) {
    checkFinite(where, `${field}.${key}`, m[key]);
  }
}

/** A `DrawPoint`. */
export function checkPoint(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const p = value as Record<string, unknown>;
  checkFinite(where, `${field}.x`, p.x);
  checkFinite(where, `${field}.y`, p.y);
}

const FILL_RULES = ["nonzero", "evenodd"];
const LINE_JOINS = ["miter", "round", "bevel"];
const LINE_CAPS = ["butt", "round", "square"];
const TEXT_ANCHORS = ["start", "middle", "end"];

/** A `DrawPaint`. */
export function checkPaint(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const paint = value as Record<string, unknown>;
  checkOptionalColor(where, `${field}.fill`, paint.fill);
  checkOptionalColor(where, `${field}.stroke`, paint.stroke);
  checkOptionalNonNegative(where, `${field}.strokeWidth`, paint.strokeWidth);
  if (paint.dash !== undefined) {
    checkArray(where, `${field}.dash`, paint.dash);
    paint.dash.forEach((length, i) => checkNonNegative(where, `${field}.dash[${i}]`, length));
  }
  checkOptionalEnum(where, `${field}.fillRule`, paint.fillRule, FILL_RULES);
  checkOptionalEnum(where, `${field}.lineJoin`, paint.lineJoin, LINE_JOINS);
  checkOptionalEnum(where, `${field}.lineCap`, paint.lineCap, LINE_CAPS);
}

/** A `DrawTextStyle`. */
export function checkTextStyle(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const style = value as Record<string, unknown>;
  checkPositive(where, `${field}.size`, style.size);
  checkOptionalString(where, `${field}.family`, style.family);
  checkOptionalBoolean(where, `${field}.bold`, style.bold);
  checkOptionalBoolean(where, `${field}.italic`, style.italic);
  checkOptionalEnum(where, `${field}.anchor`, style.anchor, TEXT_ANCHORS);
  checkOptionalColor(where, `${field}.fill`, style.fill);
}

/** A `RasterTextStyle`. */
export function checkRasterTextStyle(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const style = value as Record<string, unknown>;
  checkOptionalString(where, `${field}.family`, style.family);
  checkOptionalBoolean(where, `${field}.bold`, style.bold);
  checkOptionalBoolean(where, `${field}.italic`, style.italic);
}

/** A list of `DrawPathCommand`s. */
export function checkPathCommands(where: string, field: string, value: unknown): void {
  checkArray(where, field, value);
  value.forEach((command, i) => {
    const at = `${field}[${i}]`;
    checkObject(where, at, command);
    const c = command as Record<string, unknown>;
    switch (c.op) {
      case "move":
      case "line":
        checkFinite(where, `${at}.x`, c.x);
        checkFinite(where, `${at}.y`, c.y);
        return;
      case "cubic":
        for (const key of ["x1", "y1", "x2", "y2", "x", "y"]) {
          checkFinite(where, `${at}.${key}`, c[key]);
        }
        return;
      case "close":
        return;
      default:
        fail(where, `${at}.op`, 'one of "move", "line", "cubic", "close"', c.op);
    }
  });
}

/** A `DrawClip` / `DrawBox`: four finite numbers. */
export function checkBox(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const box = value as Record<string, unknown>;
  for (const key of ["x", "y", "width", "height"]) {
    checkFinite(where, `${field}.${key}`, box[key]);
  }
}

const NODE_KINDS = '"group", "rect", "ellipse", "line", "polyline", "path", "sector", "text"';

/** One `DrawNode` and everything beneath it. */
export function checkNode(where: string, field: string, value: unknown): void {
  checkObject(where, field, value);
  const node = value as Record<string, unknown>;
  const finite = (...keys: string[]): void => {
    for (const key of keys) {
      checkFinite(where, `${field}.${key}`, node[key]);
    }
  };
  switch (node.kind) {
    case "group":
      if (node.transform !== undefined) {
        checkMatrix(where, `${field}.transform`, node.transform);
      }
      if (node.clip !== undefined) {
        checkBox(where, `${field}.clip`, node.clip);
      }
      checkOptionalString(where, `${field}.svgFilterId`, node.svgFilterId);
      checkArray(where, `${field}.children`, node.children);
      node.children.forEach((child, i) => checkNode(where, `${field}.children[${i}]`, child));
      return;
    case "rect":
      finite("x", "y", "width", "height");
      checkOptionalNonNegative(where, `${field}.rx`, node.rx);
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "ellipse":
      finite("cx", "cy", "rx", "ry");
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "line":
      finite("x1", "y1", "x2", "y2");
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "polyline":
      checkArray(where, `${field}.points`, node.points);
      node.points.forEach((point, i) => checkPoint(where, `${field}.points[${i}]`, point));
      checkOptionalBoolean(where, `${field}.closed`, node.closed);
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "path":
      checkPathCommands(where, `${field}.commands`, node.commands);
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "sector":
      finite("cx", "cy", "startAngle", "endAngle");
      checkNonNegative(where, `${field}.radius`, node.radius);
      checkNonNegative(where, `${field}.innerRadius`, node.innerRadius);
      checkPaint(where, `${field}.paint`, node.paint);
      return;
    case "text":
      finite("x", "y");
      checkArray(where, `${field}.lines`, node.lines);
      node.lines.forEach((line, i) => {
        const at = `${field}.lines[${i}]`;
        checkObject(where, at, line);
        const l = line as Record<string, unknown>;
        checkString(where, `${at}.text`, l.text);
        checkFinite(where, `${at}.dy`, l.dy);
        checkOptionalFinite(where, `${at}.x`, l.x);
      });
      checkTextStyle(where, `${field}.style`, node.style);
      checkOptionalFinite(where, `${field}.rotate`, node.rotate);
      return;
    default:
      fail(where, `${field}.kind`, `one of ${NODE_KINDS}`, node.kind);
  }
}

/** A whole `DrawList`. */
export function checkDrawList(where: string, value: unknown): void {
  checkObject(where, "list", value);
  const list = value as Record<string, unknown>;
  checkNonNegative(where, "list.width", list.width);
  checkNonNegative(where, "list.height", list.height);
  checkArray(where, "list.children", list.children);
  list.children.forEach((child, i) => checkNode(where, `list.children[${i}]`, child));
  if (list.svgDefs !== undefined) {
    checkArray(where, "list.svgDefs", list.svgDefs);
    list.svgDefs.forEach((def, i) => checkString(where, `list.svgDefs[${i}]`, def));
  }
}

/** A `DrawSurface`: the six required drawing methods, and functions where optional ones are given. */
export function checkSurface(where: string, value: unknown): void {
  if (typeof value !== "object" || value === null) {
    fail(where, "surface", "a DrawSurface", value);
  }
  const surface = value as Record<string, unknown>;
  for (const method of ["rect", "ellipse", "sector", "polyline", "path", "text"]) {
    if (typeof surface[method] !== "function") {
      fail(where, `surface.${method}`, "a function", surface[method]);
    }
  }
  for (const method of ["pushClip", "popClip", "pushFilter", "popFilter"]) {
    if (surface[method] !== undefined && typeof surface[method] !== "function") {
      fail(where, `surface.${method}`, "a function", surface[method]);
    }
  }
}

/**
 * Throw unless `canvas` can be painted by the raster surface: a cheap structural
 * test (two of the methods it calls), so a stray object fails here, by name, rather than with
 * a `TypeError` on the first node drawn.
 */
export function checkRasterCanvas(where: string, canvas: unknown): void {
  if (
    typeof canvas !== "object" ||
    canvas === null ||
    typeof (canvas as { fillRect?: unknown }).fillRect !== "function" ||
    typeof (canvas as { pushClip?: unknown }).pushClip !== "function"
  ) {
    fail(where, "canvas", "a BasicRasterCanvas", canvas);
  }
}
