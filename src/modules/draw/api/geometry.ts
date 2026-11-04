/**
 * `Draw.Geometry` — low-level path construction, for a producer building display-list
 * nodes by hand: elliptical arcs and sectors as cubic Béziers, rounded rectangles, a
 * rect node, and flattening a path to polylines.
 *
 * `Draw.toSvg`, `Draw.toRgba` and `Draw.render` need none of this; it is what they use
 * internally, published for a producer that wants the same shapes.
 */

import {
  checkBox,
  checkFinite,
  checkNonNegative,
  checkOptionalNonNegative,
  checkPaint,
  checkPathCommands,
  checkPoint,
  checkPositive,
  fail
} from "@draw/arguments";
import {
  arcToCubics as arcToCubicsImpl,
  flattenPath as flattenPathImpl,
  rectNode,
  roundedRectToPath as roundedRectToPathImpl,
  sectorToPath as sectorToPathImpl
} from "@draw/types";
import type {
  DrawBox,
  DrawNode,
  DrawPaint,
  DrawPathCommand,
  DrawPoint,
  DrawSubpath
} from "@draw/types";

function checkBoolean(where: string, field: string, value: unknown): void {
  if (typeof value !== "boolean") {
    fail(where, field, "a boolean", value);
  }
}

/**
 * Convert an SVG-style elliptical arc (endpoint parameterisation) to cubic Béziers. A zero
 * radius, or coincident end points, degrade to a line, as SVG specifies.
 */
export function arcToCubics(
  from: DrawPoint,
  rx: number,
  ry: number,
  rotationDegrees: number,
  largeArc: boolean,
  sweep: boolean,
  to: DrawPoint
): DrawPathCommand[] {
  checkPoint("Draw.Geometry.arcToCubics", "from", from);
  checkFinite("Draw.Geometry.arcToCubics", "rx", rx);
  checkFinite("Draw.Geometry.arcToCubics", "ry", ry);
  checkFinite("Draw.Geometry.arcToCubics", "rotationDegrees", rotationDegrees);
  checkBoolean("Draw.Geometry.arcToCubics", "largeArc", largeArc);
  checkBoolean("Draw.Geometry.arcToCubics", "sweep", sweep);
  checkPoint("Draw.Geometry.arcToCubics", "to", to);
  return arcToCubicsImpl(from, rx, ry, rotationDegrees, largeArc, sweep, to);
}

/**
 * Flatten a path to polylines, one per subpath, keeping whether each closed. `scale`
 * multiplies the output points (default 1), so a supersampling backend gets finer
 * subdivision.
 */
export function flattenPath(commands: readonly DrawPathCommand[], scale?: number): DrawSubpath[] {
  checkPathCommands("Draw.Geometry.flattenPath", "commands", commands);
  if (scale !== undefined) {
    checkPositive("Draw.Geometry.flattenPath", "scale", scale);
  }
  return flattenPathImpl(commands, scale);
}

/** A rect node from a box, with an optional corner radius. */
export function createRectNode(box: DrawBox, paint: DrawPaint, rx?: number): DrawNode {
  checkBox("Draw.Geometry.createRectNode", "box", box);
  checkPaint("Draw.Geometry.createRectNode", "paint", paint);
  checkOptionalNonNegative("Draw.Geometry.createRectNode", "rx", rx);
  return rectNode(box, paint, rx);
}

/**
 * A rounded rectangle as a path. `radius` is clamped to `0 … half the shorter side`, as SVG
 * does, so a negative radius draws square corners.
 */
export function roundedRectToPath(
  x: number,
  y: number,
  width: number,
  height: number,
  radius: number
): DrawPathCommand[] {
  checkBox("Draw.Geometry.roundedRectToPath", "box", { x, y, width, height });
  checkFinite("Draw.Geometry.roundedRectToPath", "radius", radius);
  return roundedRectToPathImpl(x, y, width, height, radius);
}

/**
 * A circular sector — or an annular one when `innerRadius > 0` — as a path. Angles are
 * radians, clockwise from the positive x-axis in the Y-down space.
 */
export function sectorToPath(
  cx: number,
  cy: number,
  radius: number,
  innerRadius: number,
  startAngle: number,
  endAngle: number
): DrawPathCommand[] {
  checkFinite("Draw.Geometry.sectorToPath", "cx", cx);
  checkFinite("Draw.Geometry.sectorToPath", "cy", cy);
  checkNonNegative("Draw.Geometry.sectorToPath", "radius", radius);
  checkNonNegative("Draw.Geometry.sectorToPath", "innerRadius", innerRadius);
  checkFinite("Draw.Geometry.sectorToPath", "startAngle", startAngle);
  checkFinite("Draw.Geometry.sectorToPath", "endAngle", endAngle);
  return sectorToPathImpl(cx, cy, radius, innerRadius, startAngle, endAngle);
}
