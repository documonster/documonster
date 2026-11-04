/**
 * `Matrix` namespace surface — the 2-D affine transforms a display list's
 * `group.transform` carries.
 *
 * `import { Matrix } from "documonster/draw"` →
 *   `Matrix.IDENTITY`, `Matrix.translate(10, 0)`, `Matrix.multiply(a, b)`,
 *   `Matrix.apply(m, x, y)`, `Matrix.getRotation(m)`, …
 */

import { checkFinite, checkMatrix } from "@draw/arguments";
import {
  apply as applyImpl,
  multiply as multiplyImpl,
  rotate as rotateImpl,
  rotationOf,
  scale as scaleImpl,
  translate as translateImpl,
  uniformScale
} from "@draw/types";
import type { DrawMatrix, DrawPoint } from "@draw/types";

export { IDENTITY } from "@draw/types";

/** Apply a transform to a point. */
export function apply(m: DrawMatrix, x: number, y: number): DrawPoint {
  checkMatrix("Matrix.apply", "m", m);
  checkFinite("Matrix.apply", "x", x);
  checkFinite("Matrix.apply", "y", y);
  return applyImpl(m, x, y);
}

/** Compose two transforms: apply `inner` first, then `outer`. */
export function multiply(outer: DrawMatrix, inner: DrawMatrix): DrawMatrix {
  checkMatrix("Matrix.multiply", "outer", outer);
  checkMatrix("Matrix.multiply", "inner", inner);
  return multiplyImpl(outer, inner);
}

/** A clockwise rotation about `(cx, cy)` (default the origin), as SVG's `rotate(a cx cy)`. */
export function rotate(degrees: number, cx?: number, cy?: number): DrawMatrix {
  checkFinite("Matrix.rotate", "degrees", degrees);
  if (cx !== undefined) {
    checkFinite("Matrix.rotate", "cx", cx);
  }
  if (cy !== undefined) {
    checkFinite("Matrix.rotate", "cy", cy);
  }
  return rotateImpl(degrees, cx, cy);
}

/** A scale about the origin; `sy` defaults to `sx`. */
export function scale(sx: number, sy?: number): DrawMatrix {
  checkFinite("Matrix.scale", "sx", sx);
  if (sy !== undefined) {
    checkFinite("Matrix.scale", "sy", sy);
  }
  return scaleImpl(sx, sy);
}

/** A translation. */
export function translate(tx: number, ty: number): DrawMatrix {
  checkFinite("Matrix.translate", "tx", tx);
  checkFinite("Matrix.translate", "ty", ty);
  return translateImpl(tx, ty);
}

/**
 * The rotation a transform applies, in degrees clockwise. A sheared transform has no single
 * answer and reports the angle of its x-axis.
 */
export function getRotation(m: DrawMatrix): number {
  checkMatrix("Matrix.getRotation", "m", m);
  return rotationOf(m);
}

/**
 * The uniform-equivalent scale factor of a transform, `sqrt(|det|)` — what a backend uses
 * for a stroke width or font size. A degenerate (zero-determinant) transform reports 1.
 */
export function getUniformScale(m: DrawMatrix): number {
  checkMatrix("Matrix.getUniformScale", "m", m);
  return uniformScale(m);
}
