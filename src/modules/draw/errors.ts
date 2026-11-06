/**
 * Drawing-engine error types.
 */

import { BaseError } from "@utils/errors";
import type { BaseErrorOptions } from "@utils/errors";

/**
 * Base class for all errors raised by `documonster/draw`: a display list that cannot be
 * drawn, a raster output over the size limit, a font that cannot be read.
 */
export class DrawError extends BaseError {
  override name = "DrawError";
}

/**
 * Error thrown when a drawing function is given a value with no valid reading — a `NaN`
 * coordinate, a colour that is not `{ r, g, b, a }` in 0–1, a node `kind` the engine does
 * not know, a raster scale of zero. Thrown by the call that received the value, before
 * anything is drawn, with a message that starts with the call as the caller wrote it
 * (`Draw.toRgba: …`, `Draw.Geometry.flattenPath: …`).
 */
export class DrawArgumentError extends DrawError {
  override name = "DrawArgumentError";
  /**
   * The argument or field at fault, as a path from the parameter name in the public
   * signature: `list.children[2].paint.fill.r`, `options.scale`.
   */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/** Check if an error is a {@link DrawError}. */
export function isDrawError(err: unknown): err is DrawError {
  return err instanceof DrawError;
}
