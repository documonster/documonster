/**
 * The errors of `documonster/utils`.
 */

import { BaseError } from "../errors";
import type { BaseErrorOptions } from "../errors";

/** Base class of every error `documonster/utils` throws. */
export class UtilsError extends BaseError {
  override name = "UtilsError";
}

/**
 * Thrown when a `documonster/utils` function is given an argument or option it cannot use — a
 * value of the wrong type, outside its range, or text that is not in the format it names — before
 * anything is produced.
 */
export class UtilsArgumentError extends UtilsError {
  override name = "UtilsArgumentError";
  /** The argument or option at fault, e.g. `text` or `options.alphabet`. */
  readonly field: string;

  constructor(field: string, message: string, options?: BaseErrorOptions) {
    super(message, options);
    this.field = field;
  }
}

/** Whether `err` is an error `documonster/utils` threw. */
export function isUtilsError(err: unknown): err is UtilsError {
  return err instanceof UtilsError;
}
