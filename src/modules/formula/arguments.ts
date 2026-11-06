/**
 * The {@link SchemaCheckOptions.error} of the `Formula` entry: argument schemas from `@utils/schema`
 * throw a {@link FormulaArgumentError} through it.
 */

import { FormulaArgumentError } from "@formula/errors";
import type { SchemaCheckOptions } from "@utils/schema";

function argumentError(field: string, message: string): FormulaArgumentError {
  return new FormulaArgumentError(field, message);
}

/** Where a parameter of a `Formula` call sits, for `check`. */
export function formulaArgument(where: string, field: string): SchemaCheckOptions {
  return { where, field, error: argumentError };
}
