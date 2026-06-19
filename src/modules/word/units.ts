/**
 * DOCX Module — Unit Conversion Utilities.
 *
 * These conversions are shared across all modules and now live in the
 * dependency-free Layer-0 module `@utils/units`. This file re-exports them so
 * existing `@word/units` imports (and the public `Units` surface) keep working
 * unchanged.
 */

export * from "@utils/units";
