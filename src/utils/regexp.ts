/**
 * `text` as a regular-expression source that matches it literally: every syntax character, the
 * backslash included, is escaped. `RegExp.escape` does this natively from Node 24, which is above
 * the supported floor.
 */
export function escapeRegExp(text: string): string {
  return text.replace(/[\\^$.*+?()[\]{}|/-]/g, "\\$&");
}
