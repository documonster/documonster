/**
 * `text` as a regular-expression source that matches it literally: every syntax character, the
 * backslash included, is escaped. `RegExp.escape` does this natively from Node 24, which is above
 * the supported floor.
 *
 * `-` is written as `\x2d`, as `RegExp.escape` writes it, so the result also matches literally
 * inside a character class and stays valid under the `u` and `v` flags, which refuse `\-` outside one.
 */
export function escapeRegExp(text: string): string {
  return text.replace(/[\\^$.*+?()[\]{}|/]/g, "\\$&").replace(/-/g, "\\x2d");
}
