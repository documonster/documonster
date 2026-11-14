/**
 * Append every item of `items` to `target`, in order.
 *
 * Not `target.push(...items)`: that passes each item as an argument, and an array whose length a file decides — the
 * paragraphs of one content control, the blocks inside one `w:customXml`, the cells of a sheet a read reports on —
 * overflowed the call stack past about a hundred thousand, failing the whole read with
 * `Maximum call stack size exceeded`.
 */
export function appendAll<T>(target: T[], items: readonly T[]): void {
  // The length is read once: appending an array to itself would otherwise never end.
  const count = items.length;
  for (let i = 0; i < count; i++) {
    target.push(items[i]!);
  }
}
