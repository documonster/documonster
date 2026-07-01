/**
 * Archive-specific binary utilities.
 *
 * For common utilities (concatUint8Arrays), import directly from @stream/shared.
 */

/**
 * A reusable empty Uint8Array instance.
 *
 * Use this constant instead of creating `new Uint8Array(0)` to avoid
 * allocating new objects for empty byte arrays.
 */
export const EMPTY_UINT8ARRAY = new Uint8Array(0);

/**
 * Find the first index of `pattern` within `buffer`.
 * Returns -1 when not found.
 *
 * This is optimized for small patterns (1-4 bytes) common in ZIP parsing.
 */
export function indexOfUint8ArrayPattern(
  buffer: Uint8Array,
  pattern: Uint8Array,
  startIndex = 0
): number {
  const bufLen = buffer.length;
  const patLen = pattern.length;
  if (patLen === 0) {
    return 0;
  }
  if (patLen > bufLen) {
    return -1;
  }

  let start = startIndex | 0;
  if (start < 0) {
    start = 0;
  }
  if (start > bufLen - patLen) {
    return -1;
  }

  // Fast paths for small patterns (very common in ZIP parsing: 2/3/4-byte signatures).
  if (patLen === 1) {
    return buffer.indexOf(pattern[0], start);
  }

  if (patLen === 2) {
    const p0 = pattern[0];
    const p1 = pattern[1];
    const last = bufLen - 2;
    let i = buffer.indexOf(p0, start);
    while (i !== -1 && i <= last) {
      if (buffer[i + 1] === p1) {
        return i;
      }
      i = buffer.indexOf(p0, i + 1);
    }
    return -1;
  }

  if (patLen === 3) {
    const p0 = pattern[0];
    const p1 = pattern[1];
    const p2 = pattern[2];
    const last = bufLen - 3;
    let i = buffer.indexOf(p0, start);
    while (i !== -1 && i <= last) {
      if (buffer[i + 1] === p1 && buffer[i + 2] === p2) {
        return i;
      }
      i = buffer.indexOf(p0, i + 1);
    }
    return -1;
  }

  if (patLen === 4) {
    const p0 = pattern[0];
    const p1 = pattern[1];
    const p2 = pattern[2];
    const p3 = pattern[3];
    const last = bufLen - 4;
    let i = buffer.indexOf(p0, start);
    while (i !== -1 && i <= last) {
      if (buffer[i + 1] === p1 && buffer[i + 2] === p2 && buffer[i + 3] === p3) {
        return i;
      }
      i = buffer.indexOf(p0, i + 1);
    }
    return -1;
  }

  for (let i = start; i <= bufLen - patLen; i++) {
    let matched = true;
    for (let j = 0; j < patLen; j++) {
      if (buffer[i + j] !== pattern[j]) {
        matched = false;
        break;
      }
    }
    if (matched) {
      return i;
    }
  }
  return -1;
}
