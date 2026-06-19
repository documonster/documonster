/**
 * Stream common - end() argument parsing
 *
 * Extracts the optional chunk and callback from the overloaded end() signature.
 * Shared by browser Writable, Transform, and Duplex.
 */

export interface ParsedEndArgs<T> {
  chunk: T | undefined;
  encoding: string | undefined;
  cb: (() => void) | undefined;
}

/**
 * Parse the overloaded `end()` arguments into a normalised form.
 *
 * Supports the three standard overload shapes:
 *   end(callback?)
 *   end(chunk, callback?)
 *   end(chunk, encoding?, callback?)
 */
export function parseEndArgs<T>(
  chunkOrCallback?: T | (() => void),
  encodingOrCallback?: string | (() => void),
  callback?: () => void
): ParsedEndArgs<T> {
  const chunk = typeof chunkOrCallback === "function" ? undefined : chunkOrCallback;
  const encoding = typeof encodingOrCallback === "string" ? encodingOrCallback : undefined;
  const cb: (() => void) | undefined =
    typeof chunkOrCallback === "function"
      ? (chunkOrCallback as () => void)
      : typeof encodingOrCallback === "function"
        ? (encodingOrCallback as () => void)
        : callback;
  return { chunk, encoding, cb };
}
