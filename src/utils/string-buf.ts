/**
 * StringBuf - Cross-Platform String Buffer
 *
 * Efficient string builder that accumulates UTF-8 encoded bytes into a single
 * `Uint8Array`. Designed for hot paths that build large XML / text payloads
 * (XLSX worksheet writers, streaming DOCX writer) where naïve `string +=`
 * concatenation triggers O(n²) re-allocations.
 *
 * Works identically in Node.js and Browser environments — uses only
 * `TextEncoder` and `Uint8Array`.
 */

const encoder = new TextEncoder();

interface StringBufOptions {
  /** Initial capacity in bytes (default: 16384). */
  size?: number;
  /**
   * Encoding label, accepted for API symmetry. Only UTF-8 is supported by
   * `TextEncoder`, so this option is ignored.
   */
  encoding?: string;
}

/**
 * Efficient string builder backed by a growable `Uint8Array`.
 */
class StringBuf {
  private _buf: Uint8Array;
  private _inPos: number;
  private _buffer: Uint8Array | undefined;

  constructor(options?: StringBufOptions) {
    this._buf = new Uint8Array((options && options.size) || 16384);
    this._inPos = 0;
    this._buffer = undefined;
  }

  get length(): number {
    return this._inPos;
  }

  get capacity(): number {
    return this._buf.length;
  }

  get buffer(): Uint8Array {
    return this._buf;
  }

  /** Return a snapshot of the bytes written so far. Cached until next mutation. */
  toBuffer(): Uint8Array {
    if (!this._buffer) {
      this._buffer = this._buf.slice(0, this._inPos);
    }
    return this._buffer;
  }

  reset(position?: number): void {
    position = position ?? 0;
    this._buffer = undefined;
    this._inPos = position;
  }

  private _grow(min: number): void {
    let size = this._buf.length * 2;
    while (size < min) {
      size *= 2;
    }
    const buf = new Uint8Array(size);
    buf.set(this._buf);
    this._buf = buf;
  }

  /** Append a string, encoded as UTF-8. */
  addText(text: string): void {
    this._buffer = undefined;

    // Optimistically reserve 1 byte per char (true for ASCII; multi-byte
    // chars trigger the slow path below).
    const optimistic = this._inPos + text.length;
    if (optimistic > this._buf.length - 4) {
      this._grow(optimistic);
    }

    // `encodeInto` writes directly into the buffer, avoiding the intermediate
    // Uint8Array allocation that `encoder.encode()` creates.
    const target = this._buf.subarray(this._inPos);
    const result = encoder.encodeInto(text, target);
    if (result.read! < text.length) {
      // Multi-byte chars exceeded the optimistic estimate.
      // Worst case: 3 bytes per remaining char.
      const remaining = text.length - result.read!;
      this._grow(this._inPos + result.written + remaining * 3);
      const result2 = encoder.encodeInto(
        text.substring(result.read!),
        this._buf.subarray(this._inPos + result.written)
      );
      this._inPos += result.written + result2.written;
    } else {
      this._inPos += result.written;
    }
  }

  /** Append the contents of another StringBuf without re-encoding. */
  addStringBuf(inBuf: StringBuf): void {
    if (inBuf.length) {
      this._buffer = undefined;

      if (this.length + inBuf.length > this.capacity) {
        this._grow(this.length + inBuf.length);
      }

      this._buf.set(inBuf._buf.subarray(0, inBuf.length), this._inPos);
      this._inPos += inBuf.length;
    }
  }
}

export { StringBuf };
export type { StringBufOptions };
