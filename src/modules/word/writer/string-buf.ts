/**
 * StringBuf - Efficient String Builder for XML Serialization
 *
 * A minimal, self-contained string buffer that builds UTF-8 encoded byte arrays
 * from string input. Uses TextEncoder for cross-platform compatibility.
 *
 * This is a word-module-local implementation to avoid depending on @excel/utils.
 */

const encoder = new TextEncoder();

interface StringBufOptions {
  size?: number;
}

/**
 * Efficient string builder using Uint8Array.
 * Works identically in Node.js and Browser environments.
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

  addText(text: string): void {
    this._buffer = undefined;

    // Ensure there's room for at least the string length (optimistic: 1 byte per char).
    const optimistic = this._inPos + text.length;
    if (optimistic > this._buf.length - 4) {
      this._grow(optimistic);
    }

    // Use encodeInto to write directly into the buffer.
    const target = this._buf.subarray(this._inPos);
    const result = encoder.encodeInto(text, target);
    if (result.read! < text.length) {
      // Didn't fit — grow and retry remainder.
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
