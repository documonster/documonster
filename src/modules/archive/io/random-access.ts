/**
 * Random Access Reader Interface
 *
 * Provides an abstraction for reading arbitrary byte ranges from a data source.
 * This enables efficient access to ZIP archives without downloading the entire file.
 *
 * @module
 */

/**
 * Interface for reading arbitrary byte ranges from a data source.
 *
 * This abstraction allows ZIP parsing to work with:
 * - Remote HTTP resources (via Range requests)
 * - Local files (via fs.read with position)
 * - In-memory buffers (for testing/compatibility)
 */
export interface RandomAccessReader {
  /**
   * Total size of the data source in bytes.
   * Must be known upfront for ZIP parsing (EOCD is at the end).
   */
  readonly size: number;

  /**
   * Read a range of bytes from the data source.
   *
   * @param start - Start offset (inclusive, 0-based)
   * @param end - End offset (exclusive)
   * @returns The requested bytes
   */
  read(start: number, end: number): Promise<Uint8Array>;

  /**
   * Optional: Close the reader and release resources.
   */
  close?(): Promise<void>;
}

/**
 * Options for creating an HTTP Range reader.
 */
export interface HttpRangeReaderOptions {
  /**
   * Custom headers to include in requests.
   * Useful for authentication tokens, etc.
   */
  headers?: Record<string, string>;

  /**
   * Custom fetch function for testing or custom implementations.
   * Defaults to global fetch.
   */
  fetch?: typeof globalThis.fetch;

  /**
   * Abort signal for cancellation support.
   */
  signal?: AbortSignal;

  /**
   * Whether to use credentials (cookies) for cross-origin requests.
   * @default "same-origin"
   */
  credentials?: NonNullable<RequestInit["credentials"]>;

  /**
   * Optional: Pre-known file size to skip HEAD request.
   * If not provided, a HEAD request will be made to determine size.
   */
  size?: number;

  /**
   * Whether to validate that the server supports Range requests.
   * @default true
   */
  validateRangeSupport?: boolean;
}

/**
 * Statistics about HTTP Range reader operations.
 * Useful for debugging and performance monitoring.
 */
export interface HttpRangeReaderStats {
  /** Total number of HTTP requests made */
  requestCount: number;
  /** Total bytes downloaded */
  bytesDownloaded: number;
  /** Total file size */
  totalSize: number;
  /** Percentage of file downloaded */
  downloadedPercent: number;
}

/**
 * Error thrown when the server doesn't support Range requests.
 */
export { RangeNotSupportedError, HttpRangeError } from "@archive/shared/errors";
import { RangeNotSupportedError, HttpRangeError } from "@archive/shared/errors";

/**
 * Parse total file size from Content-Range header.
 * @example "bytes 0-0/12345" => 12345
 */
function parseTotalSizeFromContentRange(value: string | null): number | null {
  if (!value) {
    return null;
  }
  // e.g. "bytes 0-0/12345" or "bytes */12345"
  const slash = value.lastIndexOf("/");
  if (slash === -1) {
    return null;
  }
  const totalPart = value.slice(slash + 1).trim();
  if (!totalPart || totalPart === "*") {
    return null;
  }
  const total = parseInt(totalPart, 10);
  return Number.isFinite(total) && total >= 0 ? total : null;
}

/**
 * HTTP Range Reader
 *
 * Reads arbitrary byte ranges from a remote HTTP resource using Range requests.
 * Works in both Node.js (with fetch) and browsers.
 *
 * @example
 * ```ts
 * const reader = await HttpRangeReader.open("https://example.com/archive.zip");
 * console.log(`File size: ${reader.size} bytes`);
 *
 * // Read last 22 bytes (minimum EOCD size)
 * const eocd = await reader.read(reader.size - 22, reader.size);
 *
 * // Check stats
 * console.log(reader.getStats());
 *
 * await reader.close();
 * ```
 */
export class HttpRangeReader implements RandomAccessReader {
  private readonly url: string;
  private readonly headers: Record<string, string>;
  private readonly fetchFn: typeof globalThis.fetch;
  private readonly signal?: AbortSignal;
  private readonly credentials: NonNullable<RequestInit["credentials"]>;
  private _size: number;

  // If the server ignores Range and returns the whole file (200), cache it.
  private _fullData?: Uint8Array;

  // Stats tracking
  private _requestCount = 0;
  private _bytesDownloaded = 0;

  private constructor(url: string, size: number, options: HttpRangeReaderOptions = {}) {
    this.url = url;
    this._size = size;
    this.headers = options.headers ?? {};
    this.fetchFn = options.fetch ?? globalThis.fetch;
    this.signal = options.signal;
    this.credentials = options.credentials ?? "same-origin";
  }

  /**
   * Total size of the remote file in bytes.
   */
  get size(): number {
    return this._size;
  }

  /**
   * Open a remote ZIP file for random access reading.
   *
   * @param url - URL of the ZIP file
   * @param options - Reader options
   * @returns A configured HttpRangeReader instance
   *
   * @throws {RangeNotSupportedError} If the server doesn't support Range requests
   * @throws {HttpRangeError} If the HTTP request fails
   */
  static async open(url: string, options: HttpRangeReaderOptions = {}): Promise<HttpRangeReader> {
    const fetchFn = options.fetch ?? globalThis.fetch;
    const headers = options.headers ?? {};
    const credentials = options.credentials ?? "same-origin";
    const validateRangeSupport = options.validateRangeSupport ?? true;

    let requestCount = 0;
    let bytesDownloaded = 0;

    const trackedFetch = async (init: RequestInit): Promise<Response> => {
      const response = await fetchFn(url, init);
      requestCount++;
      return response;
    };

    // If size is pre-known, skip HEAD request but still validate range support if requested
    if (options.size !== undefined) {
      if (validateRangeSupport) {
        // Make a small range request to validate
        const response = await trackedFetch({
          method: "GET",
          headers: {
            ...headers,
            Range: "bytes=0-0"
          },
          signal: options.signal,
          credentials
        });

        // 206: Range supported
        // 200: server ignored Range but reader can still work (will cache on first read)
        if (response.status !== 206 && response.status !== 200) {
          throw new RangeNotSupportedError(url);
        }
      }
      const instance = new HttpRangeReader(url, options.size, options);
      instance._requestCount = requestCount;
      instance._bytesDownloaded = bytesDownloaded;
      return instance;
    }

    // Make HEAD request to get file size and check Range support
    const headResponse = await trackedFetch({
      method: "HEAD",
      headers,
      signal: options.signal,
      credentials
    });

    if (!headResponse.ok) {
      throw new HttpRangeError(url, headResponse.status, headResponse.statusText);
    }

    // Check if server supports Range requests
    const acceptRanges = headResponse.headers.get("Accept-Ranges");
    if (validateRangeSupport && acceptRanges !== "bytes") {
      // Some servers don't send Accept-Ranges but still support it
      // Try a range request to be sure
      const testResponse = await trackedFetch({
        method: "GET",
        headers: {
          ...headers,
          Range: "bytes=0-0"
        },
        signal: options.signal,
        credentials
      });

      if (testResponse.status !== 206 && testResponse.status !== 200) {
        throw new RangeNotSupportedError(url);
      }
    }

    // Determine file size.
    // Prefer Content-Length from HEAD, but fall back to Content-Range from a probe request.
    const contentLength = headResponse.headers.get("Content-Length");
    let size = contentLength ? parseInt(contentLength, 10) : NaN;

    if (!Number.isFinite(size) || size < 0) {
      // Probe with a 0-0 range request to read Content-Range: bytes 0-0/total
      const probeResponse = await trackedFetch({
        method: "GET",
        headers: {
          ...headers,
          Range: "bytes=0-0"
        },
        signal: options.signal,
        credentials
      });

      const probedTotal = parseTotalSizeFromContentRange(
        probeResponse.headers.get("Content-Range")
      );

      if (probedTotal !== null) {
        size = probedTotal;
      } else if (probeResponse.status === 200) {
        // Server ignored Range and returned the full file for the probe request.
        // Consume the probe response body and cache it to avoid a redundant full GET.
        try {
          const data = new Uint8Array(await probeResponse.arrayBuffer());
          bytesDownloaded += data.length;
          size = data.length;

          const instance = new HttpRangeReader(url, size, options);
          instance._fullData = data;
          instance._requestCount = requestCount;
          instance._bytesDownloaded = bytesDownloaded;
          return instance;
        } catch {
          // As a last resort, download once and measure.
          // This is expensive but makes open() work even without size headers.
          const full = await trackedFetch({
            method: "GET",
            headers,
            signal: options.signal,
            credentials
          });

          if (!full.ok) {
            throw new HttpRangeError(url, full.status, full.statusText);
          }

          const data = new Uint8Array(await full.arrayBuffer());
          bytesDownloaded += data.length;
          size = data.length;

          const instance = new HttpRangeReader(url, size, options);
          instance._fullData = data;
          instance._requestCount = requestCount;
          instance._bytesDownloaded = bytesDownloaded;
          return instance;
        }
      } else {
        throw new Error(
          contentLength
            ? `Invalid Content-Length "${contentLength}" for: ${url}`
            : `Server did not provide Content-Length for: ${url}`
        );
      }
    }

    const instance = new HttpRangeReader(url, size, options);
    instance._requestCount = requestCount;
    instance._bytesDownloaded = bytesDownloaded;
    return instance;
  }

  /**
   * Read a range of bytes from the remote file.
   *
   * @param start - Start offset (inclusive, 0-based)
   * @param end - End offset (exclusive)
   * @returns The requested bytes
   *
   * @throws {HttpRangeError} If the HTTP request fails
   */
  async read(start: number, end: number): Promise<Uint8Array> {
    if (start < 0 || end > this._size || start >= end) {
      throw new RangeError(`Invalid range [${start}, ${end}) for file of size ${this._size}`);
    }

    if (this._fullData) {
      return this._fullData.subarray(start, end);
    }

    // HTTP Range header uses inclusive end
    const rangeHeader = `bytes=${start}-${end - 1}`;

    const response = await this.fetchFn(this.url, {
      method: "GET",
      headers: {
        ...this.headers,
        Range: rangeHeader
      },
      signal: this.signal,
      credentials: this.credentials
    });

    // 206 Partial Content is expected for Range requests
    // 200 OK means server ignored Range header and returned full file
    if (response.status === 200) {
      // Server returned full content - extract the range we need
      const fullData = new Uint8Array(await response.arrayBuffer());
      this._fullData = fullData;
      // If the server returned a different length than expected, prefer the actual.
      // This keeps read() bounds consistent for subsequent calls.
      this._size = fullData.length;
      this._requestCount++;
      this._bytesDownloaded += fullData.length;
      return fullData.subarray(start, end);
    }

    if (response.status !== 206) {
      throw new HttpRangeError(this.url, response.status, response.statusText);
    }

    const data = new Uint8Array(await response.arrayBuffer());
    this._requestCount++;
    this._bytesDownloaded += data.length;
    // Some servers may return a larger-than-requested slice; clamp to requested length.
    const expectedLen = end - start;
    return data.length === expectedLen ? data : data.subarray(0, expectedLen);
  }

  /**
   * Get statistics about the reader's operations.
   */
  getStats(): HttpRangeReaderStats {
    return {
      requestCount: this._requestCount,
      bytesDownloaded: this._bytesDownloaded,
      totalSize: this._size,
      downloadedPercent:
        this._size > 0 ? Math.round((this._bytesDownloaded / this._size) * 10000) / 100 : 0
    };
  }

  /**
   * Close the reader.
   * Currently a no-op but included for interface compliance.
   */
  async close(): Promise<void> {
    // HTTP is stateless, nothing to close
  }
}

/**
 * In-memory random access reader.
 * Wraps a Uint8Array to provide the RandomAccessReader interface.
 * Useful for testing and as a fallback.
 */
export class BufferReader implements RandomAccessReader {
  private readonly data: Uint8Array;

  constructor(data: Uint8Array | ArrayBuffer) {
    this.data = data instanceof ArrayBuffer ? new Uint8Array(data) : data;
  }

  get size(): number {
    return this.data.length;
  }

  async read(start: number, end: number): Promise<Uint8Array> {
    if (start < 0 || end > this.data.length || start >= end) {
      throw new RangeError(
        `Invalid range [${start}, ${end}) for buffer of size ${this.data.length}`
      );
    }
    return this.data.subarray(start, end);
  }

  async close(): Promise<void> {
    // Nothing to close
  }
}
