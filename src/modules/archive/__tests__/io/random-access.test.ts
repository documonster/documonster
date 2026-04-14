import {
  HttpRangeReader,
  BufferReader,
  RangeNotSupportedError,
  HttpRangeError,
  type RandomAccessReader
} from "@archive/io/random-access";
import { describe, it, expect, vi, beforeEach, afterEach } from "vitest";

describe("BufferReader", () => {
  it("should provide correct size", () => {
    const data = new Uint8Array([1, 2, 3, 4, 5]);
    const reader = new BufferReader(data);
    expect(reader.size).toBe(5);
  });

  it("should accept ArrayBuffer", () => {
    const buffer = new ArrayBuffer(10);
    const reader = new BufferReader(buffer);
    expect(reader.size).toBe(10);
  });

  it("should read a range correctly", async () => {
    const data = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    const reader = new BufferReader(data);

    const result = await reader.read(2, 5);
    expect(result).toEqual(new Uint8Array([2, 3, 4]));
  });

  it("should read from start", async () => {
    const data = new Uint8Array([10, 20, 30, 40, 50]);
    const reader = new BufferReader(data);

    const result = await reader.read(0, 3);
    expect(result).toEqual(new Uint8Array([10, 20, 30]));
  });

  it("should read to end", async () => {
    const data = new Uint8Array([10, 20, 30, 40, 50]);
    const reader = new BufferReader(data);

    const result = await reader.read(3, 5);
    expect(result).toEqual(new Uint8Array([40, 50]));
  });

  it("should read entire buffer", async () => {
    const data = new Uint8Array([1, 2, 3]);
    const reader = new BufferReader(data);

    const result = await reader.read(0, 3);
    expect(result).toEqual(data);
  });

  it("should throw on invalid range - negative start", async () => {
    const reader = new BufferReader(new Uint8Array(10));
    await expect(reader.read(-1, 5)).rejects.toThrow(RangeError);
  });

  it("should throw on invalid range - end beyond size", async () => {
    const reader = new BufferReader(new Uint8Array(10));
    await expect(reader.read(5, 15)).rejects.toThrow(RangeError);
  });

  it("should throw on invalid range - start >= end", async () => {
    const reader = new BufferReader(new Uint8Array(10));
    await expect(reader.read(5, 5)).rejects.toThrow(RangeError);
    await expect(reader.read(6, 5)).rejects.toThrow(RangeError);
  });

  it("close should be a no-op", async () => {
    const reader = new BufferReader(new Uint8Array(10));
    await expect(reader.close()).resolves.toBeUndefined();
  });
});

describe("HttpRangeReader", () => {
  let mockFetch: any;

  beforeEach(() => {
    mockFetch = vi.fn();
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  describe("open", () => {
    it("should make HEAD request to determine size", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": "1000",
          "Accept-Ranges": "bytes"
        })
      });

      const reader = await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch
      });

      expect(reader.size).toBe(1000);
      expect(mockFetch).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledWith("https://example.com/test.zip", {
        method: "HEAD",
        headers: {},
        signal: undefined,
        credentials: "same-origin"
      });
    });

    it("should validate Range support with test request when Accept-Ranges not present", async () => {
      // HEAD response without Accept-Ranges
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": "1000"
        })
      });
      // Range test request succeeds
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        headers: new Headers({})
      });

      const reader = await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch
      });

      expect(reader.size).toBe(1000);
      expect(mockFetch).toHaveBeenCalledTimes(2);
    });

    it("should throw RangeNotSupportedError when server does not support Range", async () => {
      // HEAD response without Accept-Ranges
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": "1000"
        })
      });
      // Range test request fails (returns 416 instead of 206/200)
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 416,
        statusText: "Range Not Satisfiable",
        headers: new Headers({})
      });

      await expect(
        HttpRangeReader.open("https://example.com/test.zip", { fetch: mockFetch })
      ).rejects.toThrow(RangeNotSupportedError);
    });

    it("should throw HttpRangeError on HTTP error", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 404,
        statusText: "Not Found",
        headers: new Headers({})
      });

      await expect(
        HttpRangeReader.open("https://example.com/test.zip", { fetch: mockFetch })
      ).rejects.toThrow(HttpRangeError);
    });

    it("should skip HEAD request when size is pre-known", async () => {
      // Only validation request
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        headers: new Headers({})
      });

      const reader = await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch,
        size: 500
      });

      expect(reader.size).toBe(500);
      expect(mockFetch).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledWith(
        "https://example.com/test.zip",
        expect.objectContaining({
          method: "GET",
          headers: expect.objectContaining({ Range: "bytes=0-0" })
        })
      );
    });

    it("should skip all requests when size pre-known and validateRangeSupport is false", async () => {
      const reader = await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch,
        size: 500,
        validateRangeSupport: false
      });

      expect(reader.size).toBe(500);
      expect(mockFetch).not.toHaveBeenCalled();
    });

    it("should pass custom headers", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": "1000",
          "Accept-Ranges": "bytes"
        })
      });

      await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch,
        headers: { Authorization: "Bearer token123" }
      });

      expect(mockFetch).toHaveBeenCalledWith(
        "https://example.com/test.zip",
        expect.objectContaining({
          headers: { Authorization: "Bearer token123" }
        })
      );
    });

    it("should pass credentials option", async () => {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": "1000",
          "Accept-Ranges": "bytes"
        })
      });

      await HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch,
        credentials: "include"
      });

      expect(mockFetch).toHaveBeenCalledWith(
        "https://example.com/test.zip",
        expect.objectContaining({ credentials: "include" })
      );
    });
  });

  describe("read", () => {
    async function createReader(size = 1000): Promise<HttpRangeReader> {
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200,
        statusText: "OK",
        headers: new Headers({
          "Content-Length": String(size),
          "Accept-Ranges": "bytes"
        })
      });

      return HttpRangeReader.open("https://example.com/test.zip", {
        fetch: mockFetch
      });
    }

    it("should make Range request with correct header", async () => {
      const reader = await createReader();
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        arrayBuffer: () => Promise.resolve(new ArrayBuffer(100))
      });

      await reader.read(100, 200);

      expect(mockFetch).toHaveBeenLastCalledWith(
        "https://example.com/test.zip",
        expect.objectContaining({
          method: "GET",
          headers: expect.objectContaining({ Range: "bytes=100-199" })
        })
      );
    });

    it("should return requested data", async () => {
      const reader = await createReader();
      const expectedData = new Uint8Array([1, 2, 3, 4, 5]);
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        arrayBuffer: () => Promise.resolve(expectedData.buffer)
      });

      const result = await reader.read(0, 5);
      expect(result).toEqual(expectedData);
    });

    it("should handle server returning full content (status 200)", async () => {
      const reader = await createReader(10);
      const fullData = new Uint8Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 200, // Server returned full content
        statusText: "OK",
        arrayBuffer: () => Promise.resolve(fullData.buffer)
      });

      const result = await reader.read(2, 5);
      expect(result).toEqual(new Uint8Array([2, 3, 4]));

      // Subsequent reads should be served from cached full content.
      const beforeCalls = mockFetch.mock.calls.length;
      const result2 = await reader.read(5, 7);
      expect(result2).toEqual(new Uint8Array([5, 6]));
      expect(mockFetch.mock.calls.length).toBe(beforeCalls);
    });

    it("should throw HttpRangeError on HTTP error", async () => {
      const reader = await createReader();
      mockFetch.mockResolvedValueOnce({
        ok: false,
        status: 500,
        statusText: "Internal Server Error"
      });

      await expect(reader.read(0, 100)).rejects.toThrow(HttpRangeError);
    });

    it("should throw RangeError on invalid range", async () => {
      const reader = await createReader(100);

      await expect(reader.read(-1, 50)).rejects.toThrow(RangeError);
      await expect(reader.read(0, 200)).rejects.toThrow(RangeError);
      await expect(reader.read(50, 50)).rejects.toThrow(RangeError);
    });

    it("should track stats correctly", async () => {
      const reader = await createReader();

      // First read
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        arrayBuffer: () => Promise.resolve(new ArrayBuffer(50))
      });
      await reader.read(0, 50);

      // Second read
      mockFetch.mockResolvedValueOnce({
        ok: true,
        status: 206,
        statusText: "Partial Content",
        arrayBuffer: () => Promise.resolve(new ArrayBuffer(30))
      });
      await reader.read(100, 130);

      const stats = reader.getStats();
      // Includes the initial HEAD request in open()
      expect(stats.requestCount).toBe(3);
      expect(stats.bytesDownloaded).toBe(80);
      expect(stats.totalSize).toBe(1000);
      expect(stats.downloadedPercent).toBe(8);
    });
  });
});

describe("RandomAccessReader interface", () => {
  it("BufferReader implements RandomAccessReader", () => {
    const reader: RandomAccessReader = new BufferReader(new Uint8Array(10));
    expect(reader.size).toBe(10);
    expect(typeof reader.read).toBe("function");
  });

  it("HttpRangeReader implements RandomAccessReader", async () => {
    const mockFetch = vi.fn().mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Content-Length": "1000",
        "Accept-Ranges": "bytes"
      })
    });

    const reader: RandomAccessReader = await HttpRangeReader.open("https://example.com/test.zip", {
      fetch: mockFetch
    });
    expect(reader.size).toBe(1000);
    expect(typeof reader.read).toBe("function");
  });
});

describe("Error classes", () => {
  it("RangeNotSupportedError should have correct properties", () => {
    const error = new RangeNotSupportedError("https://example.com/test.zip");
    expect(error.name).toBe("RangeNotSupportedError");
    expect(error.message).toContain("Range requests");
    expect(error.message).toContain("https://example.com/test.zip");
  });

  it("HttpRangeError should have correct properties", () => {
    const error = new HttpRangeError("https://example.com/test.zip", 404, "Not Found");
    expect(error.name).toBe("HttpRangeError");
    expect(error.url).toBe("https://example.com/test.zip");
    expect(error.status).toBe(404);
    expect(error.statusText).toBe("Not Found");
    expect(error.message).toContain("404");
    expect(error.message).toContain("Not Found");
  });
});

describe("HttpRangeReader close", () => {
  it("should be callable and resolve", async () => {
    const mockFetch = vi.fn().mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Content-Length": "1000",
        "Accept-Ranges": "bytes"
      })
    });

    const reader = await HttpRangeReader.open("https://example.com/test.zip", {
      fetch: mockFetch
    });

    await expect(reader.close()).resolves.toBeUndefined();
  });
});

describe("HttpRangeReader with AbortSignal", () => {
  it("should pass signal to fetch requests", async () => {
    const mockFetch = vi.fn();
    const controller = new AbortController();

    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Content-Length": "1000",
        "Accept-Ranges": "bytes"
      })
    });

    await HttpRangeReader.open("https://example.com/test.zip", {
      fetch: mockFetch,
      signal: controller.signal
    });

    expect(mockFetch).toHaveBeenCalledWith(
      "https://example.com/test.zip",
      expect.objectContaining({
        signal: controller.signal
      })
    );
  });

  it("should pass signal to read requests", async () => {
    const mockFetch = vi.fn();
    const controller = new AbortController();

    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Content-Length": "1000",
        "Accept-Ranges": "bytes"
      })
    });

    const reader = await HttpRangeReader.open("https://example.com/test.zip", {
      fetch: mockFetch,
      signal: controller.signal
    });

    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 206,
      statusText: "Partial Content",
      arrayBuffer: () => Promise.resolve(new ArrayBuffer(100))
    });

    await reader.read(0, 100);

    expect(mockFetch).toHaveBeenLastCalledWith(
      "https://example.com/test.zip",
      expect.objectContaining({
        signal: controller.signal
      })
    );
  });
});

describe("HttpRangeReader Content-Length validation", () => {
  it("should fall back to Content-Range when Content-Length is missing", async () => {
    const mockFetch = vi.fn();
    // HEAD without Content-Length
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Accept-Ranges": "bytes"
      })
    });
    // Probe range request provides Content-Range total
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 206,
      statusText: "Partial Content",
      headers: new Headers({
        "Content-Range": "bytes 0-0/1234"
      })
    });

    const reader = await HttpRangeReader.open("https://example.com/test.zip", { fetch: mockFetch });
    expect(reader.size).toBe(1234);
  });

  it("should throw when Content-Length and Content-Range are missing", async () => {
    const mockFetch = vi.fn();
    // HEAD without Content-Length
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Accept-Ranges": "bytes"
      })
    });
    // Probe provides no Content-Range
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 206,
      statusText: "Partial Content",
      headers: new Headers({})
    });

    await expect(
      HttpRangeReader.open("https://example.com/test.zip", { fetch: mockFetch })
    ).rejects.toThrow(/Content-Length/);
  });

  it("should fall back to Content-Range when Content-Length is invalid", async () => {
    const mockFetch = vi.fn();
    // HEAD with invalid Content-Length
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 200,
      statusText: "OK",
      headers: new Headers({
        "Accept-Ranges": "bytes",
        "Content-Length": "not-a-number"
      })
    });
    // Probe range request provides Content-Range total
    mockFetch.mockResolvedValueOnce({
      ok: true,
      status: 206,
      statusText: "Partial Content",
      headers: new Headers({
        "Content-Range": "bytes 0-0/1000"
      })
    });

    const reader = await HttpRangeReader.open("https://example.com/test.zip", { fetch: mockFetch });
    expect(reader.size).toBe(1000);
  });
});
