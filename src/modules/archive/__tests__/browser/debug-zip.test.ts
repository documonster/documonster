import { hasSignature } from "@archive/__tests__/zip/zip-test-utils";
import { CENTRAL_DIR_HEADER_SIG, END_OF_CENTRAL_DIR_SIG } from "@archive/zip-spec/zip-records";
import { StreamingZip, ZipDeflateFile } from "@archive/zip/stream";
import { concatUint8Arrays } from "@utils/binary";
/**
 * Debug test to verify StreamingZip produces correct ZIP files
 */
import { describe, it, expect } from "vitest";

describe("Debug ZIP output", () => {
  it("should produce ZIP with data descriptor", async () => {
    const chunks: Uint8Array[] = [];

    let resolveFinish: (() => void) | undefined;
    const finishPromise = new Promise<void>(resolve => {
      resolveFinish = resolve;
    });

    const zip = new StreamingZip((err: Error | null, data: Uint8Array, final: boolean) => {
      if (err) {
        throw err;
      }
      if (data && data.length > 0) {
        chunks.push(data);
      }
      if (final) {
        resolveFinish?.();
      }
    });

    // Add a file and await completion
    const file = new ZipDeflateFile("test.txt", { level: 6 });
    zip.add(file);
    file.push(new TextEncoder().encode("Hello, World!"), true);
    await file.complete();

    zip.end();
    await finishPromise;

    const fullZip = concatUint8Arrays(chunks);

    // ZIP should be larger than just the local header (30 bytes + filename)
    expect(fullZip.length).toBeGreaterThan(50);

    // Check for local file header signature
    const localHeaderSig =
      (fullZip[0] | (fullZip[1] << 8) | (fullZip[2] << 16) | (fullZip[3] << 24)) >>> 0;
    expect(localHeaderSig).toBe(0x04034b50); // PK\x03\x04

    // Look for data descriptor signature (0x08074b50)
    let foundDataDescriptor = false;
    for (let i = 0; i < fullZip.length - 4; i++) {
      const sig =
        (fullZip[i] | (fullZip[i + 1] << 8) | (fullZip[i + 2] << 16) | (fullZip[i + 3] << 24)) >>>
        0;
      if (sig === 0x08074b50) {
        foundDataDescriptor = true;
        break;
      }
    }
    expect(foundDataDescriptor).toBe(true);

    expect(hasSignature(fullZip, CENTRAL_DIR_HEADER_SIG, 0, fullZip.length)).toBe(true);

    expect(
      hasSignature(
        fullZip,
        END_OF_CENTRAL_DIR_SIG,
        Math.max(0, fullZip.length - 256),
        fullZip.length
      )
    ).toBe(true);
  });
});
