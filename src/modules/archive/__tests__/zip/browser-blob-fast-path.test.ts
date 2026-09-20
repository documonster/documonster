/**
 * The decision table behind `ZipArchive.bytes()`'s browser Blob fast path.
 *
 * A Node test for a browser-only path: `isNode` is one of the predicate's
 * terms, which is why it is a pure function rather than an inline boolean.
 */

import { canUseBrowserBlobFastPath } from "@archive/zip/zip-archive";
import { describe, expect, it } from "vitest";

/** The combination that should take the fast path. */
function baseFacts() {
  return {
    isNode: false,
    hasBlobSource: true,
    allSourcesInMemory: true,
    smartStore: false,
    zip64Requested: false,
    needsDeflate: true,
    hasDeflateRawCompressionStream: () => true
  };
}

describe("canUseBrowserBlobFastPath", () => {
  it("takes the fast path when every term allows it", () => {
    expect(canUseBrowserBlobFastPath(baseFacts())).toBe(true);
  });

  describe("disqualifying terms", () => {
    it.each([
      ["running under Node", { isNode: true }],
      ["no Blob source to stream", { hasBlobSource: false }],
      ["a source is not in memory", { allSourcesInMemory: false }],
      ["smartStore is on", { smartStore: true }],
      ["ZIP64 was requested", { zip64Requested: true }]
    ])("refuses when %s", (_label, override) => {
      expect(canUseBrowserBlobFastPath({ ...baseFacts(), ...override })).toBe(false);
    });
  });

  describe("deflate-raw is only required by an entry that deflates", () => {
    it("refuses a deflating archive when deflate-raw is unavailable", () => {
      expect(
        canUseBrowserBlobFastPath({
          ...baseFacts(),
          needsDeflate: true,
          hasDeflateRawCompressionStream: () => false
        })
      ).toBe(false);
    });

    it("still takes the fast path for a STORE-only archive without deflate-raw", () => {
      // Chromium 80–102 and Opera 67–88: `CompressionStream` exists but the
      // `deflate-raw` format does not. A level-0 archive never constructs one,
      // so refusing it here would make it buffer every Blob for a codec it
      // was never going to call.
      expect(
        canUseBrowserBlobFastPath({
          ...baseFacts(),
          needsDeflate: false,
          hasDeflateRawCompressionStream: () => false
        })
      ).toBe(true);
    });

    it("does not even probe for deflate-raw when nothing deflates", () => {
      // The probe answers by constructing a `CompressionStream`. An archive
      // that will not deflate should not pay for that.
      let probed = 0;
      const result = canUseBrowserBlobFastPath({
        ...baseFacts(),
        needsDeflate: false,
        hasDeflateRawCompressionStream: () => {
          probed++;
          return true;
        }
      });
      expect(result).toBe(true);
      expect(probed).toBe(0);
    });

    it("probes exactly once when something deflates", () => {
      let probed = 0;
      canUseBrowserBlobFastPath({
        ...baseFacts(),
        needsDeflate: true,
        hasDeflateRawCompressionStream: () => {
          probed++;
          return true;
        }
      });
      expect(probed).toBe(1);
    });

    it("checks the cheap disqualifiers before probing", () => {
      // A Node caller must not construct a CompressionStream to be told it
      // cannot use a browser-only path.
      let probed = 0;
      canUseBrowserBlobFastPath({
        ...baseFacts(),
        isNode: true,
        hasDeflateRawCompressionStream: () => {
          probed++;
          return true;
        }
      });
      expect(probed).toBe(0);
    });
  });
});
