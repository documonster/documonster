/**
 * ZIP Streaming Module Tests - Node.js
 *
 * Runs shared streaming ZIP tests in Node.js environment.
 */

import { createDeflateStream } from "@archive";
import type { StreamingZipModuleImports } from "@archive/__tests__/zip/streaming-zip.shared";
import { runStreamingZipTests } from "@archive/__tests__/zip/streaming-zip.shared";
import { ZipParser } from "@archive/unzip/zip-parser";
import { Zip, ZipDeflate } from "@archive/zip/stream";
import { describe } from "vitest";

// =============================================================================
// Run Shared Tests
// =============================================================================
describe("ZIP Streaming - Node.js", () => {
  const imports = {
    Zip,
    ZipDeflate,
    createDeflateStream,
    ZipParser
  } as StreamingZipModuleImports;

  runStreamingZipTests(imports);
});
