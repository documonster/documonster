import { Compression } from "@archive";
import { defineFormatParitySuite } from "@archive/__tests__/compression/format-parity.shared";

defineFormatParitySuite({
  decompressAuto: Compression.decompressAuto,
  gunzip: Compression.gunzip,
  gunzipSync: Compression.gunzipSync,
  gzipSync: Compression.gzipSync,
  unzlib: Compression.unzlib,
  unzlibSync: Compression.unzlibSync,
  zlibSync: Compression.zlibSync
});
