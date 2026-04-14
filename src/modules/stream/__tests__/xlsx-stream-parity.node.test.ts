/**
 * XLSX stream vs non-stream parity tests - Node.js
 */

import { createXlsxStreamParityTests } from "@stream/__tests__/streaming/xlsx-stream-parity-tests";
import { beforeAll } from "vitest";

let Workbook: any;
let PassThrough: any;

beforeAll(async () => {
  const excelModule = await import("../../../index");
  Workbook = excelModule.Workbook;

  const streamModule = await import("@stream");
  PassThrough = streamModule.PassThrough;
});

createXlsxStreamParityTests(() => ({ Workbook, PassThrough }));
