/**
 * WorkbookWriter/WorkbookReader Accuracy Tests - Node.js
 */

import { PassThrough, Readable } from "@stream";
import { createWorkbookRoundtripAccuracyTests } from "@stream/__tests__/streaming/workbook-roundtrip-accuracy-tests";
import { beforeAll } from "vitest";

let WorkbookWriter: any;
let WorkbookReader: any;

beforeAll(async () => {
  const excelModule = await import("../../../index");
  WorkbookWriter = excelModule.WorkbookWriter;
  WorkbookReader = excelModule.WorkbookReader;
});

function getNodeContext() {
  return {
    isBrowser: false,

    createWorkbookWriter: async (
      options: { useSharedStrings: boolean; useStyles: boolean; trueStreaming: boolean },
      onData: (chunk: Uint8Array) => void
    ) => {
      const stream = new PassThrough();
      stream.on("data", (chunk: Uint8Array) => {
        onData(chunk);
      });

      const workbook = new WorkbookWriter({
        stream,
        trueStreaming: options.trueStreaming,
        useSharedStrings: options.useSharedStrings,
        useStyles: options.useStyles
      });

      return {
        addWorksheet: (name: string) => {
          const worksheet = workbook.addWorksheet(name);
          return {
            addRow: (data: (string | number)[]) => worksheet.addRow(data),
            commit: () => worksheet.commit()
          };
        },
        commit: () => workbook.commit()
      };
    },

    createWorkbookReader: async (
      data: Uint8Array,
      onRow: (sheetName: string, rowNumber: number, values: unknown[]) => void
    ) => {
      const stream = Readable.from(Buffer.from(data));
      const reader = new WorkbookReader(stream);

      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          onRow(worksheet.name, row.number, row.values);
        }
      }
    }
  };
}

createWorkbookRoundtripAccuracyTests(getNodeContext);
