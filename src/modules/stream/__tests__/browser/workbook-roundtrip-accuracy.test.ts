/**
 * WorkbookWriter/WorkbookReader Accuracy Tests - Browser
 */

import { createWorkbookRoundtripAccuracyTests } from "@stream/__tests__/streaming/workbook-roundtrip-accuracy-tests";
import { beforeAll } from "vitest";

let WorkbookWriter: any;
let WorkbookReader: any;

beforeAll(async () => {
  const excelModule = await import("../../../../index.browser");
  WorkbookWriter = excelModule.WorkbookWriter;
  WorkbookReader = excelModule.WorkbookReader;
});

function getBrowserContext() {
  return {
    isBrowser: true,

    createWorkbookWriter: async (
      options: { useSharedStrings: boolean; useStyles: boolean; trueStreaming: boolean },
      onData: (chunk: Uint8Array) => void
    ) => {
      const writable = new WritableStream<Uint8Array>({
        write(chunk) {
          onData(chunk);
        }
      });

      const workbook = new WorkbookWriter({
        stream: writable,
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
      const reader = new WorkbookReader(data);

      for await (const worksheet of reader) {
        for await (const row of worksheet) {
          onRow(worksheet.name, row.number, row.values);
        }
      }
    }
  };
}

createWorkbookRoundtripAccuracyTests(getBrowserContext);
