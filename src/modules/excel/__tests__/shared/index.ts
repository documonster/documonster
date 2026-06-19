import conditionalFormattingData from "@excel/__tests__/shared/data/conditional-formatting.json" with { type: "json" };
import headerFooter from "@excel/__tests__/shared/data/header-footer.json" with { type: "json" };
import pageSetup from "@excel/__tests__/shared/data/page-setup.json" with { type: "json" };
import properties from "@excel/__tests__/shared/data/sheet-properties.json" with { type: "json" };
import testValues from "@excel/__tests__/shared/data/sheet-values.json" with { type: "json" };
import styles from "@excel/__tests__/shared/data/styles.json" with { type: "json" };
import views from "@excel/__tests__/shared/data/views.json" with { type: "json" };
import { conditionalFormatting } from "@excel/__tests__/shared/test-conditional-formatting-sheet";
import { dataValidations } from "@excel/__tests__/shared/test-data-validation-sheet";
import { splice } from "@excel/__tests__/shared/test-spliced-sheet";
import { values } from "@excel/__tests__/shared/test-values-sheet";
import { testWorkbookReader } from "@excel/__tests__/shared/test-workbook-reader";
import { fix } from "@excel/__tests__/shared/tools";
import { addWorksheet, createWorkbook } from "@excel/core/workbook";
import { expect } from "vitest";

/** Local test helper: dotted-path getter (replaces the retired under-dash `get`). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function get<T = any>(obj: unknown, path: string, defaultValue?: T): T {
  let result: unknown = obj;
  for (const key of path.split(".")) {
    if (result == null) {
      return defaultValue as T;
    }
    result = (result as Record<string, unknown>)[key];
  }
  return (result ?? defaultValue) as T;
}

const testSheets = {
  dataValidations,
  conditionalFormatting,
  values,
  splice
};

function getOptions(docType: string, options?: any) {
  let result: any;
  switch (docType) {
    case "xlsx":
      result = {
        sheetName: "values",
        checkFormulas: true,
        checkMerges: true,
        checkStyles: true,
        checkBadAlignments: true,
        checkSheetProperties: true,
        dateAccuracy: 3,
        checkViews: true
      };
      break;
    case "csv":
      result = {
        sheetName: "sheet1",
        checkFormulas: false,
        checkMerges: false,
        checkStyles: false,
        checkBadAlignments: false,
        checkSheetProperties: false,
        dateAccuracy: 1000,
        checkViews: false
      };
      break;
    default:
      throw new Error(`Bad doc-type: ${docType}`);
  }
  return Object.assign(result, options);
}

const testUtils = {
  views: fix(views),
  testValues: fix(testValues),
  styles: fix(styles),
  properties: fix(properties),
  pageSetup: fix(pageSetup),
  conditionalFormatting: fix(conditionalFormattingData),
  headerFooter: fix(headerFooter),

  createTestBook(workbook: any, docType?: string, sheets?: string[]) {
    const options = getOptions(docType || "xlsx");
    sheets = sheets || ["values"];

    workbook.views = [{ x: 1, y: 2, width: 10000, height: 20000, firstSheet: 0, activeTab: 0 }];

    sheets.forEach(sheet => {
      const testSheet = get(testSheets, sheet);
      testSheet.addSheet(workbook, options);
    });

    return workbook;
  },

  checkTestBook(workbook: any, docType?: string, sheets?: string[], options?: any) {
    options = getOptions(docType || "xlsx", options);
    sheets = sheets || ["values"];

    expect(workbook).toBeDefined();

    if (options.checkViews) {
      expect(workbook.views).toEqual([
        {
          x: 1,
          y: 2,
          width: 10000,
          height: 20000,
          firstSheet: 0,
          activeTab: 0,
          visibility: "visible"
        }
      ]);
    }

    sheets.forEach(sheet => {
      const testSheet = get(testSheets, sheet);
      testSheet.checkSheet(workbook, options);
    });
  },

  checkTestBookReader: testWorkbookReader.checkBook,

  createSheetMock(): any {
    // Return a real worksheet record so the flat worksheet/row/column/cell
    // functions (which read `_rows` / `_columns`) work against it.
    return addWorksheet(createWorkbook(), "mock");
  }
};

export { testUtils };
