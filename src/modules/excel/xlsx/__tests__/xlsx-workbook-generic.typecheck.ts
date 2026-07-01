// Type-only regression test for the Node-vs-browser workbook type contract.
//
// This file is typechecked by `pnpm type` (tsgo) but is NOT executed by Vitest.
// It pins down the type contract: under the Node
// entry, `Workbook.xlsx` must be the Node `XLSX` (with `readFile`/`writeFile`)
// and any method that returns the workbook (`load`, `read`, `loadFromFiles`,
// etc.) must return the *Node* `Workbook` so that chaining off the result
// keeps the file-system APIs reachable.
//
// Regression history: the bug was that `XLSX.load()` (defined on the browser
// base class) was hard-coded to `Promise<BrowserWorkbook>`. The Node subclass
// inherited that signature and `await wb.xlsx.load(buf)` produced a value
// typed as the *browser* `Workbook`, which omits `xlsx.readFile` /
// `xlsx.writeFile` and is therefore not assignable to the Node `Workbook`.
// The fix makes `XLSX` generic over its workbook type so the Node subclass
// extends `XLSX<NodeWorkbook>` and the inherited methods narrow naturally.

import type { Workbook as NodeWorkbook } from "@excel/core/workbook";
import { getXlsxIo } from "@excel/core/workbook";
import type { Workbook as BrowserWorkbook } from "@excel/core/workbook.browser";
import type { XLSX as NodeXlsx } from "@excel/xlsx/xlsx";
import type { XLSX as BrowserXlsx } from "@excel/xlsx/xlsx.browser";

type Assert<T extends true> = T;

type IsAny<T> = 0 extends 1 & T ? true : false;
type IsEqual<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
type IsEqualStrict<A, B> =
  IsAny<A> extends true ? false : IsAny<B> extends true ? false : IsEqual<A, B>;

// ---------------------------------------------------------------------------
// Node entry: every method that resolves to the workbook must return the
// Node Workbook (so callers can chain Node-only file APIs off the result).
// ---------------------------------------------------------------------------

declare const nodeXlsx: NodeXlsx;

type NodeLoadResult = Awaited<ReturnType<typeof nodeXlsx.load>>;
type NodeReadResult = Awaited<ReturnType<typeof nodeXlsx.read>>;
type NodeReadFileResult = Awaited<ReturnType<typeof nodeXlsx.readFile>>;
type NodeLoadFromFilesResult = Awaited<ReturnType<typeof nodeXlsx.loadFromFiles>>;

// All four must be the *Node* `Workbook`, not the browser base.
type _NodeLoadIsNode = Assert<IsEqualStrict<NodeLoadResult, NodeWorkbook>>;
type _NodeReadIsNode = Assert<IsEqualStrict<NodeReadResult, NodeWorkbook>>;
type _NodeReadFileIsNode = Assert<IsEqualStrict<NodeReadFileResult, NodeWorkbook>>;
type _NodeLoadFromFilesIsNode = Assert<IsEqualStrict<NodeLoadFromFilesResult, NodeWorkbook>>;

// `xlsx.workbook` itself is the Node Workbook.
type _NodeWorkbookFieldIsNode = Assert<IsEqualStrict<typeof nodeXlsx.workbook, NodeWorkbook>>;

// Sanity: Node-only file APIs must be present on the Node XLSX surface.
type _HasReadFile = Assert<
  typeof nodeXlsx.readFile extends (...args: never) => unknown ? true : false
>;
type _HasWriteFile = Assert<
  typeof nodeXlsx.writeFile extends (...args: never) => unknown ? true : false
>;

// The screenshot scenario: `loaded` must accept Node-only chained calls.
async function _issue160Screenshot(ab: ArrayBuffer): Promise<NodeWorkbook> {
  // We cannot `new Workbook()` here without circular imports, but the type of
  // `wb.xlsx.load(...)` is exactly what the user sees in their IDE.
  const loaded = await nodeXlsx.load(ab);
  await getXlsxIo(loaded).writeFile("out.xlsx");
  return loaded;
}
void _issue160Screenshot;

// ---------------------------------------------------------------------------
// Browser entry: methods must return the browser base Workbook (Node-only
// APIs stay invisible so consumers get a fail-fast type error instead of a
// runtime "not supported" exception).
// ---------------------------------------------------------------------------

declare const browserXlsx: BrowserXlsx;

type BrowserLoadResult = Awaited<ReturnType<typeof browserXlsx.load>>;
type BrowserReadResult = Awaited<ReturnType<typeof browserXlsx.read>>;
type BrowserLoadFromFilesResult = Awaited<ReturnType<typeof browserXlsx.loadFromFiles>>;

type _BrowserLoadIsBrowser = Assert<IsEqualStrict<BrowserLoadResult, BrowserWorkbook>>;
type _BrowserReadIsBrowser = Assert<IsEqualStrict<BrowserReadResult, BrowserWorkbook>>;
type _BrowserLoadFromFilesIsBrowser = Assert<
  IsEqualStrict<BrowserLoadFromFilesResult, BrowserWorkbook>
>;

// The browser surface must NOT expose Node-only file methods.
// Using `keyof` indirection keeps the assertion robust against whatever
// optional / overload shape `XLSX.readFile` / `writeFile` might take.
type _BrowserOmitsReadFile = Assert<"readFile" extends keyof BrowserXlsx ? false : true>;
type _BrowserOmitsWriteFile = Assert<"writeFile" extends keyof BrowserXlsx ? false : true>;

// Liskov: NodeXlsx must remain assignable to BrowserXlsx (it's a subclass).
type _NodeXlsxAssignable = Assert<NodeXlsx extends BrowserXlsx ? true : false>;
