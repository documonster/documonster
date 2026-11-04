/**
 * Every public function of the document-model namespaces that takes a handle checks it at the boundary: a wrong one —
 * `42`, `null`, `{}` — is an `ExcelArgumentError` naming the argument, never a raw `TypeError` from inside, and never a
 * silent `undefined`.
 */
import {
  Cell,
  Column,
  DefinedNames,
  ExcelArgumentError,
  Note,
  Range,
  Row,
  Workbook,
  Worksheet
} from "@excel/index";
import { describe, expect, it } from "vitest";

type Fn = (...args: unknown[]) => unknown;

/** Functions whose first argument is not a handle. */
const NO_HANDLE: Record<string, readonly string[]> = {
  Workbook: [
    "create",
    "read",
    "readFile",
    "readStream",
    "readWithDiagnostics",
    "createStreamReader",
    "createStreamWriter"
  ],
  DefinedNames: ["create"],
  Note: ["create", "isNote", "fromModel"],
  Range: ["create"],
  // Deprecated, and pinned to its baseline: "anything but an index or a predicate clears, whatever `ws` holds".
  Worksheet: ["removeConditionalFormatting"]
};

const NAMESPACES: Record<string, Record<string, unknown>> = {
  Workbook,
  Worksheet,
  Cell,
  Row,
  Column,
  DefinedNames,
  Note,
  Range
};

const BAD = [42, null, undefined, {}, "S"];

function cases(): [string, Fn, unknown][] {
  const out: [string, Fn, unknown][] = [];
  for (const [ns, members] of Object.entries(NAMESPACES)) {
    for (const [name, value] of Object.entries(members)) {
      if (typeof value !== "function" || NO_HANDLE[ns]?.includes(name)) {
        continue;
      }
      for (const bad of BAD) {
        out.push([`${ns}.${name}(${JSON.stringify(bad) ?? "undefined"})`, value as Fn, bad]);
      }
    }
  }
  return out;
}

describe("a wrong handle", () => {
  it.each(cases())("%s is an ExcelArgumentError", async (_name, fn, bad) => {
    let error: unknown;
    try {
      // Plausible arguments after the handle, so only the handle is wrong.
      await fn(bad, "A1", 1, 1);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
  });

  it("is checked for a removed worksheet's handle too, as before", () => {
    const wb = Workbook.create();
    const ws = Workbook.addWorksheet(wb, "S");
    Workbook.removeWorksheet(wb, ws);
    expect(() => Cell.setValue(ws, "A1", 1)).toThrow(ExcelArgumentError);
    expect(Worksheet.getName(ws)).toBe("S");
  });
});
