/**
 * X10 for the feature namespaces: every public function that takes a handle checks it at the boundary — a wrong one
 * (`42`, `null`, `{}`, a name) is an `ExcelArgumentError` naming the argument, never a raw `TypeError` from inside or a
 * silent result. The document-model namespaces are covered by `handle-checks.test.ts`.
 */
import {
  Anchor,
  Chart,
  Chartsheet,
  ConditionalFormatting,
  DataValidation,
  ExcelArgumentError,
  Form,
  HeaderFooterImage,
  Image,
  Pivot,
  Sparkline,
  Stream,
  Table
} from "@excel/index";
import { describe, expect, it } from "vitest";

type Fn = (...args: unknown[]) => unknown;

/** Functions whose first argument is not a handle. */
const NO_HANDLE: Record<string, readonly string[]> = {
  // A predicate, and a conversion of a model or input into an anchor.
  Anchor: ["isAnchor", "asInstance"],
  Chart: ["applyPreset", "listPresets", "getPreset"],
  // Deprecated builders from a model.
  Chartsheet: ["create"],
  DataValidation: ["create"],
  Form: ["isCheckbox"],
  Stream: [
    "WorkbookWriter",
    "WorkbookReader",
    "WorksheetWriter",
    "WorksheetReader",
    "createWriter",
    "createReader"
  ]
};

const NAMESPACES: Record<string, Record<string, unknown>> = {
  Anchor,
  Chart,
  Chartsheet,
  ConditionalFormatting,
  DataValidation,
  Form,
  HeaderFooterImage,
  Image,
  Pivot,
  Sparkline,
  Stream,
  Table
};

const BAD = [42, null, undefined, {}, "S"];

/** Functions whose handle may be left out: `undefined` is a valid first argument. */
const OPTIONAL: readonly string[] = ["Anchor.create"];

function cases(): [string, Fn, unknown][] {
  const out: [string, Fn, unknown][] = [];
  for (const [ns, members] of Object.entries(NAMESPACES)) {
    for (const [name, value] of Object.entries(members)) {
      if (
        typeof value !== "function" ||
        NO_HANDLE[ns]?.includes(name) ||
        /^[A-Z]/.test(name) // classes and sub-namespaces
      ) {
        continue;
      }
      for (const bad of BAD) {
        if (bad === undefined && OPTIONAL.includes(`${ns}.${name}`)) {
          continue;
        }
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
      await fn(bad, "A1", 1, 1);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
  });
});
