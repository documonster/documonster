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

// The sweep above passes `"A1", 1, 1` after the handle, which some of these refuse as options first — so the handle
// went unchecked: `Workbook.toBuffer()` failed with a `TypeError`, and `Worksheet.protect({})` wrote onto the `{}`.
describe("a wrong handle, with every other argument valid", () => {
  const calls: [string, string, (bad: unknown) => unknown][] = [
    ["Workbook.toBuffer", "wb", bad => Workbook.toBuffer(bad as never)],
    ["Workbook.toStream", "wb", bad => Workbook.toStream(bad as never)],
    ["Workbook.writeStream", "wb", bad => Workbook.writeStream(bad as never, {} as never)],
    ["Workbook.protect", "wb", bad => Workbook.protect(bad as never, { lockStructure: true })],
    ["Worksheet.protect", "ws", bad => Worksheet.protect(bad as never)],
    ["Worksheet.protect (password)", "ws", bad => Worksheet.protect(bad as never, null as never)],
    ["Worksheet.freeze", "ws", bad => Worksheet.freeze(bad as never, { rows: 1 })],
    ["Worksheet.freeze (positional)", "ws", bad => Worksheet.freeze(bad as never, 1)],
    ["Worksheet.split", "ws", bad => Worksheet.split(bad as never, { x: 100 })]
  ];

  it.each(
    calls.flatMap(([name, field, call]) =>
      [undefined, 42, {}].map(bad => [name, field, call, bad] as const)
    )
  )("%s refuses %j", async (_name, field, call, bad) => {
    const handle = bad === undefined ? undefined : typeof bad === "object" ? {} : bad;
    let error: unknown;
    try {
      await call(handle);
    } catch (e) {
      error = e;
    }
    expect(error).toBeInstanceOf(ExcelArgumentError);
    expect((error as ExcelArgumentError).field).toBe(field);
    expect((error as Error).message).toMatch(new RegExp(`^${field} must be a work`));
    if (typeof handle === "object") {
      expect(handle).toEqual({});
    }
  });
});
