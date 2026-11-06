/**
 * Ported code has to produce what the original produced.
 *
 * `verify-exceljs-guide.ts` proves the exceljs guide names every exceljs API, and `verify-guide-tables.ts` that its
 * replacements compile; neither runs anything. This runs the guide's mappings: each scenario is written twice, once
 * against exceljs as an exceljs user wrote it and once against `documonster` as the guide ports it, both are saved, and
 * both files are read back by the same reader — exceljs — into a snapshot of what a user sees: every cell's value,
 * type and style (font, fill, border, alignment, number format), notes, merges, data validations, column widths and
 * keys, row heights, defined names, tables and the sheet's own settings. The two snapshots must be equal.
 *
 * The scenarios are the mappings most likely to differ: values by row and column (the 1-based arrays), styles at
 * cell, row and column level, `row.splice`, merges with and without style, notes, validations, conditional formats,
 * tables, defined names, page setup and views, column definitions with keyed rows, inserting and splicing rows,
 * dates, and the streaming writer; and for CSV, reading and writing.
 *
 * Where documonster deliberately differs from exceljs, the scenario says so and checks the documented difference.
 *
 * Needs `dist` (`pnpm build:esm`) and network access (it installs exceljs and its dependencies into a temporary
 * directory), so it is a command to run when the guide or the Excel module changes rather than part of `pnpm check`.
 *
 * `--against excelts` runs the same workbook scenarios against `@cj-tech-master/excelts@10.2.0`, which kept exceljs's
 * object model (`wb.xlsx`, `ws.getCell`, …) and is its own reader there. Its CSV methods moved off `wb.csv`, so the CSV
 * scenarios are exceljs's only.
 *
 * Usage:
 *   node scripts/verify-exceljs-parity.ts
 *   node scripts/verify-exceljs-parity.ts --against excelts
 */
import fs from "node:fs";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";
import { isDeepStrictEqual } from "node:util";

import { run } from "./api-changes/package.ts";

// exceljs reads a CSV date with no zone in the machine's local time, documonster in UTC (the guide's §17): compare
// them where the two agree, so that the documented difference is not reported as a new one.
process.env.TZ = "UTC";

const AGAINST = process.argv.includes("--against")
  ? process.argv[process.argv.indexOf("--against") + 1]
  : "exceljs";
const PACKAGE = AGAINST === "excelts" ? "@cj-tech-master/excelts@10.2.0" : "exceljs@4.4.0";

/* eslint-disable @typescript-eslint/no-explicit-any -- exceljs is loaded at run time, untyped */
type Any = any;

interface Snapshot {
  sheets: Record<string, Record<string, unknown>>;
  names: unknown;
}

/** What a user of the file sees, as exceljs reads it. */
async function snapshot(ExcelJS: Any, bytes: Uint8Array): Promise<Snapshot> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(bytes));
  const sheets: Snapshot["sheets"] = {};
  const style = (s: Any): unknown => {
    const out: Record<string, unknown> = {};
    for (const k of ["font", "fill", "border", "alignment", "numFmt", "protection"]) {
      if (s?.[k] !== undefined && JSON.stringify(s[k]) !== "{}") {
        out[k] = s[k];
      }
    }
    return out;
  };
  wb.eachSheet((ws: Any) => {
    const cells: Record<string, unknown> = {};
    ws.eachRow({ includeEmpty: false }, (row: Any) => {
      row.eachCell({ includeEmpty: true }, (cell: Any) => {
        const value = cell.value instanceof Date ? cell.value.toISOString() : cell.value;
        const note = cell.note === undefined ? undefined : JSON.parse(JSON.stringify(cell.note));
        const entry = { value, type: cell.type, style: style(cell.style), note };
        if (value !== null || Object.keys(entry.style as object).length || note !== undefined) {
          cells[cell.address] = entry;
        }
      });
    });
    const rows: Record<number, unknown> = {};
    ws.eachRow({ includeEmpty: true }, (row: Any, n: number) => {
      if (row.height !== undefined || row.hidden || row.outlineLevel) {
        rows[n] = { height: row.height, hidden: row.hidden, outlineLevel: row.outlineLevel };
      }
    });
    const columns = (ws.columns ?? []).map((c: Any) => ({
      width: c.width,
      hidden: c.hidden || undefined,
      outlineLevel: c.outlineLevel || undefined
    }));
    sheets[ws.name] = {
      cells,
      rows,
      columns,
      merges: Object.keys(ws._merges ?? {}).sort(),
      validations: JSON.parse(JSON.stringify(ws.dataValidations?.model ?? {})),
      conditionalFormats: JSON.parse(JSON.stringify(ws.conditionalFormattings ?? [])),
      tables: Object.values(ws.tables ?? {}).map((t: Any) => ({
        name: t.table?.name ?? t.name,
        ref: t.table?.tableRef ?? t.table?.ref
      })),
      pageSetup: {
        orientation: ws.pageSetup?.orientation,
        paperSize: ws.pageSetup?.paperSize,
        fitToPage: ws.pageSetup?.fitToPage,
        printArea: ws.pageSetup?.printArea
      },
      views: JSON.parse(JSON.stringify(ws.views ?? [])),
      autoFilter: ws.autoFilter,
      state: ws.state
    };
  });
  return { sheets, names: JSON.parse(JSON.stringify(wb.definedNames.model ?? [])) };
}

interface Scenario {
  name: string;
  exceljs(ExcelJS: Any): Promise<Uint8Array>;
  documonster(): Promise<Uint8Array>;
}

async function main(): Promise<void> {
  const work = fs.mkdtempSync(path.join(os.tmpdir(), "documonster-exceljs-parity-"));
  try {
    console.log(`verify:exceljs-parity — installing ${PACKAGE}`);
    fs.writeFileSync(path.join(work, "package.json"), "{}");
    run("npm", ["install", PACKAGE, "--no-audit", "--no-fund", "--silent"], work);
    const require = createRequire(path.join(work, "package.json"));
    // excelts exports its classes by name, and its streaming writer as `WorkbookWriter`: give the scenarios the
    // exceljs shape they are written against.
    const ExcelJS: Any =
      AGAINST === "excelts"
        ? (mod => ({ ...mod, stream: { xlsx: { WorkbookWriter: mod.WorkbookWriter } } }))(
            require("@cj-tech-master/excelts")
          )
        : require("exceljs");
    const { scenarios, csvScenarios } = await import("./exceljs-parity/scenarios.ts");

    const problems: string[] = [];
    for (const s of scenarios as Scenario[]) {
      const [a, b] = await Promise.all([
        s.exceljs(ExcelJS).then(bytes => snapshot(ExcelJS, bytes)),
        s.documonster().then(bytes => snapshot(ExcelJS, bytes))
      ]);
      if (!isDeepStrictEqual(a, b)) {
        problems.push(`${s.name}:\n${describeDiff(a, b)}`);
      }
    }
    for (const s of (AGAINST === "excelts" ? [] : csvScenarios) as {
      name: string;
      exceljs(ExcelJS: Any, readable: (text: string) => Readable): Promise<unknown>;
      documonster(): Promise<unknown>;
    }[]) {
      const a = await s.exceljs(ExcelJS, text => Readable.from([text]));
      const b = await s.documonster();
      if (!isDeepStrictEqual(a, b)) {
        problems.push(
          `${s.name}:\n  exceljs:     ${JSON.stringify(a)}\n  documonster: ${JSON.stringify(b)}`
        );
      }
    }
    const total = scenarios.length + (AGAINST === "excelts" ? 0 : csvScenarios.length);
    if (problems.length) {
      console.error(`✗ verify:exceljs-parity — ${problems.length} of ${total} scenario(s) differ:`);
      for (const p of problems) {
        console.error(`  ${p}`);
      }
      process.exit(1);
    }
    console.log(`✓ verify:exceljs-parity — ${total} scenario(s) produce what ${PACKAGE} produced.`);
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

/** The paths at which two snapshots differ, with both values. */
function describeDiff(a: unknown, b: unknown, at = ""): string {
  if (isDeepStrictEqual(a, b)) {
    return "";
  }
  if (
    a &&
    b &&
    typeof a === "object" &&
    typeof b === "object" &&
    !Array.isArray(a) === !Array.isArray(b)
  ) {
    const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
    return [...keys]
      .map(k =>
        describeDiff(
          (a as Record<string, unknown>)[k],
          (b as Record<string, unknown>)[k],
          `${at}.${k}`
        )
      )
      .filter(Boolean)
      .join("\n");
  }
  return `    ${at || "(root)"}: exceljs ${JSON.stringify(a)} ≠ documonster ${JSON.stringify(b)}`;
}

/* eslint-enable @typescript-eslint/no-explicit-any */

await main();
