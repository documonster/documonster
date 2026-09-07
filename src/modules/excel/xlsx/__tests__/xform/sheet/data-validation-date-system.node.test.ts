/**
 * The same date-validation bound, through the streaming writers.
 *
 * Separated from the buffered suite because `Stream.WorkbookWriter` writes to a path: this needs a temporary
 * directory and Node's `fs`, and a browser has neither — importing `node:fs/promises` there fails before a
 * single assertion runs. The buffered half of the claim stays platform-free in the sibling file and is therefore
 * re-run in a real browser.
 */
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { Cell, Stream, Workbook } from "@excel";
import type { WorkbookFormat } from "@excel";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

import { boundSerial, SERIAL_1900, SERIAL_1904, WHEN } from "./data-validation-date-system.helpers";

describe.each<WorkbookFormat>(["xlsx", "xlsb"])(
  "%s date validation bounds, through the streaming writer",
  format => {
    let dir: string;

    beforeAll(async () => {
      dir = await mkdtemp(join(tmpdir(), `documonster-dv-stream-${format}-`));
    });

    afterAll(async () => {
      await rm(dir, { recursive: true, force: true });
    });

    /** One dated cell and one date bound at the same address, written forward. */
    async function streamed(date1904: boolean): Promise<Uint8Array> {
      const path = join(dir, `dv-${String(date1904)}.${format}`);
      const writer = new Stream.WorkbookWriter({
        filename: path,
        format,
        date1904,
        useStyles: true
      });
      const sheet = writer.addWorksheet("S");
      Stream.setCellValue(sheet.getCell("A1"), WHEN);
      (sheet as { dataValidations: { model: Record<string, unknown> } }).dataValidations.model[
        "A1"
      ] = { type: "date", operator: "greaterThanOrEqual", formulae: [WHEN] };
      Stream.commitRow(sheet.getRow(1));
      sheet.commit();
      await writer.commit();
      return Uint8Array.from(await readFile(path));
    }

    it.each([false, true])(
      "puts the bound on the same day as the cell, at date1904=%s",
      async date1904 => {
        // **The streaming writers reached the bound by a different route and were missed.** The buffered path
        // goes through `WorksheetXform.render`, which was given the epoch; streaming XLSX calls the validation
        // xform directly and streaming XLSB assembles its own sheet options, so both were still converting
        // against 1900 while their cells used 1904. The assertion is `bound === cell` inside one file, which
        // no pair of matching mistakes can satisfy.
        const reopened = Workbook.create();
        await Workbook.read(reopened, await streamed(date1904));
        const sheet = Workbook.getWorksheet(reopened, "S")!;
        const rule = Cell.getValidation(sheet, "A1");
        expect(rule?.type).toBe("date");
        const bound = boundSerial(
          (rule as { formulae?: unknown[] }).formulae?.[0] as never,
          date1904
        );
        expect(bound).toBe(date1904 ? SERIAL_1904 : SERIAL_1900);
        expect(Cell.getDateParts(sheet, "A1")).toMatchObject({ year: 2020, month: 1, day: 15 });
      }
    );
  }
);
