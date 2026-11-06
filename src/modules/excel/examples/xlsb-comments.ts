/**
 * Comments and notes through XLSB, and the record that makes them visible.
 *
 * The comment text lives in `comments{N}.bin` and the *box* that displays it lives in a VML drawing the sheet
 * points at with `BrtLegacyDrawing`. Without that record Excel opens a workbook whose comment part is present
 * and correct and shows nothing — so the interesting question is not whether the text survived but whether the
 * sheet still refers to the drawing.
 *
 * Run: pnpm example --filter xlsb-comments
 */

import { report, reportDropped, reportFiles, writeBoth } from "@excel/examples/utils/both-formats";
import { buildComments, sampleSheet } from "@excel/examples/utils/features";
import { Cell, Workbook } from "@excel/index";

const workbook = Workbook.create();
buildComments(sampleSheet(workbook));

const { results, dropped } = await writeBoth(workbook, "xlsb-comments");

/** The text of what `Cell.getNote` hands back: a plain string, or a note configuration's run list. */
function textOf(value: ReturnType<typeof Cell.getNote>): string | undefined {
  if (value === undefined || typeof value === "string") {
    return value;
  }
  const runs = value.texts;
  return runs === undefined ? undefined : runs.map(run => run.text ?? "").join("");
}

report("comment text", results, reloaded => {
  const sheet = Workbook.getWorksheet(reloaded, "Data");
  return textOf(Cell.getNote(sheet!, "B2")) ?? "no comment";
});

report("note text", results, reloaded => {
  const sheet = Workbook.getWorksheet(reloaded, "Data");
  return textOf(Cell.getNote(sheet!, "B3")) ?? "no note";
});

reportDropped(dropped);
reportFiles(results);
