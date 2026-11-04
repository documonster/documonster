/**
 * Panes and the auto-filter, through the public surface.
 *
 * Both were model fields with no setter. The only way to freeze a header row was
 * `Worksheet.setModel(sheet, { ...Worksheet.getModel(sheet), views: [{ state: "frozen", xSplit: 1, … }] })`,
 * which asks a caller to know three things that are not obvious:
 *
 * - `xSplit` counts **columns** and `ySplit` **rows**, despite reading like coordinates;
 * - for a *split* view the same two fields are twips, not counts;
 * - `topLeftCell` is not an independent choice — it must be the first cell outside the frozen region, or the
 *   sheet scrolls somewhere the author did not ask for.
 *
 * Two examples in this repository did exactly that and each carried a comment naming the member that would
 * close the gap. These tests are what those comments were owed.
 */
import { Workbook, Worksheet } from "@excel";
import { autoFilter } from "@excel/core/worksheet-core";
import { describe, expect, it } from "vitest";

/**
 * The sheet's first view as the file holds it — `xSplit`, `topLeftCell`, `activePane` — which is what `freeze` and
 * `split` are responsible for. (`Worksheet.getPanes` reports the panes in the setters' own words; see
 * `surface/__tests__/r4-contracts.test.ts`.)
 */
function view(ws: Worksheet.Handle): Record<string, unknown> | undefined {
  return ws.views[0] as Record<string, unknown> | undefined;
}

/** A sheet with a little data, so a pane has something to hold. */
function sheet(): Worksheet.Handle {
  const handle = Workbook.create();
  const created = Workbook.addWorksheet(handle, "S");
  Worksheet.addAoa(created, [
    ["Region", "Units", "Cost"],
    ["APAC", 1, 2],
    ["EMEA", 3, 4]
  ]);
  return created;
}

describe("Worksheet.freeze", () => {
  it("holds the given number of columns and rows", () => {
    const target = sheet();
    Worksheet.freeze(target, { columns: 1, rows: 1 });
    expect(view(target)).toMatchObject({ state: "frozen", xSplit: 1, ySplit: 1 });
  });

  it("derives topLeftCell as the first cell outside the frozen region", () => {
    // Not a parameter, because any other value scrolls the sheet somewhere the author did not ask for.
    const target = sheet();
    Worksheet.freeze(target, { columns: 2, rows: 3 });
    expect(view(target)).toMatchObject({ topLeftCell: "C4", activeCell: "C4" });
  });

  it("omits the axis that is not frozen", () => {
    // `xSplit: 0` and "no horizontal freeze" are the same intent and Excel spells the second by omission.
    const target = sheet();
    Worksheet.freeze(target, { rows: 1 });
    const found = view(target) as Record<string, unknown>;
    expect(found.ySplit).toBe(1);
    expect("xSplit" in found).toBe(false);
  });

  it("treats freezing nothing as unfreezing", () => {
    const target = sheet();
    Worksheet.freeze(target, { columns: 1, rows: 1 });
    Worksheet.freeze(target, {});
    expect(view(target)).toBeUndefined();
  });

  it("replaces a previous pane rather than adding one", () => {
    // `views` is a list, and two entries is a workbook with two sheet views — not what a second call means.
    const target = sheet();
    Worksheet.freeze(target, { columns: 1, rows: 1 });
    Worksheet.freeze(target, { columns: 2, rows: 2 });
    expect(view(target)).toMatchObject({ xSplit: 2, ySplit: 2 });
    expect(Worksheet.getModel(target).views).toHaveLength(1);
  });
});

describe("Worksheet.split", () => {
  it("is a split view, not a frozen one", () => {
    // A different record and a different behaviour: a split pane can be dragged, a frozen one cannot.
    const target = sheet();
    Worksheet.split(target, { x: 2000, y: 1000 });
    expect(view(target)).toMatchObject({
      state: "split",
      xSplit: 2000,
      ySplit: 1000
    });
  });

  it.each([
    [2000, 1000, "bottomRight"],
    [2000, 0, "topRight"],
    [0, 1000, "bottomLeft"]
  ])("puts the cursor in the pane Excel does for (%i, %i)", (x, y, pane) => {
    const target = sheet();
    Worksheet.split(target, { x, y });
    expect(view(target)).toMatchObject({ activePane: pane });
  });
});

describe("Worksheet.setAutoFilter", () => {
  it("sets and clears the range", () => {
    const target = sheet();
    Worksheet.setAutoFilter(target, "A1:C3");
    expect(autoFilter(target)).toBe("A1:C3");
    Worksheet.setAutoFilter(target, null);
    expect(autoFilter(target)).toBeNull();
  });

  it("reports null rather than undefined for a sheet with none", () => {
    expect(autoFilter(sheet())).toBeNull();
  });
});

describe("through both containers", () => {
  it.each(["xlsx", "xlsb"] as const)("round-trips a frozen pane through %s", async format => {
    const handle = Workbook.create();
    const created = Workbook.addWorksheet(handle, "S");
    Worksheet.addAoa(created, [
      ["a", "b"],
      [1, 2]
    ]);
    Worksheet.freeze(created, { columns: 1, rows: 1 });

    const back = Workbook.create();
    await Workbook.read(back, await Workbook.toBuffer(handle, { format }));
    expect(view(Workbook.listWorksheets(back)[0]!)).toMatchObject({
      state: "frozen",
      xSplit: 1,
      ySplit: 1,
      topLeftCell: "B2"
    });
  });

  it.each(["xlsx", "xlsb"] as const)("round-trips an auto-filter through %s", async format => {
    const handle = Workbook.create();
    const created = Workbook.addWorksheet(handle, "S");
    Worksheet.addAoa(created, [
      ["a", "b"],
      [1, 2]
    ]);
    Worksheet.setAutoFilter(created, "A1:B2");

    const back = Workbook.create();
    await Workbook.read(back, await Workbook.toBuffer(handle, { format }));
    expect(autoFilter(Workbook.listWorksheets(back)[0]!)).toBe("A1:B2");
  });

  it("agrees between the containers about the pane", async () => {
    // The two readers fill in different amounts of surrounding view state — zoom, gridlines — so the
    // comparison is over the fields `freeze` sets, which are the ones it is responsible for.
    const handle = Workbook.create();
    const created = Workbook.addWorksheet(handle, "S");
    Worksheet.addAoa(created, [
      ["a", "b"],
      [1, 2]
    ]);
    Worksheet.freeze(created, { columns: 2, rows: 1 });

    const read = async (format: "xlsx" | "xlsb") => {
      const back = Workbook.create();
      await Workbook.read(back, await Workbook.toBuffer(handle, { format }));
      const found = view(Workbook.listWorksheets(back)[0]!) as Record<string, unknown>;
      return {
        state: found.state,
        xSplit: found.xSplit,
        ySplit: found.ySplit,
        topLeftCell: found.topLeftCell
      };
    };
    expect(await read("xlsb")).toEqual(await read("xlsx"));
  });
});
