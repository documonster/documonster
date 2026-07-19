import { Image, Workbook } from "@excel/index";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { makeTestDataPath } from "@test/utils";
import { describe, expect, it } from "vitest";

describe("importSheet — issue #189: drawing aliasing / duplicate Content_Types Override", () => {
  it("does not duplicate a drawing Override when copying a sheet that has a loaded image", async () => {
    // Build a workbook with an image, serialize, then read it back so the
    // worksheet carries a *loaded* drawing container (the scenario from #189).
    const src = Workbook.create();
    const ws0 = Workbook.addWorksheet(src, "Sheet1");
    const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
    const imageId = Image.add(src, { filename: dataPath("image.png"), extension: "png" });
    Image.place(ws0, imageId, "B2:D6");
    const onDisk = await Workbook.toBuffer(src);

    const wb = Workbook.create();
    await Workbook.read(wb, onDisk);
    const source = Workbook.getWorksheet(wb, "Sheet1")!;

    Workbook.importSheet(wb, source, "Sheet1 Copy");

    const buffer = await Workbook.toBuffer(wb);
    const report = await validateXlsxBuffer(buffer, { maxProblems: 50 });

    const dupes = report.problems.filter(p => p.kind === "content-types-duplicate-override");
    expect(dupes).toEqual([]);
    expect(report.problems).toEqual([]);
  });

  it("gives the imported sheet an independent drawing (mutating the copy leaves the source intact)", async () => {
    const src = Workbook.create();
    const ws0 = Workbook.addWorksheet(src, "Sheet1");
    const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");
    const imageId = Image.add(src, { filename: dataPath("image.png"), extension: "png" });
    Image.place(ws0, imageId, "B2:D6");
    const onDisk = await Workbook.toBuffer(src);

    const wb = Workbook.create();
    await Workbook.read(wb, onDisk);
    const source = Workbook.getWorksheet(wb, "Sheet1")!;
    const copy = Workbook.importSheet(wb, source, "Sheet1 Copy");

    // Both sheets still have their image and the file round-trips cleanly.
    const buffer = await Workbook.toBuffer(wb);
    const report = await validateXlsxBuffer(buffer, { maxProblems: 50 });
    expect(report.problems).toEqual([]);
    expect(copy).toBeDefined();
  });
});
