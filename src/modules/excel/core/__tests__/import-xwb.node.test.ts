import { getImages } from "@excel/core/worksheet";
import { Image, Workbook } from "@excel/index";
import { validateXlsxBuffer } from "@excel/utils/ooxml-validator";
import { makeTestDataPath } from "@test/utils";
import { describe, it, expect } from "vitest";

describe("importSheet across workbooks with an image", () => {
  it("copies an image sheet into a different workbook without duplicate/invalid parts", async () => {
    const dataPath = makeTestDataPath(import.meta.url, "../../__tests__/data");

    const src = Workbook.create();
    const ws = Workbook.addWorksheet(src, "WithImage");
    const imageId = Image.add(src, { filename: dataPath("image.png"), extension: "png" });
    Image.place(ws, imageId, "B2:D6");
    const onDisk = await Workbook.toBuffer(src);

    const source = Workbook.create();
    await Workbook.read(source, onDisk);
    const srcSheet = Workbook.getWorksheet(source, "WithImage")!;

    // Import into a DIFFERENT workbook.
    const dest = Workbook.create();
    Workbook.addWorksheet(dest, "Existing");
    Workbook.importSheet(dest, srcSheet, "Imported");

    const imported = Workbook.getWorksheet(dest, "Imported")!;
    expect(getImages(imported)).toHaveLength(1);
    expect(dest.media).toHaveLength(1);

    const out = await Workbook.toBuffer(dest);
    const report = await validateXlsxBuffer(out, { maxProblems: 50 });
    expect(report.problems).toEqual([]);

    const roundTripped = Workbook.create();
    await Workbook.read(roundTripped, out);
    const outputImages = getImages(Workbook.getWorksheet(roundTripped, "Imported")!);
    expect(outputImages).toHaveLength(1);
    expect(roundTripped.media).toHaveLength(1);
  });
});
