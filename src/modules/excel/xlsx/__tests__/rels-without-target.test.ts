/**
 * A workbook relationship with no `Target` names no part. The reader treats it as the missing relationship it is,
 * rather than failing with a `TypeError` while resolving a path out of `undefined`.
 */
import { extractAll } from "@archive/unzip/extract";
import { createZip } from "@archive/zip/zip-bytes";
import { Row, Workbook } from "@excel";
import { describe, expect, it } from "vitest";

describe("xl/_rels/workbook.xml.rels", () => {
  it("tolerates a sheet relationship with no Target", async () => {
    const wb = Workbook.create();
    Row.addMany(Workbook.addWorksheet(wb, "A"), [[1]]);
    Row.addMany(Workbook.addWorksheet(wb, "B"), [[2]]);
    const files = await extractAll(await Workbook.toBuffer(wb));
    const entries = [...files.values()]
      .filter(file => file.type === "file")
      .map(file => {
        if (file.path !== "xl/_rels/workbook.xml.rels") {
          return { name: file.path, data: file.data };
        }
        const xml = new TextDecoder()
          .decode(file.data)
          .replace(/(<Relationship[^>]*Target=)"worksheets\/sheet1\.xml"/, (_, head: string) =>
            head.replace(/\s*Target=$/, "")
          );
        expect(xml).not.toContain("worksheets/sheet1.xml");
        return { name: file.path, data: new TextEncoder().encode(xml) };
      });
    const reopened = Workbook.create();
    await Workbook.read(reopened, await createZip(entries));
    expect(Workbook.getWorksheet(reopened, "B")).toBeDefined();
  });
});
