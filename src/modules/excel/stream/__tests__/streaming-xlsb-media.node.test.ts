/**
 * Pictures written through the *streaming* XLSB writer reach the sheet.
 *
 * A background is a relationship named by `BrtBkHim`, and a placed picture a drawing named by `BrtDrawing` — both
 * records live **inside** the sheet part, which a streamed sheet closes before the package allocates relationships. The
 * streamed sheet therefore declares the ids and the package honours them, as it already did for the notes' VML. Before
 * that, the medium was written (as `image0.png.png`, the streaming writer's name already carrying its extension) and
 * nothing in the sheet pointed at it, and a placed picture beside a note got a drawing relationship sharing `rId1` with
 * the VML.
 */
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { extractAll } from "@archive/unzip/extract";
import { Stream, Workbook } from "@excel";
import { iterateInterpretableRecords } from "@excel/xlsb/binary";
import { recordSpec } from "@excel/xlsb/spec/records";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const PNG = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
  ),
  c => c.charCodeAt(0)
);

let dir: string;

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), "documonster-stream-media-"));
});

afterAll(async () => {
  await rm(dir, { recursive: true, force: true });
});

async function streamed(
  name: string,
  build: (
    writer: Stream.WorkbookWriter,
    sheet: ReturnType<Stream.WorkbookWriter["addWorksheet"]>,
    image: number
  ) => void
): Promise<Uint8Array> {
  const path = join(dir, `${name}.xlsb`);
  const writer = new Stream.WorkbookWriter({ filename: path, format: "xlsb" });
  const image = writer.addImage({ buffer: PNG, extension: "png" });
  const sheet = writer.addWorksheet("S");
  build(writer, sheet, image);
  sheet.commit();
  await writer.commit();
  return Uint8Array.from(await readFile(path));
}

/** The relationship id each named record carries, keyed by record name. */
function namedRelationships(sheet: Uint8Array, path: string): Map<string, string> {
  const named = new Map<string, string>();
  for (const entry of iterateInterpretableRecords(sheet, path)) {
    const recordName = recordSpec(entry.id)?.name;
    if (
      recordName === "BrtDrawing" ||
      recordName === "BrtLegacyDrawing" ||
      recordName === "BrtBkHim"
    ) {
      named.set(recordName, new TextDecoder("utf-16le").decode(entry.payload.subarray(4)));
    }
  }
  return named;
}

function relationships(rels: string): Map<string, string> {
  return new Map([...rels.matchAll(/Id="(\w+)"[^>]*Target="([^"]+)"/g)].map(m => [m[1]!, m[2]!]));
}

describe("pictures through the streaming XLSB writer", () => {
  it("names a background image from the sheet, with a relationship to a medium the package holds", async () => {
    const bytes = await streamed("background", (_, sheet, image) =>
      sheet.addBackgroundImage(image)
    );
    const parts = await extractAll(bytes);
    const path = "xl/worksheets/sheet1.bin";
    const named = namedRelationships(parts.get(path)!.data, path);
    expect(named.has("BrtBkHim"), "the sheet must carry BrtBkHim").toBe(true);
    const rels = relationships(
      new TextDecoder().decode(parts.get("xl/worksheets/_rels/sheet1.bin.rels")!.data)
    );
    expect(rels.get(named.get("BrtBkHim")!)).toBe("../media/image0.png");
    expect([...parts.keys()]).toContain("xl/media/image0.png");
    expect([...parts.keys()].some(key => key.endsWith(".png.png"))).toBe(false);

    const reopened = Workbook.create();
    await Workbook.read(reopened, bytes);
    const media = Workbook.getModel(reopened).worksheets[0]!.media ?? [];
    expect(media.map(medium => medium.type)).toEqual(["background"]);
  });

  it("names a placed picture's drawing from the sheet, beside a note, with distinct ids", async () => {
    const bytes = await streamed("picture-and-note", (_, sheet, image) => {
      sheet.addImage(image, "B2:D6");
      Stream.setCellValue(sheet.getCell("A1"), "x");
      Stream.setCellNote(sheet.getCell("A1"), "note");
      Stream.commitRow(sheet.getRow(1));
    });
    const parts = await extractAll(bytes);
    const path = "xl/worksheets/sheet1.bin";
    const named = namedRelationships(parts.get(path)!.data, path);
    const relsXml = new TextDecoder().decode(
      parts.get("xl/worksheets/_rels/sheet1.bin.rels")!.data
    );
    const ids = [...relsXml.matchAll(/Id="(\w+)"/g)].map(m => m[1]);
    expect(new Set(ids).size, relsXml).toBe(ids.length);
    const rels = relationships(relsXml);
    expect(rels.get(named.get("BrtDrawing")!)).toBe("../drawings/drawing1.xml");
    expect(rels.get(named.get("BrtLegacyDrawing")!)).toBe("../drawings/vmlDrawing1.vml");

    const reopened = Workbook.create();
    await Workbook.read(reopened, bytes);
    const media = Workbook.getModel(reopened).worksheets[0]!.media ?? [];
    expect(media.map(medium => medium.type)).toEqual(["image"]);
  });
});
