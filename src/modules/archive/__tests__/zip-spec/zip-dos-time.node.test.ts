/**
 * A ZIP entry's DOS date and time are local wall-clock fields, written from the local fields
 * of `modTime` and read back as local time, so a date in the DOS range round-trips in any
 * time zone. The DOS fields hold 1980–2107 and the Info-ZIP UT field 1970-01-01 to
 * 2106-02-07: a date only the UT field holds gets one even in `"dos"` mode, a date only the
 * DOS fields hold gets none even in `"dos+utc"` mode, and a `modTime` neither holds is
 * refused rather than wrapped.
 */
import { Archive, ArchiveError } from "@archive";
import { afterAll, beforeAll, describe, expect, it } from "vitest";

const originalTz = process.env.TZ;
beforeAll(() => {
  // A zone well away from UTC, so reading local fields as UTC shows.
  process.env.TZ = "Australia/Sydney";
});
afterAll(() => {
  if (originalTz === undefined) {
    delete process.env.TZ;
  } else {
    process.env.TZ = originalTz;
  }
});

const roundTrip = async (modTime: Date, timestamps?: "dos" | "dos+utc") => {
  const bytes = Archive.createZip(timestamps ? { timestamps } : {})
    .add("a.txt", "x", { modTime })
    .toBufferSync();
  const sync = Archive.unpackZipSync(bytes).get("a.txt")!.info.lastModified;
  const remote = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(bytes));
  const fromRemote = remote.getEntry("a.txt")!.lastModified;
  await remote.close();
  return [sync.getTime(), fromRemote.getTime()];
};

describe("ZIP DOS timestamps", () => {
  it.each([
    new Date(2020, 5, 1, 12, 30, 44),
    new Date(1980, 0, 1, 0, 0, 0),
    new Date(2107, 11, 31, 23, 59, 58)
  ])("round-trips %s through the DOS fields", async date => {
    expect(await roundTrip(date)).toEqual([date.getTime(), date.getTime()]);
  });

  it("reads the DOS fields of a streamed entry as local time too", async () => {
    const date = new Date(2020, 5, 1, 12, 30, 44);
    const bytes = Archive.packZipSync([["a.txt", "x"]], { modTime: date });
    for await (const entry of Archive.openZip(
      (async function* () {
        yield bytes;
      })()
    ).entries()) {
      expect(entry.info.lastModified.getTime()).toBe(date.getTime());
    }
  });

  it.each([new Date(Date.UTC(1979, 11, 31, 12)), new Date(Date.UTC(1970, 0, 1))])(
    'stores %s, which only the UT field holds, exactly even in "dos" mode',
    async date => {
      expect(await roundTrip(date)).toEqual([date.getTime(), date.getTime()]);
    }
  );

  it('stores a date only the DOS fields hold without a UT field in "dos+utc" mode', async () => {
    const date = new Date(2107, 0, 1, 10, 0, 0);
    expect(await roundTrip(date, "dos+utc")).toEqual([date.getTime(), date.getTime()]);
  });

  it.each([new Date(Date.UTC(1969, 11, 31)), new Date(2108, 0, 1), new Date(1601, 0, 1)])(
    "refuses %s, which no ZIP field holds",
    date => {
      expect(() => Archive.createZip().add("a.txt", "x", { modTime: date })).toThrow(ArchiveError);
      expect(() => Archive.createZip({ modTime: date })).toThrow(ArchiveError);
      expect(() => Archive.packZipSync([["a.txt", "x"]], { modTime: date })).toThrow(ArchiveError);
      expect(() =>
        Archive.createZip({ timestamps: "dos+utc" }).add("a.txt", "x", { modTime: date })
      ).toThrow(ArchiveError);
    }
  );

  it("refuses an out-of-range modTime in the editor and ArchiveFile too", async () => {
    const zip = Archive.packZipSync([["a.txt", "x"]]);
    const editor = await Archive.editZip(zip);
    const tooEarly = new Date(Date.UTC(1969, 0, 1));
    expect(() => editor.set("b.txt", "y", { modTime: tooEarly })).toThrow(ArchiveError);
    expect(() => new Archive.ArchiveFile().addText("y", "b.txt", { modTime: tooEarly })).toThrow(
      ArchiveError
    );
    expect(() => new Archive.ArchiveFile({ modTime: tooEarly })).toThrow(ArchiveError);
  });
});

describe("ZIP DOS timestamps across a DST change", () => {
  it.each([
    ["the first 01:30 of the repeated hour", "2024-11-03T05:30:00Z"],
    ["the second 01:30 of the repeated hour", "2024-11-03T06:30:00Z"],
    ["the hour after the spring-forward gap", "2024-03-10T07:30:00Z"]
  ])('round-trips %s in "dos" mode', async (_name, iso) => {
    process.env.TZ = "America/New_York";
    const date = new Date(iso);
    expect(await roundTrip(date)).toEqual([date.getTime(), date.getTime()]);
  });

  it("adds no UT field to a date the DOS fields hold unambiguously", () => {
    process.env.TZ = "America/New_York";
    const bytes = Archive.createZip()
      .add("a.txt", "x", { modTime: new Date(2024, 5, 1, 12, 0, 0) })
      .toBufferSync();
    expect(Archive.unpackZipSync(bytes).get("a.txt")!.info.extraField?.length ?? 0).toBe(0);
  });
});
