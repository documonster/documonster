/**
 * ZIP edits validate before they change anything: `ZipEditor.apply(plan)` checks every
 * operation first, so a plan that fails leaves the editor as it was, and every entry source
 * or comment is checked where it is given rather than when the archive is written. Readers
 * hand out copies: changing what they return changes neither the plan nor the reader.
 */
import { Archive, ArchiveError } from "@archive";
import { describe, expect, it } from "vitest";

const zip = Archive.packZipSync([
  ["a.txt", "A"],
  ["b.txt", "B"]
]);

describe("ZipEditor.apply", () => {
  it("applies nothing when an operation of the plan is refused", async () => {
    const editor = await Archive.ZipEditor.open(zip);
    const plan = new Archive.ZipEditPlan()
      .delete("a.txt")
      .set("new.txt", "N")
      .set("../evil", "E")
      .rename("b.txt", "c.txt");
    expect(() => editor.apply(plan)).toThrow(ArchiveError);
    expect(editor.getOutputEntryNames()).toEqual(["a.txt", "b.txt"]);
  });

  it.each<[string, (plan: InstanceType<typeof Archive.ZipEditPlan>) => unknown]>([
    ["an invalid level", plan => plan.set("x", "1", { level: 42 })],
    ["a source that is not one", plan => plan.set("x", 42 as never)],
    ["a comment that is not a string", plan => plan.setComment(42 as never)]
  ])("ZipEditPlan refuses %s when it is given", (_name, call) => {
    expect(() => call(new Archive.ZipEditPlan())).toThrow(ArchiveError);
  });

  it("refuses a rename onto a traversal name before applying anything", async () => {
    const editor = await Archive.ZipEditor.open(zip);
    const plan = new Archive.ZipEditPlan().delete("a.txt").rename("b.txt", "../c.txt");
    expect(() => editor.apply(plan)).toThrow(ArchiveError);
    expect(editor.getOutputEntryNames()).toEqual(["a.txt", "b.txt"]);
  });
});

describe("sources and comments are checked when they are given", () => {
  it.each([42, {}, undefined])("refuses %s as an entry source", async source => {
    expect(() => Archive.createZip().add("a", source as never)).toThrow(ArchiveError);
    expect(() => Archive.createTar().add("a", source as never)).toThrow(ArchiveError);
    expect(() => Archive.createTarGz().add("a", source as never)).toThrow(ArchiveError);
    expect(() => Archive.packZipSync([["a", source as never]])).toThrow(ArchiveError);
    const editor = await Archive.ZipEditor.open(zip);
    expect(() => editor.set("a", source as never)).toThrow(ArchiveError);
  });

  it("TAR keeps null as a directory", () => {
    const bytes = Archive.createTar().add("dir", null).toBufferSync();
    expect([...Archive.unpackTarSync(bytes).keys()]).toEqual(["dir/"]);
  });

  it("refuses a symlink target that is not a string", () => {
    expect(() => Archive.createZip().addSymlink("l", 42 as never)).toThrow(ArchiveError);
    expect(() => Archive.createTar().addSymlink("l", 42 as never)).toThrow(ArchiveError);
  });

  it("refuses an editor comment that is not a string", async () => {
    const editor = await Archive.ZipEditor.open(zip);
    expect(() => editor.setComment(42 as never)).toThrow(ArchiveError);
    expect(() => editor.setComment(undefined)).not.toThrow();
  });
});

describe("readers return copies", () => {
  it("ZipEditPlan.getOperations", async () => {
    const plan = new Archive.ZipEditPlan().set("x", "1");
    const ops = plan.getOperations() as unknown as unknown[];
    expect(() => ops.push({ type: "delete", name: "x" })).toThrow(TypeError);
    expect(plan.getOperations()).toHaveLength(1);
    expect(() => {
      (plan.getOperations()[0] as { name: string }).name = "y";
    }).toThrow(TypeError);
    expect(JSON.parse(await plan.toJson()).ops[0].name).toBe("x");
  });

  it("ZipEditPlan does not share the array it was built from", () => {
    const ops = [{ type: "delete" as const, name: "x" }];
    const plan = new Archive.ZipEditPlan(ops);
    ops.push({ type: "delete", name: "y" });
    expect(plan.getOperations()).toHaveLength(1);
  });

  it("RemoteZipReader.getEntries / getEntry / filterEntries / findEntries", async () => {
    const reader = await Archive.RemoteZipReader.fromReader(Archive.createBufferReader(zip));
    const list = reader.getEntries() as unknown as unknown[];
    expect(() => {
      list.length = 0;
    }).toThrow(TypeError);
    expect(() => {
      (reader.getEntry("a.txt") as { path: string }).path = "zzz";
    }).toThrow(TypeError);
    expect(() => {
      (reader.filterEntries(() => true)[0] as { path: string }).path = "zzz";
    }).toThrow(TypeError);
    expect(reader.getEntries().map(e => e.path)).toEqual(["a.txt", "b.txt"]);
    expect(reader.getFileCount()).toBe(2);
    expect(reader.hasEntry("a.txt")).toBe(true);
    await reader.close();
  });

  it("ZipEditor.getEntries", async () => {
    const editor = await Archive.ZipEditor.open(zip);
    expect(() => {
      (editor.getEntries()[0] as { path: string }).path = "zzz";
    }).toThrow(TypeError);
    expect(editor.getEntries().map(e => e.path)).toEqual(["a.txt", "b.txt"]);
  });
});

describe("unpackZipSync sources", () => {
  it("refuses a source the synchronous form cannot read, naming the async one", () => {
    expect(() => Archive.unpackZipSync(new Blob([zip as BlobPart]) as never)).toThrow(
      /unpackZip\b/
    );
    expect(() => Archive.unpackZipSync(null as never)).toThrow(ArchiveError);
    expect([...Archive.unpackZipSync(zip.slice().buffer).keys()]).toEqual(["a.txt", "b.txt"]);
  });
});
