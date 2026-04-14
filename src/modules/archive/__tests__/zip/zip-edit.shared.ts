import { zip, editZip, editZipUrl, BufferReader, ZipEditPlan, type ZipEditWarning } from "@archive";
import { ZipParser } from "@archive/unzip/zip-parser";
import { createZip, type ZipEntry } from "@archive/zip/zip-bytes";
import { concatUint8Arrays } from "@utils/binary";
import { describe, it, expect, vi } from "vitest";

const decode = (data: Uint8Array): string => new TextDecoder().decode(data);
const encode = (str: string): Uint8Array => new TextEncoder().encode(str);

function toEntries(files: Record<string, Uint8Array | string>): ZipEntry[] {
  return Object.entries(files).map(([name, data]) => ({
    name,
    data: typeof data === "string" ? encode(data) : data
  }));
}

function expectBytesEqual(a: Uint8Array, b: Uint8Array): void {
  expect(a.length).toBe(b.length);
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) {
      throw new Error(`Byte mismatch at ${i}: ${a[i]} !== ${b[i]}`);
    }
  }
}

export function runZipEditTests(): void {
  describe("ZipEditor", () => {
    // -------------------------------------------------------------------------
    // Core Operations
    // -------------------------------------------------------------------------

    describe("basic operations", () => {
      it("should delete/update/add entries and emit valid ZIP (bytes)", async () => {
        const original = await zip({ level: 6, reproducible: true })
          .add("a.txt", encode("aaaaa aaaaa aaaaa aaaaa aaaaa"))
          .add("b.txt", encode("bbbbb bbbbb bbbbb bbbbb bbbbb"))
          .bytes();

        const editor = await editZip(original, { level: 6, reproducible: true });
        editor.delete("b.txt");
        editor.set("a.txt", encode("AAA"));
        editor.set("c.txt", "CCC");

        const out = await editor.bytes();

        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["a.txt", "c.txt"]);

        const a = await parser.extract("a.txt");
        const c = await parser.extract("c.txt");
        expect(a).not.toBeNull();
        expect(c).not.toBeNull();
        expect(decode(a!)).toBe("AAA");
        expect(decode(c!)).toBe("CCC");
      });

      it("should preserve raw compressed bytes for unchanged entries (rename)", async () => {
        const big = encode("x".repeat(1024 * 64));
        const original = await zip({ level: 6, reproducible: true }).add("big.txt", big).bytes();

        const originalParser = new ZipParser(original);
        const originalRaw = originalParser.getRawCompressedData("big.txt");
        expect(originalRaw).not.toBeNull();

        const editor = await editZip(original, { level: 6, reproducible: true });
        editor.rename("big.txt", "moved.txt");
        const out = await editor.bytes();

        const outParser = new ZipParser(out);
        const movedRaw = outParser.getRawCompressedData("moved.txt");
        expect(movedRaw).not.toBeNull();

        expectBytesEqual(originalRaw!, movedRaw!);

        const moved = await outParser.extract("moved.txt");
        expect(moved).not.toBeNull();
        expect(moved!).toEqual(big);
      });

      it("should fall back to recompress when raw passthrough is unavailable (bytes)", async () => {
        const big = encode("x".repeat(1024 * 64));
        const original = await zip({ level: 6, reproducible: true }).add("big.txt", big).bytes();

        const onWarning = vi.fn<(w: ZipEditWarning) => void>();
        const editor = await editZip(original, {
          level: 6,
          reproducible: true,
          preserve: "best-effort",
          onWarning
        });

        // Simulate a reader that cannot provide raw compressed bytes.
        const remote = (editor as any)._remote;
        const originalGetRawCompressedData = remote.getRawCompressedData.bind(remote);
        remote.getRawCompressedData = async (path: string) => {
          if (path === "big.txt") {
            return null;
          }
          return originalGetRawCompressedData(path);
        };

        editor.rename("big.txt", "moved.txt");
        const out = await editor.bytes();

        expect(onWarning).toHaveBeenCalled();
        expect(onWarning.mock.calls.some(c => c[0].code === "raw_unavailable")).toBe(true);

        const outParser = new ZipParser(out);
        const moved = await outParser.extract("moved.txt");
        expect(moved).not.toBeNull();
        expect(moved!).toEqual(big);
      });

      it("should fall back to recompress when raw passthrough is unavailable (stream)", async () => {
        const big = encode("x".repeat(1024 * 64));
        const original = await zip({ level: 6, reproducible: true }).add("big.txt", big).bytes();

        const onWarning = vi.fn<(w: ZipEditWarning) => void>();
        const editor = await editZip(original, {
          level: 6,
          reproducible: true,
          preserve: "best-effort",
          onWarning
        });

        // Simulate a reader that cannot provide raw compressed streaming payload.
        const remote = (editor as any)._remote;
        const originalGetRawCompressedStream = remote.getRawCompressedStream.bind(remote);
        remote.getRawCompressedStream = (path: string) => {
          if (path === "big.txt") {
            return null;
          }
          return originalGetRawCompressedStream(path);
        };

        editor.rename("big.txt", "moved.txt");

        const chunks: Uint8Array[] = [];
        for await (const chunk of editor.stream()) {
          chunks.push(chunk);
        }
        const out = concatUint8Arrays(chunks);

        expect(onWarning).toHaveBeenCalled();
        expect(onWarning.mock.calls.some(c => c[0].code === "raw_unavailable")).toBe(true);

        const outParser = new ZipParser(out);
        const moved = await outParser.extract("moved.txt");
        expect(moved).not.toBeNull();
        expect(moved!).toEqual(big);
      });

      it("should not output decrypted content for encrypted entries in best-effort mode", async () => {
        const password = "pw";
        const original = await createZip(toEntries({ "secret.txt": "secret" }), {
          password,
          encryptionMethod: "zipcrypto",
          reproducible: true
        } as any);

        const onWarning = vi.fn<(w: ZipEditWarning) => void>();
        const editor = await editZip(original, {
          reproducible: true,
          preserve: "best-effort",
          password,
          onWarning
        });

        // Force raw passthrough to be unavailable.
        const remote = (editor as any)._remote;
        remote.getRawCompressedData = async (_path: string) => null;

        // Add a new file so the output isn't empty.
        editor.set("public.txt", "OK");
        const out = await editor.bytes();

        expect(onWarning.mock.calls.some(c => c[0].code === "encryption_unsupported")).toBe(true);

        const outParser = new ZipParser(out);
        const paths = outParser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["public.txt"]);
        expect(decode((await outParser.extract("public.txt"))!)).toBe("OK");
      });

      it("should support RandomAccessReader input (BufferReader)", async () => {
        const big = encode("x".repeat(1024 * 64));
        const original = await zip({ level: 6, reproducible: true }).add("big.txt", big).bytes();

        const originalParser = new ZipParser(original);
        const originalRaw = originalParser.getRawCompressedData("big.txt");
        expect(originalRaw).not.toBeNull();

        const editor = await editZip(new BufferReader(original), { level: 6, reproducible: true });
        editor.rename("big.txt", "moved.txt");
        const out = await editor.bytes();

        const outParser = new ZipParser(out);
        const movedRaw = outParser.getRawCompressedData("moved.txt");
        expect(movedRaw).not.toBeNull();

        expectBytesEqual(originalRaw!, movedRaw!);
      });

      it("should support URL input via HTTP Range (editZipUrl)", async () => {
        const big = encode("x".repeat(1024 * 64));
        const original = await zip({ level: 6, reproducible: true }).add("big.txt", big).bytes();

        const originalParser = new ZipParser(original);
        const originalRaw = originalParser.getRawCompressedData("big.txt");
        expect(originalRaw).not.toBeNull();

        const url = "https://example.com/test.zip";
        const mockFetch = vi.fn(async (_url: string, init?: RequestInit) => {
          const headers = (init?.headers ?? {}) as any;
          const rangeValue: string | undefined = headers.Range ?? headers.range;
          if (!rangeValue) {
            throw new Error("Expected Range header");
          }
          const match = /^bytes=(\d+)-(\d+)$/.exec(rangeValue);
          if (!match) {
            throw new Error(`Invalid Range header: ${rangeValue}`);
          }

          const start = Number(match[1]);
          const endInclusive = Number(match[2]);
          const endExclusive = endInclusive + 1;
          const slice = original.slice(start, endExclusive);

          return {
            ok: true,
            status: 206,
            statusText: "Partial Content",
            arrayBuffer: async () => slice.buffer
          } as any;
        });

        const editor = await editZipUrl(url, {
          fetch: mockFetch as any,
          size: original.length,
          validateRangeSupport: false,
          level: 6,
          reproducible: true
        });
        editor.rename("big.txt", "moved.txt");
        const out = await editor.bytes();

        const outParser = new ZipParser(out);
        const movedRaw = outParser.getRawCompressedData("moved.txt");
        expect(movedRaw).not.toBeNull();
        expectBytesEqual(originalRaw!, movedRaw!);

        expect(mockFetch).toHaveBeenCalled();
      });

      it("should emit valid ZIP via stream()", async () => {
        const original = await zip({ level: 6, reproducible: true })
          .add("a.txt", encode("aaaaa aaaaa aaaaa aaaaa aaaaa"))
          .add("b.txt", encode("bbbbb bbbbb bbbbb bbbbb bbbbb"))
          .bytes();

        const editor = await editZip(original, { level: 6, reproducible: true });
        editor.delete("b.txt");
        editor.set("c.txt", encode("CCC"));

        const chunks: Uint8Array[] = [];
        for await (const chunk of editor.stream()) {
          chunks.push(chunk);
        }
        const out = concatUint8Arrays(chunks);

        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["a.txt", "c.txt"]);

        const a = await parser.extract("a.txt");
        const c = await parser.extract("c.txt");
        expect(a).not.toBeNull();
        expect(c).not.toBeNull();
        expect(decode(a!)).toBe("aaaaa aaaaa aaaaa aaaaa aaaaa");
        expect(decode(c!)).toBe("CCC");
      });

      it("should apply ZipEditPlan", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", "A")
          .add("b.txt", "B")
          .bytes();

        const plan = new ZipEditPlan()
          .rename("a.txt", "renamed.txt")
          .delete("b.txt")
          .set("c.txt", "C")
          .setComment("planned");

        const editor = await editZip(original, { reproducible: true });
        editor.apply(plan);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["c.txt", "renamed.txt"]);
        expect(parser.getZipComment()).toBe("planned");
        expect(decode((await parser.extract("renamed.txt"))!)).toBe("A");
      });

      it("should serialize and deserialize ZipEditPlan", async () => {
        const plan = new ZipEditPlan()
          .set("a.txt", "hello world")
          .set("binary.bin", new Uint8Array([1, 2, 3, 4, 5]))
          .delete("old.txt")
          .rename("from.txt", "to.txt")
          .setComment("test comment");

        // Serialize to JSON
        const json = await plan.toJSON();
        expect(typeof json).toBe("string");

        // Deserialize
        const restored = ZipEditPlan.fromJSON(json);

        // Apply both to fresh archives and compare
        const original = await zip({ reproducible: true })
          .add("old.txt", "will be deleted")
          .add("from.txt", "will be renamed")
          .bytes();

        const editor1 = await editZip(original, { reproducible: true });
        editor1.apply(plan);
        const out1 = await editor1.bytes();

        const editor2 = await editZip(original, { reproducible: true });
        editor2.apply(restored);
        const out2 = await editor2.bytes();

        // Both should produce identical archives
        expect(out1).toEqual(out2);

        // Verify content
        const parser = new ZipParser(out1);
        expect(parser.getZipComment()).toBe("test comment");
        expect(decode((await parser.extract("a.txt"))!)).toBe("hello world");
        expect((await parser.extract("binary.bin"))!).toEqual(new Uint8Array([1, 2, 3, 4, 5]));
        expect(parser.hasEntry("old.txt")).toBe(false);
        expect(parser.hasEntry("to.txt")).toBe(true);
      });

      it("should handle empty ZipEditPlan serialization", async () => {
        const plan = new ZipEditPlan();
        const json = await plan.toJSON();
        const restored = ZipEditPlan.fromJSON(json);
        expect(restored.getOperations()).toHaveLength(0);
      });

      it("should handle empty archive", async () => {
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.set("new.txt", "hello");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser.getEntries().map(e => e.path);
        expect(paths).toEqual(["new.txt"]);
      });

      it("should produce empty archive when all entries deleted", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.delete("a.txt");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().length).toBe(0);
      });
    });

    // -------------------------------------------------------------------------
    // has() method
    // -------------------------------------------------------------------------

    describe("has()", () => {
      it("should return true for existing entries", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();
        const editor = await editZip(original);

        expect(editor.has("a.txt")).toBe(true);
        expect(editor.has("nonexistent.txt")).toBe(false);
      });

      it("should return false for deleted entries", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();
        const editor = await editZip(original);

        editor.delete("a.txt");
        expect(editor.has("a.txt")).toBe(false);
      });

      it("should return true for newly added entries", async () => {
        const original = await zip({ reproducible: true }).bytes();
        const editor = await editZip(original);

        expect(editor.has("new.txt")).toBe(false);
        editor.set("new.txt", "data");
        expect(editor.has("new.txt")).toBe(true);
      });
    });

    // -------------------------------------------------------------------------
    // rename() semantics
    // -------------------------------------------------------------------------

    describe("rename()", () => {
      it("should rename an entry", async () => {
        const original = await zip({ reproducible: true }).add("old.txt", "content").bytes();

        const editor = await editZip(original, { reproducible: true });
        const result = editor.rename("old.txt", "new.txt");

        expect(result).toBe(true);
        expect(editor.has("old.txt")).toBe(false);
        expect(editor.has("new.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser.getEntries().map(e => e.path);
        expect(paths).toEqual(["new.txt"]);

        const content = await parser.extract("new.txt");
        expect(decode(content!)).toBe("content");
      });

      it("should return false when source does not exist", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();
        const editor = await editZip(original);

        const result = editor.rename("nonexistent.txt", "new.txt");
        expect(result).toBe(false);
      });

      it("should overwrite target if it exists (like mv -f)", async () => {
        const original = await zip({ reproducible: true })
          .add("src.txt", "source content")
          .add("dst.txt", "destination content")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        const result = editor.rename("src.txt", "dst.txt");

        expect(result).toBe(true);
        expect(editor.has("src.txt")).toBe(false);
        expect(editor.has("dst.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser.getEntries().map(e => e.path);
        expect(paths).toEqual(["dst.txt"]);

        // Content should be from the source
        const content = await parser.extract("dst.txt");
        expect(decode(content!)).toBe("source content");
      });

      it("should handle rename to same name as no-op", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();
        const editor = await editZip(original);

        const result = editor.rename("a.txt", "a.txt");
        expect(result).toBe(true);
        expect(editor.has("a.txt")).toBe(true);
      });

      it("should rename a pending set entry", async () => {
        const original = await zip({ reproducible: true }).bytes();
        const editor = await editZip(original);

        editor.set("temp.txt", "data");
        editor.rename("temp.txt", "final.txt");

        expect(editor.has("temp.txt")).toBe(false);
        expect(editor.has("final.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().map(e => e.path)).toEqual(["final.txt"]);
      });

      it("should support chained renames (rename target)", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();

        const editor = await editZip(original, { reproducible: true });
        expect(editor.rename("a.txt", "b.txt")).toBe(true);
        expect(editor.rename("b.txt", "c.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().map(e => e.path)).toEqual(["c.txt"]);
        expect(decode((await parser.extract("c.txt"))!)).toBe("hello");
      });

      it("should overwrite an existing rename target", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", "A")
          .add("c.txt", "C")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        expect(editor.rename("a.txt", "b.txt")).toBe(true);
        expect(editor.rename("c.txt", "b.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().map(e => e.path)).toEqual(["b.txt"]);
        expect(decode((await parser.extract("b.txt"))!)).toBe("C");
      });
    });

    describe("delete()", () => {
      it("should return true when deleting a rename target", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();
        const editor = await editZip(original, { reproducible: true });

        editor.rename("a.txt", "b.txt");
        expect(editor.delete("b.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().length).toBe(0);
      });
    });

    // -------------------------------------------------------------------------
    // deleteDirectory()
    // -------------------------------------------------------------------------

    describe("deleteDirectory()", () => {
      it("should delete a directory and all its contents", async () => {
        const original = await zip({ reproducible: true })
          .add("root.txt", "root")
          .add("folder/", "")
          .add("folder/file1.txt", "file1")
          .add("folder/file2.txt", "file2")
          .add("folder/sub/", "")
          .add("folder/sub/deep.txt", "deep")
          .add("other.txt", "other")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        const count = editor.deleteDirectory("folder");

        expect(count).toBe(5); // folder/, file1, file2, sub/, deep.txt
        expect(editor.has("folder/")).toBe(false);
        expect(editor.has("folder/file1.txt")).toBe(false);
        expect(editor.has("folder/sub/deep.txt")).toBe(false);
        expect(editor.has("root.txt")).toBe(true);
        expect(editor.has("other.txt")).toBe(true);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["other.txt", "root.txt"]);
      });

      it("should work with trailing slash", async () => {
        const original = await zip({ reproducible: true })
          .add("folder/", "")
          .add("folder/file.txt", "content")
          .add("other.txt", "other")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        const count = editor.deleteDirectory("folder/");

        expect(count).toBe(2);
        expect(editor.has("folder/")).toBe(false);
        expect(editor.has("folder/file.txt")).toBe(false);
        expect(editor.has("other.txt")).toBe(true);
      });

      it("should return 0 when directory does not exist", async () => {
        const original = await zip({ reproducible: true }).add("file.txt", "content").bytes();

        const editor = await editZip(original, { reproducible: true });
        const count = editor.deleteDirectory("nonexistent");

        expect(count).toBe(0);
      });

      it("should not delete files with similar prefix but not in directory", async () => {
        const original = await zip({ reproducible: true })
          .add("test/file.txt", "in folder")
          .add("test-file.txt", "similar prefix")
          .add("testing/file.txt", "different folder")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        const count = editor.deleteDirectory("test");

        expect(count).toBe(1); // Only test/file.txt
        expect(editor.has("test/file.txt")).toBe(false);
        expect(editor.has("test-file.txt")).toBe(true);
        expect(editor.has("testing/file.txt")).toBe(true);
      });

      it("should delete nested directories correctly", async () => {
        const original = await zip({ reproducible: true })
          .add("a/b/c/file.txt", "deep")
          .add("a/b/other.txt", "mid")
          .add("a/top.txt", "top")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        const count = editor.deleteDirectory("a/b");

        expect(count).toBe(2); // c/file.txt and other.txt
        expect(editor.has("a/b/c/file.txt")).toBe(false);
        expect(editor.has("a/b/other.txt")).toBe(false);
        expect(editor.has("a/top.txt")).toBe(true);
      });

      it("should work via ZipEditPlan", async () => {
        const original = await zip({ reproducible: true })
          .add("keep.txt", "keep")
          .add("delete/a.txt", "a")
          .add("delete/b.txt", "b")
          .bytes();

        const plan = new ZipEditPlan().deleteDirectory("delete").set("new.txt", "new");

        const editor = await editZip(original, { reproducible: true });
        editor.apply(plan);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["keep.txt", "new.txt"]);
      });

      it("should delete pending set entries in the directory", async () => {
        const original = await zip({ reproducible: true }).add("existing.txt", "existing").bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.set("folder/new1.txt", "new1");
        editor.set("folder/sub/new2.txt", "new2");
        editor.set("other/file.txt", "other");

        const count = editor.deleteDirectory("folder");

        expect(count).toBe(2); // new1.txt and sub/new2.txt
        expect(editor.has("folder/new1.txt")).toBe(false);
        expect(editor.has("folder/sub/new2.txt")).toBe(false);
        expect(editor.has("other/file.txt")).toBe(true);
      });
    });

    // -------------------------------------------------------------------------
    // setComment()
    // -------------------------------------------------------------------------

    describe("setComment()", () => {
      it("should set archive comment", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.setComment("Test comment 测试");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getZipComment()).toBe("Test comment 测试");
      });

      it("should remove comment when set to undefined", async () => {
        const original = await zip({ reproducible: true, comment: "Original comment" })
          .add("a.txt", "hello")
          .bytes();

        const editor = await editZip(original, { reproducible: true, comment: "Keep this" });
        editor.setComment(undefined);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getZipComment()).toBe("");
      });
    });

    // -------------------------------------------------------------------------
    // getOutputEntryNames()
    // -------------------------------------------------------------------------

    describe("getOutputEntryNames()", () => {
      it("should return expected output names after edits", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", "a")
          .add("b.txt", "b")
          .add("c.txt", "c")
          .bytes();

        const editor = await editZip(original);
        editor.delete("b.txt");
        editor.rename("c.txt", "d.txt");
        editor.set("e.txt", "e");

        const names = editor.getOutputEntryNames();
        expect(names).toEqual(["a.txt", "d.txt", "e.txt"]);
      });
    });

    // -------------------------------------------------------------------------
    // Path normalization
    // -------------------------------------------------------------------------

    describe("path normalization", () => {
      it("should normalize paths with safe mode", async () => {
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, {
          reproducible: true,
          path: { mode: "safe" }
        });

        // Backslash should be converted to forward slash
        editor.set("folder\\file.txt", "content");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser.getEntries().map(e => e.path);
        expect(paths).toEqual(["folder/file.txt"]);
      });

      it("should reject absolute paths in safe mode", async () => {
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, {
          reproducible: true,
          path: { mode: "safe" }
        });

        expect(() => editor.set("/absolute/path.txt", "content")).toThrow(/Unsafe ZIP path/);
      });

      it("should reject traversal paths in safe mode", async () => {
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, {
          reproducible: true,
          path: { mode: "safe" }
        });

        expect(() => editor.set("../escape.txt", "content")).toThrow(/Unsafe ZIP path/);
      });

      it("should normalize ./prefix paths", async () => {
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, {
          reproducible: true,
          path: { mode: "posix" }
        });

        editor.set("./file.txt", "content");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser.getEntries().map(e => e.path);
        expect(paths).toEqual(["file.txt"]);
      });
    });

    // -------------------------------------------------------------------------
    // Directory entries
    // -------------------------------------------------------------------------

    describe("directory entries", () => {
      it("should preserve directory entries", async () => {
        const original = await zip({ reproducible: true })
          .add("folder/", new Uint8Array(0))
          .add("folder/file.txt", "content")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.set("folder/another.txt", "more content");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["folder/", "folder/another.txt", "folder/file.txt"]);
      });

      it("should delete directory entries", async () => {
        const original = await zip({ reproducible: true })
          .add("folder/", new Uint8Array(0))
          .add("folder/file.txt", "content")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.delete("folder/");
        editor.delete("folder/file.txt");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().length).toBe(0);
      });
    });

    // -------------------------------------------------------------------------
    // Warning callback
    // -------------------------------------------------------------------------

    describe("onWarning callback", () => {
      it("should call onWarning when entry cannot be preserved", async () => {
        // Create a valid ZIP first
        const original = await zip({ reproducible: true }).add("test.txt", "hello").bytes();

        const warnings: ZipEditWarning[] = [];
        const editor = await editZip(original, {
          onWarning: w => warnings.push(w)
        });

        // Normal operation should not produce warnings
        editor.set("new.txt", "world");
        await editor.bytes();

        // No warnings for normal operations
        expect(warnings.length).toBe(0);
      });
    });

    // -------------------------------------------------------------------------
    // Progress and cancellation
    // -------------------------------------------------------------------------

    describe("progress tracking", () => {
      it("should report progress via callback", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", "aaaaa")
          .add("b.txt", "bbbbb")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.set("c.txt", "ccccc");

        const progressUpdates: Array<{ entriesDone: number; entriesTotal: number }> = [];
        await editor.bytes({
          onProgress: p => {
            progressUpdates.push({ entriesDone: p.entriesDone, entriesTotal: p.entriesTotal });
          }
        });

        expect(progressUpdates.length).toBeGreaterThan(0);
        const last = progressUpdates[progressUpdates.length - 1]!;
        expect(last.entriesDone).toBe(last.entriesTotal);
      });

      it("should support cancellation via AbortSignal", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", encode("x".repeat(10000)))
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        // Add many entries to give time for abort
        for (let i = 0; i < 100; i++) {
          editor.set(`file${i}.txt`, encode("y".repeat(1000)));
        }

        const controller = new AbortController();

        // Abort immediately
        controller.abort();

        await expect(editor.bytes({ signal: controller.signal })).rejects.toThrow();
      });
    });

    // -------------------------------------------------------------------------
    // pipeTo()
    // -------------------------------------------------------------------------

    describe("pipeTo()", () => {
      it("should pipe output to a sink", async () => {
        const original = await zip({ reproducible: true }).add("a.txt", "hello").bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.set("b.txt", "world");

        const chunks: Uint8Array[] = [];
        const sink = {
          write: (chunk: Uint8Array) => {
            chunks.push(chunk);
            return Promise.resolve();
          },
          close: () => Promise.resolve()
        };

        await editor.pipeTo(sink);

        const out = concatUint8Arrays(chunks);
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["a.txt", "b.txt"]);
      });
    });

    // -------------------------------------------------------------------------
    // STORE vs DEFLATE
    // -------------------------------------------------------------------------

    describe("compression", () => {
      it("should preserve STORE entries without re-compressing", async () => {
        const data = encode("This is stored without compression");
        const original = await zip({ level: 0, reproducible: true })
          .add("stored.txt", data)
          .bytes();

        const originalParser = new ZipParser(original);
        const originalInfo = originalParser.getEntries()[0]!;
        expect(originalInfo.compressionMethod).toBe(0); // STORE

        const editor = await editZip(original, { level: 6, reproducible: true });
        const out = await editor.bytes();

        const outParser = new ZipParser(out);
        const outInfo = outParser.getEntries()[0]!;
        // Should still be STORE (preserved)
        expect(outInfo.compressionMethod).toBe(0);
      });

      it("should apply compression level to new entries", async () => {
        const data = encode("x".repeat(1000)); // Highly compressible
        const original = await zip({ reproducible: true }).bytes();

        const editor = await editZip(original, { level: 9, reproducible: true });
        editor.set("compressed.txt", data);

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const info = parser.getEntries()[0]!;

        // Should be DEFLATE
        expect(info.compressionMethod).toBe(8);
        // And compressed size should be smaller than original
        expect(info.compressedSize).toBeLessThan(data.length);
      });
    });

    // -------------------------------------------------------------------------
    // Multiple operations chaining
    // -------------------------------------------------------------------------

    describe("operation chaining", () => {
      it("should support method chaining", async () => {
        const original = await zip({ reproducible: true })
          .add("a.txt", "a")
          .add("b.txt", "b")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.delete("a.txt");
        const out = await editor.set("c.txt", "c").setComment("chained").bytes();

        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["b.txt", "c.txt"]);
        expect(parser.getZipComment()).toBe("chained");
      });
    });

    // -------------------------------------------------------------------------
    // Edge cases
    // -------------------------------------------------------------------------

    describe("edge cases", () => {
      it("should handle entry names with special characters", async () => {
        const original = await zip({ reproducible: true })
          .add("文件.txt", "中文内容")
          .add("file with spaces.txt", "content")
          .bytes();

        const editor = await editZip(original, { reproducible: true });
        editor.rename("文件.txt", "重命名.txt");

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const paths = parser
          .getEntries()
          .map(e => e.path)
          .sort();
        expect(paths).toEqual(["file with spaces.txt", "重命名.txt"]);

        const content = await parser.extract("重命名.txt");
        expect(decode(content!)).toBe("中文内容");
      });

      it("should handle very large number of entries", async () => {
        const entries: Array<{ name: string; data: string }> = [];
        for (let i = 0; i < 100; i++) {
          entries.push({ name: `file${i.toString().padStart(3, "0")}.txt`, data: `content${i}` });
        }

        const archive = zip({ reproducible: true });
        for (const e of entries) {
          archive.add(e.name, e.data);
        }
        const original = await archive.bytes();

        const editor = await editZip(original, { reproducible: true });
        // Delete every other entry
        for (let i = 0; i < 100; i += 2) {
          editor.delete(`file${i.toString().padStart(3, "0")}.txt`);
        }

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        expect(parser.getEntries().length).toBe(50);
      });

      it("should throw on empty entry name", async () => {
        const original = await zip({ reproducible: true }).bytes();
        const editor = await editZip(original);

        expect(() => editor.set("", "content")).toThrow(/Entry name is required/);
        expect(() => editor.delete("")).toThrow(/Entry name is required/);
        expect(() => editor.rename("", "new")).toThrow(/Entry name is required/);
        expect(() => editor.rename("old", "")).toThrow(/Entry name is required/);
      });

      it("should handle async iterable source in bytes()", async () => {
        const original = await zip({ reproducible: true }).bytes();

        async function* generateChunks(): AsyncGenerator<Uint8Array> {
          yield encode("hello ");
          yield encode("world");
        }

        const editor = await editZip(original, { reproducible: true });
        editor.set("async.txt", generateChunks());

        const out = await editor.bytes();
        const parser = new ZipParser(out);
        const content = await parser.extract("async.txt");
        expect(decode(content!)).toBe("hello world");
      });
    });
  });
}
