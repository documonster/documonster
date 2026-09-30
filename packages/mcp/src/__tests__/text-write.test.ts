/**
 * `text_write` — uploading a file in numbered parts.
 *
 * The tool exists because one tool argument cannot be longer than one model reply.
 * These tests pin what makes an upload a safe substitute for one large argument:
 * nothing is published with a part missing, order comes from the numbers, a resend
 * is harmless, a late call from another job cannot leak in, every failure leaves
 * the previous state intact, and the size limit is enforced before anything is
 * written.
 */

import {
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  rm,
  stat,
  symlink,
  utimes,
  writeFile
} from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { Pdf } from "documonster/pdf";

import { resolveConfig, type ServerConfig } from "../config.js";
import { docWriteTool } from "../tools/doc-write.js";
import { textWriteTool } from "../tools/text-write.js";

interface Fixture {
  readonly config: ServerConfig;
  readonly root: string;
}

async function fixture(args: readonly string[] = []): Promise<Fixture> {
  const root = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-mcp-textwrite-")));
  const base = resolveConfig(args, { cwd: root });
  return { config: { ...base, outputRoot: root, allowInPlace: true }, root };
}

async function run(
  tool: typeof textWriteTool,
  fx: Fixture,
  args: Record<string, unknown>
): Promise<string> {
  const result = await tool.handler(args as never, { config: fx.config });
  const text = result.content
    .filter((block): block is { type: "text"; text: string } => block.type === "text")
    .map(block => block.text)
    .join("\n");
  if (result.isError === true) {
    throw new Error(text);
  }
  return text;
}

function upload(fx: Fixture, args: Record<string, unknown>): Promise<string> {
  return run(textWriteTool, fx, args);
}

function idOf(reply: string): string {
  const match = /Upload \*\*([0-9a-f]{20})\*\*/.exec(reply);
  if (match?.[1] === undefined) {
    throw new Error(`no upload id in: ${reply}`);
  }
  return match[1];
}

async function start(fx: Fixture, total: number, extra: Record<string, unknown> = {}) {
  return idOf(await upload(fx, { path: "doc.md", total, ...extra }));
}

function file(fx: Fixture, name = "doc.md"): Promise<string> {
  return readFile(path.join(fx.root, name), "utf8");
}

async function present(fx: Fixture, name = "doc.md"): Promise<boolean> {
  return stat(path.join(fx.root, name)).then(
    () => true,
    () => false
  );
}

describe("text_write — assembly", () => {
  it("publishes only when every part is present, in number order", async () => {
    const fx = await fixture();
    const id = await start(fx, 3);
    await upload(fx, { upload: id, part: 3, text: "C" });
    const waiting = await upload(fx, { upload: id, part: 1, text: "A" });
    expect(waiting).toMatch(/not published yet.*part\(s\) 2/);
    // A partial file is exactly what must never be handed to the next tool.
    expect(await present(fx)).toBe(false);

    const done = await upload(fx, { upload: id, part: 2, text: "B" });
    expect(done).toContain("**published**");
    expect(await file(fx)).toBe("ABC");
  });

  it("can start with the first part in the same call", async () => {
    const fx = await fixture();
    const reply = await upload(fx, { path: "doc.md", total: 1, part: 1, text: "only" });
    expect(reply).toContain("**published**");
    expect(await file(fx)).toBe("only");
  });

  it("reports progress when called with just the id", async () => {
    const fx = await fixture();
    const id = await start(fx, 5);
    await upload(fx, { upload: id, part: 1, text: "a" });
    await upload(fx, { upload: id, part: 4, text: "d" });
    expect(await upload(fx, { upload: id })).toMatch(/2 of 5 part\(s\).*2–3, 5/s);
  });

  it("loses nothing when many parts arrive concurrently", async () => {
    const fx = await fixture();
    const count = 40;
    const id = await start(fx, count);
    await Promise.all(
      Array.from({ length: count }, (_, index) =>
        upload(fx, { upload: id, part: count - index, text: `[${count - index}]` })
      )
    );
    expect(await file(fx)).toBe(
      Array.from({ length: count }, (_, index) => `[${index + 1}]`).join("")
    );
  });

  it("joins bytes exactly and keeps its scaffolding out of sight", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "第一部分\n\n" });
    await upload(fx, { upload: id, part: 2, text: '{"k": "ü"}' });
    expect(await file(fx)).toBe('第一部分\n\n{"k": "ü"}');
    expect((await readdir(fx.root)).filter(name => !name.startsWith("."))).toEqual(["doc.md"]);
  });

  it("follows a change of plan in total", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "A" });
    await upload(fx, { upload: id, part: 2, text: "B" });
    expect(await file(fx)).toBe("AB");

    // More parts than planned: unpublished content is not half-applied.
    await upload(fx, { upload: id, total: 3 });
    await upload(fx, { upload: id, part: 3, text: "C" });
    expect(await file(fx)).toBe("ABC");

    // Fewer: parts beyond the new total are left out and said to be.
    const reply = await upload(fx, { upload: id, total: 2 });
    expect(reply).toMatch(/part\(s\) 3 are beyond total 2/);
    expect(await file(fx)).toBe("AB");
  });
});

describe("text_write — resends and isolation", () => {
  it("a resend of identical content changes nothing", async () => {
    const fx = await fixture();
    const id = await start(fx, 1);
    await upload(fx, { upload: id, part: 1, text: "same" });
    const before = await stat(path.join(fx.root, "doc.md"));
    await new Promise(resolve => setTimeout(resolve, 20));
    const reply = await upload(fx, { upload: id, part: 1, text: "same" });
    expect(reply).toContain("up to date");
    expect((await stat(path.join(fx.root, "doc.md"))).mtimeMs).toBe(before.mtimeMs);
  });

  it("a corrected part republishes the file", async () => {
    const fx = await fixture();
    const id = await start(fx, 3);
    await upload(fx, { upload: id, part: 1, text: "one " });
    await upload(fx, { upload: id, part: 2, text: "tow " });
    await upload(fx, { upload: id, part: 3, text: "three" });
    await upload(fx, { upload: id, part: 2, text: "two " });
    expect(await file(fx)).toBe("one two three");
  });

  it("a late call from an earlier upload cannot overwrite a newer one", async () => {
    const fx = await fixture();
    const first = await start(fx, 1);
    await upload(fx, { upload: first, part: 1, text: "old" });
    const second = await start(fx, 1, { overwrite: true });
    await upload(fx, { upload: second, part: 1, text: "new" });

    expect(await upload(fx, { upload: first, part: 1, text: "stale" })).toMatch(
      /publishing failed.*changed by something else/s
    );
    expect(await file(fx)).toBe("new");
  });

  it("refuses to republish over a file someone else changed", async () => {
    const fx = await fixture();
    const id = await start(fx, 1);
    await upload(fx, { upload: id, part: 1, text: "mine" });
    await writeFile(path.join(fx.root, "doc.md"), "theirs, edited afterwards");
    expect(await upload(fx, { upload: id, part: 1, text: "mine again" })).toMatch(
      /publishing failed.*changed by something else/s
    );
    expect(await file(fx)).toBe("theirs, edited afterwards");
  });
});

describe("text_write — refusals leave state intact", () => {
  it("refuses an existing file unless overwrite is given, before any part is sent", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "doc.md"), "someone's file");
    await expect(start(fx, 2)).rejects.toThrow(/already exists/);
    expect(await file(fx)).toBe("someone's file");

    const id = await start(fx, 1, { overwrite: true });
    await upload(fx, { upload: id, part: 1, text: "x" });
    expect(await file(fx)).toBe("x");
  });

  it("checks the size limit before storing, and stores nothing when over it", async () => {
    const fx = await fixture(["--max-file-size", "10"]);
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "12345" });
    await expect(upload(fx, { upload: id, part: 2, text: "678901" })).rejects.toThrow(
      /over the 10 byte limit/
    );
    expect(await upload(fx, { upload: id })).toMatch(/1 of 2 part\(s\)/);
    await upload(fx, { upload: id, part: 2, text: "67890" });
    expect(await file(fx)).toBe("1234567890");
  });

  it("a failed publish keeps every stored part, so the upload can finish later", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "A" });
    // Something occupies the destination between start and publish.
    await mkdir(path.join(fx.root, "doc.md"));
    expect(await upload(fx, { upload: id, part: 2, text: "B" })).toMatch(
      /2 of 2 part\(s\) stored.*publishing failed/s
    );

    // Once the obstruction is gone, a bare status call publishes.
    await rm(path.join(fx.root, "doc.md"), { recursive: true });
    expect(await upload(fx, { upload: id })).toContain("**published**");
    expect(await file(fx)).toBe("AB");
  });

  it("rejects malformed and unknown sessions and inconsistent calls", async () => {
    const fx = await fixture();
    await expect(upload(fx, { path: "doc.md" })).rejects.toThrow(/needs `path` and `total`/);
    await expect(upload(fx, { upload: "../x" })).rejects.toThrow(/not an upload id/);
    await expect(upload(fx, { upload: "0".repeat(20) })).rejects.toThrow(/no upload session/);

    const id = await start(fx, 2);
    await expect(upload(fx, { upload: id, part: 3, text: "x" })).rejects.toThrow(/beyond total 2/);
    await expect(upload(fx, { upload: id, part: 1 })).rejects.toThrow(/go together/);
    await expect(upload(fx, { upload: id, path: "other.md" })).rejects.toThrow(/not @output/);
  });

  it("stays inside the output root", async () => {
    const fx = await fixture();
    await expect(upload(fx, { path: "../escape.md", total: 1 })).rejects.toThrow(/outside/);

    const outside = await realpath(await mkdtemp(path.join(tmpdir(), "documonster-outside-")));
    await symlink(outside, path.join(fx.root, ".documonster-uploads"));
    await expect(upload(fx, { path: "doc.md", total: 1 })).rejects.toThrow(/outside/);
    expect(await readdir(outside)).toEqual([]);
  });
});

describe("text_write — state stays consistent across failures", () => {
  function sessionDir(fx: Fixture, id: string): string {
    return path.join(fx.root, ".documonster-uploads", id);
  }

  async function manifest(fx: Fixture, id: string): Promise<Record<string, unknown>> {
    return JSON.parse(await readFile(path.join(sessionDir(fx, id), "meta.json"), "utf8"));
  }

  it("a republish that failed is retried by a status call", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "A" });
    await upload(fx, { upload: id, part: 2, text: "B" });
    expect(await file(fx)).toBe("AB");

    // The destination is unusable while the correction arrives.
    await rm(path.join(fx.root, "doc.md"));
    await mkdir(path.join(fx.root, "doc.md"));
    expect(await upload(fx, { upload: id, part: 2, text: "b" })).toMatch(/publishing failed/);
    // Nothing about "was a part just changed" may decide whether to publish: the
    // published revision is behind, so a bare status call must publish.
    expect(await upload(fx, { upload: id })).toMatch(/publishing failed/);

    await rm(path.join(fx.root, "doc.md"), { recursive: true });
    expect(await upload(fx, { upload: id })).toContain("**published**");
    expect(await file(fx)).toBe("Ab");
    expect(await upload(fx, { upload: id })).toContain("up to date");
  });

  it("recognises its own file after a crash between installing it and recording it", async () => {
    const fx = await fixture();
    const id = await start(fx, 1);
    await upload(fx, { upload: id, part: 1, text: "first" });

    // Rewind the manifest to the state the install step leaves: the file is in
    // place and `pending` names it, but `published` was never written.
    const meta = await manifest(fx, id);
    const { published, ...rest } = meta;
    await writeFile(
      path.join(sessionDir(fx, id), "meta.json"),
      JSON.stringify({ ...rest, pending: published })
    );

    const reply = await upload(fx, { upload: id, part: 1, text: "second" });
    expect(reply).toContain("**published**");
    expect(await file(fx)).toBe("second");
  });

  it("an unreferenced part file left by an interrupted call changes nothing", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "A" });
    // What a crash after writing a part's bytes but before committing leaves behind.
    await writeFile(path.join(sessionDir(fx, id), "00002.deadbeef0000"), "GARBAGE");
    expect(await upload(fx, { upload: id })).toMatch(/1 of 2 part\(s\)/);
    await upload(fx, { upload: id, part: 2, text: "B" });
    expect(await file(fx)).toBe("AB");
  });

  it("a replaced part's old bytes are removed once the new ones are committed", async () => {
    const fx = await fixture();
    const id = await start(fx, 1);
    await upload(fx, { upload: id, part: 1, text: "old" });
    await upload(fx, { upload: id, part: 1, text: "new" });
    const files = (await readdir(sessionDir(fx, id))).filter(name => name !== "meta.json");
    expect(files).toHaveLength(1);
  });

  it("a rejected call changes nothing, including the total it carried", async () => {
    const fx = await fixture();
    const id = await start(fx, 2);
    await expect(upload(fx, { upload: id, total: 1, part: 2, text: "x" })).rejects.toThrow(
      /beyond total 1/
    );
    expect((await manifest(fx, id)).total).toBe(2);
    expect((await manifest(fx, id)).revision).toBe(0);
  });

  it("raising total cannot take the file over the size limit", async () => {
    const fx = await fixture(["--max-file-size", "10"]);
    const id = await start(fx, 2);
    await upload(fx, { upload: id, part: 1, text: "12345" });
    await upload(fx, { upload: id, part: 2, text: "67890" });
    // Shrink total so part 2 no longer counts, then grow part 1 into the room.
    await upload(fx, { upload: id, total: 1 });
    await upload(fx, { upload: id, part: 1, text: "1234567890" });
    // Bringing part 2 back would make 15 bytes.
    await expect(upload(fx, { upload: id, total: 2 })).rejects.toThrow(/over the 10 byte limit/);
    expect((await manifest(fx, id)).total).toBe(1);
    // Nor by pairing the new total with a resend of an unchanged part.
    await expect(upload(fx, { upload: id, total: 2, part: 1, text: "1234567890" })).rejects.toThrow(
      /over the 10 byte limit/
    );
    expect((await manifest(fx, id)).total).toBe(1);
    expect(await file(fx)).toBe("1234567890");
  });

  it("two uploads publishing one path in parallel leave one complete file", async () => {
    const fx = await fixture();
    const a = await start(fx, 1, { overwrite: true });
    const b = await start(fx, 1, { overwrite: true });
    const big = (letter: string) => letter.repeat(200_000);
    await Promise.all([
      upload(fx, { upload: a, part: 1, text: big("a") }),
      upload(fx, { upload: b, part: 1, text: big("b") })
    ]);
    expect([big("a"), big("b")]).toContain(await file(fx));
  });

  it("reports a damaged manifest instead of crashing", async () => {
    const fx = await fixture();
    const id = await start(fx, 1);
    await writeFile(path.join(sessionDir(fx, id), "meta.json"), "{ not json");
    await expect(upload(fx, { upload: id })).rejects.toThrow(/unreadable manifest/);
  });

  it("a session in use is never swept; an abandoned one is", async () => {
    const fx = await fixture();
    const active = await start(fx, 2);
    const abandoned = await start(fx, 2);
    const old = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
    await utimes(sessionDir(fx, active), old, old);
    await utimes(sessionDir(fx, abandoned), old, old);

    // A status call is use, even though it changes no content.
    await upload(fx, { upload: active });
    await upload(fx, { path: "other.md", total: 1 });

    await expect(upload(fx, { upload: active })).resolves.toMatch(/0 of 2/);
    await expect(upload(fx, { upload: abandoned })).rejects.toThrow(/no upload session/);
  });
});

describe("text_write → doc_write", () => {
  it("turns a document too long for one call into one PDF", async () => {
    const fx = await fixture();
    const sections = 60;
    const id = await start(fx, sections);
    for (let index = 1; index <= sections; index += 1) {
      await upload(fx, {
        upload: id,
        part: index,
        // Latin only: this is about size, not scripts. A CJK PDF needs a CJK face on
        // the host, which a CI runner does not have — font coverage is pdf-fonts' job,
        // and CJK bytes are covered, font-free, by the byte-exact test above.
        text: `## Section ${index}\n\n${"content ".repeat(300)}\n\n`
      });
    }
    await run(docWriteTool, fx, { path: "doc.pdf", from: "doc.md" });

    const parsed = await Pdf.read(new Uint8Array(await readFile(path.join(fx.root, "doc.pdf"))), {
      extractText: true
    });
    const text = parsed.pages.map(page => page.text).join("\n");
    expect(parsed.pages.length).toBeGreaterThan(10);
    expect(text).toContain("Section 1");
    expect(text).toContain(`Section ${sections}`);
  });

  it("requires exactly one of markdown and from", async () => {
    const fx = await fixture();
    await expect(run(docWriteTool, fx, { path: "a.docx" })).rejects.toThrow(
      /`markdown` or `from` is required/
    );
    await writeFile(path.join(fx.root, "a.md"), "# A\n");
    await expect(
      run(docWriteTool, fx, { path: "a.docx", markdown: "# B", from: "a.md" })
    ).rejects.toThrow(/not both/);
  });

  it("refuses an empty source file", async () => {
    const fx = await fixture();
    await writeFile(path.join(fx.root, "empty.md"), "  \n");
    await expect(run(docWriteTool, fx, { path: "a.docx", from: "empty.md" })).rejects.toThrow(
      /is empty/
    );
  });
});
