/**
 * Overlapping writes on one editor. A write keeps its working state on the editor across awaits,
 * so the editor runs them one after another; the outputs must be the ones sequential calls give.
 */
import { describe, expect, it } from "vitest";

import { generateSelfSignedCertificate } from "../examples/utils/self-signed-certificate";
import { Pdf } from "../index";

async function makeForm(): Promise<Uint8Array> {
  // A fixed date: two editors built from forms made a second apart would otherwise differ.
  const at = new Date("2024-01-01T00:00:00Z");
  const builder = new Pdf.Builder().setMetadata({ creationDate: at, modDate: at });
  const page = builder.addPage();
  page.drawText("Application", { x: 60, y: 740, fontSize: 18 });
  page.addFormField({ type: "text", name: "fullName", rect: [60, 690, 300, 712] });
  builder.addPage().drawText("Second", { x: 60, y: 740, fontSize: 18 });
  return builder.toBuffer();
}

async function edited() {
  const editor = Pdf.Editor.read(await makeForm()).setFormField("fullName", "Ada");
  editor.getPage(0).drawText("Overlay", { x: 60, y: 600, fontSize: 12 });
  editor.getPage(1).drawText("More", { x: 60, y: 600, fontSize: 12 });
  return editor;
}

type PdfEditor = Awaited<ReturnType<typeof edited>>;
type Write = (editor: PdfEditor) => Promise<Uint8Array>;
const writes: Write[] = [
  editor => editor.toBuffer(),
  editor => editor.toBuffer({ incremental: true }),
  editor => editor.toBuffer(),
  editor => editor.toBuffer({ incremental: "required" }),
  editor => editor.toBuffer({ incremental: true })
];

describe("PdfEditor — overlapping writes", () => {
  it("full and incremental writes in flight together equal the same writes run in turn", async () => {
    const sequentialEditor = await edited();
    const sequential: Uint8Array[] = [];
    for (const write of writes) {
      sequential.push(await write(sequentialEditor));
    }
    const concurrent = await Promise.all(writes.map(write => write(sequentialEditor)));
    const fresh = await edited();
    const overlapping = await Promise.all(writes.map(write => write(fresh)));
    expect(overlapping).toEqual(sequential);
    expect(concurrent).toEqual(sequential);
  });

  it("a failed write does not fail the writes queued behind it", async () => {
    const editor = await edited();
    editor.removePage(1);
    const reference = await (await edited()).toBuffer();
    const [refused, after] = await Promise.allSettled([
      editor.toBuffer({ incremental: "required" }),
      editor.toBuffer()
    ]);
    expect(refused.status).toBe("rejected");
    expect(after.status).toBe("fulfilled");
    expect((after as PromiseFulfilledResult<Uint8Array>).value.length).toBeLessThan(
      reference.length
    );
  });

  describe("a write reflects the edits made before it was called", () => {
    const calls: Array<[string, Write]> = [
      ["toBuffer()", editor => editor.toBuffer()],
      ["toBuffer({ incremental: true })", editor => editor.toBuffer({ incremental: true })],
      ['toBuffer({ incremental: "required" })', e => e.toBuffer({ incremental: "required" })],
      ["save()", editor => editor.save()],
      ["saveIncremental()", editor => editor.saveIncremental()]
    ];

    /** Edits made after the call: none of them may reach that call's output. */
    function editLater(editor: PdfEditor, value: string): void {
      editor.setFormField("fullName", value);
      editor.getPage(0).drawText("LATE", { x: 60, y: 500, fontSize: 12 });
    }

    it.each(calls)("%s with no write in flight", async (_, write) => {
      const reference = await write(await edited());
      const editor = await edited();
      const pending = write(editor);
      editLater(editor, "AFTER");
      const output = await pending;
      expect(output).toEqual(reference);
      const read = await Pdf.extract(output);
      expect(read.formFields.find(f => f.name === "fullName")?.value).toBe("Ada");
      expect(read.text).not.toContain("LATE");
    });

    it.each(calls)("%s queued behind another write", async (_, write) => {
      const midEditor = await edited();
      midEditor.setFormField("fullName", "MID");
      const reference = await write(midEditor);

      const editor = await edited();
      const first = editor.toBuffer();
      editor.setFormField("fullName", "MID");
      const queued = write(editor);
      editLater(editor, "AFTER");
      const [, output] = await Promise.all([first, queued]);
      expect(output).toEqual(reference);
      const read = await Pdf.extract(output);
      expect(read.formFields.find(f => f.name === "fullName")?.value).toBe("MID");
      expect(read.text).not.toContain("LATE");
    });

    it("a structural edit made after the call does not reach the output", async () => {
      for (const queue of [false, true]) {
        const editor = await edited();
        const first = queue ? editor.toBuffer() : undefined;
        const pending = editor.toBuffer({ incremental: true });
        editor.removePage(1);
        editor.rotatePage(0, 90);
        editor.addPage().drawText("ADDED", { x: 60, y: 600, fontSize: 12 });
        await first;
        const read = await Pdf.extract(await pending);
        expect(read.pages).toHaveLength(2);
        expect(read.text).not.toContain("ADDED");
      }
    });

    it("sign() signs the state at its call", async () => {
      const { certificate, privateKey } = await generateSelfSignedCertificate("Call time");
      for (const queue of [false, true]) {
        const editor = await edited();
        const first = queue ? editor.toBuffer() : undefined;
        const pending = editor.sign({ certificate, privateKey });
        editLater(editor, "AFTER");
        editor.removePage(1);
        await first;
        const signed = await pending;
        const read = await Pdf.extract(signed);
        expect(read.pages).toHaveLength(2);
        expect(read.formFields.find(f => f.name === "fullName")?.value).toBe("Ada");
        expect(read.text).not.toContain("LATE");
        expect((await Pdf.verifySignatures(signed))[0].intact).toBe(true);
      }
    });

    it("the edits still apply to the next write", async () => {
      const editor = await edited();
      const pending = editor.toBuffer();
      editLater(editor, "AFTER");
      await pending;
      const read = await Pdf.extract(await editor.toBuffer());
      expect(read.formFields.find(f => f.name === "fullName")?.value).toBe("AFTER");
      expect(read.text).toContain("LATE");
    });
  });

  it("sign() in flight leaves its placeholder out of an overlapping toBuffer()", async () => {
    const { certificate, privateKey } = await generateSelfSignedCertificate("Concurrent");
    const reference = await (await edited()).toBuffer();
    const editor = await edited();
    const [signed, plain] = await Promise.all([
      editor.sign({ certificate, privateKey }),
      editor.toBuffer()
    ]);
    expect(plain).toEqual(reference);
    expect((await Pdf.verifySignatures(signed))[0].intact).toBe(true);
  });
});
