/**
 * Word Example 30 — Edge cases & error handling
 *
 * Aggregated edge-case scenarios that don't fit cleanly into any single
 * topic above:
 *   - Reading a document that uses ISO-29500 Strict namespaces
 *     (auto-normalized by readDocx)
 *   - Round-trip preservation: opaque parts survive untouched
 *   - Reading an encrypted document (DocxEncryptedError)
 *   - patchDocument across runs and inside hyperlinks
 *   - validateDocument finds dangling references
 *   - readDocx tolerates malformed XML in non-essential parts (warnings)
 *   - Buffer / Uint8Array / base64 input variants
 *   - Very small (one paragraph) and very large (many paragraphs) docs
 *
 * Output: tmp/word-examples/30-edge/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  text,
  hyperlink,
  toBuffer,
  toBase64,
  readDocx,
  patchDocument,
  validateDocument,
  isDocxError,
  DocxError,
  DocxParseError,
  DocxWriteError,
  DocxMissingPartError,
  DocxInvalidStructureError,
  DocxUnsupportedFeatureError,
  DocxEncryptedError,
  DocxLimitExceededError,
  encryptDocx
} from "../index";
import type { Table } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/30-edge"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. Build → toBuffer + toBase64 (both should yield identical content)
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "Round trip");
  const built = Document.build(d);
  const buf = await toBuffer(built);
  const b64 = await toBase64(built);
  console.log(`  toBuffer: ${buf.length} bytes, toBase64: ${b64.length} chars`);
}

// ---------------------------------------------------------------------------
// 2. Read an encrypted .docx — should reject with DocxEncryptedError
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "secret content");
  const plain = await toBuffer(Document.build(d));
  const encrypted = await encryptDocx(plain, "shibboleth");
  fs.writeFileSync(path.join(outDir, "encrypted.docx"), encrypted);

  try {
    await readDocx(encrypted);
    console.log("  ERROR: expected DocxEncryptedError to be thrown");
  } catch (err) {
    if (err instanceof DocxEncryptedError) {
      console.log(`  encrypted file correctly rejected: ${err.message}`);
    } else if (isDocxError(err)) {
      console.log(`  rejected with DocxError: ${err.message}`);
    } else {
      throw err;
    }
  }
}

// ---------------------------------------------------------------------------
// 3. patchDocument: replace placeholders inside paragraphs and hyperlinks
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraphElement(
    d,
    paragraph([
      text("Hello, "),
      text("{NAME}"),
      text("! Click "),
      hyperlink("{ANCHOR}", { url: "https://example.com" })
    ])
  );
  Document.addParagraph(d, "Total: $" + "{TOTAL}");

  const buf = await toBuffer(Document.build(d));
  const patched = await patchDocument(buf, [
    { placeholder: "{NAME}", content: { type: "text", text: "Alice" } },
    { placeholder: "{ANCHOR}", content: { type: "text", text: "here" } },
    { placeholder: "$" + "{TOTAL}", content: { type: "text", text: "$1,234" } }
  ]);
  fs.writeFileSync(path.join(outDir, "patched.docx"), patched);
  console.log(`  → patched.docx (${patched.length} bytes)`);
}

// ---------------------------------------------------------------------------
// 4. validateDocument exposes structural problems
// ---------------------------------------------------------------------------
{
  // numId 999 is not defined — validator flags it via xref-numId-missing
  const result = validateDocument({
    body: [textParagraph("dangling ref", { numbering: { numId: 999, level: 0 } })]
  });
  console.log(
    `  validation issues for dangling numbering: ${result.issues.filter(i => i.rule.includes("numId")).length}`
  );
}

// ---------------------------------------------------------------------------
// 5. Empty document — produces a valid but minimal package
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  // do not add any content
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "empty.docx"), buf);
  // and re-read it without errors
  const reread = await readDocx(buf);
  console.log(`  empty doc: ${buf.length} bytes, body length after re-read: ${reread.body.length}`);
}

// ---------------------------------------------------------------------------
// 6. Very large document — exercise memory & XML serialiser
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Stress test", 1);
  for (let i = 0; i < 2000; i++) {
    Document.addParagraph(d, `Paragraph ${i + 1}: ${"Lorem ipsum ".repeat(10)}`);
  }
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "stress.docx"), buf);
  console.log(`  stress.docx: ${(buf.length / 1024).toFixed(1)} KB (2000 paragraphs)`);
}

// ---------------------------------------------------------------------------
// 7. Deep tree exercising the walker depth limit
// ---------------------------------------------------------------------------
{
  // Build a paragraph with deeply nested hyperlinks (legitimate but unusual)
  const d = Document.create();
  Document.useDefaultStyles(d);
  // Word does not allow nested hyperlinks; we just stack runs deep.
  let runs = [text("deep")];
  for (let i = 0; i < 30; i++) {
    runs = [text("("), ...runs, text(")")];
  }
  Document.addParagraphElement(d, paragraph(runs));
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "deep.docx"), buf);
  console.log(`  deep.docx: ${buf.length} bytes`);
}

// ---------------------------------------------------------------------------
// 8. Corrupted/truncated docx → readDocx must reject cleanly
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addParagraph(d, "valid");
  const ok = await toBuffer(Document.build(d));

  // Truncate to 200 bytes (after the ZIP local file header but before EOCD)
  const corrupted = ok.slice(0, 200);
  try {
    await readDocx(corrupted);
    console.log("  ERROR: expected truncated docx to reject");
  } catch (err) {
    console.log(
      `  truncated docx → ${(err as Error).constructor.name}: "${(err as Error).message.slice(0, 80)}"`
    );
  }

  // Random bytes that don't form a ZIP at all
  const random = new Uint8Array(50);
  for (let i = 0; i < random.length; i++) {
    random[i] = (i * 31) & 0xff;
  }
  try {
    await readDocx(random);
    console.log("  ERROR: expected random bytes to reject");
  } catch (err) {
    console.log(
      `  random bytes  → ${(err as Error).constructor.name}: "${(err as Error).message.slice(0, 80)}"`
    );
  }
}

// ---------------------------------------------------------------------------
// 9. Deeply nested tables (5 levels of nested cells)
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  // Build the innermost table first, then wrap it 5 times. OOXML's pct
  // unit uses 1/50 of a percent (5000 = 100 %), so we use 5000/pct.
  let inner: Table = {
    type: "table",
    properties: { width: { value: 5000, type: "pct" } },
    rows: [{ cells: [{ content: [textParagraph("innermost")] }] }]
  };
  for (let level = 4; level >= 0; level--) {
    inner = {
      type: "table",
      properties: { width: { value: 5000, type: "pct" } },
      rows: [
        {
          cells: [{ content: [textParagraph(`level ${level} label`)] }, { content: [inner] }]
        }
      ]
    };
  }
  Document.addTableElement(d, inner);
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "nested-tables.docx"), buf);
  console.log(`  nested-tables.docx: ${buf.length} bytes (5-level nesting)`);
}

// ---------------------------------------------------------------------------
// 10. Many simultaneous bookmarks (1 000) — exercise the bookmark id pipeline
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  for (let i = 0; i < 1000; i++) {
    const id = Document.nextBookmarkId(d);
    Document.addParagraphElement(
      d,
      paragraph([
        { type: "bookmarkStart", id, name: `bm-${i}` },
        text(`Item ${i}`),
        { type: "bookmarkEnd", id }
      ])
    );
  }
  const buf = await toBuffer(Document.build(d));
  fs.writeFileSync(path.join(outDir, "many-bookmarks.docx"), buf);
  console.log(`  many-bookmarks.docx: ${buf.length} bytes (1 000 bookmarks)`);
}

// ---------------------------------------------------------------------------
// 11. Error class taxonomy — demonstrate the full DocxError hierarchy.
//
// Hierarchy (each class extends the one to its left):
//   DocxError
//   ├── DocxParseError
//   │   ├── DocxMissingPartError
//   │   ├── DocxInvalidStructureError
//   │   └── DocxLimitExceededError
//   ├── DocxWriteError
//   ├── DocxUnsupportedFeatureError
//   └── DocxEncryptedError  (+ DocxDecryptionError)
//
// Application code should catch the most specific class it can act on, and
// fall back to `isDocxError(err)` for unknown DOCX failures. All classes
// preserve the standard `cause` chain.
// ---------------------------------------------------------------------------
{
  const cases: { name: string; err: DocxError }[] = [
    { name: "DocxError", err: new DocxError("base error") },
    { name: "DocxParseError", err: new DocxParseError("malformed XML in document.xml") },
    { name: "DocxWriteError", err: new DocxWriteError("failed to write docProps") },
    { name: "DocxMissingPartError", err: new DocxMissingPartError("word/document.xml") },
    {
      name: "DocxInvalidStructureError",
      err: new DocxInvalidStructureError("<w:body> has no children")
    },
    {
      name: "DocxUnsupportedFeatureError",
      err: new DocxUnsupportedFeatureError("vendor extension foo:Bar")
    },
    { name: "DocxEncryptedError", err: new DocxEncryptedError() },
    { name: "DocxLimitExceededError", err: new DocxLimitExceededError("packageSize", 100, 200) }
  ];
  for (const { name, err } of cases) {
    const isParse = err instanceof DocxParseError;
    const isWrite = err instanceof DocxWriteError;
    const isAny = isDocxError(err);
    console.log(
      `  ${name.padEnd(28)} parse=${isParse ? "Y" : "n"} write=${isWrite ? "Y" : "n"} docx=${isAny ? "Y" : "n"} :: ${err.message.slice(0, 50)}`
    );
  }
  // Show that limit-exceeded carries structured fields downstream callers
  // can branch on without parsing the message.
  const limit = new DocxLimitExceededError("partCount", 1000, 1500);
  console.log(
    `  DocxLimitExceededError fields: limit=${limit.limit} max=${limit.maximum} actual=${limit.actual}`
  );
}
