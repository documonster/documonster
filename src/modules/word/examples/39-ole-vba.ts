/**
 * Word Example 39 — OLE objects & VBA macros (.docm round-trip)
 *
 * Covers:
 *   - extractOleObjects / hasOleObjects / getOleObjectData
 *   - createOleEmbedding (with optional preview image)
 *   - hasVbaProject / getVbaProjectInfo / getVbaProjectData
 *   - addVbaProject / removeVbaProject
 *   - listVbaParts
 *
 * Note: the OLE / VBA binary contents in this example are synthetic stubs.
 * Real OLE objects need a valid OLE2 compound stream and real VBA macros
 * need a properly compiled `vbaProject.bin`.  Word will reject malformed
 * binaries when opening the file, but the writer/reader API surface still
 * round-trips the bytes faithfully.
 *
 * Output: tmp/word-examples/39-ole-vba/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  toBuffer,
  readDocx,
  extractOleObjects,
  hasOleObjects,
  getOleObjectData,
  createOleEmbedding,
  hasVbaProject,
  getVbaProjectInfo,
  getVbaProjectData,
  addVbaProject,
  removeVbaProject,
  listVbaParts
} from "../index";
import type { DocxDocument } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/39-ole-vba"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// 1. createOleEmbedding — produce opaque parts representing an OLE object
// ---------------------------------------------------------------------------
const fakeOleBytes = new TextEncoder().encode("FAKE-OLE2-STUB-DATA");
// 1×1 PNG to use as preview image
const previewPng = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x03, 0x00, 0x01, 0x5b, 0x6e, 0x5e, 0x49, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82
]);

const ole = createOleEmbedding(fakeOleBytes, "Excel.Sheet.12", {
  previewImage: previewPng,
  previewContentType: "image/png"
});
console.log(
  `  OLE part: ${ole.olePart.path} (${ole.olePart.data.length} bytes), rId=${ole.oleRId}`
);
console.log(
  `  Preview part: ${ole.previewPart?.path}, rId=${ole.previewRId} (${ole.previewPart?.data.length ?? 0} bytes)`
);

// ---------------------------------------------------------------------------
// 2. Build a doc and inject the OLE part as opaqueParts
// ---------------------------------------------------------------------------
const baseDoc = (() => {
  const dd = Document.create();
  Document.useDefaultStyles(dd);
  Document.addParagraphElement(
    dd,
    paragraph([text("Document with an embedded OLE object (Excel sheet stub).")])
  );
  return Document.build(dd);
})();
const docWithOle: DocxDocument = {
  ...baseDoc,
  opaqueParts: [ole.olePart, ...(ole.previewPart ? [ole.previewPart] : [])]
};
fs.writeFileSync(path.join(outDir, "01-with-ole.docx"), await toBuffer(docWithOle));
console.log(`  → 01-with-ole.docx`);

// ---------------------------------------------------------------------------
// 3. Read it back and inspect the OLE objects
// ---------------------------------------------------------------------------
const reread = await readDocx(await toBuffer(docWithOle));
console.log(`  hasOleObjects(reread): ${hasOleObjects(reread)}`);
const extraction = extractOleObjects(reread);
console.log(`  extractOleObjects: ${extraction.objects.length} OLE part(s)`);
for (const obj of extraction.objects) {
  console.log(
    `    rId=${obj.rId}, progId=${obj.progId ?? "(none)"}, dataSize=${obj.data?.length ?? 0}`
  );
}
const dataBack = getOleObjectData(reread, ole.oleRId);
console.log(`  getOleObjectData(rId): ${dataBack ? `${dataBack.length} bytes` : "undefined"}`);

// ---------------------------------------------------------------------------
// 4. VBA macros — addVbaProject / hasVbaProject / removeVbaProject
//    addVbaProject() validates the file's OLE2 compound-document magic
//    (D0 CF 11 E0 A1 B1 1A E1), so we hand-craft a stub starting with
//    those 8 bytes plus a zero-padded body. Word will still complain when
//    actually trying to execute the macro, but the round-trip works.
// ---------------------------------------------------------------------------
const ole2Magic = Uint8Array.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const fakeVbaBin = new Uint8Array(512);
fakeVbaBin.set(ole2Magic, 0);
// Fill the remainder with deterministic noise
for (let i = ole2Magic.length; i < fakeVbaBin.length; i++) {
  fakeVbaBin[i] = (i * 13) & 0xff;
}
const docWithVba = addVbaProject(baseDoc, fakeVbaBin);
console.log(`  hasVbaProject after add: ${hasVbaProject(docWithVba)}`);
console.log(`  getVbaProjectInfo: ${JSON.stringify(getVbaProjectInfo(docWithVba))}`);
console.log(`  getVbaProjectData length: ${getVbaProjectData(docWithVba)?.length ?? 0}`);
console.log(`  listVbaParts count: ${listVbaParts(docWithVba).length}`);

// Save .docm file
const docmBytes = await toBuffer(docWithVba);
fs.writeFileSync(path.join(outDir, "02-macro-enabled.docm"), docmBytes);
console.log(`  → 02-macro-enabled.docm (${docmBytes.length} bytes)`);

// Round-trip then strip
const rereadVba = await readDocx(docmBytes);
console.log(`  re-read .docm has VBA: ${hasVbaProject(rereadVba)}`);
const stripped = removeVbaProject(rereadVba);
console.log(`  after remove: ${hasVbaProject(stripped)}`);
fs.writeFileSync(path.join(outDir, "03-macros-stripped.docx"), await toBuffer(stripped));
console.log(`  → 03-macros-stripped.docx`);
