/**
 * Word Example 46 — Security policy & render context
 *
 * The reader / writer apply hard limits and content filters from a
 * `DocxSecurityPolicy`. Two presets ship out of the box:
 *   - DEFAULT_SECURITY_POLICY — preserve everything, generous limits.
 *   - STRICT_SECURITY_POLICY  — strip VBA/OLE/altChunks, tighter limits,
 *     no external relationship targets.
 *
 * Covers:
 *   - DEFAULT_SECURITY_POLICY / STRICT_SECURITY_POLICY constants
 *   - Security.resolvePolicy — fill missing fields with defaults
 *   - A partial policy — unspecified fields fall back to the defaults
 *   - Document.read(bytes, { securityPolicy }) — enforced limits during read
 *   - RenderContext.createIds / RenderContext.create — used when writing
 *     auxiliary parts manually (not common, but the public surface is
 *     here for advanced consumers)
 *
 * Output: tmp/word-examples/46-policy.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, isDocxError, RenderContext, Security } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const lines: string[] = [];
function log(s: string): void {
  console.log(s);
  lines.push(s);
}

// ---------------------------------------------------------------------------
// 1. Show the two preset policies
// ---------------------------------------------------------------------------
log(`  DEFAULT_SECURITY_POLICY:`);
for (const [k, v] of Object.entries(Security.DEFAULT_POLICY)) {
  log(`    ${k.padEnd(24)} = ${v}`);
}
log(`\n  STRICT_SECURITY_POLICY:`);
for (const [k, v] of Object.entries(Security.STRICT_POLICY)) {
  log(`    ${k.padEnd(24)} = ${v}`);
}

// ---------------------------------------------------------------------------
// 2. Security.resolvePolicy fills in missing fields
// ---------------------------------------------------------------------------
const partial = Security.resolvePolicy({ maxPackageSize: 1024 * 1024 }); // 1 MB cap
log(
  `\n  resolvePolicy({ maxPackageSize: 1MB }) merged with defaults: maxPackageSize=${partial.maxPackageSize}, preserveVbaProject=${partial.preserveVbaProject}, allowExternalTargets=${partial.allowExternalTargets}`
);

// ---------------------------------------------------------------------------
// 3. Build a small doc and read it back under both policies
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addParagraph(d, "Tiny doc");
const bytes = await Document.toBuffer(d);
fs.writeFileSync(path.join(outDir, "46-policy.docx"), bytes);

// Default policy: succeeds
const defaultRead = await Document.read(bytes, { securityPolicy: Security.DEFAULT_POLICY });
log(`\n  Document.read(DEFAULT) body length: ${defaultRead.body.length}`);

// Strict policy: still succeeds for this benign file
const strictRead = await Document.read(bytes, { securityPolicy: Security.STRICT_POLICY });
log(`  Document.read(STRICT)  body length: ${strictRead.body.length}`);

// A partial policy with a comically tight cap forces a rejection; every field
// it leaves out keeps its default.
const tightPolicy = { maxPackageSize: 100 };
try {
  await Document.read(bytes, { securityPolicy: tightPolicy });
  log(`  ERROR: expected size-cap to throw`);
} catch (err) {
  if (isDocxError(err)) {
    log(`  Document.read({maxPackageSize:100}) → ${err.constructor.name}: "${err.message}"`);
  } else {
    throw err;
  }
}

// ---------------------------------------------------------------------------
// 4. RenderContext.createIds — incrementing counters with optional seed
// ---------------------------------------------------------------------------
const ids = RenderContext.createIds({ drawingId: 100, sdtId: 50 });
log(
  `\n  RenderContext.createIds(drawing=100, sdt=50): drawing=${ids.nextDrawingId()}, sdt=${ids.nextSdtId()}, bookmark=${ids.nextBookmarkId()}, docPr=${ids.nextDocPrId()}, chart=${ids.nextChartId()}, image=${ids.nextImagePartId()}`
);
log(
  `  next round increments:                       drawing=${ids.nextDrawingId()}, sdt=${ids.nextSdtId()}`
);

// ---------------------------------------------------------------------------
// 5. RenderContext.create — usually used internally; exposed for advanced
//    callers that build auxiliary parts manually.
// ---------------------------------------------------------------------------
const ctx = RenderContext.create({
  partName: "/word/custom.xml",
  securityPolicy: Security.STRICT_POLICY,
  ids
});
log(
  `\n  RenderContext.create: part=${ctx.partName}, rawXmlPolicy=${ctx.rawXmlPolicy}, securityPolicy.preserveVbaProject=${ctx.securityPolicy.preserveVbaProject}`
);

fs.writeFileSync(path.join(outDir, "46-policy.txt"), lines.join("\n"));
console.log(`\n  → 46-policy.txt`);
