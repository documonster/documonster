/**
 * Word Example 46 — Security policy & render context
 *
 * The reader / writer apply hard limits and content filters from a
 * `WordSecurityPolicy`. Two presets ship out of the box:
 *   - DEFAULT_SECURITY_POLICY — preserve everything, generous limits.
 *   - STRICT_SECURITY_POLICY  — strip VBA/OLE/altChunks, tighter limits,
 *     no external relationship targets.
 *
 * Covers:
 *   - DEFAULT_SECURITY_POLICY / STRICT_SECURITY_POLICY constants
 *   - resolveSecurityPolicy — fill missing fields with defaults
 *   - readDocx({ securityPolicy }) — enforced limits during read
 *   - createIdGenerators / createRenderContext — used when writing
 *     auxiliary parts manually (not common, but the public surface is
 *     here for advanced consumers)
 *
 * Output: tmp/word-examples/46-policy.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, isDocxError, Io, RenderContext, Security } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const lines: string[] = [];
const log = (s: string): void => {
  console.log(s);
  lines.push(s);
};

// ---------------------------------------------------------------------------
// 1. Show the two preset policies
// ---------------------------------------------------------------------------
log(`  DEFAULT_SECURITY_POLICY:`);
for (const [k, v] of Object.entries(Security.DEFAULT_SECURITY_POLICY)) {
  log(`    ${k.padEnd(24)} = ${v}`);
}
log(`\n  STRICT_SECURITY_POLICY:`);
for (const [k, v] of Object.entries(Security.STRICT_SECURITY_POLICY)) {
  log(`    ${k.padEnd(24)} = ${v}`);
}

// ---------------------------------------------------------------------------
// 2. resolveSecurityPolicy fills in missing fields
// ---------------------------------------------------------------------------
const partial = Security.resolveSecurityPolicy({ maxPackageSize: 1024 * 1024 }); // 1 MB cap
log(
  `\n  resolveSecurityPolicy({ maxPackageSize: 1MB }) merged with defaults: maxPackageSize=${partial.maxPackageSize}, preserveVbaProject=${partial.preserveVbaProject}, allowExternalTargets=${partial.allowExternalTargets}`
);

// ---------------------------------------------------------------------------
// 3. Build a small doc and read it back under both policies
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addParagraph(d, "Tiny doc");
const bytes = await Io.toBuffer(Document.build(d));
fs.writeFileSync(path.join(outDir, "46-policy.docx"), bytes);

// Default policy: succeeds
const defaultRead = await Io.read(bytes, { securityPolicy: Security.DEFAULT_SECURITY_POLICY });
log(`\n  readDocx(DEFAULT) body length: ${defaultRead.body.length}`);

// Strict policy: still succeeds for this benign file
const strictRead = await Io.read(bytes, { securityPolicy: Security.STRICT_SECURITY_POLICY });
log(`  readDocx(STRICT)  body length: ${strictRead.body.length}`);

// Custom policy with a comically tight cap forces a rejection
const tightPolicy = Security.resolveSecurityPolicy({ maxPackageSize: 100 });
try {
  await Io.read(bytes, { securityPolicy: tightPolicy });
  log(`  ERROR: expected size-cap to throw`);
} catch (err) {
  if (isDocxError(err)) {
    log(`  readDocx({maxPackageSize:100}) → ${err.constructor.name}: "${err.message}"`);
  } else {
    throw err;
  }
}

// ---------------------------------------------------------------------------
// 4. createIdGenerators — incrementing counters with optional seed
// ---------------------------------------------------------------------------
const ids = RenderContext.createIds({ drawingId: 100, sdtId: 50 });
log(
  `\n  createIdGenerators(drawing=100, sdt=50): drawing=${ids.nextDrawingId()}, sdt=${ids.nextSdtId()}, bookmark=${ids.nextBookmarkId()}, docPr=${ids.nextDocPrId()}, chart=${ids.nextChartId()}, image=${ids.nextImagePartId()}`
);
log(
  `  next round increments:                  drawing=${ids.nextDrawingId()}, sdt=${ids.nextSdtId()}`
);

// ---------------------------------------------------------------------------
// 5. createRenderContext — usually used internally; exposed for advanced
//    callers that build auxiliary parts manually.
// ---------------------------------------------------------------------------
const ctx = RenderContext.create({
  partName: "/word/custom.xml",
  securityPolicy: Security.STRICT_SECURITY_POLICY,
  ids
});
log(
  `\n  createRenderContext: part=${ctx.partName}, rawXmlPolicy=${ctx.rawXmlPolicy}, securityPolicy.preserveVbaProject=${ctx.securityPolicy.preserveVbaProject}`
);

fs.writeFileSync(path.join(outDir, "46-policy.txt"), lines.join("\n"));
console.log(`\n  → 46-policy.txt`);
