/**
 * Word Example 21 — Validation
 *
 * Covers:
 *   - validateDocument on a clean doc (should pass)
 *   - validateDocument on an intentionally malformed doc — exercises rules
 *     for missing IDs, broken bookmarks, broken hyperlinks, malformed
 *     numbering refs, etc.
 *   - Strict mode (warnings → errors)
 *   - Per-version compatibility check (Word 2007)
 *   - maxSeverity filter, maxErrors short-circuit
 *
 * Output: tmp/word-examples/21-validation.txt — a human-readable report
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  text,
  bookmarkStart,
  bookmarkEnd,
  hyperlink,
  validateDocument,
  toBuffer
} from "../index";
import type { DocxDocument, BodyContent } from "../index";

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
// 1. Clean document — should produce no errors
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Clean", 1);
  Document.addParagraph(d, "Hello world.");
  Document.addParagraphElement(
    d,
    paragraph([bookmarkStart(0, "intro"), text("anchored content"), bookmarkEnd(0)])
  );

  const result = validateDocument(Document.build(d));
  log(`# Clean document`);
  log(`  valid: ${result.valid}`);
  log(`  errors: ${result.errorCount}, warnings: ${result.warningCount}`);
  for (const issue of result.issues) {
    log(`  [${issue.severity}] ${issue.rule} @ ${issue.path}: ${issue.message}`);
  }
}

// ---------------------------------------------------------------------------
// 2. Document with structural anomalies (dangling bookmark, anchor pointing
//    to a missing target, numbering reference without a definition).
//    Note: the current validator does not enforce every cross-reference
//    rule — it focuses on schema-level violations. This sample is mostly a
//    smoke test for the public API surface.
// ---------------------------------------------------------------------------
{
  const body: BodyContent[] = [
    paragraph([
      // Hyperlink → anchor that doesn't exist
      hyperlink("link to missing", { anchor: "no-such-bookmark" }),
      text(" "),
      // bookmarkStart with no matching end
      bookmarkStart(99, "lonely"),
      text("dangling bookmark text")
    ]),
    // Paragraph that references numId 12345 (undefined)
    textParagraph("Numbered without numbering def", {
      numbering: { numId: 12345, level: 0 }
    })
  ];
  const malformed: DocxDocument = { body };
  const result = validateDocument(malformed);
  log(`\n# Malformed document`);
  log(`  valid: ${result.valid}`);
  log(`  errors: ${result.errorCount}, warnings: ${result.warningCount}`);
  for (const issue of result.issues) {
    log(`  [${issue.severity}] ${issue.rule} @ ${issue.path}: ${issue.message}`);
  }
}

// ---------------------------------------------------------------------------
// 3. Strict mode — treat warnings as errors
// ---------------------------------------------------------------------------
{
  const body: BodyContent[] = [
    // Empty paragraph normally is fine, but combined with strict=true some
    // warnings get promoted.
    textParagraph(""),
    paragraph([hyperlink("dangling", { anchor: "missing" }), text(" tail")])
  ];
  const r = validateDocument({ body }, { strict: true });
  log(`\n# Strict mode (warnings → errors)`);
  log(`  valid: ${r.valid}, errors: ${r.errorCount}, warnings: ${r.warningCount}`);
}

// ---------------------------------------------------------------------------
// 4. Word 2007 compatibility check (some 2016+ features are flagged)
// ---------------------------------------------------------------------------
{
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Compat check", 1);
  Document.addParagraph(d, "ChartEx (Office 2016+) is flagged when targeting Word 2007.");
  // Insert a chartEx block (Office 2016+ feature). The XML is opaque; the
  // validator only flags it for compat purposes.
  Document.addContent(d, {
    type: "chartEx",
    chartExXml:
      '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>',
    altText: "treemap"
  });

  const r = validateDocument(Document.build(d), { compatibilityMode: "word2007" });
  log(`\n# Word 2007 compatibility`);
  log(`  valid: ${r.valid}, errors: ${r.errorCount}, warnings: ${r.warningCount}`);
  for (const issue of r.issues.slice(0, 5)) {
    log(`  [${issue.severity}] ${issue.rule}: ${issue.message}`);
  }
}

// ---------------------------------------------------------------------------
// 5. maxErrors short-circuit
// ---------------------------------------------------------------------------
{
  const body: BodyContent[] = Array.from({ length: 50 }, (_, i) =>
    textParagraph(`numbered ${i}`, { numbering: { numId: 999 + i, level: 0 } })
  );
  const r = validateDocument({ body }, { maxErrors: 5 });
  log(`\n# maxErrors=5`);
  log(`  reported issues: ${r.issues.length} (capped)`);
}

fs.writeFileSync(path.join(outDir, "21-validation.txt"), lines.join("\n"));
console.log(`  → 21-validation.txt`);

// Pretend to write a docx to keep parity with other examples
const ok = Document.create();
Document.useDefaultStyles(ok);
Document.addParagraph(ok, "(See 21-validation.txt for the report.)");
fs.writeFileSync(path.join(outDir, "21-validation.docx"), await toBuffer(Document.build(ok)));
