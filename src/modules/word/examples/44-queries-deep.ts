/**
 * Word Example 44 — Deep queries: bookmarks, comments, format-aware search
 *
 * Picks up the smaller-but-useful query helpers that 18-read-modify only
 * skimmed:
 *   - Query.getBookmark — locate a bookmark range by name (returns surrounding
 *     paragraph index + range)
 *   - Query.getComment — locate a CommentDef by id
 *   - searchByFormat with the full FormatCriteria surface
 *     (paragraphStyle / characterStyle / textMatch / multiple flags AND'd)
 *   - Query.listUsedFormats — distinct RunProperties shapes encountered
 *
 * Output: tmp/word-examples/44-queries.txt
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Query, Units } from "../index";

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
// Build a sample document
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);
Document.addStyle(d, {
  type: "character",
  styleId: "Strong",
  name: "Strong",
  runProperties: { bold: true, color: "C00000" }
});

Document.addHeading(d, "Section A", 1);
Document.addParagraph(
  d,
  Build.paragraph([
    Build.bookmarkStart(0, "intro"),
    Build.text("Plain "),
    Build.bold("bold "),
    Build.italic("italic "),
    Build.text("RED ", { color: "FF0000" }),
    Build.text("LARGE", { size: Units.ptToHalfPoint(18) }),
    Build.text(" "),
    Build.text("strong-styled", { style: "Strong" }),
    Build.bookmarkEnd(0)
  ])
);

const cId = Document.addComment(d, "Reviewer", "Important note", { initials: "RV" });
Document.addParagraph(
  d,
  Build.paragraph([
    Build.commentRangeStart(cId),
    Build.text("Sentence with comment annotation."),
    Build.commentRangeEnd(cId),
    Build.commentReference(cId)
  ])
);

Document.addHeading(d, "Section B", 1);
Document.addParagraph(
  d,
  Build.paragraph([
    Build.bookmarkStart(1, "TODO"),
    Build.bold("TODO: revise this section."),
    Build.bookmarkEnd(1)
  ])
);
Document.addParagraph(d, "TODO: also reconcile the figures.");

fs.writeFileSync(path.join(outDir, "44-queries.docx"), await Document.toBuffer(d));

// ---------------------------------------------------------------------------
// 1. Query.getBookmark — returns { bookmark, paragraphIndex, childIndex }
// ---------------------------------------------------------------------------
const bm = Query.getBookmark(d, "intro");
log(
  `  Query.getBookmark("intro"): ${bm ? `id=${bm.bookmark.id}, paragraphIndex=${bm.paragraphIndex}, childIndex=${bm.childIndex}` : "missing"}`
);
const bm2 = Query.getBookmark(d, "TODO");
log(`  Query.getBookmark("TODO"): ${bm2 ? `id=${bm2.bookmark.id}` : "missing"}`);
const bmMissing = Query.getBookmark(d, "no-such");
log(`  Query.getBookmark("no-such"): ${bmMissing}`);

// ---------------------------------------------------------------------------
// 2. Query.getComment
// ---------------------------------------------------------------------------
const c0 = Query.getComment(d, cId);
log(`  Query.getComment(${cId}): author=${c0?.author}, initials=${c0?.initials}`);
const cMissing = Query.getComment(d, 999);
log(`  Query.getComment(999): ${cMissing}`);

// ---------------------------------------------------------------------------
// 3. searchByFormat — multiple criteria AND-ed
// ---------------------------------------------------------------------------
log(`\n  searchByFormat:`);
log(`    bold:                ${Query.searchByFormat(d, { bold: true }).length}`);
log(`    italic:              ${Query.searchByFormat(d, { italic: true }).length}`);
log(`    color FF0000:        ${Query.searchByFormat(d, { color: "FF0000" }).length}`);
log(
  `    size 36 (=18pt):     ${Query.searchByFormat(d, { size: Units.ptToHalfPoint(18) }).length}`
);
log(
  `    bold AND TODO regex: ${Query.searchByFormat(d, { bold: true, textMatch: /^TODO:/ }).length}`
);
log(`    Strong char style:   ${Query.searchByFormat(d, { characterStyle: "Strong" }).length}`);
log(`    Heading1 para style: ${Query.searchByFormat(d, { paragraphStyle: "Heading1" }).length}`);

// ---------------------------------------------------------------------------
// 4. countByFormat (no allocations)
// ---------------------------------------------------------------------------
log(`\n  countByFormat({ bold:true }):  ${Query.getCountByFormat(d, { bold: true })}`);
log(`  countByFormat({ italic:true }): ${Query.getCountByFormat(d, { italic: true })}`);

// ---------------------------------------------------------------------------
// 5. Query.listUsedFormats — distinct property shapes
// ---------------------------------------------------------------------------
const used = Query.listUsedFormats(d);
log(`\n  getUsedFormats: ${used.length} distinct shapes`);
for (const f of used.slice(0, 10)) {
  log(`    · ${JSON.stringify(f)}`);
}

// ---------------------------------------------------------------------------
// 6. Edge: regex with no matches still returns array
// ---------------------------------------------------------------------------
log(
  `\n  searchByFormat regex with no matches: ${Query.searchByFormat(d, { textMatch: /XYZNEVER/ }).length}`
);

fs.writeFileSync(path.join(outDir, "44-queries.txt"), lines.join("\n"));
console.log(`\n  → 44-queries.txt`);
