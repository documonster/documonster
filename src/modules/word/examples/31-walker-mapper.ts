/**
 * Word Example 31 — Walker & Mapper (transform / collect)
 *
 * Covers the low-level traversal API surface:
 *   - walkDocument / walkBlocks — visitor pattern with VisitAction
 *     ("continue" / "skip" / "stop")
 *   - WalkPath context (section, depth, inHeader, inFooter, …)
 *   - collectParagraphs / collectRuns / collectTables — convenience
 *   - mapDocument with DocxTransformer — pure functional rewrite
 *     (replace / remove / expand body content)
 *   - Visiting headers, footers, footnotes, endnotes via WalkOptions
 *
 * Edge cases:
 *   - "skip" prunes a subtree without aborting traversal
 *   - "stop" aborts cleanly mid-walk
 *   - mapDocument never mutates the input
 *   - Removing every paragraph leaves an empty (but valid) document
 *
 * Output: tmp/word-examples/31-walker/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  textParagraph,
  text,
  bold,
  italic,
  hyperlink,
  toBuffer,
  walkDocument,
  walkBlocks,
  collectParagraphs,
  collectRuns,
  collectTables,
  mapDocument
} from "../index";
import type { DocxDocument, DocxTransformer, DocxVisitor, BodyContent } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/31-walker"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Build a sample document with headers, footers, footnotes, comments
// ---------------------------------------------------------------------------
function makeDoc(): DocxDocument {
  const d = Document.create();
  Document.useDefaultStyles(d);
  Document.addHeading(d, "Walker demo", 1);
  Document.addHeading(d, "First section", 2);
  Document.addParagraph(d, "Body paragraph 1.");
  Document.addParagraphElement(
    d,
    paragraph([
      text("Body paragraph 2 with "),
      bold("bold"),
      text(", "),
      italic("italic"),
      text(", and "),
      hyperlink("a link", { url: "https://example.com" }),
      text(".")
    ])
  );
  Document.addTable(
    d,
    [
      ["A", "B"],
      ["1", "2"]
    ],
    { headerRow: true, borders: true }
  );

  // Footnote
  const fnId = Document.addFootnote(d, "Footnote text");
  Document.addParagraphElement(
    d,
    paragraph([
      text("Cite"),
      { properties: { vertAlign: "superscript" }, content: [{ type: "footnoteRef", id: fnId }] }
    ])
  );

  // Header
  Document.setHeader(d, "default", { children: [textParagraph("Header line")] });
  Document.setFooter(d, "default", { children: [textParagraph("Footer line")] });

  return Document.build(d);
}

const doc = makeDoc();

// ---------------------------------------------------------------------------
// 1. Convenience collectors
// ---------------------------------------------------------------------------
console.log(`  collectParagraphs(): ${collectParagraphs(doc).length}`);
console.log(`  collectRuns():       ${collectRuns(doc).length}`);
console.log(`  collectTables():     ${collectTables(doc).length}`);

// Without headers/footers/notes
const bodyOnlyParas = collectParagraphs(doc, {
  includeHeaders: false,
  includeFooters: false,
  includeFootnotes: false
});
console.log(`  body-only paragraphs: ${bodyOnlyParas.length}`);

// ---------------------------------------------------------------------------
// 2. Custom visitor — count by element kind, with "skip" pruning tables
// ---------------------------------------------------------------------------
const counts = { paragraphs: 0, tables: 0, runs: 0, hyperlinks: 0, footnoteRefs: 0 };
const visitor: DocxVisitor = {
  enterParagraph(_, path) {
    counts.paragraphs++;
    if (path.inFootnote) {
      // Don't double-count footnote paragraphs as body paragraphs.
    }
  },
  enterTable() {
    counts.tables++;
    return "skip"; // Don't descend into rows/cells/runs of tables
  },
  enterRun() {
    counts.runs++;
  },
  enterHyperlink() {
    counts.hyperlinks++;
  },
  visitRunContent(content) {
    if (content.type === "footnoteRef") {
      counts.footnoteRefs++;
    }
  }
};
walkDocument(doc, visitor, {
  includeHeaders: true,
  includeFooters: true,
  includeFootnotes: true
});
console.log(`  visitor counts: ${JSON.stringify(counts)}`);

// ---------------------------------------------------------------------------
// 3. "stop" — abort early
// ---------------------------------------------------------------------------
let visitedBeforeStop = 0;
walkDocument(doc, {
  enterParagraph() {
    visitedBeforeStop++;
    if (visitedBeforeStop >= 2) {
      return "stop";
    }
    return "continue";
  }
});
console.log(`  paragraphs visited before stop: ${visitedBeforeStop}`);

// ---------------------------------------------------------------------------
// 4. walkBlocks on a sub-tree (no document context required)
// ---------------------------------------------------------------------------
const fragment: BodyContent[] = [textParagraph("a"), textParagraph("b"), textParagraph("c")];
let visited = 0;
walkBlocks(fragment, {
  enterParagraph() {
    visited++;
  }
});
console.log(`  walkBlocks(fragment): ${visited} paragraphs`);

// ---------------------------------------------------------------------------
// 5. mapDocument — uppercase every text run, drop tables entirely
// ---------------------------------------------------------------------------
const transformer: DocxTransformer = {
  transformRun(run) {
    if (!run.content.some(c => c.type === "text")) {
      return run;
    }
    return {
      ...run,
      content: run.content.map(c => (c.type === "text" ? { ...c, text: c.text.toUpperCase() } : c))
    };
  },
  transformTable() {
    // Drop the table entirely from the output
    return null;
  }
};
const transformed = mapDocument(doc, transformer);

// Verify: input was not mutated
const originalParas = collectParagraphs(doc, { includeFootnotes: false });
const originalText = originalParas.map(p =>
  p.children
    .map(c =>
      "content" in c
        ? c.content
            .filter(rc => rc.type === "text")
            .map(rc => (rc as { text: string }).text)
            .join("")
        : ""
    )
    .join("")
);
console.log(`  pristine input still has lowercase: ${originalText.some(t => /[a-z]/.test(t))}`);

const transformedBuf = await toBuffer(transformed);
fs.writeFileSync(path.join(outDir, "01-uppercased-no-tables.docx"), transformedBuf);
console.log(`  → 01-uppercased-no-tables.docx (${transformedBuf.length} bytes)`);

// ---------------------------------------------------------------------------
// 6. mapDocument with body content expansion (one paragraph → many)
// ---------------------------------------------------------------------------
const expander: DocxTransformer = {
  transformBodyContent(block) {
    if (block.type === "paragraph") {
      const txt = block.children
        .filter(c => "content" in c)
        .flatMap(c =>
          (c as { content: { type: string; text?: string }[] }).content
            .filter(rc => rc.type === "text")
            .map(rc => rc.text ?? "")
        )
        .join("");
      if (txt.startsWith("EXPAND:")) {
        // Replace this paragraph with three new ones
        return [
          textParagraph(`A — ${txt.slice(7)}`),
          textParagraph(`B — ${txt.slice(7)}`),
          textParagraph(`C — ${txt.slice(7)}`)
        ];
      }
    }
    return block;
  }
};
const seed = Document.create();
Document.useDefaultStyles(seed);
Document.addParagraph(seed, "Plain");
Document.addParagraph(seed, "EXPAND:hello");
Document.addParagraph(seed, "Plain again");
const expanded = mapDocument(Document.build(seed), expander);
console.log(`  expander: 3 → ${expanded.body.length} body items`);
fs.writeFileSync(path.join(outDir, "02-expanded.docx"), await toBuffer(expanded));
console.log(`  → 02-expanded.docx`);

// ---------------------------------------------------------------------------
// 7. Edge case: removing every paragraph still produces a valid empty doc
// ---------------------------------------------------------------------------
const emptied = mapDocument(doc, {
  transformBodyContent() {
    return null;
  }
});
console.log(`  emptied body length: ${emptied.body.length}`);
fs.writeFileSync(path.join(outDir, "03-empty-after-map.docx"), await toBuffer(emptied));
console.log(`  → 03-empty-after-map.docx`);
