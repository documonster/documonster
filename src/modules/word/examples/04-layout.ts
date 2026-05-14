/**
 * Word Example 04 — Page Layout
 *
 * Covers section / page-level configuration:
 *   - Standard sizes (A4 / Letter / Legal / B5 / A3) and orientation
 *   - Margins (mirror / gutter for binding)
 *   - Multi-column sections with separator
 *   - Multiple sections via section breaks (nextPage / continuous)
 *   - Page borders (artistic, plain, decorative-on-first-page-only)
 *   - Page vertical alignment (top/center/bottom)
 *   - Line numbering, page numbering format & start value
 *   - Document grid (East-Asian)
 *   - Bidi (right-to-left section)
 *
 * Output: tmp/word-examples/04-layout.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  Document,
  paragraph,
  text,
  pageBreak,
  columnBreak,
  pageNumberField,
  cmToTwips,
  inchesToTwips,
  toBuffer
} from "../index";
import type { SectionProperties } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Helpers — common page sizes (twips). Word stores landscape by swapping
// width/height AND setting orientation="landscape".
// ---------------------------------------------------------------------------
const A4_PORTRAIT = { width: cmToTwips(21), height: cmToTwips(29.7) };
const A4_LANDSCAPE = {
  width: cmToTwips(29.7),
  height: cmToTwips(21),
  orientation: "landscape" as const
};
const LETTER = { width: inchesToTwips(8.5), height: inchesToTwips(11) };
const A3 = { width: cmToTwips(29.7), height: cmToTwips(42) };

// ---------------------------------------------------------------------------
// Document 1: A4 portrait, mirror margins, gutter
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "A4 portrait — bindable layout", 1);
  Document.addParagraph(
    doc,
    "Mirrored margins via gutter; the inner edge is wider for binding. " +
      "Body text fills the printable area."
  );
  Document.addParagraph(doc, "Lorem ipsum… ".repeat(60));

  const sect: SectionProperties = {
    pageSize: A4_PORTRAIT,
    margins: {
      top: cmToTwips(2.5),
      bottom: cmToTwips(2.5),
      left: cmToTwips(3.0),
      right: cmToTwips(2.0),
      gutter: cmToTwips(1.0),
      header: cmToTwips(1.25),
      footer: cmToTwips(1.25)
    }
  };
  Document.setSectionProperties(doc, sect);
  const buf = await toBuffer(Document.build(doc));
  fs.writeFileSync(path.join(outDir, "04-layout-a4.docx"), buf);
  console.log(`  → 04-layout-a4.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// Document 2: Multi-section document
//   Section 1: Letter portrait, single column
//   Section 2: Letter landscape, two equal columns with separator
//   Section 3: A3 portrait, three unequal columns
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);

  // Section 1 content
  Document.addHeading(doc, "Section 1 — Letter Portrait", 1);
  Document.addParagraph(doc, "Plain single-column section. " + "Lorem ipsum… ".repeat(20));

  // End of section 1: attach SectionProperties via addSectionBreak
  Document.addSectionBreak(doc, {
    pageSize: LETTER,
    margins: {
      top: inchesToTwips(1),
      bottom: inchesToTwips(1),
      left: inchesToTwips(1),
      right: inchesToTwips(1)
    },
    breakType: "nextPage"
  });

  // Section 2 content
  Document.addHeading(doc, "Section 2 — Letter Landscape, 2 columns", 1);
  for (let i = 0; i < 3; i++) {
    Document.addParagraph(
      doc,
      `Left column paragraph ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. ` +
        "Vivamus lacinia odio vitae vestibulum vestibulum. ".repeat(2)
    );
  }
  // Force a break to the second column
  Document.addParagraphElement(doc, paragraph([columnBreak()]));
  for (let i = 0; i < 3; i++) {
    Document.addParagraph(
      doc,
      `Right column paragraph ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. ` +
        "Vivamus lacinia odio vitae vestibulum vestibulum. ".repeat(2)
    );
  }
  Document.addSectionBreak(doc, {
    pageSize: {
      ...LETTER,
      width: inchesToTwips(11),
      height: inchesToTwips(8.5),
      orientation: "landscape"
    },
    margins: {
      top: inchesToTwips(1),
      bottom: inchesToTwips(1),
      left: inchesToTwips(1),
      right: inchesToTwips(1)
    },
    breakType: "nextPage",
    columns: { count: 2, space: 720, equalWidth: true, separator: true }
  });

  // Section 3 content
  Document.addHeading(doc, "Section 3 — A3 Portrait, 3 unequal columns", 1);
  for (let i = 0; i < 9; i++) {
    Document.addParagraph(
      doc,
      `A3 column ${i + 1}. ` + "Sed do eiusmod tempor incididunt ut labore et dolore. ".repeat(3)
    );
  }
  // Final section properties live on the document, not via section break
  Document.setSectionProperties(doc, {
    pageSize: A3,
    margins: {
      top: cmToTwips(2),
      bottom: cmToTwips(2),
      left: cmToTwips(2),
      right: cmToTwips(2)
    },
    columns: {
      equalWidth: false,
      space: 360,
      columns: [
        { width: cmToTwips(8), space: 360 },
        { width: cmToTwips(8), space: 360 },
        { width: cmToTwips(8.7) }
      ],
      separator: false
    }
  });

  const buf = await toBuffer(Document.build(doc));
  fs.writeFileSync(path.join(outDir, "04-layout-multi-section.docx"), buf);
  console.log(`  → 04-layout-multi-section.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// Document 3: Page borders, page vertical alignment, line numbering, page-num format
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);

  Document.addHeading(doc, "Decorative page", 1);
  Document.addParagraph(
    doc,
    "This page has artistic borders, vertical-centered text, and Roman page numbers."
  );
  Document.addParagraphElement(doc, paragraph([pageBreak()]));
  Document.addParagraph(doc, "Second page — line numbers should appear on the left margin.");
  for (let i = 0; i < 30; i++) {
    Document.addParagraph(doc, `Numbered line ${i + 1}.`);
  }

  // pgNumType only declares the *format* of any page number rendered in
  // a footer — it does NOT auto-emit the number. Without an explicit footer
  // containing a PAGE field, no page numbers appear in the document at all.
  Document.setFooter(doc, "default", {
    children: [paragraph([pageNumberField()], { alignment: "center", style: "Footer" })]
  });

  Document.setSectionProperties(doc, {
    pageSize: A4_PORTRAIT,
    margins: {
      top: cmToTwips(3),
      bottom: cmToTwips(3),
      left: cmToTwips(3),
      right: cmToTwips(3)
    },
    pageBorders: {
      display: "allPages",
      offsetFrom: "page",
      top: { style: "double", size: 12, color: "1F4E79", space: 24 },
      left: { style: "double", size: 12, color: "1F4E79", space: 24 },
      bottom: { style: "double", size: 12, color: "1F4E79", space: 24 },
      right: { style: "double", size: 12, color: "1F4E79", space: 24 }
    },
    verticalAlign: "top",
    pageNumbering: { start: 1, format: "lowerRoman" },
    lineNumbers: { countBy: 1, start: 1, restart: "newPage", distance: cmToTwips(0.5) }
  });

  const buf = await toBuffer(Document.build(doc));
  fs.writeFileSync(path.join(outDir, "04-layout-decorative.docx"), buf);
  console.log(`  → 04-layout-decorative.docx (${buf.length} bytes)`);
}

// ---------------------------------------------------------------------------
// Document 4: Landscape A4, document grid (CJK), RTL section
// ---------------------------------------------------------------------------
{
  const doc = Document.create();
  Document.useDefaultStyles(doc);
  Document.addHeading(doc, "横向 A4 + 文档网格", 1);
  Document.addParagraph(
    doc,
    "本节启用 East-Asian 文档网格 — 字符按固定列宽对齐,这是中文/日文排版常用设置。" +
      "示例段落填充以观察网格效果。".repeat(8)
  );

  Document.addSectionBreak(doc, {
    pageSize: A4_LANDSCAPE,
    margins: { top: cmToTwips(2), bottom: cmToTwips(2), left: cmToTwips(2), right: cmToTwips(2) },
    breakType: "nextPage",
    docGrid: { type: "linesAndChars", linePitch: 312, charSpace: 0 }
  });

  // Note: section-level bidi sets the section's default direction, but each
  // paragraph still needs its own bidi flag to be rendered RTL by Word —
  // the section default is not auto-inherited by paragraphs that lack the
  // property. Apply bidi on every paragraph in this RTL section, including
  // the heading.
  Document.addParagraphElement(
    doc,
    paragraph([text("Right-to-left section")], { style: "Heading1", bidi: true })
  );
  Document.addParagraphElement(
    doc,
    paragraph([text("هذه فقرة في قسم RTL.  Latin words وكلمات عربية mixed.")], { bidi: true })
  );
  Document.addParagraphElement(doc, paragraph([text("النص الثاني في هذا القسم.")], { bidi: true }));

  Document.setSectionProperties(doc, {
    pageSize: A4_PORTRAIT,
    margins: { top: cmToTwips(2), bottom: cmToTwips(2), left: cmToTwips(2), right: cmToTwips(2) },
    bidi: true
  });

  const buf = await toBuffer(Document.build(doc));
  fs.writeFileSync(path.join(outDir, "04-layout-landscape-grid-rtl.docx"), buf);
  console.log(`  → 04-layout-landscape-grid-rtl.docx (${buf.length} bytes)`);
}
