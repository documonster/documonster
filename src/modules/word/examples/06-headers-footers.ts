/**
 * Word Example 06 — Headers & Footers
 *
 * Covers:
 *   - Default header / footer (applies to all pages without other overrides)
 *   - First-page-different header (titlePage)
 *   - Even / odd page headers
 *   - Header with logo (image), aligned right
 *   - Footer with PAGE / NUMPAGES fields ("Page X of Y")
 *   - Footer with date field
 *   - 3-cell footer table (left/center/right alignment trick)
 *   - Different headers per section (continuous section break)
 *
 * Output: tmp/word-examples/06-headers-footers.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Units } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

// Body content — multiple pages so headers/footers are visible
Document.addHeading(doc, "Headers & Footers demo", 1);
Document.addParagraph(doc, "First page body. ");
Document.addParagraph(doc, "Lorem ipsum… ".repeat(50));
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "Page 2", 2);
Document.addParagraph(doc, "Even page; check the even-page header.");
Document.addParagraph(doc, "Lorem ipsum… ".repeat(50));
Document.addParagraphElement(doc, Build.paragraph([Build.pageBreak()]));
Document.addHeading(doc, "Page 3", 2);
Document.addParagraph(doc, "Odd page; default odd-page header should appear.");
Document.addParagraph(doc, "Lorem ipsum… ".repeat(50));

// ---------------------------------------------------------------------------
// Default (odd-page) header — title + tab-aligned date on the right
// ---------------------------------------------------------------------------
Document.setHeader(doc, "default", {
  children: [
    Build.paragraph(
      [
        Build.bold("My Report"),
        Build.tab(),
        Build.tab(),
        Build.text("Date: "),
        Build.dateField("MMMM d, yyyy")
      ],
      {
        tabs: [{ position: 9072, type: "right" }],
        style: "Header"
      }
    )
  ]
});

// ---------------------------------------------------------------------------
// Even-page header — mirrored, document title on the right
// ---------------------------------------------------------------------------
Document.setHeader(doc, "even", {
  children: [
    Build.paragraph([Build.text("v1.0"), Build.tab(), Build.bold("My Report")], {
      tabs: [{ position: 9072, type: "right" }],
      style: "Header"
    })
  ]
});

// ---------------------------------------------------------------------------
// First-page header — "Draft" stamp
// ---------------------------------------------------------------------------
Document.setHeader(doc, "first", {
  children: [
    Build.paragraph(
      [
        Build.text("DRAFT — INTERNAL USE ONLY", {
          color: "C00000",
          bold: true,
          caps: true
        })
      ],
      { alignment: "center", style: "Header" }
    )
  ]
});

// ---------------------------------------------------------------------------
// Footer (default) — three-cell table for left/center/right alignment.
// Real-world DOCX style for "Page X of Y" with author on the left.
// ---------------------------------------------------------------------------
const footerTable = Build.table(
  [
    Build.row([
      Build.cell([Build.textParagraph("© 2026 Example Co.")], {
        width: { value: 33, type: "pct" },
        borders: {
          top: Build.border("nil"),
          bottom: Build.border("nil"),
          left: Build.border("nil"),
          right: Build.border("nil")
        }
      }),
      Build.cell(
        [
          Build.paragraph(
            [
              Build.text("Page "),
              Build.pageNumberField(),
              Build.text(" of "),
              Build.totalPagesField()
            ],
            {
              alignment: "center"
            }
          )
        ],
        {
          width: { value: 34, type: "pct" },
          borders: {
            top: Build.border("nil"),
            bottom: Build.border("nil"),
            left: Build.border("nil"),
            right: Build.border("nil")
          }
        }
      ),
      Build.cell([Build.paragraph([Build.text("Confidential")], { alignment: "right" })], {
        width: { value: 33, type: "pct" },
        borders: {
          top: Build.border("nil"),
          bottom: Build.border("nil"),
          left: Build.border("nil"),
          right: Build.border("nil")
        }
      })
    ])
  ],
  {
    width: { value: 5000, type: "pct" },
    borders: {
      top: Build.border("nil"),
      bottom: Build.border("nil"),
      left: Build.border("nil"),
      right: Build.border("nil"),
      insideH: Build.border("nil"),
      insideV: Build.border("nil")
    }
  }
);
Document.setFooter(doc, "default", { children: [footerTable] });

// First-page footer (different)
Document.setFooter(doc, "first", {
  children: [
    Build.textParagraph("(this footer appears only on the title page)", { alignment: "center" })
  ]
});

// Even-page footer
Document.setFooter(doc, "even", {
  children: [
    Build.paragraph([Build.pageNumberField(), Build.tab(), Build.text("My Report")], {
      tabs: [{ position: 9072, type: "right" }],
      style: "Footer"
    })
  ]
});

// Tell the section to honour title-page (first) and even/odd headers.
Document.setSectionProperties(doc, {
  pageSize: { width: Units.cmToTwips(21), height: Units.cmToTwips(29.7) },
  margins: {
    top: Units.cmToTwips(2.5),
    bottom: Units.cmToTwips(2.5),
    left: Units.cmToTwips(2.5),
    right: Units.cmToTwips(2.5),
    header: Units.cmToTwips(1.25),
    footer: Units.cmToTwips(1.25)
  },
  titlePage: true
});

// Tell the document settings to enable evenAndOddHeaders globally.
Document.setSettings(doc, { evenAndOddHeaders: true });

const buf = await Io.toBuffer(Document.build(doc));
fs.writeFileSync(path.join(outDir, "06-headers-footers.docx"), buf);
console.log(`  → 06-headers-footers.docx (${buf.length} bytes)`);
