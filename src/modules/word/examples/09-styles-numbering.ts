/**
 * Word Example 09 — Styles & numbering definitions
 *
 * Covers:
 *   - Document defaults (docDefaults)
 *   - Custom paragraph style (with basedOn / next)
 *   - Custom character (run) style
 *   - Custom table style with banded rows/cols
 *   - Style by reference (paragraph.style = "MyQuote")
 *   - Custom numbering (legal/upperRoman/asian/picture-like bullets)
 *   - Linked styles (paragraph + character pair)
 *   - Edge case: style with same id as built-in (override)
 *
 * Output: tmp/word-examples/09-styles-numbering.docx
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Units } from "../index";
import type { StyleDef, AbstractNumbering, NumberingInstance } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples"
);
fs.mkdirSync(outDir, { recursive: true });

const doc = Document.create();
Document.useDefaultStyles(doc);

// ---------------------------------------------------------------------------
// 1. Document defaults — applied to every paragraph that doesn't override
// ---------------------------------------------------------------------------
Document.setDocDefaults(doc, {
  runProperties: {
    font: { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "Microsoft YaHei", cs: "Arial" },
    size: Units.ptToHalfPoint(11),
    color: "262626"
  },
  paragraphProperties: {
    spacing: { before: 0, after: 120, line: 276, lineRule: "auto" }
  }
});

// ---------------------------------------------------------------------------
// 2. Custom paragraph style — "Quote"
// ---------------------------------------------------------------------------
const quoteStyle: StyleDef = {
  type: "paragraph",
  styleId: "MyQuote",
  name: "My Quote",
  basedOn: "Normal",
  next: "Normal",
  qFormat: true,
  uiPriority: 30,
  paragraphProperties: {
    indent: { left: Units.cmToTwips(1), right: Units.cmToTwips(1) },
    spacing: { before: 240, after: 240 },
    borders: {
      left: { style: "single", size: 16, color: "1F4E79", space: 8 }
    }
  },
  runProperties: {
    italic: true,
    color: "1F4E79",
    size: Units.ptToHalfPoint(11)
  }
};
Document.addStyle(doc, quoteStyle);

// ---------------------------------------------------------------------------
// 3. Custom character style — "MyEmphasis"
// ---------------------------------------------------------------------------
const emphasisStyle: StyleDef = {
  type: "character",
  styleId: "MyEmphasis",
  name: "MyEmphasis",
  uiPriority: 40,
  qFormat: true,
  runProperties: {
    bold: true,
    color: "C00000",
    underline: "single"
  }
};
Document.addStyle(doc, emphasisStyle);

// ---------------------------------------------------------------------------
// 4. Custom table style with banding & header
// ---------------------------------------------------------------------------
const tableStyle: StyleDef = {
  type: "table",
  styleId: "MyGrid",
  name: "My Grid",
  basedOn: "TableNormal",
  uiPriority: 39,
  tableProperties: {
    borders: Build.gridBorders(4, "BFBFBF"),
    // One row per banding stripe so band1Horz/band2Horz alternate every row.
    rowBandSize: 1,
    cellMargins: {
      top: { value: 60, type: "dxa" },
      bottom: { value: 60, type: "dxa" },
      left: { value: 100, type: "dxa" },
      right: { value: 100, type: "dxa" }
    }
  },
  tableStyleConditions: [
    {
      type: "firstRow",
      runProperties: { bold: true, color: "FFFFFF" },
      cellProperties: { shading: { fill: "1F4E79", pattern: "clear" } }
    },
    {
      // Odd stripes (rows 1,3,… of the body) — white, for explicit contrast.
      type: "oddRowBanding",
      cellProperties: { shading: { fill: "FFFFFF", pattern: "clear" } }
    },
    {
      // Even stripes (rows 2,4,…) — light grey zebra banding.
      type: "evenRowBanding",
      cellProperties: { shading: { fill: "F2F2F2", pattern: "clear" } }
    }
  ]
};
Document.addStyle(doc, tableStyle);

// ---------------------------------------------------------------------------
// 5. Linked styles (paragraph "Lead" + char "LeadChar")
// ---------------------------------------------------------------------------
Document.addStyle(doc, {
  type: "paragraph",
  styleId: "Lead",
  name: "Lead",
  basedOn: "Normal",
  next: "Normal",
  link: "LeadChar",
  qFormat: true,
  paragraphProperties: { spacing: { before: 0, after: 240 }, alignment: "both" },
  runProperties: { size: Units.ptToHalfPoint(13), color: "595959" }
});
Document.addStyle(doc, {
  type: "character",
  styleId: "LeadChar",
  name: "Lead Char",
  link: "Lead",
  basedOn: "Normal",
  runProperties: { size: Units.ptToHalfPoint(13), color: "595959" }
});

// ---------------------------------------------------------------------------
// Body content using the styles above
// ---------------------------------------------------------------------------
Document.addHeading(doc, "Word — Custom Styles & Numbering", 1);

Document.addParagraphElement(
  doc,
  Build.paragraph(
    [
      Build.text("This first paragraph uses the "),
      Build.text("Lead", { style: "LeadChar" }),
      Build.text(" paragraph style.")
    ],
    { style: "Lead" }
  )
);

// MyQuote
Document.addParagraphElement(
  doc,
  Build.paragraph(
    [
      Build.text(
        "“The only way to learn a new programming language is by writing programs in it.” — Brian Kernighan"
      )
    ],
    { style: "MyQuote" }
  )
);

// Inline character style
Document.addParagraphElement(
  doc,
  Build.paragraph([
    Build.text("Run formatting via "),
    Build.text("MyEmphasis", { style: "MyEmphasis" }),
    Build.text(" character style; this is independent of any paragraph style.")
  ])
);

// Custom table style
Document.addHeading(doc, "Custom table style with banding", 2);
Document.addTableElement(
  doc,
  Build.table(
    [
      Build.row([Build.cell("Quarter"), Build.cell("Revenue"), Build.cell("Profit")]),
      Build.row([Build.cell("Q1"), Build.cell("$1.2M"), Build.cell("$200K")]),
      Build.row([Build.cell("Q2"), Build.cell("$1.5M"), Build.cell("$280K")]),
      Build.row([Build.cell("Q3"), Build.cell("$1.8M"), Build.cell("$340K")]),
      Build.row([Build.cell("Q4"), Build.cell("$2.1M"), Build.cell("$410K")])
    ],
    {
      style: "MyGrid",
      width: { value: 5000, type: "pct" },
      look: { firstRow: true, lastRow: false, firstColumn: false, noHBand: false }
    }
  )
);

// ---------------------------------------------------------------------------
// 6. Custom numbering — Asian-numeral list & picture-like Wingdings bullets
// ---------------------------------------------------------------------------
const asianNumbering: AbstractNumbering = {
  abstractNumId: 500,
  multiLevelType: "singleLevel",
  levels: [
    {
      level: 0,
      start: 1,
      format: "ideographDigital",
      text: "%1、",
      justification: "left",
      paragraphProperties: { indent: { left: 720, hanging: 360 } }
    }
  ]
};
const asianInstance: NumberingInstance = { numId: 500, abstractNumId: 500 };

const fancyBullets: AbstractNumbering = {
  abstractNumId: 501,
  multiLevelType: "singleLevel",
  levels: [
    {
      level: 0,
      start: 1,
      format: "bullet",
      text: "\uF0D8", // Wingdings ➤-style chevron
      justification: "left",
      paragraphProperties: { indent: { left: 720, hanging: 360 } },
      runProperties: { font: { ascii: "Wingdings", hAnsi: "Wingdings" }, color: "C00000" }
    }
  ]
};
const fancyInstance: NumberingInstance = { numId: 501, abstractNumId: 501 };

const built = Document.build(doc);
const final = {
  ...built,
  abstractNumberings: [...(built.abstractNumberings ?? []), asianNumbering, fancyBullets],
  numberingInstances: [...(built.numberingInstances ?? []), asianInstance, fancyInstance],
  body: [
    ...built.body,
    Build.textParagraph("Asian-numeral list:", { style: "Heading2" }),
    Build.textParagraph("第一项内容", { numbering: { numId: 500, level: 0 } }),
    Build.textParagraph("第二项内容", { numbering: { numId: 500, level: 0 } }),
    Build.textParagraph("第三项内容", { numbering: { numId: 500, level: 0 } }),

    Build.textParagraph("Fancy chevron bullets:", { style: "Heading2" }),
    Build.textParagraph("Foo", { numbering: { numId: 501, level: 0 } }),
    Build.textParagraph("Bar", { numbering: { numId: 501, level: 0 } }),
    Build.textParagraph("Baz", { numbering: { numId: 501, level: 0 } })
  ]
};

const buf = await Io.toBuffer(final);
fs.writeFileSync(path.join(outDir, "09-styles-numbering.docx"), buf);
console.log(`  → 09-styles-numbering.docx (${buf.length} bytes)`);
