/**
 * Word Example 32 — Style resolution & style maps
 *
 * Covers:
 *   - resolveStyle: walk the basedOn chain + docDefaults to compute the
 *     effective ParagraphProperties + RunProperties.
 *   - resolveRunStyle: same for run-level character styles.
 *   - resolveNumberingLevel: turn a paragraph's numbering ref into the full
 *     NumberingLevel record.
 *   - resolveTableStyle: merge table style + conditional formats by
 *     row/col position.
 *   - resolveThemeColor: convert a theme color spec into a literal hex.
 *   - parseStyleMap / createStyleMap / mergeStyleMaps / matchStyleMap +
 *     DEFAULT_STYLE_MAP — the mammoth-style mapping DSL used by HTML/MD
 *     converters.
 *
 * Output: tmp/word-examples/32-styles/...
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { Document, Build, Io, Query, Styles, Theme, Units } from "../index";
import type { StyleDef, AbstractNumbering, NumberingInstance, DocumentTheme } from "../index";

const outDir = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  "../../../../tmp/word-examples/32-styles"
);
fs.mkdirSync(outDir, { recursive: true });

// ---------------------------------------------------------------------------
// Build a doc with a multi-level style chain and a theme so resolution has
// something to chew on.
// ---------------------------------------------------------------------------
const d = Document.create();
Document.useDefaultStyles(d);

// A grandparent → parent → child style chain. The child overrides only one
// property; the chain merge should fill in the rest.
const grandparent: StyleDef = {
  type: "paragraph",
  styleId: "Base",
  name: "Base",
  basedOn: "Normal",
  paragraphProperties: { spacing: { before: 240, after: 120 } },
  runProperties: { font: "Calibri", size: Units.ptToHalfPoint(11), color: "262626" }
};
const parent: StyleDef = {
  type: "paragraph",
  styleId: "ColoredBase",
  name: "Colored Base",
  basedOn: "Base",
  runProperties: { color: "1F4E79" }
};
const child: StyleDef = {
  type: "paragraph",
  styleId: "Lead",
  name: "Lead Paragraph",
  basedOn: "ColoredBase",
  paragraphProperties: { alignment: "both" },
  runProperties: { italic: true, size: Units.ptToHalfPoint(12) }
};
Document.addStyle(d, grandparent);
Document.addStyle(d, parent);
Document.addStyle(d, child);

// Linked character style on top
Document.addStyle(d, {
  type: "character",
  styleId: "Strong",
  name: "Strong",
  runProperties: { bold: true, color: "C00000" }
});

// Numbering definitions — for resolveNumberingLevel
const numAbs: AbstractNumbering = {
  abstractNumId: 50,
  multiLevelType: "multilevel",
  levels: [
    {
      level: 0,
      start: 1,
      format: "decimal",
      text: "%1.",
      paragraphProperties: { indent: { left: 720, hanging: 360 } }
    },
    {
      level: 1,
      start: 1,
      format: "lowerLetter",
      text: "%2)",
      paragraphProperties: { indent: { left: 1440, hanging: 360 } }
    }
  ]
};
const numInst: NumberingInstance = { numId: 50, abstractNumId: 50 };

Document.addParagraphElement(
  d,
  Build.paragraph(
    [
      Build.text("This paragraph uses the resolved Lead style; "),
      Build.text("Strong run", { style: "Strong" })
    ],
    { style: "Lead" }
  )
);

Document.addParagraphElement(
  d,
  Build.textParagraph("List item one", { numbering: { numId: 50, level: 0 } })
);
Document.addParagraphElement(
  d,
  Build.textParagraph("Sub-item a", { numbering: { numId: 50, level: 1 } })
);

// Add custom theme so resolveThemeColor returns a useful colour
const theme: DocumentTheme = {
  name: "Demo",
  colorScheme: {
    name: "Demo",
    colors: {
      dk1: "1F1F1F",
      lt1: "FFFFFF",
      dk2: "303030",
      lt2: "F2F2F2",
      accent1: "C0392B",
      accent2: "E67E22",
      accent3: "F1C40F",
      accent4: "27AE60",
      accent5: "2980B9",
      accent6: "8E44AD",
      hlink: "0563C1",
      folHlink: "954F72"
    }
  },
  fontScheme: {
    name: "Demo",
    majorFont: "Cambria",
    minorFont: "Calibri"
  }
};

// Build with theme attached and numbering merged in
const built = Document.build(d);
const docModel = {
  ...built,
  abstractNumberings: [...(built.abstractNumberings ?? []), numAbs],
  numberingInstances: [...(built.numberingInstances ?? []), numInst],
  theme
};

// ---------------------------------------------------------------------------
// 1. resolveStyle on a paragraph — walks the chain and merges
// ---------------------------------------------------------------------------
const leadPara = docModel.body.find(
  b => "type" in b && b.type === "paragraph" && b.properties?.style === "Lead"
);
if (leadPara && leadPara.type === "paragraph") {
  const resolved = Query.resolveStyle(docModel, leadPara);
  console.log(`  resolveStyle("Lead") chain: ${resolved.chain.join(" → ")}`);
  console.log(`    paragraphProperties: ${JSON.stringify(resolved.paragraphProperties)}`);
  console.log(`    runProperties:       ${JSON.stringify(resolved.runProperties)}`);
}

// ---------------------------------------------------------------------------
// 2. resolveRunStyle — character style chain + paragraph fallback
// ---------------------------------------------------------------------------
if (leadPara && leadPara.type === "paragraph") {
  const strongRun = leadPara.children.find(
    c => "properties" in c && c.properties?.style === "Strong"
  );
  if (strongRun && "content" in strongRun) {
    const resolved = Query.resolveRunStyle(docModel, strongRun);
    console.log(`  resolveRunStyle("Strong") chain: ${resolved.chain.join(" → ")}`);
    console.log(`    runProperties: ${JSON.stringify(resolved.runProperties)}`);
  }
}

// ---------------------------------------------------------------------------
// 3. resolveNumberingLevel — turn { numId, level } into the level definition
// ---------------------------------------------------------------------------
const listPara = docModel.body.find(
  b => "type" in b && b.type === "paragraph" && b.properties?.numbering?.numId === 50
);
if (listPara && listPara.type === "paragraph") {
  const lvl = Query.resolveNumberingLevel(docModel, listPara);
  console.log(
    `  resolveNumberingLevel: level=${lvl?.level}, format=${lvl?.format}, text="${lvl?.text}"`
  );
}

// ---------------------------------------------------------------------------
// 4. resolveTableStyle — walks the basedOn chain of a table style and
//    merges its paragraph / run / table properties.  Conditional formats
//    (firstRow / banding) are stored on the StyleDef but applied at render
//    time; resolveTableStyle returns the base merged properties.
// ---------------------------------------------------------------------------
const customTableStyle: StyleDef = {
  type: "table",
  styleId: "BandedGrid",
  name: "Banded Grid",
  basedOn: "TableNormal",
  tableProperties: {
    width: { value: 5000, type: "pct" },
    cellMargins: {
      top: { value: 60, type: "dxa" },
      bottom: { value: 60, type: "dxa" },
      left: { value: 100, type: "dxa" },
      right: { value: 100, type: "dxa" }
    }
  },
  runProperties: { font: "Calibri", size: Units.ptToHalfPoint(10) },
  tableStyleConditions: [
    { type: "firstRow", runProperties: { bold: true, color: "FFFFFF" } },
    { type: "evenRowBanding", cellProperties: { shading: { fill: "F2F2F2", pattern: "clear" } } }
  ]
};
const docWithTableStyle = { ...docModel, styles: [...(docModel.styles ?? []), customTableStyle] };
const resolvedTbl = Query.resolveTableStyle(docWithTableStyle, "BandedGrid");
console.log(`  resolveTableStyle("BandedGrid") chain: ${resolvedTbl.chain.join(" → ")}`);
console.log(`    runProperties: ${JSON.stringify(resolvedTbl.runProperties)}`);
console.log(`    tableProperties.width: ${JSON.stringify(resolvedTbl.tableProperties?.width)}`);

// ---------------------------------------------------------------------------
// 5. resolveThemeColor — resolves a ColorSpec against the document theme,
//    applying theme-color lookup + tint/shade.
// ---------------------------------------------------------------------------
console.log(
  `  resolveThemeColor(accent1):              ${Theme.resolveColor({ val: "auto", themeColor: "accent1" }, theme)}`
);
console.log(
  `  resolveThemeColor(accent1, tint 7F):     ${Theme.resolveColor({ val: "auto", themeColor: "accent1", themeTint: "7F" }, theme)}`
);
console.log(
  `  resolveThemeColor(accent1, shade 7F):    ${Theme.resolveColor({ val: "auto", themeColor: "accent1", themeShade: "7F" }, theme)}`
);
console.log(`  resolveThemeColor(plain hex "FF8800"):   ${Theme.resolveColor("FF8800", theme)}`);
console.log(`  resolveThemeColor(undefined):              ${Theme.resolveColor(undefined, theme)}`);

// ---------------------------------------------------------------------------
// 6. Style mapping DSL — used by HTML/Markdown converters
// ---------------------------------------------------------------------------
const userMap = Styles.parse(
  `
  // mammoth-style DSL
  p[style-name='Lead'] => p.lead
  p[style-name='Heading 1'] => h1.title
  r[style-name='Strong'] => strong.attention
`,
  { includeDefaults: true }
);
console.log(`  parseStyleMap rules: ${userMap.rules.length}`);

const programmaticMap = Styles.create([
  {
    source: "p",
    conditions: [{ attribute: "style-name", value: "Quote" }],
    target: { tagName: "blockquote", className: "fancy" },
    priority: 5
  }
]);
const merged = Styles.merge(Styles.DEFAULT, userMap, programmaticMap);
console.log(`  mergeStyleMaps: ${merged.rules.length} total rules`);

const matched = Styles.match(merged, "p", { "style-name": "Lead" });
console.log(`  matchStyleMap('p[Lead]') → ${matched?.tagName}.${matched?.className ?? ""}`);
const matched2 = Styles.match(merged, "r", { "style-name": "Strong" });
console.log(`  matchStyleMap('r[Strong]') → ${matched2?.tagName}.${matched2?.className ?? ""}`);
const noMatch = Styles.match(merged, "p", { "style-name": "NoSuchStyle" });
console.log(`  matchStyleMap unknown → ${noMatch}`);

// Save the underlying document for visual inspection
const buf = await Io.toBuffer(docModel);
fs.writeFileSync(path.join(outDir, "01-styled.docx"), buf);
console.log(`  → 01-styled.docx (${buf.length} bytes)`);
