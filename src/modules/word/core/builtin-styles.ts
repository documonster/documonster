/**
 * Word's built-in styles, defined as Word defines them when one is first used
 * (its latent styles made real): the set `Document.useDefaultStyles` adds, and
 * the definitions the writer supplies for a built-in style id that content
 * references but the document does not define — so a `Heading1` paragraph is
 * a heading to Word (its outline level, its TOC entry) rather than a dangling
 * reference Word reads as Normal.
 */

import { walkDocument } from "@word/core/walker";
import type { Border, DocxDocument, StyleDef, TableBorders } from "@word/types";

const GRID_LINE: Border = { style: "single", size: 4, space: 0, color: "auto" };
const GRID_BORDERS: TableBorders = {
  top: GRID_LINE,
  left: GRID_LINE,
  bottom: GRID_LINE,
  right: GRID_LINE,
  insideH: GRID_LINE,
  insideV: GRID_LINE
};

/** The style set `Document.useDefaultStyles` adds. */
export function defaultStyleSet(): StyleDef[] {
  return [
    { type: "paragraph", styleId: "Normal", name: "Normal", isDefault: true, qFormat: true },
    {
      type: "paragraph",
      styleId: "Heading1",
      name: "heading 1",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      paragraphProperties: {
        outlineLevel: 0,
        keepNext: true,
        keepLines: true,
        spacing: { before: 240, after: 0 }
      },
      runProperties: { font: "Calibri Light", color: "2F5496", size: 32 }
    },
    {
      type: "paragraph",
      styleId: "Heading2",
      name: "heading 2",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      paragraphProperties: {
        outlineLevel: 1,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { font: "Calibri Light", color: "2F5496", size: 26 }
    },
    {
      type: "paragraph",
      styleId: "Heading3",
      name: "heading 3",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      paragraphProperties: {
        outlineLevel: 2,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { font: "Calibri Light", color: "1F3763", size: 24 }
    },
    // Heading 4-9 — Word's built-in latent styles. Required so that every
    // call to addHeading(doc, text, level) where level > 3 resolves to a
    // defined paragraph style instead of leaving an undefined pStyle ref.
    {
      type: "paragraph",
      styleId: "Heading4",
      name: "heading 4",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 3,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { italic: true, color: "2F5496", size: 22 }
    },
    {
      type: "paragraph",
      styleId: "Heading5",
      name: "heading 5",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 4,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { color: "2F5496", size: 22 }
    },
    {
      type: "paragraph",
      styleId: "Heading6",
      name: "heading 6",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 5,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { italic: true, color: "1F3763", size: 22 }
    },
    {
      type: "paragraph",
      styleId: "Heading7",
      name: "heading 7",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 6,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { color: "1F3763", size: 22 }
    },
    {
      type: "paragraph",
      styleId: "Heading8",
      name: "heading 8",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 7,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { color: "272727", size: 21 }
    },
    {
      type: "paragraph",
      styleId: "Heading9",
      name: "heading 9",
      basedOn: "Normal",
      next: "Normal",
      qFormat: true,
      uiPriority: 9,
      unhideWhenUsed: true,
      paragraphProperties: {
        outlineLevel: 8,
        keepNext: true,
        keepLines: true,
        spacing: { before: 40, after: 0 }
      },
      runProperties: { italic: true, color: "272727", size: 21 }
    },
    {
      // Word's zero-formatting base character style. Every other character
      // style (Hyperlink, HeaderChar, FooterChar, MyEmphasis, …) bases on
      // it; without this anchor, Word may decline to apply rStyle
      // references emitted by field/run writers.
      type: "character",
      styleId: "DefaultParagraphFont",
      name: "Default Paragraph Font",
      isDefault: true,
      uiPriority: 1,
      semiHidden: true,
      unhideWhenUsed: true
    },
    {
      type: "character",
      styleId: "Hyperlink",
      name: "Hyperlink",
      basedOn: "DefaultParagraphFont",
      uiPriority: 99,
      unhideWhenUsed: true,
      runProperties: { color: "0563C1", underline: "single" }
    },
    {
      // Word's built-in "visited hyperlink" character style. Referenced by
      // hyperlinks created with `history: true` so visited links render in
      // the standard purple instead of the unvisited blue.
      type: "character",
      styleId: "FollowedHyperlink",
      name: "FollowedHyperlink",
      basedOn: "DefaultParagraphFont",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      runProperties: { color: "954F72", underline: "single" }
    },
    {
      // Word's built-in zero-formatting base table style. Required so
      // that styles like TableGrid (which sets `basedOn: "TableNormal"`)
      // resolve cleanly without dangling-reference validation warnings.
      // It is the default table style: without `w:default`, Word gives a
      // table with no style of its own no cell padding at all.
      type: "table",
      styleId: "TableNormal",
      name: "Normal Table",
      isDefault: true,
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      tableProperties: {
        cellMargins: {
          top: { value: 0, type: "dxa" },
          bottom: { value: 0, type: "dxa" },
          left: { value: 108, type: "dxa" },
          right: { value: 108, type: "dxa" }
        }
      }
    },
    {
      type: "table",
      styleId: "TableGrid",
      name: "Table Grid",
      basedOn: "TableNormal",
      uiPriority: 39,
      tableProperties: { borders: GRID_BORDERS }
    },
    {
      type: "numbering",
      styleId: "NoList",
      name: "No List",
      isDefault: true,
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true
    },
    // Header / Footer paragraph styles. Word ships these as latent
    // built-ins; emitting them explicitly avoids "unknown style" warnings
    // when callers attach `style: "Header"` / `"Footer"` to header/footer
    // paragraphs (e.g. page-number footers, multi-cell title-bars).
    // The default tabs (center at 4536 twips, right at 9072 twips) are
    // what Word's "Insert > Header" preset uses so headers/footers laid
    // out with center+right tabs render in the expected positions.
    {
      type: "paragraph",
      styleId: "Header",
      name: "header",
      basedOn: "Normal",
      link: "HeaderChar",
      uiPriority: 99,
      unhideWhenUsed: true,
      paragraphProperties: {
        spacing: { after: 0, line: 240, lineRule: "auto" },
        tabs: [
          { position: 4536, type: "center" },
          { position: 9072, type: "right" }
        ]
      }
    },
    {
      type: "character",
      styleId: "HeaderChar",
      name: "Header Char",
      basedOn: "DefaultParagraphFont",
      link: "Header",
      uiPriority: 99
    },
    {
      type: "paragraph",
      styleId: "Footer",
      name: "footer",
      basedOn: "Normal",
      link: "FooterChar",
      uiPriority: 99,
      unhideWhenUsed: true,
      paragraphProperties: {
        spacing: { after: 0, line: 240, lineRule: "auto" },
        tabs: [
          { position: 4536, type: "center" },
          { position: 9072, type: "right" }
        ]
      }
    },
    {
      type: "character",
      styleId: "FooterChar",
      name: "Footer Char",
      basedOn: "DefaultParagraphFont",
      link: "Footer",
      uiPriority: 99
    }
  ];
}

/** A `toc N` paragraph style (`TOC1`..`TOC9`). */
export function tocStyle(level: number): StyleDef {
  return {
    type: "paragraph",
    styleId: `TOC${level}`,
    name: `toc ${level}`,
    basedOn: "Normal",
    next: "Normal",
    autoRedefine: true,
    uiPriority: 39,
    unhideWhenUsed: true,
    paragraphProperties: {
      spacing: { after: 100 },
      ...(level > 1 ? { indent: { left: (level - 1) * 220 } } : {})
    }
  };
}

/** Built-in styles beyond the default set, as Word defines them on first use. */
function moreBuiltInStyles(): StyleDef[] {
  return [
    {
      type: "paragraph",
      styleId: "Title",
      name: "Title",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 10,
      qFormat: true,
      paragraphProperties: {
        spacing: { after: 0, line: 240, lineRule: "auto" },
        contextualSpacing: true
      },
      runProperties: { font: "Calibri Light", spacing: -10, kern: 28, size: 56 }
    },
    {
      type: "paragraph",
      styleId: "Subtitle",
      name: "Subtitle",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 11,
      qFormat: true,
      runProperties: { color: "5A5A5A", spacing: 15 }
    },
    {
      type: "paragraph",
      styleId: "Quote",
      name: "Quote",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 29,
      qFormat: true,
      paragraphProperties: {
        spacing: { before: 200, after: 160 },
        indent: { left: 864, right: 864 },
        alignment: "center"
      },
      runProperties: { italic: true, color: "404040" }
    },
    {
      type: "paragraph",
      styleId: "IntenseQuote",
      name: "Intense Quote",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 30,
      qFormat: true,
      paragraphProperties: {
        spacing: { before: 360, after: 360 },
        indent: { left: 864, right: 864 },
        alignment: "center"
      },
      runProperties: { italic: true, color: "4472C4" }
    },
    {
      type: "paragraph",
      styleId: "NoSpacing",
      name: "No Spacing",
      uiPriority: 1,
      qFormat: true,
      paragraphProperties: { spacing: { after: 0, line: 240, lineRule: "auto" } }
    },
    {
      type: "paragraph",
      styleId: "ListParagraph",
      name: "List Paragraph",
      basedOn: "Normal",
      uiPriority: 34,
      qFormat: true,
      paragraphProperties: { indent: { left: 720 }, contextualSpacing: true }
    },
    {
      type: "paragraph",
      styleId: "Caption",
      name: "caption",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 35,
      unhideWhenUsed: true,
      qFormat: true,
      paragraphProperties: { spacing: { after: 200, line: 240, lineRule: "auto" } },
      runProperties: { italic: true, color: "44546A", size: 18 }
    },
    ...[1, 2, 3, 4, 5, 6, 7, 8, 9].map(tocStyle),
    {
      type: "paragraph",
      styleId: "TOCHeading",
      name: "TOC Heading",
      basedOn: "Heading1",
      next: "Normal",
      uiPriority: 39,
      unhideWhenUsed: true,
      qFormat: true,
      paragraphProperties: { outlineLevel: 9 }
    },
    {
      type: "paragraph",
      styleId: "TableofFigures",
      name: "table of figures",
      basedOn: "Normal",
      next: "Normal",
      uiPriority: 99,
      unhideWhenUsed: true,
      paragraphProperties: { spacing: { after: 0 } }
    },
    {
      type: "paragraph",
      styleId: "FootnoteText",
      name: "footnote text",
      basedOn: "Normal",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      paragraphProperties: { spacing: { after: 0, line: 240, lineRule: "auto" } },
      runProperties: { size: 20 }
    },
    {
      type: "character",
      styleId: "FootnoteReference",
      name: "footnote reference",
      basedOn: "DefaultParagraphFont",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      runProperties: { vertAlign: "superscript" }
    },
    {
      type: "paragraph",
      styleId: "EndnoteText",
      name: "endnote text",
      basedOn: "Normal",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      paragraphProperties: { spacing: { after: 0, line: 240, lineRule: "auto" } },
      runProperties: { size: 20 }
    },
    {
      type: "character",
      styleId: "EndnoteReference",
      name: "endnote reference",
      basedOn: "DefaultParagraphFont",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      runProperties: { vertAlign: "superscript" }
    },
    {
      type: "paragraph",
      styleId: "CommentText",
      name: "annotation text",
      basedOn: "Normal",
      uiPriority: 99,
      unhideWhenUsed: true,
      paragraphProperties: { spacing: { line: 240, lineRule: "auto" } },
      runProperties: { size: 20 }
    },
    {
      type: "character",
      styleId: "CommentReference",
      name: "annotation reference",
      basedOn: "DefaultParagraphFont",
      uiPriority: 99,
      semiHidden: true,
      unhideWhenUsed: true,
      runProperties: { size: 16 }
    }
  ];
}

let catalogue: ReadonlyMap<string, StyleDef> | undefined;

/** The built-in style Word defines under `styleId`, or `undefined` when it is not one. */
export function builtInStyle(styleId: string): StyleDef | undefined {
  catalogue ??= new Map([...defaultStyleSet(), ...moreBuiltInStyles()].map(s => [s.styleId, s]));
  return catalogue.get(styleId);
}

/**
 * `styles` with a definition added for every built-in style id in `referenced`
 * it lacks — and for the built-in styles those are based on, link to or are
 * followed by — or `styles` itself when none is missing. A defined style is
 * never replaced, a built-in already defined under another id (a localized
 * document's `heading 1`) is not defined twice, an added default style (`Normal`,
 * `TableNormal`, …) is not made the default when another style of its type is, and an id that is not a
 * built-in one is left to the caller.
 */
export function withReferencedBuiltInStyles(
  styles: readonly StyleDef[] | undefined,
  referenced: Iterable<string>
): readonly StyleDef[] | undefined {
  const defined = new Set((styles ?? []).map(s => s.styleId));
  const names = new Set((styles ?? []).map(s => s.name.toLowerCase()));
  const defaultTypes = new Set((styles ?? []).filter(s => s.isDefault).map(s => s.type));
  const added: StyleDef[] = [];
  const add = (id: string | undefined): void => {
    if (id === undefined || defined.has(id)) {
      return;
    }
    const builtIn = builtInStyle(id);
    if (!builtIn || names.has(builtIn.name.toLowerCase())) {
      return;
    }
    const style =
      builtIn.isDefault && defaultTypes.has(builtIn.type)
        ? { ...builtIn, isDefault: false }
        : builtIn;
    if (style.isDefault) {
      defaultTypes.add(style.type);
    }
    defined.add(id);
    added.push(style);
    add(style.basedOn);
    add(style.link);
    add(style.next);
  };
  for (const id of referenced) {
    add(id);
  }
  return added.length === 0 ? styles : [...(styles ?? []), ...added];
}

/**
 * Every style id the content of `doc` references — paragraph, run and table
 * styles in every part, list levels' paragraph styles, and `Header` for the
 * paragraph the writer draws a watermark in.
 */
export function referencedStyleIds(doc: DocxDocument): Set<string> {
  const ids = new Set<string>();
  const add = (id: string | undefined): void => {
    if (id) {
      ids.add(id);
    }
  };
  walkDocument(
    doc,
    {
      enterParagraph(para) {
        add(para.properties?.style);
      },
      enterRun(run) {
        add(run.properties?.style);
      },
      enterTable(table) {
        add(table.properties?.style);
      },
      visitTableOfContents(toc) {
        for (const para of toc.cachedParagraphs ?? []) {
          add(para.properties?.style);
        }
      }
    },
    { includeComments: true }
  );
  for (const abstract of doc.abstractNumberings ?? []) {
    for (const level of abstract.levels) {
      add(level.paragraphStyle);
    }
  }
  if (doc.watermark) {
    ids.add("Header");
  }
  return ids;
}
