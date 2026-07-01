/**
 * DOCX Reader - Paragraph and Section Properties Parsers
 *
 * Pure parsers for `w:pPr` (paragraph properties) and `w:sectPr` (section
 * properties). Both are large and have inter-dependencies (paragraphs may
 * embed section properties); they're co-located here.
 */

import type { Mutable } from "@word/core/internal-utils";
import {
  attrInt,
  attrVal,
  boolToggle,
  findChildNs,
  findChildrenNs,
  parseNoteProperties
} from "@word/reader/parse-utils";
import {
  parseBorder,
  parseRevisionInfo,
  parseRunProperties,
  parseShading
} from "@word/reader/properties-parsers";
import type {
  Border,
  HeaderFooterRef,
  Indentation,
  LineSpacing,
  ParagraphBorders,
  ParagraphFrame,
  ParagraphProperties,
  PageBorders,
  SectionColumns,
  SectionProperties,
  SectionPropertyChange,
  TabStop
} from "@word/types";
import type { XmlElement } from "@xml/types";

// =============================================================================
// Paragraph Properties
// =============================================================================

function parseParagraphProperties(pPrEl: XmlElement): ParagraphProperties {
  const pPr: Mutable<ParagraphProperties> = {};

  const pStyleEl = findChildNs(pPrEl, "pStyle");
  if (pStyleEl) {
    pPr.style = attrVal(pStyleEl, "val");
  }

  const jcEl = findChildNs(pPrEl, "jc");
  if (jcEl) {
    pPr.alignment = attrVal(jcEl, "val") as ParagraphProperties["alignment"];
  }

  if (findChildNs(pPrEl, "keepNext")) {
    pPr.keepNext = true;
  }
  if (findChildNs(pPrEl, "keepLines")) {
    pPr.keepLines = true;
  }
  if (findChildNs(pPrEl, "pageBreakBefore")) {
    pPr.pageBreakBefore = true;
  }
  if (findChildNs(pPrEl, "bidi")) {
    pPr.bidi = true;
  }

  // New boolean toggles
  const ctxSp = boolToggle(pPrEl, "contextualSpacing");
  if (ctxSp !== undefined) {
    pPr.contextualSpacing = ctxSp;
  }
  const suppLn = boolToggle(pPrEl, "suppressLineNumbers");
  if (suppLn !== undefined) {
    pPr.suppressLineNumbers = suppLn;
  }
  const suppHyph = boolToggle(pPrEl, "suppressAutoHyphens");
  if (suppHyph !== undefined) {
    pPr.suppressAutoHyphens = suppHyph;
  }
  const mirr = boolToggle(pPrEl, "mirrorIndents");
  if (mirr !== undefined) {
    pPr.mirrorIndents = mirr;
  }
  const wc = boolToggle(pPrEl, "widowControl");
  if (wc !== undefined) {
    pPr.widowControl = wc;
  }
  const ww = boolToggle(pPrEl, "wordWrap");
  if (ww !== undefined) {
    pPr.wordWrap = ww;
  }
  const stg = boolToggle(pPrEl, "snapToGrid");
  if (stg !== undefined) {
    pPr.snapToGrid = stg;
  }
  const ofp = boolToggle(pPrEl, "overflowPunct");
  if (ofp !== undefined) {
    pPr.overflowPunctuation = ofp;
  }
  const topLinePunct = boolToggle(pPrEl, "topLinePunct");
  if (topLinePunct !== undefined) {
    pPr.topLinePunctuation = topLinePunct;
  }
  const kinsoku = boolToggle(pPrEl, "kinsoku");
  if (kinsoku !== undefined) {
    pPr.kinsoku = kinsoku;
  }
  const asd = boolToggle(pPrEl, "autoSpaceDE");
  if (asd !== undefined) {
    pPr.autoSpaceEastAsianText = asd;
  }
  const asdn = boolToggle(pPrEl, "autoSpaceDN");
  if (asdn !== undefined) {
    pPr.autoSpaceEastAsianDigit = asdn;
  }

  const textAlignEl = findChildNs(pPrEl, "textAlignment");
  if (textAlignEl) {
    pPr.textAlignment = attrVal(textAlignEl, "val") as ParagraphProperties["textAlignment"];
  }

  const outlineLvlEl = findChildNs(pPrEl, "outlineLvl");
  if (outlineLvlEl) {
    pPr.outlineLevel = attrInt(outlineLvlEl, "val");
  }

  const textDirEl = findChildNs(pPrEl, "textDirection");
  if (textDirEl) {
    pPr.textDirection = attrVal(textDirEl, "val") as ParagraphProperties["textDirection"];
  }

  // Paragraph frame
  const framePrEl = findChildNs(pPrEl, "framePr");
  if (framePrEl) {
    const f: Partial<Mutable<ParagraphFrame>> = {};
    const dropCap = attrVal(framePrEl, "dropCap");
    if (dropCap) {
      f.dropCap = dropCap as ParagraphFrame["dropCap"];
    }
    const lines = attrInt(framePrEl, "lines");
    if (lines !== undefined) {
      f.lines = lines;
    }
    const fw = attrInt(framePrEl, "w");
    if (fw !== undefined) {
      f.width = fw;
    }
    const fh = attrInt(framePrEl, "h");
    if (fh !== undefined) {
      f.height = fh;
    }
    const hSpace = attrInt(framePrEl, "hSpace");
    if (hSpace !== undefined) {
      f.hSpace = hSpace;
    }
    const vSpace = attrInt(framePrEl, "vSpace");
    if (vSpace !== undefined) {
      f.vSpace = vSpace;
    }
    const wrap = attrVal(framePrEl, "wrap");
    if (wrap) {
      f.wrap = wrap as ParagraphFrame["wrap"];
    }
    const hAnchor = attrVal(framePrEl, "hAnchor");
    if (hAnchor) {
      f.hAnchor = hAnchor as ParagraphFrame["hAnchor"];
    }
    const vAnchor = attrVal(framePrEl, "vAnchor");
    if (vAnchor) {
      f.vAnchor = vAnchor as ParagraphFrame["vAnchor"];
    }
    const x = attrInt(framePrEl, "x");
    if (x !== undefined) {
      f.x = x;
    }
    const xAlign = attrVal(framePrEl, "xAlign");
    if (xAlign) {
      f.xAlign = xAlign as ParagraphFrame["xAlign"];
    }
    const y = attrInt(framePrEl, "y");
    if (y !== undefined) {
      f.y = y;
    }
    const yAlign = attrVal(framePrEl, "yAlign");
    if (yAlign) {
      f.yAlign = yAlign as ParagraphFrame["yAlign"];
    }
    pPr.frame = f as ParagraphFrame;
  }

  // Thematic break: check for bottom border with special pattern
  const pBdrEl = findChildNs(pPrEl, "pBdr");
  if (pBdrEl) {
    const borders: Partial<ParagraphBorders> = {};
    for (const side of ["top", "bottom", "left", "right", "between", "bar"] as const) {
      const sideEl = findChildNs(pBdrEl, side);
      if (sideEl) {
        (borders as Record<string, Border>)[side] = parseBorder(sideEl);
      }
    }
    pPr.borders = borders;
  }

  const spacingEl = findChildNs(pPrEl, "spacing");
  if (spacingEl) {
    const spacing: Partial<Mutable<LineSpacing>> = {};
    const before = attrInt(spacingEl, "before");
    const after = attrInt(spacingEl, "after");
    const line = attrInt(spacingEl, "line");
    const lineRule = attrVal(spacingEl, "lineRule");
    if (before !== undefined) {
      spacing.before = before;
    }
    if (after !== undefined) {
      spacing.after = after;
    }
    if (line !== undefined) {
      spacing.line = line;
      // Per ECMA-376, w:lineRule defaults to "auto" when line is set
      spacing.lineRule = (lineRule ?? "auto") as LineSpacing["lineRule"];
    } else if (lineRule) {
      spacing.lineRule = lineRule as LineSpacing["lineRule"];
    }
    const beforeAuto = attrVal(spacingEl, "beforeAutospacing");
    if (beforeAuto === "1" || beforeAuto === "true") {
      spacing.beforeAutoSpacing = true;
    }
    const afterAuto = attrVal(spacingEl, "afterAutospacing");
    if (afterAuto === "1" || afterAuto === "true") {
      spacing.afterAutoSpacing = true;
    }
    pPr.spacing = spacing;
  }

  const indEl = findChildNs(pPrEl, "ind");
  if (indEl) {
    const indent: Partial<Mutable<Indentation>> = {};
    const left = attrInt(indEl, "left");
    const right = attrInt(indEl, "right");
    const hanging = attrInt(indEl, "hanging");
    const firstLine = attrInt(indEl, "firstLine");
    const start = attrInt(indEl, "start");
    const end = attrInt(indEl, "end");
    if (left !== undefined) {
      indent.left = left;
    }
    if (right !== undefined) {
      indent.right = right;
    }
    if (hanging !== undefined) {
      indent.hanging = hanging;
    }
    if (firstLine !== undefined) {
      indent.firstLine = firstLine;
    }
    if (start !== undefined) {
      indent.start = start;
    }
    if (end !== undefined) {
      indent.end = end;
    }
    pPr.indent = indent;
  }

  const numPrEl = findChildNs(pPrEl, "numPr");
  if (numPrEl) {
    const ilvlEl = findChildNs(numPrEl, "ilvl");
    const numIdEl = findChildNs(numPrEl, "numId");
    // Per OOXML schema, numId is required but ilvl is optional (defaults to 0).
    if (numIdEl) {
      pPr.numbering = {
        level: ilvlEl ? (attrInt(ilvlEl, "val") ?? 0) : 0,
        numId: attrInt(numIdEl, "val") ?? 0
      };
    }
  }

  const tabsEl = findChildNs(pPrEl, "tabs");
  if (tabsEl) {
    const tabs: TabStop[] = [];
    for (const tabEl of findChildrenNs(tabsEl, "tab")) {
      tabs.push({
        type: (attrVal(tabEl, "val") ?? "left") as TabStop["type"],
        position: attrInt(tabEl, "pos") ?? 0,
        leader: attrVal(tabEl, "leader") as TabStop["leader"]
      });
    }
    if (tabs.length > 0) {
      pPr.tabs = tabs;
    }
  }

  const shdEl = findChildNs(pPrEl, "shd");
  if (shdEl) {
    pPr.shading = parseShading(shdEl);
  }

  const rPrEl = findChildNs(pPrEl, "rPr");
  if (rPrEl) {
    pPr.markRunProperties = parseRunProperties(rPrEl);
  }

  const sectPrEl = findChildNs(pPrEl, "sectPr");
  if (sectPrEl) {
    pPr.sectionProperties = parseSectionProperties(sectPrEl);
  }

  // Conditional formatting style mask
  const cnfStyleEl = findChildNs(pPrEl, "cnfStyle");
  if (cnfStyleEl) {
    pPr.cnfStyle = attrVal(cnfStyleEl, "val");
  }

  // Paragraph property change
  const pPrChangeEl = findChildNs(pPrEl, "pPrChange");
  if (pPrChangeEl) {
    const rev = parseRevisionInfo(pPrChangeEl);
    if (rev) {
      const prevPPrEl = findChildNs(pPrChangeEl, "pPr");
      pPr.propertyChange = {
        revision: rev,
        previousProperties: prevPPrEl ? parseParagraphProperties(prevPPrEl) : undefined
      };
    }
  }

  // Paragraph mark insertion/deletion (w:pPr > w:rPr > w:ins/w:del)
  const rPrInPPr = findChildNs(pPrEl, "rPr");
  if (rPrInPPr) {
    const insEl = findChildNs(rPrInPPr, "ins");
    if (insEl) {
      const rev = parseRevisionInfo(insEl);
      if (rev) {
        pPr.paragraphInsertion = rev;
      }
    }
    const delEl = findChildNs(rPrInPPr, "del");
    if (delEl) {
      const rev = parseRevisionInfo(delEl);
      if (rev) {
        pPr.paragraphDeletion = rev;
      }
    }
  }

  return pPr;
}

// =============================================================================
// Section Properties
// =============================================================================

function parseSectionProperties(sectPrEl: XmlElement): SectionProperties {
  const sect: Mutable<SectionProperties> = {};

  const pgSzEl = findChildNs(sectPrEl, "pgSz");
  if (pgSzEl) {
    // Per ECMA-376, w:orient defaults to "portrait" when absent
    const orient = attrVal(pgSzEl, "orient");
    // Per ECMA-376 17.6.13, w:w/h are required attributes; default to US Letter
    // (12240 × 15840 twips) when malformed input omits them.
    sect.pageSize = {
      width: attrInt(pgSzEl, "w") ?? 12240,
      height: attrInt(pgSzEl, "h") ?? 15840,
      orientation: orient === "landscape" ? "landscape" : "portrait"
    };
  }

  const pgMarEl = findChildNs(sectPrEl, "pgMar");
  if (pgMarEl) {
    sect.margins = {
      top: attrInt(pgMarEl, "top") ?? 1440,
      right: attrInt(pgMarEl, "right") ?? 1440,
      bottom: attrInt(pgMarEl, "bottom") ?? 1440,
      left: attrInt(pgMarEl, "left") ?? 1440,
      header: attrInt(pgMarEl, "header"),
      footer: attrInt(pgMarEl, "footer"),
      gutter: attrInt(pgMarEl, "gutter")
    };
  }

  const typeEl = findChildNs(sectPrEl, "type");
  if (typeEl) {
    sect.breakType = attrVal(typeEl, "val") as SectionProperties["breakType"];
  }

  const colsEl = findChildNs(sectPrEl, "cols");
  if (colsEl) {
    const cols: Mutable<SectionColumns> = {};
    cols.space = attrInt(colsEl, "space");
    cols.count = attrInt(colsEl, "num");
    const eqw = attrVal(colsEl, "equalWidth");
    if (eqw !== undefined) {
      cols.equalWidth = eqw === "1";
    }
    const sep = attrVal(colsEl, "sep");
    if (sep === "1" || sep === "true") {
      cols.separator = true;
    }
    const colDefs = findChildrenNs(colsEl, "col");
    if (colDefs.length > 0) {
      cols.columns = colDefs.map(c => ({
        width: attrInt(c, "w") ?? 0,
        space: attrInt(c, "space")
      }));
    }
    sect.columns = cols;
  }

  if (findChildNs(sectPrEl, "titlePg")) {
    sect.titlePage = true;
  }

  const pgNumEl = findChildNs(sectPrEl, "pgNumType");
  if (pgNumEl) {
    sect.pageNumbering = {
      start: attrInt(pgNumEl, "start"),
      format: attrVal(pgNumEl, "fmt") as NonNullable<SectionProperties["pageNumbering"]>["format"]
    };
  }

  // Page borders
  const pgBordersEl = findChildNs(sectPrEl, "pgBorders");
  if (pgBordersEl) {
    const pb: Mutable<PageBorders> = {};
    for (const side of ["top", "left", "bottom", "right"] as const) {
      const sideEl = findChildNs(pgBordersEl, side);
      if (sideEl) {
        pb[side] = parseBorder(sideEl);
      }
    }
    const display = attrVal(pgBordersEl, "display");
    if (display) {
      pb.display = display as PageBorders["display"];
    }
    const offsetFrom = attrVal(pgBordersEl, "offsetFrom");
    if (offsetFrom) {
      pb.offsetFrom = offsetFrom as PageBorders["offsetFrom"];
    }
    const zOrder = attrVal(pgBordersEl, "zOrder");
    if (zOrder) {
      pb.zOrder = zOrder as PageBorders["zOrder"];
    }
    sect.pageBorders = pb;
  }

  // Vertical alignment
  const vAlignEl = findChildNs(sectPrEl, "vAlign");
  if (vAlignEl) {
    sect.verticalAlign = attrVal(vAlignEl, "val") as SectionProperties["verticalAlign"];
  }

  // Text direction
  const textDirEl = findChildNs(sectPrEl, "textDirection");
  if (textDirEl) {
    sect.textDirection = attrVal(textDirEl, "val") as SectionProperties["textDirection"];
  }

  // Bidi
  const bidiToggle = boolToggle(sectPrEl, "bidi");
  if (bidiToggle !== undefined) {
    sect.bidi = bidiToggle;
  }

  // RTL gutter
  const rtlGutterEl = findChildNs(sectPrEl, "rtlGutter");
  if (rtlGutterEl) {
    sect.rtlGutter = true;
  }

  // Form protection
  const formProtEl = findChildNs(sectPrEl, "formProt");
  if (formProtEl) {
    const v = attrVal(formProtEl, "val");
    sect.formProtection = v === "1" || v === "true";
  }

  // Document grid
  const docGridEl = findChildNs(sectPrEl, "docGrid");
  if (docGridEl) {
    type DocGrid = NonNullable<SectionProperties["docGrid"]>;
    const dg: Mutable<DocGrid> = {};
    const linePitch = attrInt(docGridEl, "linePitch");
    if (linePitch !== undefined) {
      dg.linePitch = linePitch;
    }
    const charSpace = attrInt(docGridEl, "charSpace");
    if (charSpace !== undefined) {
      dg.charSpace = charSpace;
    }
    const gridType = attrVal(docGridEl, "type") as DocGrid["type"];
    if (gridType) {
      dg.type = gridType;
    }
    sect.docGrid = dg;
  }

  // Line numbers
  const lnNumEl = findChildNs(sectPrEl, "lnNumType");
  if (lnNumEl) {
    type LineNumbers = NonNullable<SectionProperties["lineNumbers"]>;
    const ln: Mutable<LineNumbers> = {};
    const countBy = attrInt(lnNumEl, "countBy");
    if (countBy !== undefined) {
      ln.countBy = countBy;
    }
    const start = attrInt(lnNumEl, "start");
    if (start !== undefined) {
      ln.start = start;
    }
    const restart = attrVal(lnNumEl, "restart") as LineNumbers["restart"];
    if (restart) {
      ln.restart = restart;
    }
    const distance = attrInt(lnNumEl, "distance");
    if (distance !== undefined) {
      ln.distance = distance;
    }
    sect.lineNumbers = ln;
  }

  // Footnote properties
  const fnPrEl = findChildNs(sectPrEl, "footnotePr");
  if (fnPrEl) {
    sect.footnoteProperties = parseNoteProperties(fnPrEl) as
      | SectionProperties["footnoteProperties"]
      | undefined;
  }

  // Endnote properties
  const enPrEl = findChildNs(sectPrEl, "endnotePr");
  if (enPrEl) {
    sect.endnoteProperties = parseNoteProperties(enPrEl) as
      | SectionProperties["endnoteProperties"]
      | undefined;
  }

  // Headers/Footers refs
  const headerRefs: HeaderFooterRef[] = [];
  for (const hRef of findChildrenNs(sectPrEl, "headerReference")) {
    headerRefs.push({
      type: (attrVal(hRef, "type") ?? "default") as HeaderFooterRef["type"],
      rId: hRef.attributes["r:id"] ?? ""
    });
  }
  if (headerRefs.length > 0) {
    sect.headers = headerRefs;
  }

  const footerRefs: HeaderFooterRef[] = [];
  for (const fRef of findChildrenNs(sectPrEl, "footerReference")) {
    footerRefs.push({
      type: (attrVal(fRef, "type") ?? "default") as HeaderFooterRef["type"],
      rId: fRef.attributes["r:id"] ?? ""
    });
  }
  if (footerRefs.length > 0) {
    sect.footers = footerRefs;
  }

  // sectPrChange (track changes for section properties)
  const sectPrChangeEl = findChildNs(sectPrEl, "sectPrChange");
  if (sectPrChangeEl) {
    const rev = parseRevisionInfo(sectPrChangeEl);
    if (rev) {
      const prevSectPrEl = findChildNs(sectPrChangeEl, "sectPr");
      const change: SectionPropertyChange = {
        revision: rev,
        previousProperties: prevSectPrEl ? parseSectionProperties(prevSectPrEl) : undefined
      };
      sect.propertyChange = change;
    }
  }

  return sect;
}

export { parseParagraphProperties, parseSectionProperties };
