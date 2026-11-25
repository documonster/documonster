/**
 * The record grammar of the XLSB parts this library writes, from MS-XLSB, and the positions a record takes in it.
 *
 * Each part's record sequence is an ABNF production in the specification (§2.1.7.x). It is copied here
 * **verbatim** — production names, brackets and repetition counts exactly as published, link targets removed — and
 * parsed, rather than restated as a hand-made table. A list someone rewrote from the grammar is a second
 * implementation of it that can be wrong the same way the writer is; a copy can be checked against the page by
 * eye, character by character. Three of the positions this check first caught had been placed by inference from a
 * reference workbook that lacked the neighbour in question: the merge cells written ahead of the auto filter, the
 * sparklines ahead of the drawing, and the pivot cache ids after `BrtFileRecover`.
 *
 * Every scope whose production is published is ranked, the root's and those inside it (`grammarScopes`): a view's
 * pane before its selections, a book's supporting links before `BrtExternSheet`, a table's auto filter before its
 * columns. Within a scope, a production that opens a scope of its own (`MERGECELLS = BrtBeginMergeCells …
 * BrtEndMergeCells`) occupies one position, named by its opening record, and so does a scope written inline; a
 * production that does not (`FRTWORKSHEET`, a run of optional future-record collections) is inlined.
 * `BrtFRTBegin` and `BrtACBegin` wrap rather than open — the scope check treats them as transparent for the same
 * reason — so a collection behind one is named by the record after it.
 *
 * Flattening follows the XML content models' (`ooxml-validator/ooxml-sequences.ts`): alternatives are merged
 * position by position, a repeated group of several positions collapses to one, `FRT` and `ACUID` (which
 * may appear almost anywhere) contribute nothing, and a record that the grammar places at two positions is left
 * unranked — lenient, never wrong.
 *
 * The common productions of §2.1.8 are not published as text (the section names an internal `.abnf` file). The few
 * a top-level grammar uses are listed in {@link COMMON_OPENERS} by the record that opens them — the one part written
 * by hand, held to the record dictionary by a test and to Excel's own files by the corpus — and a scope whose
 * content uses one is not ranked inside, since there is nothing published to rank it by.
 */

import { ExcelError } from "@excel/errors";
import { MS_XLSB_RECORD_NAMES } from "@excel/xlsb/spec/record-names";
import { withoutAmbiguousNames } from "@xml/sequence-order";

/** A part's grammar: the root production, and the productions it uses that its own section defines. */
export interface XlsbGrammar {
  /** Where it is published. */
  readonly source: string;
  readonly root: string;
  readonly productions: Readonly<Record<string, string>>;
}

export const WORKSHEET_GRAMMAR: XlsbGrammar = {
  source:
    "MS-XLSB §2.1.7.62 Worksheet, https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/d9599ff2-ac37-4651-8d96-f36097f8a09a",
  root: "WORKSHEET",
  productions: {
    WORKSHEET:
      "BrtBeginSheet [BrtWsProp] [BrtWsDim] [WSVIEWS2] [WSFMTINFO] *COLINFOS CELLTABLE [BrtSheetCalcProp] [[BrtSheetProtectionIso] BrtSheetProtection] *([BrtRangeProtectionIso] BrtRangeProtection) [SCENMAN] [AUTOFILTER] [SORTSTATE] [DCON] [USERSHVIEWS] [MERGECELLS] [BrtPhoneticInfo] *CONDITIONALFORMATTING [DVALS] *([ACUID] BrtHLink) [BrtPrintOptions] [BrtMargins] [BrtPageSetup] [HEADERFOOTER] [RWBRK] [COLBRK] *BrtBigName [CELLWATCHES] [IGNOREECS] [SMARTTAGS] [BrtDrawing] [BrtLegacyDrawing] [BrtLegacyDrawingHF] [BrtBkHim] [OLEOBJECTS] [ACTIVEXCONTROLS] [WEBPUBITEMS] [LISTPARTS] FRTWORKSHEET [ACUID] BrtEndSheet",
    WSVIEWS2: "BrtBeginWsViews 1*WSVIEW2 *FRT BrtEndWsViews",
    WSVIEW2: "[ACUID] BrtBeginWsView [BrtPane] *4BrtSel *4SXSELECT *FRT BrtEndWsView",
    SXSELECT: "BrtBeginSxSelect PIVOTRULE BrtEndSxSelect",
    SCT: "BrtBeginSct 1*32BrtSlc BrtEndSct",
    SCENMAN: "BrtBeginScenMan 1*32767SCT BrtEndScenMan",
    MERGECELLS: "BrtBeginMergeCells 1*BrtMergeCell BrtEndMergeCells",
    DVALS: "BrtBeginDVals 1*65534([ACUID] [ACDVALLIST] BrtDVal) BrtEndDVals",
    CELLWATCHES: "BrtBeginCellWatches 1*BrtCellWatch BrtEndCellWatches",
    IGNOREECS: "BrtBeginCellIgnoreECs 1*BrtCellIgnoreEC *FRT BrtEndCellIgnoreECs",
    SMARTTAGS: "BrtBeginSmartTags 1*CELLSMARTTAGS BrtEndSmartTags",
    CELLSMARTTAGS: "BrtBeginCellSmartTags 1*CELLSMARTTAG BrtEndCellSmartTags",
    CELLSMARTTAG: "BrtBeginCellSmartTag *BrtCellSmartTagProperty BrtEndCellSmartTag",
    ACTIVEXCONTROLS: "BrtBeginActiveXControls 1*BrtActiveX BrtEndActiveXControls",
    LISTPARTS: "BrtBeginListParts *BrtListPart BrtEndListParts",
    FRTWORKSHEET:
      "[CONDITIONALFORMATTINGS] [DVALS14] [SPARKLINEGROUPS] [SLICERSEX] [RANGEPROTECTION14] [IGNOREECS14] [WEBEXTENSIONS] [TABLESLICERSEX] [TIMELINESEX] *FRT",
    CONDITIONALFORMATTINGS:
      "BrtFRTBegin BrtBeginConditionalFormattings *2147483647CONDITIONALFORMATTING14 BrtEndConditionalFormattings BrtFRTEnd",
    DVALS14: "BrtFRTBegin BrtBeginDVals14 1*65534([ACDVALLIST] BrtDVal14) BrtEndDVals14 BrtFRTEnd",
    SPARKLINEGROUPS:
      "BrtFRTBegin BrtBeginSparklineGroups 1*2147483647SPARKLINEGROUP BrtEndSparklineGroups BrtFRTEnd",
    SPARKLINEGROUP:
      "BrtBeginSparklineGroup [ACUID] BrtBeginSparklines 1*2147483647BrtSparkline BrtEndSparklines BrtEndSparklineGroup",
    ACDVALLIST: "BrtACBegin BrtDValList BrtACEnd",
    SLICEREX: "BrtBeginSlicerEx BrtEndSlicerEx",
    TABLESLICEREX: "BrtBeginSlicerEx BrtEndSlicerEx",
    TIMELINEEX: "BrtBeginTimelineEx BrtEndTimelineEx",
    SLICERSEX: "BrtFRTBegin BrtBeginSlicersEx SLICEREX BrtEndSlicersEx BrtFRTEnd",
    TABLESLICERSEX: "BrtFRTBegin BrtBeginSlicersEx TABLESLICEREX BrtEndSlicersEx BrtFRTEnd",
    RANGEPROTECTION14: "BrtFRTBegin 1*([BrtRangeProtectionIso14] BrtRangeProtection14) BrtFRTEnd",
    IGNOREECS14:
      "BrtFRTBegin BrtBeginCellIgnoreECs14 1*BrtCellIgnoreEC14 BrtEndCellIgnoreECs14 BrtFRTEnd",
    TIMELINESEX: "BrtFRTBegin BrtBeginTimelinesEx TIMELINEEX BrtEndTimelinesEx BrtFRTEnd",
    WEBEXTENSIONS:
      "BrtFRTBegin BrtBeginWebExtensions 1*BrtWebExtension BrtEndWebExtensions BrtFRTEnd"
  }
};

export const WORKBOOK_GRAMMAR: XlsbGrammar = {
  source:
    "MS-XLSB §2.1.7.61 Workbook, https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/be874d90-2b52-4141-b7d7-d33426a615f3",
  root: "WORKBOOK",
  productions: {
    WORKBOOK:
      "BrtBeginBook [BrtFileVersion] [[BrtFileSharingIso] BrtFileSharing] [BrtWbProp] [ACABSPATH] [ACREVISIONPTR] [[BrtBookProtectionIso] BrtBookProtection] [BOOKVIEWS] BUNDLESHS [FNGROUP] [EXTERNALS] *BrtName [BrtCalcProp] [BrtOleSize] *(BrtUserBookView *FRT) [PIVOTCACHEIDS] [BrtWbFactoid] [SMARTTAGTYPES] [BrtWebOpt] *BrtFileRecover [WEBPUBITEMS] [CRERRS] FRTWORKBOOK BrtEndBook",
    ACABSPATH: "BrtACBegin BrtAbsPath15 BrtACEnd",
    ACREVISIONPTR: "BrtACBegin BrtRevisionPtr BrtACEnd",
    BOOKVIEWS: "BrtBeginBookViews 1*([ACUID] BrtBookView *FRT) BrtEndBookViews",
    BUNDLESHS: "BrtBeginBundleShs 1*BrtBundleSh BrtEndBundleShs",
    FNGROUP: "BrtBeginFnGroup *BrtFnGroup BrtEndFnGroup",
    EXTERNALS: "BrtBeginExternals *SUP BrtExternSheet BrtEndExternals",
    SUP: "BrtSupSelf / SUPSAME / SUPADDIN / BrtSupBookSrc",
    SUPSAME: "BrtSupSame *BrtPlaceholderName",
    SUPADDIN: "BrtSupAddin 1*BrtPlaceholderName",
    PIVOTCACHEIDS: "BrtBeginPivotCacheIDs 1*PIVOTCACHEID BrtEndPivotCacheIDs",
    PIVOTCACHEID: "BrtBeginPivotCacheID BrtEndPivotCacheID",
    SMARTTAGTYPES: "BrtBeginSmartTagTypes 1*BrtSmartTagType BrtEndSmartTagTypes",
    CRERRS: "BrtBeginCRErrs *BrtCrashRecErr BrtEndCRErrs",
    FRTWORKBOOK:
      "[NAMEEXT] [SLICERCACHESPIVOTCACHEIDS] [SLICERCACHEIDS] [WORKBOOKPR14] [DECOUPLEDPIVOTCACHEIDS] [PIVOTTABLEREFS] [TIMELINECACHEPIVOTCACHEIDS] [TIMELINECACHEIDS] [TABLESLICERCACHEIDS] [WORKBOOKPR15] [DATAMODEL] [CALCFEATURES] [EXTERNALLINKSPR] [EXTERNALCODESERVICE] *FRT",
    NAMEEXT: "BrtFRTBegin 1*BrtNameExt BrtFRTEnd",
    SLICERCACHESPIVOTCACHEID: "BrtBeginSlicerCachesPivotCacheID BrtEndSlicerCachesPivotCacheID",
    SLICERCACHEID: "BrtBeginSlicerCacheID BrtEndSlicerCacheID",
    TIMELINECACHEPIVOTCACHEID: "BrtTimelineCachePivotCacheID",
    TIMELINECACHEID: "BrtBeginTimelineCacheID BrtEndTimelineCacheID",
    SLICERCACHESPIVOTCACHEIDS:
      "BrtFRTBegin BrtBeginSlicerCachesPivotCacheIDs 1*SLICERCACHESPIVOTCACHEID BrtEndSlicerCachesPivotCacheIDs BrtFRTEnd",
    SLICERCACHEIDS:
      "BrtFRTBegin BrtBeginSlicerCacheIDs 1*SLICERCACHEID BrtEndSlicerCacheIDs BrtFRTEnd",
    WORKBOOKPR14: "BrtFRTBegin BrtWbProp14 BrtFRTEnd",
    DECOUPLEDPIVOTCACHEIDS:
      "BrtFRTBegin BrtBeginDecoupledPivotCacheIDs 1*BrtDecoupledPivotCacheID BrtEndDecoupledPivotCacheIDs BrtFRTEnd",
    PIVOTTABLEREFS:
      "BrtFRTBegin BrtBeginPivotTableRefs 1*BrtPivotTableRef BrtEndPivotTableRefs BrtFRTEnd",
    TIMELINECACHEPIVOTCACHEIDS:
      "BrtFRTBegin BrtBeginTimelineCachePivotCacheIDs 1*TIMELINECACHEPIVOTCACHEID BrtEndTimelineCachePivotCacheIDs BrtFRTEnd",
    TIMELINECACHEIDS:
      "BrtFRTBegin BrtBeginTimelineCacheIDs 1*TIMELINECACHEID BrtEndTimelineCacheIDs BrtFRTEnd",
    TABLESLICERCACHEIDS: "BrtFRTBegin BrtTableSlicerCacheIDs 1*TABLESLICERCACHEID BrtFRTEnd",
    TABLESLICERCACHEID: "BrtTableSlicerCacheID",
    WORKBOOKPR15: "BrtFRTBegin BrtWorkBookPr15 BrtFRTEnd",
    DATAMODEL: "BrtBeginDataModel MODELTABLES MODELRELATIONSHIPS FRTDATAMODEL BrtEndDataModel",
    MODELTABLES: "BrtBeginModelTables 1*BrtModelTable BrtEndModelTables",
    MODELRELATIONSHIPS:
      "brtBeginModelRelationships 1*brtModelRelationship brtEndModelRelationships",
    FRTDATAMODEL: "[MODELTIMEGROUPINGS] *FRT",
    MODELTIMEGROUPINGS:
      "BrtFRTBegin brtBeginModelTimeGroupings 1*MODELTIMEGROUPING brtEndModelTimeGroupings BrtFRTEnd",
    MODELTIMEGROUPING:
      "brtBeginModelTimeGrouping 1*brtModelTimeGroupingCalcCol brtEndModelTimeGrouping",
    CALCFEATURES: "BrtFRTBegin BrtBeginCalcFeatures 1*BrtCalcFeature BrtEndCalcFeatures BrtFRTEnd",
    EXTERNALLINKSPR: "BrtFRTBegin BrtExternalLinksPr BrtFRTEnd",
    EXTERNALCODESERVICE: "BrtFRTBegin BrtExternalCodeService BrtFRTEnd"
  }
};

export const STYLESHEET_GRAMMAR: XlsbGrammar = {
  source:
    "MS-XLSB §2.1.7.50 Styles, https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/b0da927e-fefc-45b0-bdb9-38883737957f",
  root: "STYLESHEET",
  productions: {
    STYLESHEET:
      "BrtBeginStyleSheet [FMTS] [FONTS] [FILLS] [BORDERS] CELLSTYLEXFS CELLXFS STYLES DXFS TABLESTYLES [COLORPALETTE] FRTSTYLESHEET BrtEndStyleSheet",
    FMTS: "BrtBeginFmts 1*(2ACFMT / BrtFmt) BrtEndFmts",
    ACFMT: "BrtACBegin BrtFmt BrtACEnd",
    ACFONTS: "BrtACBegin BrtKnownFonts BrtACEnd",
    XFCOMPLEMENTID: "BrtFRTBegin BrtXfComplementId BrtFRTEnd",
    DXF: "((BrtACBegin BrtDXF BrtACEnd BrtACBegin FRTDXF BrtACEnd) / (FRTDXF))",
    FRTDXF: "BrtDXF *FRT",
    TABLESTYLE: "[ACUID] BrtBeginTableStyle *28BrtTableStyleElement BrtEndTableStyle",
    INDEXEDCOLORS: "BrtBeginIndexedColors 64BrtIndexedColor BrtEndIndexedColors",
    MRUCOLORS: "BrtBeginMRUColors 1*10BrtMRUColor BrtEndMRUColors",
    SLICERSTYLES: "BrtBeginSlicerStyles *2147483647SLICERSTYLE BrtEndSlicerStyles",
    SLICERSTYLE:
      "BrtBeginSlicerStyle BrtBeginSlicerStyleElements *8BrtSlicerStyleElement BrtEndSlicerStyleElements BrtEndSlicerStyle",
    TIMELINESTYLES: "BrtBeginTimelineStyles *2147483647TIMELINESTYLE BrtEndTimelineStyles",
    TIMELINESTYLE:
      "BrtBeginTimelineStyle BrtBeginTimelineStyleElements *7BrtTimelineStyleElement BrtEndTimelineStyleElements BrtEndTimelineStyle",
    FONTS: "BrtBeginFonts 1*65491BrtFont [ACFONTS] BrtEndFonts",
    FILLS: "BrtBeginFills 1*65431BrtFill BrtEndFills",
    BORDERS: "BrtBeginBorders 1*65430BrtBorder BrtEndBorders",
    CELLSTYLEXFS: "BrtBeginCellStyleXFs 1*65430(BrtXF [XFCOMPLEMENTID] *FRT) BrtEndCellStyleXFs",
    CELLXFS: "BrtBeginCellXFs 1*65430(BrtXF [XFCOMPLEMENTID] *FRT) BrtEndCellXFs",
    STYLES: "BrtBeginStyles 1*65430([ACUID] BrtStyle *FRT) BrtEndStyles",
    DXFS: "BrtBeginDXFs *2147483647DXF BrtEndDXFs",
    TABLESTYLES: "BrtBeginTableStyles *TABLESTYLE BrtEndTableStyles",
    COLORPALETTE: "BrtBeginColorPalette [INDEXEDCOLORS] [MRUCOLORS] BrtEndColorPalette",
    FRTSTYLESHEET: "[STYLESHEET14] [DXF14S] [STYLESHEET15] [DXFS15] *FRT",
    STYLESHEET14:
      "BrtFRTBegin BrtBeginStyleSheetExt14 [SLICERSTYLES] BrtEndStyleSheetExt14 BrtFRTEnd",
    DXF14S: "BrtFRTBegin BrtBeginDXF14s 1*2147483647BrtDXF14 BrtEndDXF14s BrtFRTEnd",
    STYLESHEET15:
      "BrtFRTBegin BrtBeginTimelineStyleSheetExt15 TIMELINESTYLES BrtEndTimelineStyleSheetExt15 BrtFRTEnd",
    DXFS15: "BrtFRTBegin BrtBeginDXFs15 1*2147483647BrtDXF15 BrtEndDXFs15 BrtFRTEnd"
  }
};

export const TABLE_GRAMMAR: XlsbGrammar = {
  source:
    "MS-XLSB §2.1.7.51 Table, https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/6e4302ea-b82b-4581-9765-519ac3475299",
  root: "TABLE",
  productions: {
    TABLE:
      "[ACUID] BrtBeginList [AUTOFILTER] [SORTSTATE] LISTCOLS BrtTableStyleClient FRTTABLE BrtEndList",
    LISTCOLS: "BrtBeginListCols 1*LISTCOL BrtEndListCols",
    LISTCOL:
      "[ACUID] BrtBeginListCol [BrtListCCFmla] [BrtListTrFmla] [LISTXMLCPR] *FRT BrtEndListCol",
    LISTXMLCPR: "BrtBeginListXmlCPr *FRT BrtEndListXmlCPr",
    FRTTABLE: "[BrtFRTBegin BrtList14 BrtFRTEnd] *FRT"
  }
};

export const COMMENTS_GRAMMAR: XlsbGrammar = {
  source:
    "MS-XLSB §2.1.7.8 Comments, https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/f60856f3-f77b-47d1-8389-b50c2c864d63",
  root: "COMMENTS",
  productions: {
    COMMENTS: "BrtBeginComments COMMENTAUTHORS COMMENTLIST *FRT BrtEndComments",
    COMMENTAUTHORS: "BrtBeginCommentAuthors *BrtCommentAuthor BrtEndCommentAuthors",
    COMMENTLIST: "BrtBeginCommentList *COMMENT BrtEndCommentList",
    COMMENT: "[ACUID] BrtBeginComment [BrtCommentText] BrtEndComment"
  }
};

/**
 * The §2.1.8 common productions the grammars above use at their top level, by the record that opens each. They
 * are not published as text, so this is the one hand-written part; every name is a scope's `Begin` record (or,
 * for `WSFMTINFO`, the single record it is), and a test holds each to the record dictionary.
 */
export const COMMON_OPENERS: Readonly<Record<string, string>> = {
  WSFMTINFO: "BrtWsFmtInfo",
  COLINFOS: "BrtBeginColInfos",
  CELLTABLE: "BrtBeginSheetData",
  AUTOFILTER: "BrtBeginAFilter",
  SORTSTATE: "BrtBeginSortState",
  DCON: "BrtBeginDCon",
  USERSHVIEWS: "BrtBeginUserShViews",
  CONDITIONALFORMATTING: "BrtBeginConditionalFormatting",
  HEADERFOOTER: "BrtBeginHeaderFooter",
  RWBRK: "BrtBeginRwBrk",
  COLBRK: "BrtBeginColBrk",
  OLEOBJECTS: "BrtBeginOleObjects",
  WEBPUBITEMS: "BrtBeginWebPubItems"
};

/**
 * The productions these grammars name that MS-XLSB does not publish: §2.1.8 lists them, but its text is a reference to
 * an `.abnf` file rather than the productions, and neither has a record of its own to stand for it as the openers in
 * {@link COMMON_OPENERS} do. A scope that needs one is not ranked. Listed by name so that any other name that does
 * not resolve — a typo in a copied production — is an error rather than a scope quietly left out.
 */
const UNPUBLISHED: ReadonlySet<string> = new Set(["CONDITIONALFORMATTING14", "PIVOTRULE"]);

/** Productions that may appear between almost any two records, so they take no position. */
const POSITIONLESS = new Set(["FRT", "ACUID"]);

/** Records that only delimit: an `End` closes what its `Begin` opened, and the two wrappers are transparent. */
export function isDelimiterOnly(name: string): boolean {
  return /^[Bb]rtEnd/.test(name) || isWrapperEnd(name);
}

/** The record that closes a `BrtFRTBegin` or `BrtACBegin` wrapper. */
export function isWrapperEnd(name: string): boolean {
  return name === "BrtFRTEnd" || name === "BrtACEnd";
}

export function isWrapper(name: string): boolean {
  return name === "BrtFRTBegin" || name === "BrtACBegin";
}

// -----------------------------------------------------------------------------
// ABNF → positions
// -----------------------------------------------------------------------------

type Node =
  | { kind: "name"; name: string; repeats: boolean }
  | { kind: "seq"; items: Node[]; repeats: boolean }
  | { kind: "alt"; branches: Node[]; repeats: boolean };

/** Tokens: brackets, `/`, and an element with its repetition prefix (`*`, `1*`, `*4`, `1*65534`, `2`). */
function tokenize(text: string): string[] {
  return text
    .match(/\[|\]|\(|\)|\/|(?:\d*\*\d*|\d+)?(?:[A-Za-z][A-Za-z0-9]*|(?=\())/g)!
    .filter(t => t !== "");
}

function parse(text: string): Node {
  const tokens = tokenize(text);
  let pos = 0;

  function repeatsOf(prefix: string): boolean {
    if (prefix === "") {
      return false;
    }
    if (prefix.includes("*")) {
      const [, max] = prefix.split("*");
      return max === "" || Number(max) > 1;
    }
    return Number(prefix) > 1;
  }

  function parseAlt(close?: string): Node {
    const branches: Node[] = [parseSeq(close)];
    while (tokens[pos] === "/") {
      pos++;
      branches.push(parseSeq(close));
    }
    return branches.length === 1 ? branches[0]! : { kind: "alt", branches, repeats: false };
  }

  function parseSeq(close?: string): Node {
    const items: Node[] = [];
    while (pos < tokens.length && tokens[pos] !== close && tokens[pos] !== "/") {
      const token = tokens[pos++]!;
      if (token === "[") {
        items.push(parseAlt("]"));
        pos++;
        continue;
      }
      const prefix = /^(\d*\*\d*|\d+)?/.exec(token)![0];
      const rest = token.slice(prefix.length);
      if (rest === "") {
        // A repetition prefix in front of a group: the group follows.
        if (tokens[pos++] !== "(") {
          throw new ExcelError(`expected a group after ${prefix} in: ${text}`);
        }
        const group = parseAlt(")");
        pos++;
        items.push({ ...group, repeats: repeatsOf(prefix) });
        continue;
      }
      if (token === "(") {
        items.push(parseAlt(")"));
        pos++;
        continue;
      }
      items.push({ kind: "name", name: rest, repeats: repeatsOf(prefix) });
    }
    return { kind: "seq", items, repeats: false };
  }

  const tree = parseAlt();
  if (pos !== tokens.length) {
    throw new ExcelError(`unparsed input at "${tokens.slice(pos).join(" ")}" in: ${text}`);
  }
  return tree;
}

function collapse(slots: string[][], repeats: boolean): string[][] {
  return repeats && slots.length > 1 ? [[...new Set(slots.flat())]] : slots;
}

function mergeBranches(branches: string[][][]): string[][] {
  const merged: string[][] = [];
  for (const branch of branches) {
    branch.forEach((slot, k) => {
      merged[k] = [...new Set([...(merged[k] ?? []), ...slot])];
    });
  }
  return merged;
}

/** The first record a production's definition names, past a wrapper. */
function openerOf(definition: string): string | undefined {
  return tokenize(definition)
    .map(t => t.replace(/^(\d*\*\d*|\d+)/, ""))
    .find(t => /^[Bb]rt/.test(t) && !isWrapper(t));
}

/**
 * A scope written inline — `BrtBeginSparklines 1*BrtSparkline BrtEndSparklines` inside `SPARKLINEGROUP` — is one
 * position, named by its `Begin`: what it holds is a level further down, so it is dropped here.
 */
function foldInlineScopes(items: readonly Node[]): Node[] {
  const folded: Node[] = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i]!;
    folded.push(item);
    if (item.kind !== "name" || !/^[Bb]rtBegin/.test(item.name)) {
      continue;
    }
    const end = item.name.replace(/^([Bb]rt)Begin/, "$1End");
    const close = items.findIndex(
      (candidate, j) => j > i && candidate.kind === "name" && candidate.name === end
    );
    if (close > i) {
      i = close;
    }
  }
  return folded;
}

function flattenNode(grammar: XlsbGrammar, node: Node, stack: readonly string[]): string[][] {
  if (node.kind === "seq") {
    return collapse(
      foldInlineScopes(node.items).flatMap(item => flattenNode(grammar, item, stack)),
      node.repeats
    );
  }
  if (node.kind === "alt") {
    return collapse(
      mergeBranches(node.branches.map(b => flattenNode(grammar, b, stack))),
      node.repeats
    );
  }
  const { name } = node;
  if (/^[Bb]rt/.test(name)) {
    return isWrapper(name) || isDelimiterOnly(name) ? [] : [[name]];
  }
  if (POSITIONLESS.has(name)) {
    return [];
  }
  const definition = grammar.productions[name];
  if (definition === undefined) {
    const opener = COMMON_OPENERS[name];
    if (opener === undefined) {
      throw new ExcelError(
        `${grammar.root}: production ${name} is neither defined here nor in COMMON_OPENERS`
      );
    }
    return [[opener]];
  }
  const opener = openerOf(definition);
  // A production that opens a scope is one position; the records inside it are a level down.
  if (opener !== undefined && /^[Bb]rtBegin/.test(opener)) {
    return [[opener]];
  }
  if (stack.includes(name)) {
    throw new ExcelError(`${grammar.root}: production ${name} is recursive`);
  }
  return collapse(flattenNode(grammar, parse(definition), [...stack, name]), node.repeats);
}

/**
 * The specification's grammar and its record table do not always agree on case — `BrtBeginDXF14s` in the one,
 * `BrtBeginDxF14s` in the other — and the dictionary is what a record's id resolves to.
 */
function dictionaryName(name: string): string {
  if (!DICTIONARY_BY_FOLDED_NAME) {
    DICTIONARY_BY_FOLDED_NAME = new Map(
      [...MS_XLSB_RECORD_NAMES.values()].map(known => [foldRecordName(known), known])
    );
  }
  return DICTIONARY_BY_FOLDED_NAME.get(foldRecordName(name)) ?? name;
}

/** Record names are ASCII identifiers; folding their case is not a name comparison in Office's sense. */
function foldRecordName(name: string): string {
  return name.toLowerCase();
}

let DICTIONARY_BY_FOLDED_NAME: Map<string, string> | undefined;

/**
 * The positions of the records directly inside every scope the grammar defines, by the record that opens the scope —
 * the root's included: `BrtBeginWsView` ranks a pane before a selection, `BrtBeginExternals` its supporting books
 * before `BrtExternSheet`. A production wrapped whole in `BrtFRTBegin … BrtFRTEnd` is the scope inside the wrapper.
 *
 * A scope whose content uses an {@link UNPUBLISHED} production is left out: there is nothing to rank it by.
 * Anything else that does not resolve — a typo in a copied production — throws, and so does a
 * second production that opens the same scope with different content.
 */
export function grammarScopes(
  grammar: XlsbGrammar
): ReadonlyMap<string, ReadonlyMap<string, number>> {
  const scopes = new Map<string, ReadonlyMap<string, number>>();
  for (const [name, definition] of Object.entries(grammar.productions)) {
    const node = parse(definition);
    if (node.kind !== "seq") {
      continue;
    }
    const items = withoutWrapper(node.items);
    const opener = items.findIndex(item => item.kind === "name" && /^[Bb]rtBegin/.test(item.name));
    const last = items[items.length - 1];
    if (opener < 0 || last?.kind !== "name" || !/^[Bb]rtEnd/.test(last.name)) {
      continue;
    }
    const inner: Node = { ...node, items: items.slice(opener + 1, -1) };
    if (usesUnpublished(grammar, inner, [name])) {
      if (name === grammar.root) {
        throw new ExcelError(`${grammar.root}: the root production uses one that is not published`);
      }
      continue;
    }
    const slots = withoutAmbiguousNames(
      flattenNode(grammar, inner, [name]).map(slot => slot.map(dictionaryName))
    );
    if (slots.length < 2) {
      continue;
    }
    const openerName = dictionaryName((items[opener] as { name: string }).name);
    const ranks = new Map(slots.flatMap((slot, i) => slot.map(record => [record, i] as const)));
    const existing = scopes.get(openerName);
    if (
      existing &&
      (existing.size !== ranks.size ||
        [...ranks].some(([record, rank]) => existing.get(record) !== rank))
    ) {
      throw new ExcelError(
        `${grammar.root}: two productions open ${openerName} with different content`
      );
    }
    scopes.set(openerName, ranks);
  }
  return scopes;
}

/** A production's items without a wrapper around the whole of it (`BrtFRTBegin … BrtFRTEnd`). */
function withoutWrapper(items: readonly Node[]): readonly Node[] {
  const first = items[0];
  const last = items[items.length - 1];
  return first?.kind === "name" &&
    isWrapper(first.name) &&
    last?.kind === "name" &&
    isWrapperEnd(last.name)
    ? items.slice(1, -1)
    : items;
}

/**
 * Whether ranking `node` would need one of the {@link UNPUBLISHED} productions. Walks exactly what
 * {@link flattenNode} walks, branch for branch: an inline scope is folded and a defined production that opens a scope
 * is one position, so neither is looked into; a defined name is looked up before a common one. A recursive production
 * is left for {@link flattenNode} to refuse.
 */
function usesUnpublished(grammar: XlsbGrammar, node: Node, stack: readonly string[]): boolean {
  if (node.kind === "seq") {
    return foldInlineScopes(node.items).some(item => usesUnpublished(grammar, item, stack));
  }
  if (node.kind === "alt") {
    return node.branches.some(branch => usesUnpublished(grammar, branch, stack));
  }
  const { name } = node;
  if (/^[Bb]rt/.test(name) || POSITIONLESS.has(name)) {
    return false;
  }
  const definition = grammar.productions[name];
  if (definition === undefined) {
    return UNPUBLISHED.has(name);
  }
  const opener = openerOf(definition);
  if ((opener !== undefined && /^[Bb]rtBegin/.test(opener)) || stack.includes(name)) {
    return false;
  }
  return usesUnpublished(grammar, parse(definition), [...stack, name]);
}
