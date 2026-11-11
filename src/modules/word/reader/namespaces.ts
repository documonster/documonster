/**
 * The namespaces of a Word package, by the prefixes the reader matches.
 *
 * The reader matches element and attribute names with the prefixes Word writes — `w:p`, `wp:inline`,
 * `a:blip` — and a prefix is only a local alias for a namespace. Each part is read with its names
 * rewritten to these prefixes by the namespace they are in (`@xml/namespace-prefixes`): a part that binds
 * the main namespace to `z`, or makes it the default one, reads as Word wrote it; and a part that binds
 * `w` to some other namespace has its elements renamed out of the way, so they are not read as Word's.
 * A Strict document's namespaces take the prefixes of their Transitional counterparts, as Word reads
 * them.
 */
import {
  NS_A,
  NS_ASVG,
  NS_C_CHART,
  NS_CONTENT_TYPES,
  NS_CP,
  NS_CUSTOM,
  NS_CX_CHART,
  NS_DC,
  NS_DCMITYPE,
  NS_DCTERMS,
  NS_EP,
  NS_M,
  NS_MC,
  NS_O,
  NS_PIC,
  NS_PKG_RELS,
  NS_R,
  NS_V,
  NS_VT,
  NS_W,
  NS_W10,
  NS_W14,
  NS_W15,
  NS_WP,
  NS_WP14,
  NS_WPC,
  NS_WPG,
  NS_WPI,
  NS_WPS,
  NS_XSI,
  STRICT_TO_TRANSITIONAL_NS
} from "@word/constants";
import { createPrefixVocabulary } from "@xml/namespace-prefixes";
import type { PrefixVocabulary } from "@xml/namespace-prefixes";

/** Namespace URI → the prefix the reader matches; `""` is "unprefixed". */
const TRANSITIONAL: ReadonlyArray<readonly [string, string]> = [
  [NS_CONTENT_TYPES, ""],
  [NS_PKG_RELS, ""],
  [NS_EP, ""],
  [NS_CUSTOM, ""],
  [NS_W, "w"],
  [NS_R, "r"],
  [NS_WP, "wp"],
  [NS_A, "a"],
  [NS_PIC, "pic"],
  [NS_MC, "mc"],
  [NS_V, "v"],
  [NS_O, "o"],
  [NS_M, "m"],
  [NS_W10, "w10"],
  [NS_W14, "w14"],
  [NS_W15, "w15"],
  ["http://schemas.microsoft.com/office/word/2015/wordml/symex", "w16se"],
  ["http://schemas.microsoft.com/office/word/2016/wordml/cid", "w16cid"],
  ["http://schemas.microsoft.com/office/word/2018/wordml", "w16"],
  ["http://schemas.microsoft.com/office/word/2018/wordml/cex", "w16cex"],
  ["http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash", "w16sdtdh"],
  ["http://schemas.microsoft.com/office/word/2006/wordml", "wne"],
  [NS_WPG, "wpg"],
  [NS_WPI, "wpi"],
  [NS_WPS, "wps"],
  [NS_WPC, "wpc"],
  [NS_WP14, "wp14"],
  ["http://schemas.microsoft.com/office/drawing/2010/main", "a14"],
  [NS_ASVG, "asvg"],
  [NS_C_CHART, "c"],
  [NS_CX_CHART, "cx"],
  ["http://schemas.microsoft.com/office/drawing/2007/8/2/chart", "c14"],
  ["http://schemas.microsoft.com/office/drawing/2014/chart", "c16"],
  ["http://schemas.microsoft.com/office/drawing/2015/06/chart", "c16r2"],
  ["http://schemas.microsoft.com/office/drawing/2017/03/chart", "c16r3"],
  ["http://schemas.microsoft.com/office/drawing/2012/chartStyle", "cs"],
  ["http://schemas.openxmlformats.org/officeDocument/2006/customXml", "ds"],
  [NS_CP, "cp"],
  [NS_DC, "dc"],
  [NS_DCTERMS, "dcterms"],
  [NS_DCMITYPE, "dcmitype"],
  [NS_VT, "vt"],
  [NS_XSI, "xsi"]
];

function wordPrefixes(): ReadonlyMap<string, string> {
  const prefixes = new Map(TRANSITIONAL);
  for (const [strict, transitional] of STRICT_TO_TRANSITIONAL_NS) {
    const prefix = prefixes.get(transitional);
    if (prefix !== undefined) {
      prefixes.set(strict, prefix);
    }
  }
  return prefixes;
}

/** Word's namespaces, its prefixes reserved: see the module comment. */
export const WORD_NAMESPACES: PrefixVocabulary = /* @__PURE__ */ createPrefixVocabulary(
  wordPrefixes(),
  {
    reservePrefixes: true
  }
);
