/**
 * Namespace prefixes, normalised on read, for the parts of an Excel package.
 *
 * Every xform matches element names with the prefix Excel writes — `c:chartSpace`, `xdr:twoCellAnchor`,
 * an unprefixed `worksheet` — but a prefix is only a local alias for a namespace URI. So the reader
 * renames each element and prefixed attribute to the canonical prefix of its namespace before any xform
 * sees it: see `@xml/namespace-prefixes`, which does that for any format given its namespaces. This
 * module is that, with Excel's.
 */
import { createPartDecoder } from "@excel/utils/part-decoder";
import {
  canonicalizeXml as canonicalizeWith,
  canonicalNamespaceOf as canonicalNamespaceWith,
  createPrefixNormalizer as createNormalizerWith,
  createPrefixVocabulary,
  needsPrefixNormalization as needsNormalizationWith,
  normalizeElementPrefixes as normalizeElementsWith
} from "@xml/namespace-prefixes";
import type { PrefixNormalizer, PrefixNormalizerOptions } from "@xml/namespace-prefixes";
import type { ParseOpenTag } from "@xml/types";

export type { PrefixNormalizer, PrefixNormalizerOptions } from "@xml/namespace-prefixes";

/** Namespace URI → the prefix the xforms match. `""` is "unprefixed". */
const EXCEL_PREFIXES: ReadonlyMap<string, string> = new Map([
  ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", ""],
  ["http://schemas.openxmlformats.org/package/2006/content-types", ""],
  ["http://schemas.openxmlformats.org/package/2006/relationships", ""],
  ["http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", ""],
  ["http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments", ""],
  ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r"],
  ["http://schemas.openxmlformats.org/drawingml/2006/main", "a"],
  ["http://schemas.openxmlformats.org/drawingml/2006/chart", "c"],
  ["http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xdr"],
  ["http://schemas.openxmlformats.org/drawingml/2006/chartDrawing", "cdr"],
  ["http://schemas.openxmlformats.org/markup-compatibility/2006", "mc"],
  ["http://schemas.microsoft.com/office/drawing/2014/chartex", "cx"],
  ["http://schemas.microsoft.com/office/drawing/2012/chartStyle", "cs"],
  ["http://schemas.microsoft.com/office/drawing/2007/8/2/chart", "c14"],
  ["http://schemas.microsoft.com/office/drawing/2014/chart", "c16"],
  ["http://schemas.microsoft.com/office/drawing/2015/06/chart", "c16r2"],
  ["http://schemas.microsoft.com/office/drawing/2010/main", "a14"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "x14"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "x14ac"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2010/11/main", "x15"],
  ["http://schemas.microsoft.com/office/excel/2006/main", "xm"],
  ["urn:schemas-microsoft-com:vml", "v"],
  ["urn:schemas-microsoft-com:office:office", "o"],
  ["urn:schemas-microsoft-com:office:excel", "x"],
  ["http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "cp"],
  ["http://purl.org/dc/elements/1.1/", "dc"],
  ["http://purl.org/dc/terms/", "dcterms"],
  ["http://purl.org/dc/dcmitype/", "dcmitype"],
  ["http://www.w3.org/2001/XMLSchema-instance", "xsi"],
  ["http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", "vt"],
  // ChartEx, one namespace per revision; `mc:Choice Requires` names them, so each must keep its prefix.
  ["http://schemas.microsoft.com/office/drawing/2015/9/8/chartex", "cx1"],
  ["http://schemas.microsoft.com/office/drawing/2015/10/21/chartex", "cx2"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/9/chartex", "cx3"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/10/chartex", "cx4"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/11/chartex", "cx5"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/12/chartex", "cx6"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/13/chartex", "cx7"],
  ["http://schemas.microsoft.com/office/drawing/2016/5/14/chartex", "cx8"],
  // Office extensions the reader matches by prefix, or keeps and writes back by prefix.
  ["http://schemas.microsoft.com/office/drawing/2012/chart", "c15"],
  ["http://schemas.microsoft.com/office/drawing/2017/03/chart", "c16r3"],
  ["http://schemas.microsoft.com/office/drawing/2012/main", "a15"],
  ["http://schemas.microsoft.com/office/drawing/2014/main", "a16"],
  ["http://schemas.microsoft.com/office/drawing/2016/SVG/main", "asvg"],
  ["http://schemas.microsoft.com/office/drawing/2016/ink", "aink"],
  ["http://schemas.microsoft.com/office/drawing/2017/model3d", "am3d"],
  ["http://schemas.microsoft.com/office/thememl/2012/main", "thm15"],
  ["http://schemas.openxmlformats.org/drawingml/2006/picture", "pic"],
  ["http://schemas.openxmlformats.org/drawingml/2006/diagram", "dgm"],
  ["http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas", "lc"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac", "x15ac"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2011/1/ac", "x12ac"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2015/02/main", "x16r2"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "xr"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", "xr2"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision3", "xr3"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision6", "xr6"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2016/revision10", "xr10"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures", "xcalcf"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray", "xda"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2017/richdata", "xlrd"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2", "xlrd2"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout", "xpdl"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag", "xfpb"],
  ["http://schemas.microsoft.com/office/spreadsheetml/2019/extlinksprops", "xxlnp"],
  ["http://schemas.microsoft.com/office/mac/excel/2008/main", "mx"],
  ["urn:schemas-microsoft-com:mac:vml", "mv"],
  ["http://schemas.openxmlformats.org/officeDocument/2006/customXml", "ds"]
]);

/** Excel's namespaces. An unknown namespace keeps its prefix unless a known one in scope takes it. */
const EXCEL = /* @__PURE__ */ createPrefixVocabulary(EXCEL_PREFIXES);

export function createPrefixNormalizer(options?: PrefixNormalizerOptions): PrefixNormalizer {
  return createNormalizerWith(EXCEL, options);
}

export function needsPrefixNormalization(xml: string): boolean {
  return needsNormalizationWith(EXCEL, xml);
}

/**
 * The namespace `prefix` stands for in every part this library reads and writes — the reader renames a
 * known namespace to it — or `undefined` for a prefix that is no namespace's canonical one.
 */
export function canonicalNamespaceOf(prefix: string): string | undefined {
  return canonicalNamespaceWith(EXCEL, prefix);
}

export function normalizeElementPrefixes(
  root: Parameters<typeof normalizeElementsWith>[1]
): ReadonlyMap<Parameters<typeof normalizeElementsWith>[1], ParseOpenTag> {
  return normalizeElementsWith(EXCEL, root);
}

/**
 * `xml` rewritten with Excel's canonical prefixes: see `canonicalizeXml` in `@xml/namespace-prefixes`.
 * Parsed under the policy the xforms read a part with (`parseXformStream`): a character XML forbids is
 * skipped, and a part that does not parse throws.
 */
export function canonicalizeXml(xml: string): string {
  return canonicalizeWith(EXCEL, xml, XFORM_PARSE_POLICY);
}

const XFORM_PARSE_POLICY = { invalidCharHandling: "skip" } as const;

/**
 * `chunks` decoded to text and shown to `normalizer` before they are yielded — for a lazy normaliser
 * whose parser is fed by a helper that does not expose the text.
 */
export async function* scanningChunks(
  chunks: AsyncIterable<unknown> | Iterable<unknown>,
  normalizer: PrefixNormalizer
): AsyncGenerator<string> {
  const decoder = createPartDecoder();
  for await (const chunk of chunks) {
    const text =
      typeof chunk === "string" ? chunk : decoder.decode(chunk as Uint8Array, { stream: true });
    normalizer.scan(text);
    yield text;
  }
  const trailing = decoder.decode();
  if (trailing) {
    normalizer.scan(trailing);
    yield trailing;
  }
}
