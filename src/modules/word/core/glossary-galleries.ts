/**
 * Building-block galleries: the model's {@link BuildingBlockGallery} names and
 * the `ST_DocPartGallery` values (ECMA-376 §17.18.23) they are written as and
 * read from. Shared by the glossary writer, the reader and `Glossary.*`.
 */

import type { BuildingBlockGallery } from "@word/types";

/**
 * The `ST_DocPartGallery` value each gallery is written as. Values MUST be
 * exact enum members — Word rejects (and silently discards) the entire
 * glossary if it sees an out-of-enum gallery value. There is no plain
 * "quickParts" value; Quick Parts map to "custQuickParts".
 */
export const GALLERY_TO_OOXML: Readonly<Record<BuildingBlockGallery, string>> = {
  autoText: "autoTxt",
  quickParts: "custQuickParts",
  coverPages: "coverPg",
  tableOfContents: "tblOfContents",
  headers: "hdrs",
  footers: "ftrs",
  pageNumbers: "pgNum",
  tables: "tbls",
  textBoxes: "txtBox",
  watermarks: "watermarks",
  equations: "eq",
  bibliographies: "bib",
  custom1: "custom1",
  custom2: "custom2",
  custom3: "custom3",
  custom4: "custom4",
  custom5: "custom5"
};

/**
 * The gallery a `ST_DocPartGallery` value is read as. Word keeps built-in and
 * custom entries of one kind in two values (`autoTxt` and `custAutoTxt`), and
 * Quick Parts under `docParts` as well as `custQuickParts`; each pair is one
 * gallery here. `placeholder`, `any` and `default` — content-control
 * placeholder text and Word's internal entries — are no gallery a user
 * browses and have none.
 */
export const OOXML_TO_GALLERY: ReadonlyMap<string, BuildingBlockGallery> = new Map<
  string,
  BuildingBlockGallery
>([
  ...Object.entries(GALLERY_TO_OOXML).map(
    ([gallery, value]) => [value, gallery as BuildingBlockGallery] as const
  ),
  ["docParts", "quickParts"],
  ["custAutoTxt", "autoText"],
  ["custCoverPg", "coverPages"],
  ["custTblOfContents", "tableOfContents"],
  ["custHdrs", "headers"],
  ["custFtrs", "footers"],
  ["custPgNum", "pageNumbers"],
  ["pgNumT", "pageNumbers"],
  ["pgNumB", "pageNumbers"],
  ["pgNumMargins", "pageNumbers"],
  ["custPgNumT", "pageNumbers"],
  ["custPgNumB", "pageNumbers"],
  ["custPgNumMargins", "pageNumbers"],
  ["custTbls", "tables"],
  ["custTxtBox", "textBoxes"],
  ["custWatermarks", "watermarks"],
  ["custEq", "equations"],
  ["custBib", "bibliographies"]
]);
