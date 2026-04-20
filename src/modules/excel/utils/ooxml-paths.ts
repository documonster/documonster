export const OOXML_PATHS = {
  contentTypes: "[Content_Types].xml",
  rootRels: "_rels/.rels",

  docPropsApp: "docProps/app.xml",
  docPropsCore: "docProps/core.xml",

  xlWorkbook: "xl/workbook.xml",
  xlWorkbookRels: "xl/_rels/workbook.xml.rels",
  xlSharedStrings: "xl/sharedStrings.xml",
  xlStyles: "xl/styles.xml",
  xlTheme1: "xl/theme/theme1.xml",
  xlFeaturePropertyBag: "xl/featurePropertyBag/featurePropertyBag.xml",
  xlMetadata: "xl/metadata.xml"
} as const;

const worksheetXmlRegex = /^xl\/worksheets\/sheet(\d+)[.]xml$/;
const worksheetRelsXmlRegex = /^xl\/worksheets\/_rels\/sheet(\d+)[.]xml[.]rels$/;
const themeXmlRegex = /^xl\/theme\/[a-zA-Z0-9]+[.]xml$/;

const mediaFilenameRegex = /^xl\/media\/([a-zA-Z0-9]+[.][a-zA-Z0-9]{3,4})$/;
const drawingXmlRegex = /^xl\/drawings\/(drawing\d+)[.]xml$/;
const drawingRelsXmlRegex = /^xl\/drawings\/_rels\/(drawing\d+)[.]xml[.]rels$/;
const vmlDrawingRegex = /^xl\/drawings\/(vmlDrawing\d+)[.]vml$/;
const vmlDrawingHFRegex = /^xl\/drawings\/(vmlDrawingHF\d+)[.]vml$/;
// Matches both flat layout (xl/comments1.xml) and subdirectory layout (xl/comments/comment1.xml).
// Both are valid OOXML — the actual path is determined by .rels, not by convention.
const commentsXmlRegex = /^xl\/(?:comments(\d+)|comments\/comment(\d+))[.]xml$/;
const tableXmlRegex = /^xl\/tables\/(table\d+)[.]xml$/;

const pivotTableXmlRegex = /^xl\/pivotTables\/(pivotTable\d+)[.]xml$/;
const pivotTableRelsXmlRegex = /^xl\/pivotTables\/_rels\/(pivotTable\d+)[.]xml[.]rels$/;
const pivotCacheDefinitionXmlRegex = /^xl\/pivotCache\/(pivotCacheDefinition\d+)[.]xml$/;
const pivotCacheDefinitionRelsXmlRegex =
  /^xl\/pivotCache\/_rels\/(pivotCacheDefinition\d+)[.]xml[.]rels$/;
const pivotCacheRecordsXmlRegex = /^xl\/pivotCache\/(pivotCacheRecords\d+)[.]xml$/;

// External workbook links (xl/externalLinks/externalLink{n}.xml and its rels)
const externalLinkXmlRegex = /^xl\/externalLinks\/externalLink(\d+)[.]xml$/;
const externalLinkRelsXmlRegex = /^xl\/externalLinks\/_rels\/externalLink(\d+)[.]xml[.]rels$/;

export function normalizeZipPath(path: string): string {
  return path.startsWith("/") ? path.slice(1) : path;
}

export function getWorksheetNoFromWorksheetPath(path: string): number | undefined {
  const match = worksheetXmlRegex.exec(path);
  if (!match) {
    return undefined;
  }
  return parseInt(match[1], 10);
}

export function getWorksheetNoFromWorksheetRelsPath(path: string): number | undefined {
  const match = worksheetRelsXmlRegex.exec(path);
  if (!match) {
    return undefined;
  }
  return parseInt(match[1], 10);
}

export function isMediaPath(path: string): boolean {
  return path.startsWith("xl/media/");
}

export function getMediaFilenameFromPath(path: string): string | undefined {
  const match = mediaFilenameRegex.exec(path);
  return match ? match[1] : undefined;
}

export function isThemePath(path: string): boolean {
  return themeXmlRegex.test(path);
}

export function getThemeNameFromPath(path: string): string | undefined {
  const match = /^xl\/theme\/([a-zA-Z0-9]+)[.]xml$/.exec(path);
  return match ? match[1] : undefined;
}

export function isBinaryEntryPath(path: string): boolean {
  // Media and themes should be treated as binary in the legacy buffer-based reader.
  // Everything else is parsed as text XML.
  return isMediaPath(path) || isThemePath(path);
}

export function getDrawingNameFromPath(path: string): string | undefined {
  const match = drawingXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getDrawingNameFromRelsPath(path: string): string | undefined {
  const match = drawingRelsXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getVmlDrawingNameFromPath(path: string): string | undefined {
  const match = vmlDrawingRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getVmlDrawingHFNameFromPath(path: string): string | undefined {
  const match = vmlDrawingHFRegex.exec(path);
  return match ? match[1] : undefined;
}

/**
 * Check if a zip entry path is a comments XML file.
 * Works for both `xl/comments1.xml` and `xl/comments/comment1.xml`.
 */
export function isCommentsPath(path: string): boolean {
  return commentsXmlRegex.test(path);
}

export function getTableNameFromPath(path: string): string | undefined {
  const match = tableXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getPivotTableNameFromPath(path: string): string | undefined {
  const match = pivotTableXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getPivotTableNameFromRelsPath(path: string): string | undefined {
  const match = pivotTableRelsXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getPivotCacheDefinitionNameFromPath(path: string): string | undefined {
  const match = pivotCacheDefinitionXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getPivotCacheDefinitionNameFromRelsPath(path: string): string | undefined {
  const match = pivotCacheDefinitionRelsXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

export function getPivotCacheRecordsNameFromPath(path: string): string | undefined {
  const match = pivotCacheRecordsXmlRegex.exec(path);
  return match ? match[1] : undefined;
}

/**
 * Extract the 1-based index `N` from `xl/externalLinks/externalLink{N}.xml`.
 * Returns the raw integer (e.g. `1` for externalLink1.xml) or undefined.
 */
export function getExternalLinkIndexFromPath(path: string): number | undefined {
  const match = externalLinkXmlRegex.exec(path);
  return match ? parseInt(match[1], 10) : undefined;
}

/**
 * Extract the 1-based index `N` from
 * `xl/externalLinks/_rels/externalLink{N}.xml.rels`.
 */
export function getExternalLinkIndexFromRelsPath(path: string): number | undefined {
  const match = externalLinkRelsXmlRegex.exec(path);
  return match ? parseInt(match[1], 10) : undefined;
}

export function toContentTypesPartName(zipPath: string): string {
  // ContentTypes uses leading slash PartName (e.g. "/xl/workbook.xml").
  return zipPath.startsWith("/") ? zipPath : `/${zipPath}`;
}

export function themePath(themeName: string): string {
  return `xl/theme/${themeName}.xml`;
}

export function mediaPath(filename: string): string {
  return `xl/media/${filename}`;
}

export function worksheetPath(sheetId: number | string): string {
  return `xl/worksheets/sheet${sheetId}.xml`;
}

export function worksheetRelsPath(sheetId: number | string): string {
  return `xl/worksheets/_rels/sheet${sheetId}.xml.rels`;
}

export function worksheetRelTarget(sheetId: number | string): string {
  return `worksheets/sheet${sheetId}.xml`;
}

export function commentsPath(sheetId: number | string): string {
  return `xl/comments${sheetId}.xml`;
}

export function commentsPathFromName(commentName: string): string {
  // For comments when the caller already has the logical name (e.g. "comments1").
  return `xl/${commentName}.xml`;
}

export function vmlDrawingPath(sheetId: number | string): string {
  return `xl/drawings/vmlDrawing${sheetId}.vml`;
}

export function vmlDrawingHFPath(sheetId: number | string): string {
  return `xl/drawings/vmlDrawingHF${sheetId}.vml`;
}

export function vmlDrawingHFRelsPath(sheetId: number | string): string {
  return `xl/drawings/_rels/vmlDrawingHF${sheetId}.vml.rels`;
}

export function tablePath(target: string): string {
  return `xl/tables/${target}`;
}

export function drawingPath(drawingName: string): string {
  return `xl/drawings/${drawingName}.xml`;
}

export function drawingRelsPath(drawingName: string): string {
  return `xl/drawings/_rels/${drawingName}.xml.rels`;
}

export function pivotCacheDefinitionPath(n: number | string): string {
  return `xl/pivotCache/pivotCacheDefinition${n}.xml`;
}

export function pivotCacheDefinitionRelsPath(n: number | string): string {
  return `xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`;
}

export function pivotCacheRecordsPath(n: number | string): string {
  return `xl/pivotCache/pivotCacheRecords${n}.xml`;
}

export function pivotCacheRecordsRelTarget(n: number | string): string {
  return `pivotCacheRecords${n}.xml`;
}

export function pivotTablePath(n: number | string): string {
  return `xl/pivotTables/pivotTable${n}.xml`;
}

export function pivotTablePathFromName(name: string): string {
  return `xl/pivotTables/${name}.xml`;
}

export function pivotTableRelsPath(n: number | string): string {
  return `xl/pivotTables/_rels/pivotTable${n}.xml.rels`;
}

// -------- External links --------
export function externalLinkPath(n: number | string): string {
  return `xl/externalLinks/externalLink${n}.xml`;
}

export function externalLinkRelsPath(n: number | string): string {
  return `xl/externalLinks/_rels/externalLink${n}.xml.rels`;
}

/**
 * Build the `Target` value for an externalLink relationship inside
 * `xl/_rels/workbook.xml.rels` (base: `xl/`).
 */
export function externalLinkRelTargetFromWorkbook(n: number | string): string {
  return `externalLinks/externalLink${n}.xml`;
}

export function pivotCacheDefinitionRelTargetFromPivotTable(n: number | string): string {
  return `../pivotCache/pivotCacheDefinition${n}.xml`;
}

export const OOXML_REL_TARGETS = {
  // Targets inside xl/_rels/workbook.xml.rels (base: xl/)
  workbookStyles: "styles.xml",
  workbookSharedStrings: "sharedStrings.xml",
  workbookTheme1: "theme/theme1.xml",
  workbookFeaturePropertyBag: "featurePropertyBag/featurePropertyBag.xml",
  workbookMetadata: "metadata.xml"
} as const;

export function pivotCacheDefinitionRelTargetFromWorkbook(n: number | string): string {
  // Target inside xl/_rels/workbook.xml.rels (base: xl/)
  return `pivotCache/pivotCacheDefinition${n}.xml`;
}

export function commentsRelTargetFromWorksheet(sheetId: number | string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../comments${sheetId}.xml`;
}

export function vmlDrawingRelTargetFromWorksheet(sheetId: number | string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../drawings/vmlDrawing${sheetId}.vml`;
}

export function vmlDrawingHFRelTargetFromWorksheet(sheetId: number | string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../drawings/vmlDrawingHF${sheetId}.vml`;
}

export function drawingRelTargetFromWorksheet(drawingName: string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../drawings/${drawingName}.xml`;
}

export function pivotTableRelTargetFromWorksheet(n: number | string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../pivotTables/pivotTable${n}.xml`;
}

export function tableRelTargetFromWorksheet(target: string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../tables/${target}`;
}

export function mediaRelTargetFromRels(filename: string): string {
  // Target from a rels file located under xl/*/_rels (base is one level deeper than xl/)
  return `../media/${filename}`;
}

// Form Control (ctrlProps) path functions
export function ctrlPropPath(id: number | string): string {
  return `xl/ctrlProps/ctrlProp${id}.xml`;
}

export function ctrlPropRelTargetFromWorksheet(id: number | string): string {
  // Target inside xl/worksheets/_rels/sheetN.xml.rels (base: xl/worksheets/)
  return `../ctrlProps/ctrlProp${id}.xml`;
}

/**
 * Resolve a relationship Target (relative or absolute) to a normalized zip path.
 *
 * OOXML relationship targets may be:
 *   - Relative: `../comments1.xml` (resolved against `baseDir`)
 *   - Absolute: `/xl/comments/comment1.xml` (leading slash stripped)
 *
 * @param baseDir  The directory containing the source part (e.g. `xl/worksheets/`)
 * @param target   The raw Target value from the .rels file
 */
export function resolveRelTarget(baseDir: string, target: string): string {
  // Absolute target — strip leading slash to get the zip path.
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  // Ensure baseDir ends with "/" so the join works correctly.
  const base = baseDir.endsWith("/") ? baseDir : baseDir + "/";
  // Relative target — resolve against baseDir.
  // Simple resolution: join base + target, then resolve `.` and `..` segments.
  const parts = (base + target).split("/");
  const resolved: string[] = [];
  for (const part of parts) {
    if (part === "." || part === "") {
      continue;
    } else if (part === "..") {
      resolved.pop();
    } else {
      resolved.push(part);
    }
  }
  return resolved.join("/");
}
