import type { ExtractedFile } from "@archive/unzip/extract";
import { attr, findChildren, parseXml, walk } from "@xml/dom";
import type { XmlElement } from "@xml/types";

const RELS_SUFFIX = ".rels";

const CONTENT_TYPES: Record<string, string> = {
  "xl/workbook.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  "xl/styles.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  "xl/sharedStrings.xml":
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  "xl/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  chartEx: "application/vnd.ms-office.chartEx+xml",
  chartStyle: "application/vnd.ms-office.chartstyle+xml",
  chartColors: "application/vnd.ms-office.chartcolorstyle+xml",
  chartExStyle: "application/vnd.ms-office.chartstyle+xml",
  chartExColors: "application/vnd.ms-office.chartcolorstyle+xml",
  chartsheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
  drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
  pivotTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
  pivotCacheDefinition:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
  pivotCacheRecords:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
};

const REL_TYPES = {
  chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
  chartEx: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
  chartStyle: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
  chartColors: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
  drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
  pivotTable: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
  pivotCacheDefinition:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
  pivotCacheRecords:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
  worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  chartsheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  sharedStrings:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  theme: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
} as const;

export interface OoxmlPackageAuditOptions {
  requireReferencedParts?: boolean;
}

export interface OoxmlPackageAuditResult {
  errors: string[];
  warnings: string[];
}

export function auditOoxmlPackage(
  entries: Map<string, ExtractedFile>,
  options: OoxmlPackageAuditOptions = {}
): OoxmlPackageAuditResult {
  const errors: string[] = [];
  const warnings: string[] = [];
  const paths = new Set(entries.keys());
  const contentTypes = readContentTypes(entries, errors);

  requirePart(paths, "xl/workbook.xml", errors);
  requirePart(paths, "xl/_rels/workbook.xml.rels", errors);
  requireContentType(contentTypes, "xl/workbook.xml", CONTENT_TYPES["xl/workbook.xml"], errors);

  for (const path of paths) {
    auditKnownContentType(path, contentTypes, errors);
    if (path.endsWith(RELS_SUFFIX)) {
      auditRelationships(
        path,
        entries,
        paths,
        errors,
        warnings,
        options.requireReferencedParts !== false
      );
    }
    if (/^xl\/charts\/chart\d+[.]xml$/.test(path)) {
      auditClassicChart(path, entries, errors);
    }
    if (/^xl\/charts\/chartEx\d+[.]xml$/.test(path)) {
      auditChartEx(path, entries, errors);
    }
    if (/^xl\/drawings\/drawing\d+[.]xml$/.test(path)) {
      auditDrawing(path, entries, errors);
    }
    if (/^xl\/chartsheets\/sheet\d+[.]xml$/.test(path)) {
      auditChartsheet(path, entries, errors);
    }
  }

  return { errors, warnings };
}

function readContentTypes(
  entries: Map<string, ExtractedFile>,
  errors: string[]
): Map<string, string> {
  const entry = entries.get("[Content_Types].xml");
  const result = new Map<string, string>();
  if (!entry) {
    errors.push("Missing [Content_Types].xml");
    return result;
  }
  let root: XmlElement;
  try {
    root = parseXml(decode(entry.data)).root;
  } catch (error) {
    errors.push(`Invalid [Content_Types].xml: ${String(error)}`);
    return result;
  }
  for (const override of findChildren(root, "Override")) {
    const partName = attr(override, "PartName");
    const contentType = attr(override, "ContentType");
    if (partName && contentType) {
      result.set(normalizePartName(partName), contentType);
    }
  }
  return result;
}

function auditKnownContentType(
  path: string,
  contentTypes: Map<string, string>,
  errors: string[]
): void {
  if (path.endsWith(RELS_SUFFIX) || path === "[Content_Types].xml") {
    return;
  }
  const expected = expectedContentType(path);
  if (expected) {
    requireContentType(contentTypes, path, expected, errors);
  }
}

function expectedContentType(path: string): string | undefined {
  if (/^xl\/charts\/chart\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chart;
  }
  if (/^xl\/charts\/chartEx\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartEx;
  }
  if (/^xl\/charts\/style\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartStyle;
  }
  if (/^xl\/charts\/colors\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartColors;
  }
  if (/^xl\/charts\/styleEx\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartExStyle;
  }
  if (/^xl\/charts\/colorsEx\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartExColors;
  }
  if (/^xl\/chartsheets\/sheet\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.chartsheet;
  }
  if (/^xl\/drawings\/drawing\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.drawing;
  }
  if (/^xl\/pivotTables\/pivotTable\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.pivotTable;
  }
  if (/^xl\/pivotCache\/pivotCacheDefinition\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.pivotCacheDefinition;
  }
  if (/^xl\/pivotCache\/pivotCacheRecords\d+[.]xml$/.test(path)) {
    return CONTENT_TYPES.pivotCacheRecords;
  }
  return CONTENT_TYPES[path];
}

function auditRelationships(
  relsPath: string,
  entries: Map<string, ExtractedFile>,
  paths: Set<string>,
  errors: string[],
  warnings: string[],
  requireReferencedParts: boolean
): void {
  const rels = readRelationships(relsPath, entries, errors);
  const seenIds = new Set<string>();
  const sourceDir = relsSourceDir(relsPath);
  for (const rel of rels) {
    if (!rel.id) {
      errors.push(`${relsPath}: relationship is missing Id`);
    } else if (seenIds.has(rel.id)) {
      errors.push(`${relsPath}: duplicate relationship Id ${rel.id}`);
    } else {
      seenIds.add(rel.id);
    }
    if (!rel.type) {
      errors.push(`${relsPath}: relationship ${rel.id || "<unknown>"} is missing Type`);
    }
    if (!rel.target) {
      errors.push(`${relsPath}: relationship ${rel.id || "<unknown>"} is missing Target`);
      continue;
    }
    if (rel.targetMode === "External") {
      continue;
    }
    const resolved = resolveRelTarget(sourceDir, rel.target);
    if (requireReferencedParts && !paths.has(resolved)) {
      errors.push(`${relsPath}: ${rel.id || rel.type} targets missing part ${resolved}`);
    }
    auditRelationshipTypeTarget(relsPath, rel.type, resolved, errors, warnings);
  }
}

function auditRelationshipTypeTarget(
  relsPath: string,
  type: string | undefined,
  target: string,
  errors: string[],
  warnings: string[]
): void {
  if (!type) {
    return;
  }
  const checks: Array<[string, RegExp, string]> = [
    [REL_TYPES.chart, /^xl\/charts\/chart\d+[.]xml$/, "chart"],
    [REL_TYPES.chartEx, /^xl\/charts\/chartEx\d+[.]xml$/, "chartEx"],
    [REL_TYPES.chartStyle, /^xl\/charts\/style(?:Ex)?\d+[.]xml$/, "chart style"],
    [REL_TYPES.chartColors, /^xl\/charts\/colors(?:Ex)?\d+[.]xml$/, "chart colors"],
    [REL_TYPES.drawing, /^xl\/drawings\/drawing\d+[.]xml$/, "drawing"],
    [REL_TYPES.pivotTable, /^xl\/pivotTables\/pivotTable\d+[.]xml$/, "pivot table"],
    [
      REL_TYPES.pivotCacheDefinition,
      /^xl\/pivotCache\/pivotCacheDefinition\d+[.]xml$/,
      "pivot cache definition"
    ],
    [
      REL_TYPES.pivotCacheRecords,
      /^xl\/pivotCache\/pivotCacheRecords\d+[.]xml$/,
      "pivot cache records"
    ],
    [REL_TYPES.worksheet, /^xl\/worksheets\/sheet\d+[.]xml$/, "worksheet"],
    [REL_TYPES.chartsheet, /^xl\/chartsheets\/sheet\d+[.]xml$/, "chartsheet"],
    [REL_TYPES.styles, /^xl\/styles[.]xml$/, "styles"],
    [REL_TYPES.sharedStrings, /^xl\/sharedStrings[.]xml$/, "shared strings"],
    [REL_TYPES.theme, /^xl\/theme\/[^/]+[.]xml$/, "theme"]
  ];
  const check = checks.find(([candidate]) => candidate === type);
  if (!check) {
    return;
  }
  if (!check[1].test(target)) {
    warnings.push(`${relsPath}: ${check[2]} relationship points at unexpected target ${target}`);
  }
}

function auditClassicChart(
  path: string,
  entries: Map<string, ExtractedFile>,
  errors: string[]
): void {
  const root = readXmlRoot(path, entries, errors);
  if (!root) {
    return;
  }
  if (!hasDescendant(root, "c:chart")) {
    errors.push(`${path}: missing c:chart`);
  }
  if (!hasDescendant(root, "c:plotArea")) {
    errors.push(`${path}: missing c:plotArea`);
  }
}

function auditChartEx(path: string, entries: Map<string, ExtractedFile>, errors: string[]): void {
  const root = readXmlRoot(path, entries, errors);
  if (!root) {
    return;
  }
  if (!hasDescendant(root, "cx:chart")) {
    errors.push(`${path}: missing cx:chart`);
  }
  if (!hasDescendant(root, "cx:plotArea")) {
    errors.push(`${path}: missing cx:plotArea`);
  }
  if (!hasDescendant(root, "cx:series")) {
    errors.push(`${path}: missing cx:series`);
  }
  const dataIds = new Set(
    collectDescendants(root, "cx:data")
      .map(el => parseInt(attr(el, "id") ?? "", 10))
      .filter(Number.isFinite)
  );
  const axisIds = new Set(
    collectDescendants(root, "cx:axis")
      .map(el => parseInt(attr(el, "id") ?? "", 10))
      .filter(Number.isFinite)
  );
  for (const series of collectDescendants(root, "cx:series")) {
    if (!attr(series, "layoutId")) {
      errors.push(`${path}: cx:series is missing layoutId`);
    }
    for (const dataId of findChildren(series, "cx:dataId")) {
      const id = parseInt(attr(dataId, "val") ?? "", 10);
      if (!dataIds.has(id)) {
        errors.push(`${path}: cx:series references missing cx:data id ${attr(dataId, "val")}`);
      }
    }
    for (const axisId of findChildren(series, "cx:axisId")) {
      const id = parseInt(attr(axisId, "val") ?? "", 10);
      if (!axisIds.has(id)) {
        errors.push(`${path}: cx:series references missing cx:axis id ${attr(axisId, "val")}`);
      }
    }
  }
  const externalDataIds = collectDescendants(root, "cx:externalData")
    .map(el => attr(el, "r:id"))
    .filter((id): id is string => !!id);
  if (externalDataIds.length > 0) {
    const relsPath = chartPartRelsPath(path);
    const relIds = new Set(readRelationships(relsPath, entries, errors).map(rel => rel.id));
    for (const id of externalDataIds) {
      if (!relIds.has(id)) {
        errors.push(`${path}: cx:externalData references missing relationship ${id}`);
      }
    }
  }
}

function auditDrawing(path: string, entries: Map<string, ExtractedFile>, errors: string[]): void {
  const root = readXmlRoot(path, entries, errors);
  if (!root) {
    return;
  }
  const hasChartFrame = hasDescendant(root, "c:chart") || hasDescendant(root, "cx:chart");
  if (!hasChartFrame && hasDescendant(root, "xdr:graphicFrame")) {
    errors.push(`${path}: graphicFrame is missing c:chart or cx:chart reference`);
  }
}

function auditChartsheet(
  path: string,
  entries: Map<string, ExtractedFile>,
  errors: string[]
): void {
  const root = readXmlRoot(path, entries, errors);
  if (!root) {
    return;
  }
  if (!hasDescendant(root, "drawing")) {
    errors.push(`${path}: missing drawing reference`);
  }
}

function readRelationships(
  path: string,
  entries: Map<string, ExtractedFile>,
  errors: string[]
): Array<{ id?: string; type?: string; target?: string; targetMode?: string }> {
  const root = readXmlRoot(path, entries, errors);
  if (!root) {
    return [];
  }
  return findChildren(root, "Relationship").map(rel => ({
    id: attr(rel, "Id"),
    type: attr(rel, "Type"),
    target: attr(rel, "Target"),
    targetMode: attr(rel, "TargetMode")
  }));
}

function readXmlRoot(
  path: string,
  entries: Map<string, ExtractedFile>,
  errors: string[]
): XmlElement | undefined {
  const entry = entries.get(path);
  if (!entry) {
    errors.push(`Missing ${path}`);
    return undefined;
  }
  try {
    return parseXml(decode(entry.data)).root;
  } catch (error) {
    errors.push(`${path}: invalid XML: ${String(error)}`);
    return undefined;
  }
}

function hasDescendant(root: XmlElement, name: string): boolean {
  let found = root.name === name;
  if (found) {
    return true;
  }
  walk(root, child => {
    if (child.name === name) {
      found = true;
    }
  });
  return found;
}

function collectDescendants(root: XmlElement, name: string): XmlElement[] {
  const result: XmlElement[] = root.name === name ? [root] : [];
  walk(root, child => {
    if (child.name === name) {
      result.push(child);
    }
  });
  return result;
}

function chartPartRelsPath(path: string): string {
  const slash = path.lastIndexOf("/");
  const dir = slash >= 0 ? path.slice(0, slash) : "";
  const name = slash >= 0 ? path.slice(slash + 1) : path;
  return `${dir}/_rels/${name}.rels`;
}

function requirePart(paths: Set<string>, path: string, errors: string[]): void {
  if (!paths.has(path)) {
    errors.push(`Missing ${path}`);
  }
}

function requireContentType(
  contentTypes: Map<string, string>,
  path: string,
  expected: string,
  errors: string[]
): void {
  const actual = contentTypes.get(path);
  if (!actual) {
    errors.push(`Missing content type override for /${path}`);
  } else if (actual !== expected) {
    errors.push(`Unexpected content type for /${path}: ${actual}`);
  }
}

function relsSourceDir(relsPath: string): string {
  if (relsPath === "_rels/.rels") {
    return "";
  }
  const marker = "/_rels/";
  const index = relsPath.indexOf(marker);
  if (index < 0) {
    return dirname(relsPath);
  }
  return relsPath.slice(0, index + 1);
}

function resolveRelTarget(baseDir: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePartName(target);
  }
  const parts = `${baseDir}${target}`.split("/");
  const resolved: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      resolved.pop();
    } else {
      resolved.push(part);
    }
  }
  return resolved.join("/");
}

function dirname(path: string): string {
  const index = path.lastIndexOf("/");
  return index < 0 ? "" : path.slice(0, index + 1);
}

function normalizePartName(path: string): string {
  return path.startsWith("/") ? path.slice(1) : path;
}

function decode(data: Uint8Array): string {
  return new TextDecoder().decode(data);
}
