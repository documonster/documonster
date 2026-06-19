/**
 * Content-types check.
 *
 * OPC requires every part in the package to be resolvable to a
 * content-type via either a `<Default Extension=...>` entry or a
 * `<Override PartName=...>` entry. Excel rejects packages where this
 * mapping has gaps — the classic "repairable content" error.
 *
 * Beyond the OPC-mandatory rules, we also enforce Microsoft's
 * well-known part→type expectations for the common OOXML part layouts
 * (chart, pivot, chartsheet, drawing, etc.). A wrong content-type on
 * these parts is one of the top causes of "Excel cannot open" errors
 * because Excel uses the type string to route parsing.
 */

import type { ValidationContext } from "@excel/utils/ooxml-validator/context";
import {
  getExtension,
  isLegalPartName,
  stripLeadingSlash
} from "@excel/utils/ooxml-validator/path-utils";

// -----------------------------------------------------------------------------
// Well-known content types
// -----------------------------------------------------------------------------

const CT = {
  workbook: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  sharedStrings: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  theme: "application/vnd.openxmlformats-officedocument.theme+xml",
  worksheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  chartsheet: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
  chart: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
  chartEx: "application/vnd.ms-office.chartex+xml",
  chartStyle: "application/vnd.ms-office.chartstyle+xml",
  chartColors: "application/vnd.ms-office.chartcolorstyle+xml",
  drawing: "application/vnd.openxmlformats-officedocument.drawing+xml",
  pivotTable: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml",
  pivotCacheDefinition:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml",
  pivotCacheRecords:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
  comments: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
  threadedComments: "application/vnd.ms-excel.threadedcomments+xml",
  externalLink: "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml",
  tableDef: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
  calcChain: "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
  app: "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  core: "application/vnd.openxmlformats-package.core-properties+xml"
} as const;

/**
 * Path→expected content-type rules. The first matching rule wins. We
 * keep these in a single table so it is easy to audit and extend. Exact
 * string keys take precedence over regex rules for performance and
 * readability.
 */
const EXACT_PART_CONTENT_TYPE: Record<string, string> = {
  "xl/workbook.xml": CT.workbook,
  "xl/styles.xml": CT.styles,
  "xl/sharedStrings.xml": CT.sharedStrings,
  "xl/calcChain.xml": CT.calcChain,
  "docProps/app.xml": CT.app,
  "docProps/core.xml": CT.core
};

interface PatternRule {
  pattern: RegExp;
  contentType: string;
  label: string;
}

const PATTERN_RULES: PatternRule[] = [
  { pattern: /^xl\/theme\/[^/]+\.xml$/, contentType: CT.theme, label: "theme" },
  { pattern: /^xl\/worksheets\/sheet\d+\.xml$/, contentType: CT.worksheet, label: "worksheet" },
  { pattern: /^xl\/chartsheets\/sheet\d+\.xml$/, contentType: CT.chartsheet, label: "chartsheet" },
  { pattern: /^xl\/charts\/chart\d+\.xml$/, contentType: CT.chart, label: "chart" },
  { pattern: /^xl\/charts\/chartEx\d+\.xml$/, contentType: CT.chartEx, label: "chartEx" },
  { pattern: /^xl\/charts\/style\d+\.xml$/, contentType: CT.chartStyle, label: "chart style" },
  { pattern: /^xl\/charts\/colors\d+\.xml$/, contentType: CT.chartColors, label: "chart colors" },
  { pattern: /^xl\/charts\/styleEx\d+\.xml$/, contentType: CT.chartStyle, label: "chart styleEx" },
  {
    pattern: /^xl\/charts\/colorsEx\d+\.xml$/,
    contentType: CT.chartColors,
    label: "chart colorsEx"
  },
  { pattern: /^xl\/drawings\/drawing\d+\.xml$/, contentType: CT.drawing, label: "drawing" },
  {
    pattern: /^xl\/pivotTables\/pivotTable\d+\.xml$/,
    contentType: CT.pivotTable,
    label: "pivot table"
  },
  {
    pattern: /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/,
    contentType: CT.pivotCacheDefinition,
    label: "pivot cache definition"
  },
  {
    pattern: /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/,
    contentType: CT.pivotCacheRecords,
    label: "pivot cache records"
  },
  { pattern: /^xl\/tables\/table\d+\.xml$/, contentType: CT.tableDef, label: "table" },
  {
    pattern: /^xl\/externalLinks\/externalLink\d+\.xml$/,
    contentType: CT.externalLink,
    label: "external link"
  },
  { pattern: /^xl\/comments\d+\.xml$/, contentType: CT.comments, label: "comments" },
  {
    pattern: /^xl\/threadedComments\/threadedComment\d+\.xml$/,
    contentType: CT.threadedComments,
    label: "threaded comments"
  }
];

function expectedContentType(path: string): { contentType: string; label: string } | undefined {
  const exact = EXACT_PART_CONTENT_TYPE[path];
  if (exact) {
    return { contentType: exact, label: path };
  }
  for (const rule of PATTERN_RULES) {
    if (rule.pattern.test(path)) {
      return { contentType: rule.contentType, label: rule.label };
    }
  }
  return undefined;
}

// -----------------------------------------------------------------------------
// Checker
// -----------------------------------------------------------------------------

export function checkContentTypes(ctx: ValidationContext): void {
  if (!ctx.has("[Content_Types].xml")) {
    // checkStructure will already have reported missing-part.
    return;
  }
  const data = ctx.readContentTypes();
  if (!data.parseOk) {
    // context already reported `content-types-malformed`.
    return;
  }

  // Duplicate overrides — we keep the first and report extras.
  for (const pn of data.duplicateOverrides) {
    ctx.reporter.error(
      "content-types-duplicate-override",
      `Duplicate Override PartName: ${pn}`,
      "[Content_Types].xml"
    );
  }

  // The two OPC-mandated defaults.
  const relsDefault = data.defaults.get("rels");
  if (relsDefault !== "application/vnd.openxmlformats-package.relationships+xml") {
    ctx.reporter.error(
      "content-types-missing-default",
      "Missing/incorrect Default for .rels (expected application/vnd.openxmlformats-package.relationships+xml)",
      "[Content_Types].xml"
    );
  }
  const xmlDefault = data.defaults.get("xml");
  if (xmlDefault !== "application/xml") {
    ctx.reporter.error(
      "content-types-missing-default",
      "Missing/incorrect Default for .xml (expected application/xml)",
      "[Content_Types].xml"
    );
  }

  // Every override must point to an existing part.
  for (const [partName, _ct] of data.overrides) {
    if (ctx.reporter.capped) {
      return;
    }
    if (!ctx.has(partName)) {
      ctx.reporter.error(
        "content-types-missing",
        `Override PartName points to missing file: /${partName}`,
        "[Content_Types].xml"
      );
    }
  }

  // Every part in the zip must resolve to a content type and — when we
  // know the expected type — match it.
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || path === "[Content_Types].xml") {
      continue;
    }

    // Reject invalid OPC part names (e.g. containing `\`, ending `/`, or
    // embedded `..`). Excel refuses such packages outright.
    if (!isLegalPartName(path)) {
      ctx.reporter.error("part-name-invalid", `Illegal OPC part name: ${path}`, path);
    }

    const overrideType = data.overrides.get(path);
    const ext = getExtension(path);
    const defaultType = ext ? data.defaults.get(ext) : undefined;

    if (!overrideType && !defaultType) {
      if (!ext) {
        ctx.reporter.error(
          "content-types-missing-for-part",
          `No content type for part without extension: ${path}`,
          "[Content_Types].xml"
        );
      } else {
        ctx.reporter.error(
          "content-types-missing-for-part",
          `No Default/Override content type for part: ${path} (extension .${ext})`,
          "[Content_Types].xml"
        );
      }
      continue;
    }

    const expected = expectedContentType(path);
    if (!expected) {
      continue;
    }
    const actual = overrideType ?? defaultType;
    if (actual !== expected.contentType) {
      ctx.reporter.error(
        "content-types-wrong-for-part",
        `Unexpected content type for /${stripLeadingSlash(path)}: got "${actual}", expected "${expected.contentType}" (${expected.label})`,
        "[Content_Types].xml"
      );
    }
  }
}
