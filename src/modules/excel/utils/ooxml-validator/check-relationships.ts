/**
 * Relationships check.
 *
 * For every `.rels` file in the package we verify:
 *   - The root element is `<Relationships>` and children are `<Relationship>`.
 *   - Each rel has `Id`, `Type`, and `Target` attributes.
 *   - Relationship `Id` values are unique within the same rels file.
 *   - Non-external targets resolve to files that exist inside the package.
 *   - Resolved targets stay inside the package root (no `..` escape).
 *   - For well-known rel types, the resolved target matches the expected
 *     part-name pattern (e.g. a `Type=...relationships/chart` rel must
 *     point at `xl/charts/chartN.xml`).
 *   - Every non-root rels file has a corresponding source part in the zip.
 *   - The package root rels include an officeDocument rel to `xl/workbook.xml`.
 */

import type { ValidationContext } from "@excel/utils/ooxml-validator/context";
import {
  getRelsSourceDir,
  isSafeResolvedPath,
  posixBasename,
  resolveRelTarget,
  sourcePartForRels
} from "@excel/utils/ooxml-validator/path-utils";

// -----------------------------------------------------------------------------
// Rel-type -> expected target pattern.
// -----------------------------------------------------------------------------

interface RelTypeTargetRule {
  type: string;
  pattern: RegExp;
  label: string;
}

const REL_TYPE_TARGET_RULES: RelTypeTargetRule[] = [
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
    pattern: /^xl\/workbook\.xml$/,
    label: "workbook"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
    pattern: /^xl\/worksheets\/sheet\d+\.xml$/,
    label: "worksheet"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    pattern: /^xl\/chartsheets\/sheet\d+\.xml$/,
    label: "chartsheet"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
    pattern: /^xl\/charts\/chart\d+\.xml$/,
    label: "chart"
  },
  {
    type: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
    pattern: /^xl\/charts\/chartEx\d+\.xml$/,
    label: "chartEx"
  },
  {
    type: "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
    pattern: /^xl\/charts\/style(?:Ex)?\d+\.xml$/,
    label: "chart style"
  },
  {
    type: "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
    pattern: /^xl\/charts\/colors(?:Ex)?\d+\.xml$/,
    label: "chart colors"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
    pattern: /^xl\/drawings\/drawing\d+\.xml$/,
    label: "drawing"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    pattern: /^xl\/styles\.xml$/,
    label: "styles"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
    pattern: /^xl\/sharedStrings\.xml$/,
    label: "shared strings"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    pattern: /^xl\/theme\/[^/]+\.xml$/,
    label: "theme"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable",
    pattern: /^xl\/pivotTables\/pivotTable\d+\.xml$/,
    label: "pivot table"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
    pattern: /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/,
    label: "pivot cache definition"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords",
    pattern: /^xl\/pivotCache\/pivotCacheRecords\d+\.xml$/,
    label: "pivot cache records"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
    pattern: /^xl\/tables\/table\d+\.xml$/,
    label: "table"
  },
  {
    type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
    pattern: /^xl\/externalLinks\/externalLink\d+\.xml$/,
    label: "external link"
  }
];

const REL_TYPE_TARGET_BY_TYPE = new Map(REL_TYPE_TARGET_RULES.map(r => [r.type, r]));

// -----------------------------------------------------------------------------
// Checker
// -----------------------------------------------------------------------------

export function checkRootRelationships(ctx: ValidationContext): void {
  if (!ctx.has("_rels/.rels")) {
    return; // missing-part already reported.
  }
  const rels = ctx.readRels("_rels/.rels").rels;
  const hasOfficeDocument = rels.some(
    r => r.type.includes("/relationships/officeDocument") && r.target === "xl/workbook.xml"
  );
  if (!hasOfficeDocument) {
    ctx.reporter.error(
      "root-rels-missing-officeDocument",
      "Missing officeDocument relationship to xl/workbook.xml",
      "_rels/.rels"
    );
  }
}

export function checkRelationships(ctx: ValidationContext): void {
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || !path.endsWith(".rels")) {
      continue;
    }
    checkSingleRelsFile(ctx, path);
  }
}

function checkSingleRelsFile(ctx: ValidationContext, relsPath: string): void {
  const data = ctx.readRels(relsPath);
  // Missing attributes.
  for (const m of data.malformedEntries) {
    if (ctx.reporter.capped) {
      return;
    }
    if (m.missingId) {
      ctx.reporter.error("rels-missing-id-attr", `Relationship missing Id attribute`, relsPath);
    }
    if (m.missingType) {
      ctx.reporter.error(
        "rels-missing-type-attr",
        `Relationship ${m.id ?? "<no-id>"} missing Type attribute`,
        relsPath
      );
    }
    if (m.missingTarget) {
      ctx.reporter.error(
        "rels-missing-target",
        `Relationship ${m.id ?? "<no-id>"} missing Target attribute`,
        relsPath
      );
    }
  }

  // Duplicate IDs — scan raw array (parsed keeps first).
  const seen = new Set<string>();
  for (const rel of data.rels) {
    if (ctx.reporter.capped) {
      return;
    }
    if (seen.has(rel.id)) {
      ctx.reporter.error("rels-duplicate-id", `Duplicate relationship Id: ${rel.id}`, relsPath);
    } else {
      seen.add(rel.id);
    }

    if (rel.targetMode === "External") {
      continue;
    }
    if (rel.target === "") {
      ctx.reporter.error(
        "rels-empty-target",
        `Relationship ${rel.id} (${rel.type}) has empty Target`,
        relsPath
      );
      continue;
    }
    const resolved = resolveRelTarget(relsPath, rel.target);
    if (!isSafeResolvedPath(resolved)) {
      ctx.reporter.error(
        "rels-invalid-target-path",
        `Rel ${rel.id} (${rel.type}) target escapes package root: ${rel.target} -> ${resolved}`,
        relsPath
      );
      continue;
    }
    if (!ctx.has(resolved)) {
      ctx.reporter.error(
        "rels-missing-target",
        `Rel ${rel.id} (${rel.type}) target missing: ${rel.target} -> ${resolved}`,
        relsPath
      );
      continue;
    }
    // Rel type -> target pattern.
    const rule = REL_TYPE_TARGET_BY_TYPE.get(rel.type);
    if (rule && !rule.pattern.test(resolved)) {
      ctx.reporter.error(
        "rels-type-target-mismatch",
        `Rel ${rel.id} (${rule.label}) points at unexpected target: ${resolved}`,
        relsPath
      );
    }
  }

  // Source part must exist.
  if (relsPath !== "_rels/.rels") {
    const srcDir = getRelsSourceDir(relsPath);
    const srcName = posixBasename(relsPath).replace(/\.rels$/, "");
    const sourcePath = sourcePartForRels(relsPath) ?? (srcDir ? `${srcDir}/${srcName}` : srcName);
    if (!ctx.has(sourcePath)) {
      ctx.reporter.error(
        "rels-source-missing",
        `Relationships part has no corresponding source part: ${sourcePath}`,
        relsPath
      );
    }
  }
}
