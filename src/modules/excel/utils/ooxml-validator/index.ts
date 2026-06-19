/**
 * Public entry point for the OOXML validator.
 *
 * Validates an xlsx package buffer against the OPC spec and Excel's
 * known opening invariants. See `types.ts` for the full catalog of
 * problem kinds.
 *
 * ```ts
 * const report = await validateXlsxBuffer(await Workbook.toBuffer(wb));
 * if (!report.ok) {
 *   console.error(report.problems);
 * }
 * ```
 *
 * The implementation is composed of independent checkers under
 * `check-*.ts` — each can be unit-tested in isolation. The entry point
 * orchestrates them with a shared `ValidationContext` that caches XML
 * parsing across checks.
 */

import type { ExtractedFile } from "@archive/unzip/extract";
import { extractAll } from "@archive/unzip/extract";
import { checkChart } from "@excel/utils/ooxml-validator/check-chart";
import { checkChartSidecars } from "@excel/utils/ooxml-validator/check-chart-sidecar";
import { checkChartsheet } from "@excel/utils/ooxml-validator/check-chartsheet";
import { checkContentTypes } from "@excel/utils/ooxml-validator/check-content-types";
import { checkDrawing } from "@excel/utils/ooxml-validator/check-drawing";
import { checkPivot } from "@excel/utils/ooxml-validator/check-pivot";
import {
  checkRelationships,
  checkRootRelationships
} from "@excel/utils/ooxml-validator/check-relationships";
import { checkStructure, checkXmlWellFormed } from "@excel/utils/ooxml-validator/check-structure";
import { checkStyles } from "@excel/utils/ooxml-validator/check-styles";
import { checkTables } from "@excel/utils/ooxml-validator/check-table";
import { checkWorkbook } from "@excel/utils/ooxml-validator/check-workbook";
import { checkWorksheets } from "@excel/utils/ooxml-validator/check-worksheet";
import { ValidationContext } from "@excel/utils/ooxml-validator/context";
import { isXmlLike } from "@excel/utils/ooxml-validator/path-utils";
import { Reporter } from "@excel/utils/ooxml-validator/reporter";
import type {
  OoxmlOrderingProblemKind,
  OoxmlOrderingValidationProblem,
  OoxmlProblemKind,
  OoxmlProblemSeverity,
  OoxmlValidateOptions,
  OoxmlValidationProblem,
  OoxmlValidationReport,
  OoxmlValidationStats
} from "@excel/utils/ooxml-validator/types";

// Re-export types at the public boundary.
export type {
  OoxmlOrderingProblemKind,
  OoxmlOrderingValidationProblem,
  OoxmlProblemKind,
  OoxmlProblemSeverity,
  OoxmlValidateOptions,
  OoxmlValidationProblem,
  OoxmlValidationReport,
  OoxmlValidationStats
};

/**
 * Validate an xlsx package. Returns a report with every detected
 * problem; `ok === true` when no error-severity problems were found.
 *
 * This function never throws — malformed input is surfaced as a
 * `xml-malformed` or similar problem kind. The only exception is if
 * `extractAll` fails to unzip the buffer (not a valid ZIP at all), in
 * which case the underlying archive error propagates.
 */
export async function validateXlsxBuffer(
  xlsxBuffer: Uint8Array,
  options: OoxmlValidateOptions = {}
): Promise<OoxmlValidationReport> {
  const {
    checkXmlWellFormed: optCheckXmlWellFormed = true,
    checkRelationshipTargets = true,
    checkContentTypesOverrides = true,
    checkWorksheetControlWiring = true,
    checkChartStructure = true,
    checkStylesIntegrity = true,
    maxProblems,
    includeWarnings = false
  } = options;

  const entries = await extractAll(xlsxBuffer);
  const reporter = new Reporter({ maxProblems, includeWarnings });
  const ctx = new ValidationContext(entries, reporter);

  // Order matters: later checks rely on earlier ones NOT firing false
  // positives when structural invariants are broken. Structure first.
  checkStructure(ctx);

  if (optCheckXmlWellFormed) {
    checkXmlWellFormed(ctx);
  }

  if (checkContentTypesOverrides) {
    checkContentTypes(ctx);
  }

  checkRootRelationships(ctx);

  if (checkRelationshipTargets) {
    checkRelationships(ctx);
  }

  checkWorkbook(ctx);

  if (checkWorksheetControlWiring) {
    checkWorksheets(ctx);
  }

  if (checkChartStructure) {
    checkChart(ctx);
    checkDrawing(ctx);
    checkChartsheet(ctx);
    checkPivot(ctx);
    checkTables(ctx);
    checkChartSidecars(ctx);
  }

  if (checkStylesIntegrity) {
    checkStyles(ctx);
  }

  const stats = computeStats(entries);
  return {
    ok: !reporter.hasErrors,
    problems: reporter.problems,
    stats
  };
}

function computeStats(entries: Map<string, ExtractedFile>): OoxmlValidationStats {
  let entryCount = 0;
  let xmlLikeCount = 0;
  let relsCount = 0;
  for (const [, entry] of entries) {
    if (entry.type === "directory") {
      continue;
    }
    entryCount++;
    if (isXmlLike(entry.path)) {
      xmlLikeCount++;
    }
    if (entry.path.endsWith(".rels")) {
      relsCount++;
    }
  }
  return { entryCount, xmlLikeCount, relsCount };
}
