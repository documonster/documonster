/**
 * Workbook-level check.
 *
 * Reads `xl/workbook.xml` as DOM (replacing the previous regex-based
 * approach, which missed namespaced elements like `<x:sheet>`) and walks
 * every `<sheet>` declaration to verify:
 *
 *   - `sheetId` uniqueness (OOXML requires globally unique sheet IDs).
 *   - `r:id` uniqueness and pointing to a worksheet-type relationship.
 *   - `name` attribute presence and Excel's name constraints:
 *       · 1–31 characters
 *       · No characters in `\ / ? * [ ] :`
 *       · Not repeated case-insensitively
 *   - CT_Workbook child element ordering (ECMA-376 Part 1 §18.2.27).
 *     Excel rejects files whose direct children of `<workbook>` violate
 *     the XSD sequence order.
 */

import type { ValidationContext } from "./context";
import { attrByLocalName, findChildLocal, findChildrenLocal, localName } from "./xml-utils";

const WORKBOOK_PATH = "xl/workbook.xml";
const WORKBOOK_RELS_PATH = "xl/_rels/workbook.xml.rels";

// Excel rejects sheet names that contain any of these characters. Source:
// https://support.microsoft.com/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
const INVALID_SHEET_NAME_CHARS = /[\\/?*[\]:]/;
const SHEET_NAME_MAX_LEN = 31;

/**
 * CT_Workbook child element sequence per ECMA-376 Part 1 §18.2.27.
 *
 * Excel strictly enforces this order: any child appearing AFTER another
 * that should precede it causes the file to be rejected as corrupt
 * (no repair offered). The list below covers every element in the
 * standard sequence. Unknown / extension elements (mc:AlternateContent,
 * extLst, etc.) are ignored — they can appear anywhere without breaking
 * conformance.
 */
const WORKBOOK_CHILD_ORDER: readonly string[] = [
  "fileVersion",
  "fileSharing",
  "workbookPr",
  "workbookProtection",
  "bookViews",
  "sheets",
  "functionGroups",
  "externalReferences",
  "definedNames",
  "calcPr",
  "oleSize",
  "customWorkbookViews",
  "pivotCaches",
  "smartTagPr",
  "smartTagTypes",
  "webPublishing",
  "fileRecoveryPr",
  "webPublishObjects",
  "extLst"
];

export function checkWorkbook(ctx: ValidationContext): void {
  if (!ctx.has(WORKBOOK_PATH) || !ctx.has(WORKBOOK_RELS_PATH)) {
    return; // missing-part already reported.
  }

  const dom = ctx.readDom(WORKBOOK_PATH, err => {
    ctx.reporter.error("xml-malformed", `Malformed XML: ${err.message}`, WORKBOOK_PATH);
  });
  if (!dom) {
    return;
  }
  const rels = ctx.readRels(WORKBOOK_RELS_PATH);

  // --- Element order check ---
  checkWorkbookChildOrder(ctx, dom.root);

  // --- Sheet checks ---
  const sheets = findChildLocal(dom.root, "sheets");
  if (!sheets) {
    // workbook with zero sheets is handled elsewhere as a structural issue
    // — here we only care about sheet wiring.
    return;
  }

  const seenSheetId = new Set<string>();
  const seenSheetRid = new Set<string>();
  const seenNamesLower = new Set<string>();

  for (const sheet of findChildrenLocal(sheets, "sheet")) {
    if (ctx.reporter.capped) {
      return;
    }

    const sheetId = attrByLocalName(sheet, "sheetId");
    const rid = attrByLocalName(sheet, "id"); // r:id — local name "id"
    const name = attrByLocalName(sheet, "name");

    if (sheetId !== undefined) {
      if (seenSheetId.has(sheetId)) {
        ctx.reporter.error(
          "workbook-duplicate-sheetId",
          `Duplicate sheetId in workbook: ${sheetId}`,
          WORKBOOK_PATH
        );
      } else {
        seenSheetId.add(sheetId);
      }
    }

    if (rid !== undefined) {
      if (seenSheetRid.has(rid)) {
        ctx.reporter.error(
          "workbook-duplicate-sheet-rid",
          `Duplicate sheet r:id in workbook: ${rid}`,
          WORKBOOK_PATH
        );
      } else {
        seenSheetRid.add(rid);
      }
      const rel = rels.byId.get(rid);
      if (!rel) {
        ctx.reporter.error(
          "workbook-sheet-missing-rel",
          `Workbook <sheet> references missing relationship: ${rid} (in ${WORKBOOK_RELS_PATH})`,
          WORKBOOK_PATH
        );
      } else if (
        !rel.type.includes("/relationships/worksheet") &&
        !rel.type.includes("/relationships/chartsheet")
      ) {
        ctx.reporter.error(
          "workbook-sheet-wrong-rel-type",
          `Workbook <sheet> ${rid} relationship is neither worksheet nor chartsheet: ${rel.type}`,
          WORKBOOK_PATH
        );
      }
    }

    if (name === undefined) {
      ctx.reporter.error(
        "workbook-sheet-missing-name",
        `Workbook <sheet> (sheetId=${sheetId ?? "?"}) missing name attribute`,
        WORKBOOK_PATH
      );
      continue;
    }
    if (name.length === 0 || name.length > SHEET_NAME_MAX_LEN) {
      ctx.reporter.error(
        "workbook-sheet-name-too-long",
        `Sheet name "${name}" has ${name.length} characters; Excel limit is ${SHEET_NAME_MAX_LEN}`,
        WORKBOOK_PATH
      );
    }
    if (INVALID_SHEET_NAME_CHARS.test(name)) {
      ctx.reporter.error(
        "workbook-sheet-name-invalid-chars",
        `Sheet name "${name}" contains characters disallowed by Excel (\\ / ? * [ ] :)`,
        WORKBOOK_PATH
      );
    }
    const lower = name.toLowerCase();
    if (seenNamesLower.has(lower)) {
      ctx.reporter.error(
        "workbook-sheet-name-duplicate",
        `Duplicate sheet name (case-insensitive): "${name}"`,
        WORKBOOK_PATH
      );
    } else {
      seenNamesLower.add(lower);
    }
  }
}

// =============================================================================
// Element order validation
// =============================================================================

import type { XmlElement } from "@xml/types";

/**
 * Verify that the direct children of `<workbook>` follow the CT_Workbook
 * XSD sequence. Excel is strict about this and rejects files with
 * out-of-order elements as unrecoverable corruption.
 */
function checkWorkbookChildOrder(ctx: ValidationContext, root: XmlElement): void {
  const orderIndex = new Map<string, number>();
  for (let i = 0; i < WORKBOOK_CHILD_ORDER.length; i++) {
    orderIndex.set(WORKBOOK_CHILD_ORDER[i], i);
  }

  let maxSeenIdx = -1;
  let maxSeenName = "";

  for (const child of root.children) {
    if (child.type !== "element") {
      continue;
    }
    const local = localName(child.name);
    const idx = orderIndex.get(local);
    if (idx === undefined) {
      // Unknown / extension element — skip (mc:AlternateContent etc.)
      continue;
    }
    if (idx < maxSeenIdx) {
      ctx.reporter.error(
        "workbook-child-order",
        `Workbook child element <${local}> appears after <${maxSeenName}>, ` +
          `violating CT_Workbook sequence order (ECMA-376 §18.2.27). ` +
          `Excel will reject this file as corrupt.`,
        WORKBOOK_PATH
      );
      // Report only the first violation to avoid noise
      return;
    }
    if (idx > maxSeenIdx) {
      maxSeenIdx = idx;
      maxSeenName = local;
    }
  }
}
