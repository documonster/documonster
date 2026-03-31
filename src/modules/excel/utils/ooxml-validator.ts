import path from "node:path";

import type { ExtractedFile } from "@archive/unzip/extract";
import { extractAll } from "@archive/unzip/extract";
import { parseSax } from "@xml/sax";

export type OoxmlProblemKind =
  | "missing-part"
  | "xml-malformed"
  | "content-types-missing"
  | "content-types-malformed"
  | "content-types-missing-default"
  | "content-types-missing-for-part"
  | "content-types-duplicate-override"
  | "root-rels-missing-officeDocument"
  | "rels-malformed"
  | "rels-missing-target"
  | "rels-duplicate-id"
  | "rels-empty-target"
  | "rels-invalid-target-path"
  | "rels-source-missing"
  | "workbook-sheet-missing-rel"
  | "workbook-sheet-wrong-rel-type"
  | "workbook-duplicate-sheetId"
  | "workbook-duplicate-sheet-rid"
  | "sheet-missing-rels"
  | "sheet-controls-missing-drawing"
  | "sheet-control-missing-rel"
  | "sheet-control-wrong-rel-type"
  | "sheet-legacyDrawing-missing-rel"
  | "sheet-legacyDrawing-wrong-rel-type"
  | "sheet-drawing-missing-rel"
  | "sheet-drawing-wrong-rel-type"
  | "sheet-comments-missing-rel"
  | "sheet-comments-wrong-rel-type"
  | "sheet-hyperlink-missing-rel"
  | "sheet-hyperlink-wrong-rel-type"
  | "sheet-tablePart-missing-rel"
  | "sheet-tablePart-wrong-rel-type";

export type OoxmlOrderingProblemKind = "sheet-legacyDrawing-after-controls";

export interface OoxmlValidationProblem {
  kind: OoxmlProblemKind;
  file?: string;
  message: string;
}

export interface OoxmlOrderingValidationProblem {
  kind: OoxmlOrderingProblemKind;
  file?: string;
  message: string;
}

export interface OoxmlValidationReport {
  ok: boolean;
  problems: Array<OoxmlValidationProblem | OoxmlOrderingValidationProblem>;
  stats: {
    entryCount: number;
    xmlLikeCount: number;
    relsCount: number;
  };
}

export interface OoxmlValidateOptions {
  /**
   * Whether to check every XML-like entry (.xml, .rels, .vml) for well-formedness.
   * Default: true.
   */
  checkXmlWellFormed?: boolean;

  /**
   * Whether to validate relationship targets exist (TargetMode=External ignored).
   * Default: true.
   */
  checkRelationshipTargets?: boolean;

  /**
   * Whether to validate that every ContentTypes Override PartName exists.
   * Default: true.
   */
  checkContentTypesOverrides?: boolean;

  /**
   * Whether to validate worksheet <controls>/<legacyDrawing> r:id wiring.
   * Default: true.
   */
  checkWorksheetControlWiring?: boolean;

  /**
   * If provided, stops after this many problems.
   */
  maxProblems?: number;
}

function pushProblem(
  problems: Array<OoxmlValidationProblem | OoxmlOrderingValidationProblem>,
  problem: OoxmlValidationProblem | OoxmlOrderingValidationProblem,
  maxProblems?: number
): void {
  if (maxProblems !== undefined && problems.length >= maxProblems) {
    return;
  }
  problems.push(problem);
}

function stripLeadingSlash(p: string): string {
  return p.startsWith("/") ? p.slice(1) : p;
}

function isXmlLike(pathName: string): boolean {
  return pathName.endsWith(".xml") || pathName.endsWith(".rels") || pathName.endsWith(".vml");
}

function getRelsSourceDir(relsPath: string): string {
  // In OPC, relationship targets are resolved relative to the SOURCE part directory,
  // not the .rels part directory.
  //
  // Examples:
  // - _rels/.rels                           -> base "" (package root)
  // - xl/_rels/workbook.xml.rels            -> base "xl"
  // - xl/worksheets/_rels/sheet1.xml.rels   -> base "xl/worksheets"
  if (relsPath === "_rels/.rels") {
    return "";
  }
  const relsMarker = "/_rels/";
  const idx = relsPath.indexOf(relsMarker);
  if (idx === -1) {
    return relsPath.includes("/") ? relsPath.slice(0, relsPath.lastIndexOf("/")) : "";
  }
  return relsPath.slice(0, idx);
}

function resolveRelTarget(relsPath: string, target: string): string {
  const baseDir = getRelsSourceDir(relsPath);
  const resolved = path.posix.normalize(path.posix.join(baseDir, target));
  return resolved.replace(/^\//, "");
}

function isSafeResolvedPath(resolved: string): boolean {
  // Prevent path traversal out of package root.
  // `normalize()` can produce paths starting with "../".
  return !(resolved === ".." || resolved.startsWith("../") || resolved.includes("/../"));
}

async function assertXmlWellFormed(xmlText: string): Promise<void> {
  // parseSax throws on malformed XML.
  for await (const _events of parseSax([xmlText])) {
    // no-op
  }
}

function parseContentTypes(xml: string): {
  defaults: Array<{ extension: string; contentType: string }>;
  overrides: Array<{ partName: string; contentType: string }>;
  parseOk: boolean;
} {
  const defaults: Array<{ extension: string; contentType: string }> = [];
  const overrides: Array<{ partName: string; contentType: string }> = [];
  const defaultRe = /<Default\s+[^>]*Extension="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*\/>/g;
  const overrideRe = /<Override\s+[^>]*PartName="([^"]+)"[^>]*ContentType="([^"]+)"[^>]*\/>/g;

  for (let match = defaultRe.exec(xml); match; match = defaultRe.exec(xml)) {
    defaults.push({ extension: match[1], contentType: match[2] });
  }

  for (let match = overrideRe.exec(xml); match; match = overrideRe.exec(xml)) {
    overrides.push({ partName: match[1], contentType: match[2] });
  }

  const parseOk = xml.includes("<Types") && (defaults.length > 0 || overrides.length > 0);
  return { defaults, overrides, parseOk };
}

function getExtension(p: string): string {
  const base = path.posix.basename(p);
  const idx = base.lastIndexOf(".");
  return idx === -1 ? "" : base.slice(idx + 1);
}

function isPackagePart(pathName: string): boolean {
  // Exclude directories and the [Content_Types].xml pseudo-root.
  return pathName !== "" && pathName !== "[Content_Types].xml";
}

function parseRelationships(xml: string): {
  rels: Array<{ id: string; type: string; target: string; targetMode?: string }>;
  parseOk: boolean;
} {
  // Ensure the rels XML itself is parseable.
  // We still use regex to extract rels for speed/portability.
  const relRe =
    /<Relationship\s+[^>]*Id="([^"]+)"[^>]*Type="([^"]+)"[^>]*Target="([^"]*)"(?:[^>]*TargetMode="([^"]+)")?[^>]*\/>/g;

  const rels: Array<{ id: string; type: string; target: string; targetMode?: string }> = [];
  for (let match = relRe.exec(xml); match; match = relRe.exec(xml)) {
    rels.push({ id: match[1], type: match[2], target: match[3], targetMode: match[4] });
  }

  // Basic sanity: should have <Relationships ...> root.
  const parseOk = xml.includes("<Relationships") && xml.includes("Relationship");
  return { rels, parseOk };
}

export async function validateXlsxBuffer(
  xlsxBuffer: Uint8Array,
  options: OoxmlValidateOptions = {}
): Promise<OoxmlValidationReport> {
  const {
    checkXmlWellFormed = true,
    checkRelationshipTargets = true,
    checkContentTypesOverrides = true,
    checkWorksheetControlWiring = true,
    maxProblems
  } = options;

  const problems: Array<OoxmlValidationProblem | OoxmlOrderingValidationProblem> = [];

  const entries = await extractAll(xlsxBuffer);
  const has = (p: string) => entries.has(p);

  const mustExist = [
    "[Content_Types].xml",
    "_rels/.rels",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels"
  ];
  for (const p of mustExist) {
    if (!has(p)) {
      pushProblem(
        problems,
        { kind: "missing-part", file: p, message: `Missing required part: ${p}` },
        maxProblems
      );
    }
  }

  // XML well-formedness for all XML-like parts.
  if (checkXmlWellFormed) {
    for (const [p, entry] of entries) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      if (entry.type === "directory" || !isXmlLike(p)) {
        continue;
      }
      const xml = new TextDecoder().decode(entry.data);
      try {
        await assertXmlWellFormed(xml);
      } catch (err: any) {
        pushProblem(
          problems,
          {
            kind: "xml-malformed",
            file: p,
            message: `Malformed XML: ${err?.message || String(err)}`
          },
          maxProblems
        );
      }
    }
  }

  // Content types overrides must point to existing parts.
  if (checkContentTypesOverrides && has("[Content_Types].xml")) {
    const ctXml = new TextDecoder().decode(entries.get("[Content_Types].xml")!.data);
    const { defaults, overrides, parseOk } = parseContentTypes(ctXml);
    if (!parseOk) {
      pushProblem(
        problems,
        {
          kind: "content-types-malformed",
          file: "[Content_Types].xml",
          message: "Content types XML missing expected root/entries"
        },
        maxProblems
      );
    }

    const defaultByExt = new Map(defaults.map(d => [d.extension.toLowerCase(), d.contentType]));
    const overrideByPart = new Map<string, string>();
    for (const ov of overrides) {
      const key = stripLeadingSlash(ov.partName);
      if (overrideByPart.has(key)) {
        pushProblem(
          problems,
          {
            kind: "content-types-duplicate-override",
            file: "[Content_Types].xml",
            message: `Duplicate Override PartName: ${ov.partName}`
          },
          maxProblems
        );
      }
      overrideByPart.set(key, ov.contentType);
    }

    // RFC: .rels and .xml defaults are expected in valid packages.
    const relsDefault = defaultByExt.get("rels");
    if (relsDefault !== "application/vnd.openxmlformats-package.relationships+xml") {
      pushProblem(
        problems,
        {
          kind: "content-types-missing-default",
          file: "[Content_Types].xml",
          message:
            "Missing/incorrect Default for .rels (expected application/vnd.openxmlformats-package.relationships+xml)"
        },
        maxProblems
      );
    }
    const xmlDefault = defaultByExt.get("xml");
    if (xmlDefault !== "application/xml") {
      pushProblem(
        problems,
        {
          kind: "content-types-missing-default",
          file: "[Content_Types].xml",
          message: "Missing/incorrect Default for .xml (expected application/xml)"
        },
        maxProblems
      );
    }

    for (const ov of overrides) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      const zipPath = stripLeadingSlash(ov.partName);
      if (!has(zipPath)) {
        pushProblem(
          problems,
          {
            kind: "content-types-missing",
            file: "[Content_Types].xml",
            message: `Override PartName points to missing file: ${ov.partName}`
          },
          maxProblems
        );
      }
    }

    // Strong check: every part in the zip should have a content type via Default or Override.
    for (const [p, entry] of entries) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      if (entry.type === "directory" || !isPackagePart(p)) {
        continue;
      }

      const overrideType = overrideByPart.get(p);
      if (overrideType) {
        continue;
      }
      const ext = getExtension(p).toLowerCase();
      if (!ext) {
        pushProblem(
          problems,
          {
            kind: "content-types-missing-for-part",
            file: "[Content_Types].xml",
            message: `No content type for part without extension: ${p}`
          },
          maxProblems
        );
        continue;
      }
      if (!defaultByExt.has(ext)) {
        pushProblem(
          problems,
          {
            kind: "content-types-missing-for-part",
            file: "[Content_Types].xml",
            message: `No Default/Override content type for part: ${p} (extension .${ext})`
          },
          maxProblems
        );
      }
    }
  }

  // Root relationships must point to the workbook (OPC officeDocument).
  if (has("_rels/.rels")) {
    const rootRelsXml = new TextDecoder().decode(entries.get("_rels/.rels")!.data);
    const { rels } = parseRelationships(rootRelsXml);
    const hasOfficeDocument = rels.some(
      r => r.type.includes("/relationships/officeDocument") && r.target === "xl/workbook.xml"
    );
    if (!hasOfficeDocument) {
      pushProblem(
        problems,
        {
          kind: "root-rels-missing-officeDocument",
          file: "_rels/.rels",
          message: "Missing officeDocument relationship to xl/workbook.xml"
        },
        maxProblems
      );
    }
  }

  // Relationships: validate target existence and basic ID uniqueness.
  if (checkRelationshipTargets) {
    for (const [p, entry] of entries) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      if (entry.type === "directory" || !p.endsWith(".rels")) {
        continue;
      }

      const relsXml = new TextDecoder().decode(entry.data);
      const { rels, parseOk } = parseRelationships(relsXml);
      if (!parseOk) {
        pushProblem(
          problems,
          {
            kind: "rels-malformed",
            file: p,
            message: "Relationships XML missing expected root/entries"
          },
          maxProblems
        );
      }

      const ids = new Set<string>();
      for (const rel of rels) {
        if (maxProblems !== undefined && problems.length >= maxProblems) {
          break;
        }

        if (ids.has(rel.id)) {
          pushProblem(
            problems,
            {
              kind: "rels-duplicate-id",
              file: p,
              message: `Duplicate relationship Id: ${rel.id}`
            },
            maxProblems
          );
        }
        ids.add(rel.id);

        if (rel.targetMode === "External") {
          continue;
        }
        if (!rel.target) {
          pushProblem(
            problems,
            {
              kind: "rels-empty-target",
              file: p,
              message: `Relationship ${rel.id} (${rel.type}) has empty Target`
            },
            maxProblems
          );
          continue;
        }

        const resolvedTarget = resolveRelTarget(p, rel.target);
        if (!isSafeResolvedPath(resolvedTarget)) {
          pushProblem(
            problems,
            {
              kind: "rels-invalid-target-path",
              file: p,
              message: `Rel ${rel.id} (${rel.type}) target escapes package root: ${rel.target} -> ${resolvedTarget}`
            },
            maxProblems
          );
          continue;
        }
        if (!has(resolvedTarget)) {
          pushProblem(
            problems,
            {
              kind: "rels-missing-target",
              file: p,
              message: `Rel ${rel.id} (${rel.type}) target missing: ${rel.target} -> ${resolvedTarget}`
            },
            maxProblems
          );
        }
      }

      // Optional: ensure the source part exists for non-root rels.
      if (p !== "_rels/.rels") {
        // Convert: xl/_rels/workbook.xml.rels -> xl/workbook.xml
        // Convert: xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets/sheet1.xml
        const srcDir = getRelsSourceDir(p);
        const relsBaseName = path.posix.basename(p);
        const sourceName = relsBaseName.replace(/\.rels$/, "");
        const sourcePath = srcDir ? `${srcDir}/${sourceName}` : sourceName;
        if (!has(sourcePath)) {
          pushProblem(
            problems,
            {
              kind: "rels-source-missing",
              file: p,
              message: `Relationships part has no corresponding source part: ${sourcePath}`
            },
            maxProblems
          );
        }
      }
    }
  }

  // Workbook -> worksheets wiring.
  if (has("xl/workbook.xml") && has("xl/_rels/workbook.xml.rels")) {
    const workbookXml = new TextDecoder().decode(entries.get("xl/workbook.xml")!.data);
    const workbookRelsXml = new TextDecoder().decode(
      entries.get("xl/_rels/workbook.xml.rels")!.data
    );
    const { rels: wbRels } = parseRelationships(workbookRelsXml);
    const wbById = new Map(wbRels.map(r => [r.id, r]));

    // Uniqueness checks: sheetId and r:id should not be duplicated.
    const sheetIdRe = /<sheet\b[^>]*\bsheetId="(\d+)"[^>]*\/>/g;
    const seenSheetIds = new Set<string>();
    for (let match = sheetIdRe.exec(workbookXml); match; match = sheetIdRe.exec(workbookXml)) {
      const id = match[1];
      if (seenSheetIds.has(id)) {
        pushProblem(
          problems,
          {
            kind: "workbook-duplicate-sheetId",
            file: "xl/workbook.xml",
            message: `Duplicate sheetId in workbook: ${id}`
          },
          maxProblems
        );
      }
      seenSheetIds.add(id);
    }

    const sheetRidSeen = new Set<string>();

    const sheetRidRe = /<sheet\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
    for (let match = sheetRidRe.exec(workbookXml); match; match = sheetRidRe.exec(workbookXml)) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      const rid = match[1];
      if (sheetRidSeen.has(rid)) {
        pushProblem(
          problems,
          {
            kind: "workbook-duplicate-sheet-rid",
            file: "xl/workbook.xml",
            message: `Duplicate sheet r:id in workbook: ${rid}`
          },
          maxProblems
        );
      }
      sheetRidSeen.add(rid);
      const rel = wbById.get(rid);
      if (!rel) {
        pushProblem(
          problems,
          {
            kind: "workbook-sheet-missing-rel",
            file: "xl/workbook.xml",
            message: `Workbook <sheet> references missing relationship: ${rid} (in xl/_rels/workbook.xml.rels)`
          },
          maxProblems
        );
        continue;
      }
      if (!rel.type.includes("/relationships/worksheet")) {
        pushProblem(
          problems,
          {
            kind: "workbook-sheet-wrong-rel-type",
            file: "xl/workbook.xml",
            message: `Workbook <sheet> ${rid} relationship is not worksheet: ${rel.type}`
          },
          maxProblems
        );
      }
    }
  }

  // Worksheet <controls>/<legacyDrawing> wiring.
  if (checkWorksheetControlWiring) {
    for (const [p, entry] of entries) {
      if (maxProblems !== undefined && problems.length >= maxProblems) {
        break;
      }
      if (
        entry.type === "directory" ||
        !p.startsWith("xl/worksheets/sheet") ||
        !p.endsWith(".xml")
      ) {
        continue;
      }

      const sheetXml = new TextDecoder().decode(entry.data);
      const relsPath = `xl/worksheets/_rels/${path.posix.basename(p)}.rels`;

      // Excel is sensitive to worksheet child element ordering. In particular,
      // legacyDrawing must come before controls when both are present.
      const legacyDrawingIdx = sheetXml.indexOf("<legacyDrawing");
      const controlsIdx = sheetXml.indexOf("<controls");
      if (legacyDrawingIdx !== -1 && controlsIdx !== -1 && legacyDrawingIdx > controlsIdx) {
        pushProblem(
          problems,
          {
            kind: "sheet-legacyDrawing-after-controls",
            file: p,
            message:
              "Worksheet has <legacyDrawing> after <controls>; Excel may repair or reject this sheet"
          },
          maxProblems
        );
      }

      // Match <control ...> elements (not necessarily self-closing).
      const controlRidRe = /<control\b[^>]*\br:id="(rId\d+)"[^>]*>/g;
      const controlRids: string[] = [];
      for (let match = controlRidRe.exec(sheetXml); match; match = controlRidRe.exec(sheetXml)) {
        controlRids.push(match[1]);
      }

      // Excel Online / strict Excel builds may reject or "repair" legacy form controls
      // if the sheet doesn't also have a DrawingML <drawing> part.
      if (controlRids.length > 0 && sheetXml.indexOf("<drawing") === -1) {
        pushProblem(
          problems,
          {
            kind: "sheet-controls-missing-drawing",
            file: p,
            message:
              "Worksheet has legacy <controls> but no <drawing>; Excel may repair/reject legacy form controls"
          },
          maxProblems
        );
      }

      const legacyDrawingRidRe = /<legacyDrawing\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
      const legacyDrawingRids: string[] = [];
      for (
        let match = legacyDrawingRidRe.exec(sheetXml);
        match;
        match = legacyDrawingRidRe.exec(sheetXml)
      ) {
        legacyDrawingRids.push(match[1]);
      }

      if ((controlRids.length > 0 || legacyDrawingRids.length > 0) && !has(relsPath)) {
        pushProblem(
          problems,
          {
            kind: "sheet-missing-rels",
            file: p,
            message: `Worksheet has controls/legacyDrawing but missing rels part: ${relsPath}`
          },
          maxProblems
        );
        continue;
      }

      if (!has(relsPath)) {
        continue;
      }

      const sheetRelsXml = new TextDecoder().decode(entries.get(relsPath)!.data);
      const { rels: sheetRels } = parseRelationships(sheetRelsXml);
      const byId = new Map(sheetRels.map(r => [r.id, r]));

      const assertRidType = (
        rid: string,
        expectedTypeIncludes: string,
        kindMissing: OoxmlProblemKind,
        kindWrong: OoxmlProblemKind,
        nodeLabel: string
      ) => {
        const rel = byId.get(rid);
        if (!rel) {
          pushProblem(
            problems,
            {
              kind: kindMissing,
              file: p,
              message: `Sheet ${nodeLabel} references missing relationship: ${rid} (in ${relsPath})`
            },
            maxProblems
          );
          return;
        }
        if (!rel.type.includes(expectedTypeIncludes)) {
          pushProblem(
            problems,
            {
              kind: kindWrong,
              file: p,
              message: `Sheet ${nodeLabel} ${rid} relationship is not ${expectedTypeIncludes}: ${rel.type}`
            },
            maxProblems
          );
        }
      };

      for (const rid of controlRids) {
        if (maxProblems !== undefined && problems.length >= maxProblems) {
          break;
        }
        assertRidType(
          rid,
          "/relationships/ctrlProp",
          "sheet-control-missing-rel",
          "sheet-control-wrong-rel-type",
          "<control>"
        );
      }

      for (const rid of legacyDrawingRids) {
        if (maxProblems !== undefined && problems.length >= maxProblems) {
          break;
        }
        assertRidType(
          rid,
          "/relationships/vmlDrawing",
          "sheet-legacyDrawing-missing-rel",
          "sheet-legacyDrawing-wrong-rel-type",
          "<legacyDrawing>"
        );
      }

      // Common worksheet nodes that are wired via r:id
      // - <drawing r:id="..."/> -> drawing
      // - <comments r:id="..."/> -> comments
      // - <tableParts><tablePart r:id="..."/></tableParts> -> table
      // - <hyperlink r:id="..."/> -> hyperlink
      const drawingRidRe = /<drawing\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
      for (let match = drawingRidRe.exec(sheetXml); match; match = drawingRidRe.exec(sheetXml)) {
        assertRidType(
          match[1],
          "/relationships/drawing",
          "sheet-drawing-missing-rel",
          "sheet-drawing-wrong-rel-type",
          "<drawing>"
        );
      }

      const commentsRidRe = /<comments\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
      for (let match = commentsRidRe.exec(sheetXml); match; match = commentsRidRe.exec(sheetXml)) {
        assertRidType(
          match[1],
          "/relationships/comments",
          "sheet-comments-missing-rel",
          "sheet-comments-wrong-rel-type",
          "<comments>"
        );
      }

      const tablePartRidRe = /<tablePart\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
      for (
        let match = tablePartRidRe.exec(sheetXml);
        match;
        match = tablePartRidRe.exec(sheetXml)
      ) {
        assertRidType(
          match[1],
          "/relationships/table",
          "sheet-tablePart-missing-rel",
          "sheet-tablePart-wrong-rel-type",
          "<tablePart>"
        );
      }

      const hyperlinkRidRe = /<hyperlink\b[^>]*\br:id="(rId\d+)"[^>]*\/>/g;
      for (
        let match = hyperlinkRidRe.exec(sheetXml);
        match;
        match = hyperlinkRidRe.exec(sheetXml)
      ) {
        assertRidType(
          match[1],
          "/relationships/hyperlink",
          "sheet-hyperlink-missing-rel",
          "sheet-hyperlink-wrong-rel-type",
          "<hyperlink>"
        );
      }
    }
  }

  const stats = {
    entryCount: entries.size,
    xmlLikeCount: [...entries.values()].filter(
      (f: ExtractedFile) => f.type !== "directory" && isXmlLike(f.path)
    ).length,
    relsCount: [...entries.values()].filter(
      (f: ExtractedFile) => f.type !== "directory" && f.path.endsWith(".rels")
    ).length
  };

  return {
    ok: problems.length === 0,
    problems,
    stats
  };
}
