/**
 * DOCX Module - VBA Project (docm) Round-trip Support
 *
 * Provides support for preserving VBA macros during document read/write.
 * When a .docm file is read, the vbaProject.bin and related parts are
 * preserved in opaque parts. This module provides utilities to:
 *
 * - Detect VBA project presence
 * - Extract VBA metadata
 * - Ensure VBA parts are properly preserved in the output
 * - Convert between docx/docm content types
 *
 * Note: This module does NOT execute or compile VBA code. It only preserves
 * the binary VBA project for round-trip fidelity.
 */

import type { DocxDocument, OpaquePart } from "../types";

// =============================================================================
// Types
// =============================================================================

/** VBA project metadata extracted from the document. */
export interface VbaProjectInfo {
  /** Whether a VBA project is present. */
  readonly hasVba: boolean;
  /** Path of the vbaProject.bin in the archive. */
  readonly projectPath?: string;
  /** Content type of the VBA project part. */
  readonly contentType?: string;
  /** Module names detected (best-effort extraction). */
  readonly moduleNames?: readonly string[];
  /** Size in bytes of the VBA project binary. */
  readonly sizeBytes?: number;
}

/** Content types for macro-enabled documents. */
export const DOCM_CONTENT_TYPES = {
  /** Content type for .docm main document part. */
  document: "application/vnd.ms-word.document.macroEnabled.main+xml",
  /** Content type for vbaProject.bin. */
  vbaProject: "application/vnd.ms-office.vbaProject",
  /** Content type for vbaData.xml. */
  vbaData: "application/vnd.ms-word.vbaData+xml"
} as const;

/** VBA-related relationship type. */
export const VBA_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";

// =============================================================================
// VBA Detection & Extraction
// =============================================================================

/**
 * Check if a document contains a VBA project (is macro-enabled).
 *
 * @param doc - The parsed document.
 * @returns true if VBA macros are present.
 */
export function hasVbaProject(doc: DocxDocument): boolean {
  if (!doc.opaqueParts) {
    return false;
  }
  return doc.opaqueParts.some(p => isVbaPart(p.path));
}

/**
 * Extract VBA project metadata from a document.
 *
 * @param doc - The parsed document.
 * @returns VBA project info, or info with hasVba=false if not present.
 */
export function getVbaProjectInfo(doc: DocxDocument): VbaProjectInfo {
  if (!doc.opaqueParts) {
    return { hasVba: false };
  }

  const vbaPart = doc.opaqueParts.find(p => isVbaProjectBin(p.path));
  if (!vbaPart) {
    return { hasVba: false };
  }

  const moduleNames = extractModuleNames(vbaPart.data);

  return {
    hasVba: true,
    projectPath: vbaPart.path,
    contentType: vbaPart.contentType ?? DOCM_CONTENT_TYPES.vbaProject,
    moduleNames,
    sizeBytes: vbaPart.data.length
  };
}

/**
 * Get the raw VBA project binary data.
 *
 * @param doc - The parsed document.
 * @returns The vbaProject.bin bytes, or undefined if not present.
 */
export function getVbaProjectData(doc: DocxDocument): Uint8Array | undefined {
  if (!doc.opaqueParts) {
    return undefined;
  }
  const vbaPart = doc.opaqueParts.find(p => isVbaProjectBin(p.path));
  return vbaPart?.data;
}

/**
 * Add a VBA project to a document (making it a .docm).
 * The VBA project binary should be a valid OLE compound document
 * containing the VBA modules.
 *
 * @param doc - The document to add VBA to.
 * @param vbaProjectBin - The vbaProject.bin binary data.
 * @returns A new document with the VBA project embedded.
 */
export function addVbaProject(doc: DocxDocument, vbaProjectBin: Uint8Array): DocxDocument {
  const existingParts = doc.opaqueParts ? [...doc.opaqueParts] : [];

  // Remove any existing VBA parts
  const filteredParts = existingParts.filter(p => !isVbaPart(p.path));

  // Add the VBA project part
  const vbaPart: OpaquePart = {
    path: "word/vbaProject.bin",
    data: vbaProjectBin,
    contentType: DOCM_CONTENT_TYPES.vbaProject
  };

  filteredParts.push(vbaPart);

  return {
    ...doc,
    opaqueParts: filteredParts
  };
}

/**
 * Remove VBA project from a document (converting .docm to .docx behavior).
 *
 * @param doc - The macro-enabled document.
 * @returns A new document without VBA parts.
 */
export function removeVbaProject(doc: DocxDocument): DocxDocument {
  if (!doc.opaqueParts) {
    return doc;
  }

  const filteredParts = doc.opaqueParts.filter(p => !isVbaPart(p.path));

  return {
    ...doc,
    opaqueParts: filteredParts.length > 0 ? filteredParts : undefined
  };
}

/**
 * List all VBA-related parts in the document.
 */
export function listVbaParts(doc: DocxDocument): readonly OpaquePart[] {
  if (!doc.opaqueParts) {
    return [];
  }
  return doc.opaqueParts.filter(p => isVbaPart(p.path));
}

// =============================================================================
// Internal Helpers
// =============================================================================

function isVbaPart(path: string): boolean {
  const lower = path.toLowerCase();
  return (
    lower === "word/vbaproject.bin" ||
    lower === "word/vbadata.xml" ||
    lower.endsWith("/vbaproject.bin") ||
    lower.endsWith("/vbadata.xml") ||
    lower.includes("vbaproject")
  );
}

function isVbaProjectBin(path: string): boolean {
  const lower = path.toLowerCase();
  return lower === "word/vbaproject.bin" || lower.endsWith("vbaproject.bin");
}

/**
 * Best-effort extraction of VBA module names from vbaProject.bin.
 * The binary is an OLE2 compound document; we scan for module name strings.
 */
function extractModuleNames(data: Uint8Array): string[] {
  const names: string[] = [];

  // Module names in VBA projects are typically stored as UTF-16LE strings
  // preceded by specific markers. We do a simple heuristic scan.
  const str = tryDecodeStrings(data);

  // Look for common VBA module patterns
  const modulePatterns = [/Module\d*\s*=\s*([^\r\n]+)/g, /Attribute VB_Name\s*=\s*"([^"]+)"/g];

  for (const pattern of modulePatterns) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(str)) !== null) {
      const name = match[1]!.trim();
      if (name && !names.includes(name)) {
        names.push(name);
      }
    }
  }

  // Also look for standard module names in the dir stream
  const stdNames = ["ThisDocument", "Module1", "Module2", "Module3", "Sheet1"];
  for (const stdName of stdNames) {
    if (str.includes(stdName) && !names.includes(stdName)) {
      names.push(stdName);
    }
  }

  return names;
}

function tryDecodeStrings(data: Uint8Array): string {
  // Extract printable ASCII sequences (for heuristic module name detection)
  let result = "";
  let current = "";
  for (let i = 0; i < data.length; i++) {
    const b = data[i]!;
    if (b >= 0x20 && b < 0x7f) {
      current += String.fromCharCode(b);
    } else {
      if (current.length >= 3) {
        result += current + "\n";
      }
      current = "";
    }
  }
  if (current.length >= 3) {
    result += current;
  }
  return result;
}
