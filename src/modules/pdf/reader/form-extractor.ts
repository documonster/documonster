/**
 * PDF form field (AcroForm) extractor.
 *
 * Extracts interactive form fields from a PDF's `/AcroForm` dictionary.
 * Supports all standard field types:
 * - **Text** (`/Tx`) — Text input fields
 * - **Button** (`/Btn`) — Checkboxes, radio buttons, push buttons
 * - **Choice** (`/Ch`) — Dropdowns (combo boxes) and list boxes
 * - **Signature** (`/Sig`) — Digital signature fields
 *
 * Handles field hierarchies (parent/child), inherited values, and default appearances.
 *
 * @see PDF Reference 1.7, §12.7 - Interactive Forms
 */

import type { PdfDocument } from "./pdf-document";
import type { PdfDictValue, PdfObject } from "./pdf-parser";
import {
  isPdfArray,
  isPdfRef,
  dictGetName,
  dictGetNumber,
  decodePdfStringBytes
} from "./pdf-parser";
import { getDictStringValue } from "./reader-utils";

// =============================================================================
// Types
// =============================================================================

/** Type of form field. */
export type PdfFormFieldType =
  | "text"
  | "checkbox"
  | "radio"
  | "dropdown"
  | "listbox"
  | "button"
  | "signature"
  | "unknown";

/** A single form field extracted from the PDF. */
export interface PdfFormField {
  /** Fully qualified field name (e.g. "form1.address.city") */
  name: string;
  /** Field type */
  type: PdfFormFieldType;
  /** Current value of the field */
  value: string;
  /** Default value (/DV entry) */
  defaultValue: string;
  /** Whether the field is read-only */
  readOnly: boolean;
  /** Whether the field is required */
  required: boolean;
  /** For choice fields: the list of available options */
  options: string[];
  /** For checkboxes/radio buttons: the export value when checked */
  exportValue: string;
  /** Field flags (/Ff entry) — raw bit field */
  flags: number;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Extract form fields from a PDF document.
 *
 * Reads the `/AcroForm` dictionary from the catalog and recursively
 * traverses the field tree.
 *
 * @param doc - The PDF document
 * @returns Array of extracted form fields
 */
export function extractFormFields(doc: PdfDocument): PdfFormField[] {
  try {
    const catalog = doc.getCatalog();
    const acroFormObj = catalog.get("AcroForm");
    if (!acroFormObj) {
      return [];
    }

    const acroForm = doc.derefDict(acroFormObj);
    if (!acroForm) {
      return [];
    }

    const fieldsObj = acroForm.get("Fields");
    if (!fieldsObj) {
      return [];
    }

    const fieldsArr = doc.deref(fieldsObj);
    if (!isPdfArray(fieldsArr)) {
      return [];
    }

    const fields: PdfFormField[] = [];
    const visited = new Set<number>();

    for (const fieldRef of fieldsArr) {
      collectFields(fieldRef, doc, "", fields, visited);
    }

    return fields;
  } catch {
    return [];
  }
}

// =============================================================================
// Field Tree Traversal
// =============================================================================

/** Field flags from PDF spec §12.7.3 */
const FLAG_READ_ONLY = 1 << 0;
const FLAG_REQUIRED = 1 << 1;
// Button-specific
const FLAG_PUSHBUTTON = 1 << 16;
const FLAG_RADIO = 1 << 15;
// Choice-specific
const FLAG_COMBO = 1 << 17;

/** Maximum depth for parent traversal to prevent cycles in malformed PDFs */
const MAX_INHERIT_DEPTH = 20;

function collectFields(
  fieldObj: PdfObject,
  doc: PdfDocument,
  parentName: string,
  result: PdfFormField[],
  visited: Set<number>
): void {
  // Track visited objects to avoid cycles
  if (isPdfRef(fieldObj)) {
    if (visited.has(fieldObj.objNum)) {
      return;
    }
    visited.add(fieldObj.objNum);
  }

  const dict = doc.derefDict(fieldObj);
  if (!dict) {
    return;
  }

  // Build the fully qualified field name
  const partialName = getDictStringValue(dict, "T", doc);
  const fullName = parentName
    ? partialName
      ? `${parentName}.${partialName}`
      : parentName
    : partialName;

  // Check for children (/Kids)
  const kidsObj = dict.get("Kids");
  if (kidsObj) {
    const kids = doc.deref(kidsObj);
    if (isPdfArray(kids)) {
      // Check if kids are field nodes or widget nodes
      // If kids have /T entries, they are field nodes (continue recursion)
      // If kids don't have /T, they are widget annotations — treat parent as the field
      let hasFieldChildren = false;
      for (const kid of kids) {
        const kidDict = doc.derefDict(kid);
        if (kidDict && kidDict.has("T")) {
          hasFieldChildren = true;
          break;
        }
      }

      if (hasFieldChildren) {
        // Recurse into child fields
        for (const kid of kids) {
          collectFields(kid, doc, fullName, result, visited);
        }
        return;
      }

      // Kids are widgets — extract value from first kid or parent
      // For radio buttons, collect export values from kids
      const ft = resolveFieldType(dict, doc);
      if (ft === "Btn") {
        const ff = resolveFieldFlags(dict, doc);
        if ((ff & FLAG_RADIO) !== 0 && (ff & FLAG_PUSHBUTTON) === 0) {
          // Radio button: current value is on the parent, export values on kids
          const field = parseRadioField(dict, kids, fullName, ff, doc);
          if (field) {
            result.push(field);
          }
          return;
        }
      }
    }
  }

  // Leaf field — extract its properties
  const ft = resolveFieldType(dict, doc);
  if (!ft) {
    return; // Not a real field (no /FT)
  }

  const field = parseField(dict, fullName, ft, doc);
  if (field) {
    result.push(field);
  }
}

// =============================================================================
// Field Parsing
// =============================================================================

/**
 * Resolve /FT (field type) which may be inherited from parent.
 */
function resolveFieldType(dict: PdfDictValue, doc: PdfDocument, depth = 0): string | undefined {
  const ft = dictGetName(dict, "FT");
  if (ft) {
    return ft;
  }

  if (depth >= MAX_INHERIT_DEPTH) {
    return undefined;
  }

  // Check parent
  const parent = dict.get("Parent");
  if (parent) {
    const parentDict = doc.derefDict(parent);
    if (parentDict) {
      return resolveFieldType(parentDict, doc, depth + 1);
    }
  }

  return undefined;
}

/**
 * Resolve /Ff (field flags) which may be inherited from parent.
 */
function resolveFieldFlags(dict: PdfDictValue, doc: PdfDocument, depth = 0): number {
  const ff = dictGetNumber(dict, "Ff");
  if (ff !== undefined) {
    return ff;
  }

  if (depth >= MAX_INHERIT_DEPTH) {
    return 0;
  }

  const parent = dict.get("Parent");
  if (parent) {
    const parentDict = doc.derefDict(parent);
    if (parentDict) {
      return resolveFieldFlags(parentDict, doc, depth + 1);
    }
  }

  return 0;
}

function parseField(
  dict: PdfDictValue,
  name: string,
  ft: string,
  doc: PdfDocument
): PdfFormField | null {
  const ff = resolveFieldFlags(dict, doc);
  const value = getFieldValue(dict, doc);
  const defaultValue = getDictStringValue(dict, "DV", doc);

  const type = classifyFieldType(ft, ff);

  let options: string[] = [];
  let exportValue = "";

  if (ft === "Ch") {
    options = parseChoiceOptions(dict, doc);
  }

  if (ft === "Btn" && (ff & FLAG_PUSHBUTTON) === 0 && (ff & FLAG_RADIO) === 0) {
    // Checkbox — extract export value from /AP /N keys
    exportValue = parseCheckboxExportValue(dict, doc);
  }

  return {
    name: name || "(unnamed)",
    type,
    value,
    defaultValue,
    readOnly: (ff & FLAG_READ_ONLY) !== 0,
    required: (ff & FLAG_REQUIRED) !== 0,
    options,
    exportValue,
    flags: ff
  };
}

function parseRadioField(
  parentDict: PdfDictValue,
  kids: PdfObject[],
  name: string,
  ff: number,
  doc: PdfDocument
): PdfFormField | null {
  const value = getFieldValue(parentDict, doc);
  const defaultValue = getDictStringValue(parentDict, "DV", doc);

  // Collect export values from kid appearance dictionaries
  const options: string[] = [];
  for (const kid of kids) {
    const kidDict = doc.derefDict(kid);
    if (!kidDict) {
      continue;
    }

    const apDict = doc.derefDict(kidDict.get("AP"));
    if (!apDict) {
      continue;
    }

    const nDict = doc.derefDict(apDict.get("N"));
    if (!nDict) {
      continue;
    }

    // Keys of /AP /N are the possible states (e.g. "/Choice1", "/Off")
    for (const key of nDict.keys()) {
      if (key !== "Off" && !options.includes(key)) {
        options.push(key);
      }
    }
  }

  return {
    name: name || "(unnamed)",
    type: "radio",
    value,
    defaultValue,
    readOnly: (ff & FLAG_READ_ONLY) !== 0,
    required: (ff & FLAG_REQUIRED) !== 0,
    options,
    exportValue: "",
    flags: ff
  };
}

function classifyFieldType(ft: string, ff: number): PdfFormFieldType {
  switch (ft) {
    case "Tx":
      return "text";
    case "Btn":
      if ((ff & FLAG_PUSHBUTTON) !== 0) {
        return "button";
      }
      if ((ff & FLAG_RADIO) !== 0) {
        return "radio";
      }
      return "checkbox";
    case "Ch":
      if ((ff & FLAG_COMBO) !== 0) {
        return "dropdown";
      }
      return "listbox";
    case "Sig":
      return "signature";
    default:
      return "unknown";
  }
}

function parseChoiceOptions(dict: PdfDictValue, doc: PdfDocument): string[] {
  const optObj = dict.get("Opt");
  if (!optObj) {
    return [];
  }

  const optArr = doc.deref(optObj);
  if (!isPdfArray(optArr)) {
    return [];
  }

  const options: string[] = [];
  for (const item of optArr) {
    const resolved = doc.deref(item);
    if (typeof resolved === "string") {
      options.push(resolved);
    } else if (resolved instanceof Uint8Array) {
      options.push(decodePdfStringBytes(resolved));
    } else if (isPdfArray(resolved) && resolved.length >= 2) {
      // [exportValue, displayValue] pair
      const display = doc.deref(resolved[1]);
      if (typeof display === "string") {
        options.push(display);
      } else if (display instanceof Uint8Array) {
        options.push(decodePdfStringBytes(display));
      }
    }
  }

  return options;
}

function parseCheckboxExportValue(dict: PdfDictValue, doc: PdfDocument): string {
  // The export value is the key in /AP /N that isn't "Off"
  const apDict = doc.derefDict(dict.get("AP"));
  if (!apDict) {
    return "Yes"; // Default per spec
  }

  const nDict = doc.derefDict(apDict.get("N"));
  if (!nDict) {
    return "Yes";
  }

  for (const key of nDict.keys()) {
    if (key !== "Off") {
      return key;
    }
  }

  return "Yes";
}

// =============================================================================
// Value Extraction
// =============================================================================

/** Get the field value (/V entry), resolving from parent if needed. */
function getFieldValue(dict: PdfDictValue, doc: PdfDocument, depth = 0): string {
  const val = dict.get("V");
  if (val !== undefined) {
    return resolveValue(val, doc);
  }

  if (depth >= MAX_INHERIT_DEPTH) {
    return "";
  }

  // Inherit from parent
  const parent = dict.get("Parent");
  if (parent) {
    const parentDict = doc.derefDict(parent);
    if (parentDict) {
      return getFieldValue(parentDict, doc, depth + 1);
    }
  }

  return "";
}

function resolveValue(val: PdfObject, doc: PdfDocument): string {
  const resolved = doc.deref(val);
  if (typeof resolved === "string") {
    return resolved;
  }
  if (resolved instanceof Uint8Array) {
    return decodePdfStringBytes(resolved);
  }
  if (typeof resolved === "number") {
    return String(resolved);
  }
  if (typeof resolved === "boolean") {
    return resolved ? "true" : "false";
  }
  return "";
}
