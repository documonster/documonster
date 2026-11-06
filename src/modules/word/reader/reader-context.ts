/**
 * DOCX Reader Context
 *
 * Encapsulates per-parse session state so that concurrent readDocx() calls
 * are safe. Previously this state lived in a module-level `_session` object.
 */

import { STRICT_TO_TRANSITIONAL_REL } from "@word/constants";
import type { OpenFieldPiece } from "@word/reader/field-pieces";
import { parsePartXml } from "@word/reader/part-xml";
import type { DocxSecurityPolicy } from "@word/security/policy";
import { DEFAULT_SECURITY_POLICY } from "@word/security/policy";
import type { FormField, RangeMarker, RunProperties } from "@word/types";
import type { XmlElement } from "@xml/types";

/**
 * A parsed OPC relationship entry (from a .rels part).
 */
export interface ParsedRelationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * State for the OOXML field state machine (`<w:fldChar>`).
 *
 * OOXML allows complex fields (TOC, INDEX, SEQ, REF, …) to span multiple
 * paragraphs: `<w:fldChar fldCharType="begin">` may be in paragraph A while
 * the matching `end` is in paragraph C. This state therefore lives on the
 * shared reader context rather than as locals in `parseParagraph`.
 *
 * It is intentionally swappable: when entering a self-contained part such as
 * a header, footer, footnote, endnote or comment we save and reset the field
 * state so an unterminated field in the body never bleeds into those parts.
 */
export interface FieldParseState {
  state: "none" | "instrText" | "cached";
  instr: string;
  cached: string;
  runProps: RunProperties | undefined;
  formField: FormField | undefined;
  /** Range markers met inside the open field, in source order. */
  markers: PendingFieldMarker[];
  /**
   * The fields read in pieces that a paragraph opened and has not yet ended
   * (see `splitFieldPieces`), innermost last.
   */
  open: OpenFieldPiece[];
  /** `w:fldChar` begin characters whose fields are read in pieces whatever their shape. */
  forced?: ReadonlySet<XmlElement>;
  /**
   * Whether the paragraph being read may hold marker runs for field pieces
   * (`splitFieldPieces` changed it); when not, no run is looked up.
   */
  marked: boolean;
  /**
   * Set when the part being read holds no `fldChar` and no `fldSimple` at
   * all, so no field can be read in pieces and paragraphs skip
   * `splitFieldPieces`.
   */
  noFieldChars?: boolean;
}

/** A range marker inside an open complex field, at an offset into its raw instruction or its result. */
export interface PendingFieldMarker {
  readonly marker: RangeMarker;
  readonly part: "instruction" | "result";
  readonly offset: number;
  /** Set once the field ends and holds it. */
  embedded: boolean;
}

/**
 * Per-parse session context. Holds relationship maps and any other state
 * that needs to be shared across parse functions during a single readDocx()
 * invocation without relying on module-level mutable state.
 */
export interface ReaderContext {
  /** Relationship map for the current part being parsed (document, header, footer, etc.) */
  relMap: Map<string, ParsedRelationship>;
  /** Field state machine; shared across paragraphs to support cross-paragraph fields. */
  field: FieldParseState;
  /**
   * Resolved security policy for this read. Parsers consult specific
   * fields (e.g. `allowExternalTargets`) to decide whether to keep certain
   * pieces of information that may be dangerous in untrusted documents.
   * Defaults to `DEFAULT_SECURITY_POLICY` when not set.
   */
  securityPolicy: Required<DocxSecurityPolicy>;
}

/**
 * Create a fresh field state in the "none" state.
 */
export function createFieldState(): FieldParseState {
  return {
    state: "none",
    instr: "",
    cached: "",
    runProps: undefined,
    formField: undefined,
    markers: [],
    open: [],
    marked: true
  };
}

/**
 * Create a fresh reader context for a new parse session.
 */
export function createReaderContext(securityPolicy?: Required<DocxSecurityPolicy>): ReaderContext {
  return {
    relMap: new Map(),
    field: createFieldState(),
    securityPolicy: securityPolicy ?? DEFAULT_SECURITY_POLICY
  };
}

// =============================================================================
// Relationships Parser
// =============================================================================

/**
 * Parse a `.rels` XML string into an array of `ParsedRelationship` entries.
 *
 * Normalizes ISO 29500 Strict relationship types to Transitional equivalents
 * via `STRICT_TO_TRANSITIONAL_REL`.
 */
export function parseRelationships(xmlStr: string, maxDepth?: number): ParsedRelationship[] {
  const doc = parsePartXml(xmlStr, maxDepth);
  const rels: ParsedRelationship[] = [];

  for (const child of doc.root.children) {
    if (child.type === "element" && child.name === "Relationship") {
      let relType = child.attributes["Type"] ?? "";
      // Normalize ISO 29500 Strict relationship types to Transitional
      const transitional = STRICT_TO_TRANSITIONAL_REL.get(relType);
      if (transitional) {
        relType = transitional;
      }
      rels.push({
        id: child.attributes["Id"] ?? "",
        type: relType,
        target: child.attributes["Target"] ?? "",
        targetMode: child.attributes["TargetMode"]
      });
    }
  }

  return rels;
}
