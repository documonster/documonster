/**
 * Package structure check.
 *
 * Verifies the package contains the four OPC-mandatory parts (every xlsx
 * must have these to be considered an Open Packaging Convention file) and
 * that every XML-like part is syntactically well-formed.
 *
 * XML well-formedness is the broadest reachable check — it catches
 * unclosed tags, bad escapes, and invalid entities that would cause Excel
 * to repair-or-reject the whole package. We parse via DOM (which
 * internally delegates to SAX) so that subsequent DOM-based checkers hit
 * the same cached result.
 */

import type { ValidationContext } from "@excel/utils/ooxml-validator/context";
import { isXmlLike } from "@excel/utils/ooxml-validator/path-utils";

const REQUIRED_PARTS = [
  "[Content_Types].xml",
  "_rels/.rels",
  "xl/workbook.xml",
  "xl/_rels/workbook.xml.rels"
] as const;

export function checkStructure(ctx: ValidationContext): void {
  for (const p of REQUIRED_PARTS) {
    if (!ctx.has(p)) {
      ctx.reporter.error("missing-part", `Missing required part: ${p}`, p);
    }
  }
}

/**
 * Walk every XML-like entry and parse it to surface malformed-XML
 * problems. Leverages the context cache so DOM-based checkers that run
 * afterwards do not re-parse.
 */
export function checkXmlWellFormed(ctx: ValidationContext): void {
  for (const [path, entry] of ctx.files()) {
    if (ctx.reporter.capped) {
      return;
    }
    if (entry.type === "directory" || !isXmlLike(path)) {
      continue;
    }
    // Relationships malformed errors are reported by readRels; content
    // types malformed errors are reported by readContentTypes. Prefer
    // those specific kinds.
    if (path.endsWith(".rels")) {
      ctx.readRels(path);
      continue;
    }
    if (path === "[Content_Types].xml") {
      ctx.readContentTypes();
      continue;
    }
    ctx.readDom(path, err => {
      ctx.reporter.error("xml-malformed", `Malformed XML: ${err.message}`, path);
    });
  }
}
