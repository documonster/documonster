/**
 * DOCX Module — Glossary (Building Blocks) Part Writer
 *
 * Serialises a {@link GlossaryDocument} into the canonical
 * `word/glossary/document.xml` OOXML form:
 *
 * ```xml
 * <w:glossaryDocument>
 *   <w:docParts>
 *     <w:docPart>
 *       <w:docPartPr>
 *         <w:name w:val="…"/>
 *         <w:category><w:name w:val="…"/><w:gallery w:val="…"/></w:category>
 *         <w:behaviors><w:behavior w:val="content"/></w:behaviors>
 *         <w:guid w:val="{…}"/>
 *       </w:docPartPr>
 *       <w:docPartBody>… body content …</w:docPartBody>
 *     </w:docPart>
 *   </w:docParts>
 * </w:glossaryDocument>
 * ```
 *
 * Lives in the writer layer (not `advanced/`) so the packager can depend on
 * it without creating a `advanced/ → writer/` import cycle.
 */

import { DOCUMENT_NAMESPACES, STD_DOC_ATTRIBUTES } from "@word/constants";
import { GALLERY_TO_OOXML } from "@word/core/glossary-galleries";
import { within } from "@word/core/node-path";
import type { BodyContent, GlossaryDocument } from "@word/types";
import { renderBodyContent } from "@word/writer/document-writer";
import type { DocxRenderContext } from "@word/writer/render-context";
import { createRenderContext } from "@word/writer/render-context";
import type { XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

/**
 * Whether `glossary` is written as the markup it was read from: it has that
 * markup, and its blocks are the ones read from it (or it was not read, and
 * a caller supplied `rawXml` without blocks to compare).
 */
export function isGlossaryAsRead(glossary: GlossaryDocument): boolean {
  if (!glossary.rawXml) {
    return false;
  }
  return (
    glossary._blocksAsRead === undefined ||
    JSON.stringify(glossary.blocks) === glossary._blocksAsRead
  );
}

/** Render a {@link GlossaryDocument} to a `word/glossary/document.xml` string. */
export function renderGlossaryDocument(glossary: GlossaryDocument): string {
  // Byte-faithful round-trip: a glossary read from an existing document is
  // carried as verbatim XML and re-emitted unchanged while its blocks are.
  if (isGlossaryAsRead(glossary)) {
    return glossary.rawXml!;
  }
  const writer = new XmlWriter();
  renderGlossary(writer, glossary);
  return writer.xml;
}

function renderGlossary(xml: XmlSink, glossary: GlossaryDocument): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:glossaryDocument", DOCUMENT_NAMESPACES);
  xml.openNode("w:docParts");
  // One context for the whole glossary document: it is a document of its own,
  // so its bookmark and drawing ids are unique within it rather than within
  // the main document.
  const ctx = createRenderContext({ partName: "/word/glossary/document.xml" });
  // A refused value is named from the document root, where the glossary is `glossary`.
  glossary.blocks.forEach((block, i) =>
    within(`glossary.blocks[${i}]`, () => renderDocPart(xml, block, ctx))
  );
  // Entries of a read glossary the model has no gallery for, as they were read.
  for (const raw of glossary._unlistedDocParts ?? []) {
    xml.writeRaw(raw);
  }
  xml.closeNode(); // w:docParts
  xml.closeNode(); // w:glossaryDocument
}

function renderDocPart(
  xml: XmlSink,
  block: GlossaryDocument["blocks"][number],
  ctx: DocxRenderContext
): void {
  xml.openNode("w:docPart");

  // CT_DocPartPr — child order is fixed by the schema (ECMA-376 §17.12.1):
  // name → style → category → types → behaviors → description → guid.
  // Emitting these out of order makes Word reject the package on open.
  xml.openNode("w:docPartPr");
  xml.leafNode("w:name", { "w:val": block.name });
  xml.openNode("w:category");
  xml.leafNode("w:name", { "w:val": block.category ?? "General" });
  xml.leafNode("w:gallery", { "w:val": GALLERY_TO_OOXML[block.gallery] ?? "placeholder" });
  xml.closeNode(); // w:category
  // `<w:types>` is optional; we omit it so we never emit an out-of-enum
  // ST_DocPartType value. A docPart placed in the body inserts its content.
  xml.openNode("w:behaviors");
  xml.leafNode("w:behavior", { "w:val": "content" });
  xml.closeNode(); // w:behaviors
  if (block.description) {
    xml.leafNode("w:description", { "w:val": block.description });
  }
  if (block.guid) {
    xml.leafNode("w:guid", { "w:val": normaliseGuid(block.guid) });
  }
  xml.closeNode(); // w:docPartPr

  // docPartBody — reuse the main body renderer with the glossary's context;
  // building-block content here is plain text/tables so it does not need
  // image/chart rId remapping.
  xml.openNode("w:docPartBody");
  block.content.forEach((item, i) =>
    within(`content[${i}]`, () => renderBodyContent(xml, item as BodyContent, ctx))
  );
  // CT_Body must end with a paragraph; emit one if the block had no content.
  if (block.content.length === 0) {
    xml.openNode("w:p");
    xml.closeNode();
  }
  xml.closeNode(); // w:docPartBody

  xml.closeNode(); // w:docPart
}

/** Word expects the docPart guid wrapped in braces: `{XXXXXXXX-…}`. */
function normaliseGuid(guid: string): string {
  const trimmed = guid.trim();
  if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
    return trimmed;
  }
  return `{${trimmed}}`;
}
