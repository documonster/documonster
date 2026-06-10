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

import type { XmlSink } from "@xml/types";
import { XmlWriter } from "@xml/writer";

import { DOCUMENT_NAMESPACES, STD_DOC_ATTRIBUTES } from "../constants";
import type { BodyContent, BuildingBlockGallery, GlossaryDocument } from "../types";
import { renderBodyContent } from "./document-writer";

/**
 * Map the friendly {@link BuildingBlockGallery} token to the OOXML
 * `ST_DocPartGallery` value used in `<w:gallery w:val="…">` (ECMA-376
 * §17.18.23). Values MUST be exact enum members — Word rejects (and silently
 * discards) the entire glossary if it sees an out-of-enum gallery value.
 * There is no plain "quickParts" value; Quick Parts map to "custQuickParts".
 */
const GALLERY_TO_OOXML: Record<BuildingBlockGallery, string> = {
  autoText: "autoTxt",
  quickParts: "custQuickParts",
  coverPages: "coverPg",
  tableOfContents: "tblOfContents",
  headers: "hdrs",
  footers: "ftrs",
  pageNumbers: "pgNum",
  tables: "tbls",
  textBoxes: "txtBox",
  watermarks: "watermarks",
  equations: "eq",
  bibliographies: "bib",
  custom1: "custom1",
  custom2: "custom2",
  custom3: "custom3",
  custom4: "custom4",
  custom5: "custom5"
};

/** Render a {@link GlossaryDocument} to a `word/glossary/document.xml` string. */
export function renderGlossaryDocument(glossary: GlossaryDocument): string {
  // Byte-faithful round-trip: a glossary read from an existing document is
  // carried as verbatim XML and re-emitted unchanged.
  if (glossary.rawXml) {
    return glossary.rawXml;
  }
  const writer = new XmlWriter();
  renderGlossary(writer, glossary);
  return writer.xml;
}

function renderGlossary(xml: XmlSink, glossary: GlossaryDocument): void {
  xml.openXml(STD_DOC_ATTRIBUTES);
  xml.openNode("w:glossaryDocument", DOCUMENT_NAMESPACES);
  xml.openNode("w:docParts");
  for (const block of glossary.blocks) {
    renderDocPart(xml, block);
  }
  xml.closeNode(); // w:docParts
  xml.closeNode(); // w:glossaryDocument
}

function renderDocPart(xml: XmlSink, block: GlossaryDocument["blocks"][number]): void {
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

  // docPartBody — reuse the main body renderer. A fresh render context keeps
  // id counters local; building-block content here is plain text/tables so it
  // does not need image/chart rId remapping.
  xml.openNode("w:docPartBody");
  for (const item of block.content) {
    renderBodyContent(xml, item as BodyContent);
  }
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
