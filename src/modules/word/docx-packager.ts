/**
 * DOCX Module - Packager
 *
 * Assembles a DocxDocument model into a DOCX ZIP package using the
 * archive and XML modules. Supports all parts including comments,
 * custom properties, document background, hyperlink relationships,
 * and per-header/footer relationship files.
 */

import { zip } from "@archive/create-archive";
import { XmlWriter } from "@xml/writer";

import { ContentType, RelType, PartPath, STD_DOC_ATTRIBUTES } from "./constants";
import {
  createContentTypes,
  addContentTypeDefault,
  addContentTypeOverride,
  addImageContentTypeDefaults,
  renderContentTypes
} from "./content-types";
import {
  createRelationships,
  addRelationship,
  addRelationshipWithId,
  getRelationshipCount,
  renderRelationships
} from "./relationships";
import type { RelationshipsState } from "./relationships";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Table,
  Hyperlink,
  HeaderFooterContent,
  Run,
  StructuredDocumentTag,
  ChartContent,
  AltChunk,
  ImageDef
} from "./types";
import { renderChartPart } from "./writers/chart-writer";
import { renderComments, renderCommentsExtended } from "./writers/comment-writer";
import { renderDocument, CHART_RID_REGISTRY } from "./writers/document-writer";
import { renderFootnotes, renderEndnotes } from "./writers/footnote-writer";
import { renderHeader, renderFooter, renderWatermarkHeader } from "./writers/header-footer-writer";
import { renderNumbering } from "./writers/numbering-writer";
import {
  renderSettings,
  renderFontTable,
  renderCoreProperties,
  renderAppProperties,
  renderCustomProperties,
  renderWebSettings,
  renderPeople,
  renderTheme
} from "./writers/parts-writer";
import { renderStyles } from "./writers/styles-writer";

/** Render XML to string using XmlWriter. */
function renderXml(renderFn: (xml: XmlWriter) => void): string {
  const writer = new XmlWriter();
  renderFn(writer);
  return writer.xml;
}

/** Get the .rels path for a given part path. */
function getPartRelsPath(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash) : "";
  const name = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

/** Scan a Run for image rIds. */
function scanRunForImages(run: Run, out: Set<string>): void {
  for (const rc of run.content) {
    if (rc.type === "image" && rc.rId) {
      out.add(rc.rId);
    }
  }
}

/** Scan ParagraphChild[] for hyperlinks with URL (no rId yet). */
function scanChildrenForHyperlinks(children: readonly ParagraphChild[], out: Hyperlink[]): void {
  for (const child of children) {
    if ("type" in child && child.type === "hyperlink") {
      const h = child as Hyperlink;
      if (h.url && !h.rId) {
        out.push(h);
      }
      // Also scan hyperlink children (runs) for images — handled elsewhere
    }
  }
}

/** Scan ParagraphChild[] for image rIds. */
function scanChildrenForImages(children: readonly ParagraphChild[], out: Set<string>): void {
  for (const child of children) {
    if ("type" in child) {
      if (child.type === "hyperlink") {
        const h = child as Hyperlink;
        for (const r of h.children) {
          scanRunForImages(r, out);
        }
      }
    } else {
      // It's a Run (no type field)
      scanRunForImages(child as Run, out);
    }
  }
}

/** Recursively collect hyperlinks with URL from body content. */
function collectHyperlinks(body: readonly BodyContent[]): Hyperlink[] {
  const links: Hyperlink[] = [];

  function fromParagraph(p: Paragraph): void {
    scanChildrenForHyperlinks(p.children, links);
  }

  function fromTable(t: Table): void {
    for (const row of t.rows) {
      for (const cell of row.cells) {
        for (const c of cell.content) {
          if (c.type === "paragraph") {
            fromParagraph(c);
          } else if (c.type === "table") {
            fromTable(c);
          }
        }
      }
    }
  }

  function fromSdt(sdt: StructuredDocumentTag): void {
    for (const c of sdt.content) {
      if ("type" in c && c.type === "paragraph") {
        fromParagraph(c as Paragraph);
      } else if ("type" in c && c.type === "table") {
        fromTable(c as Table);
      }
    }
  }

  for (const content of body) {
    if (content.type === "paragraph") {
      fromParagraph(content);
    } else if (content.type === "table") {
      fromTable(content);
    } else if (content.type === "sdt") {
      fromSdt(content as StructuredDocumentTag);
    } else if (content.type === "tableOfContents") {
      // Scan TOC cached paragraphs
      const toc = content as any;
      if (toc.cachedParagraphs) {
        for (const p of toc.cachedParagraphs) {
          fromParagraph(p);
        }
      }
    }
  }

  return links;
}

/** Infer a content type for an opaque part based on its file extension. */
function inferContentType(ext: string): string | undefined {
  const map: Record<string, string> = {
    xml: "application/xml",
    rels: "application/vnd.openxmlformats-package.relationships+xml",
    png: "image/png",
    jpeg: "image/jpeg",
    jpg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    tiff: "image/tiff",
    tif: "image/tiff",
    svg: "image/svg+xml",
    webp: "image/webp",
    emf: "image/x-emf",
    wmf: "image/x-wmf",
    odttf: "application/vnd.openxmlformats-officedocument.obfuscatedFont",
    ttf: "application/x-font-ttf",
    otf: "application/x-font-otf",
    bin: "application/vnd.openxmlformats-officedocument.oleObject",
    vml: "application/vnd.openxmlformats-officedocument.vmlDrawing"
  };
  return map[ext];
}

/**
 * Generic walker: visit all paragraphs within a block list, recursing into
 * tables (including nested tables) and SDTs. This is shared by image/hyperlink
 * collection helpers to ensure consistent coverage.
 */
function walkParagraphs(
  blocks: readonly (
    | { type: string; [k: string]: unknown }
    | Paragraph
    | Table
    | StructuredDocumentTag
  )[],
  onParagraph: (p: Paragraph) => void
): void {
  for (const block of blocks) {
    if (!("type" in block)) {
      continue;
    }
    if (block.type === "paragraph") {
      onParagraph(block as Paragraph);
    } else if (block.type === "table") {
      for (const row of (block as Table).rows) {
        for (const cell of row.cells) {
          walkParagraphs(cell.content as readonly (Paragraph | Table)[], onParagraph);
        }
      }
    } else if (block.type === "sdt") {
      const sdt = block as StructuredDocumentTag;
      // SDT content may include Paragraph | Run | Table — only the first and last here
      const filtered = sdt.content.filter(
        c => "type" in c && (c.type === "paragraph" || c.type === "table")
      );
      walkParagraphs(filtered as readonly (Paragraph | Table)[], onParagraph);
    } else if (block.type === "tableOfContents") {
      const toc = block as unknown as { cachedParagraphs?: readonly Paragraph[] };
      if (toc.cachedParagraphs) {
        walkParagraphs(toc.cachedParagraphs, onParagraph);
      }
    }
  }
}

/** Collect all image rIds referenced in header/footer content. */
function collectImageRidsFromContent(content: HeaderFooterContent): Set<string> {
  const rIds = new Set<string>();
  walkParagraphs(content.children, p => {
    scanChildrenForImages(p.children, rIds);
  });
  return rIds;
}

/** Collect hyperlinks from header/footer content. */
function collectHyperlinksFromHeaderFooter(content: HeaderFooterContent): Hyperlink[] {
  const links: Hyperlink[] = [];
  walkParagraphs(content.children, p => {
    scanChildrenForHyperlinks(p.children, links);
  });
  return links;
}

/**
 * Package a DocxDocument model into a DOCX ZIP file.
 * Returns the ZIP bytes as Uint8Array.
 */
export async function packageDocx(
  doc: DocxDocument,
  compressionLevel?: number
): Promise<Uint8Array> {
  const archive = zip({ level: compressionLevel ?? 6 });

  // Managers
  const contentTypes = createContentTypes();
  const packageRels = createRelationships();
  const documentRels = createRelationships();

  // --- Package relationships ---
  addRelationship(packageRels, RelType.OfficeDocument, "word/document.xml");
  addRelationship(packageRels, RelType.CoreProperties, "docProps/core.xml");
  addRelationship(packageRels, RelType.ExtendedProperties, "docProps/app.xml");

  // Custom properties
  const hasCustomProps = doc.customProperties && doc.customProperties.length > 0;
  if (hasCustomProps) {
    addRelationship(packageRels, RelType.CustomProperties, "docProps/custom.xml");
  }

  // --- Document relationships ---
  addRelationship(documentRels, RelType.Styles, "styles.xml");
  addRelationship(documentRels, RelType.Settings, "settings.xml");
  addRelationship(documentRels, RelType.FontTable, "fontTable.xml");
  addRelationship(documentRels, RelType.Theme, "theme/theme1.xml");

  // Numbering
  const hasNumbering =
    (doc.abstractNumberings && doc.abstractNumberings.length > 0) ||
    (doc.numberingInstances && doc.numberingInstances.length > 0);
  if (hasNumbering) {
    addRelationship(documentRels, RelType.Numbering, "numbering.xml");
  }

  // Footnotes & Endnotes
  const hasFootnotes = doc.footnotes && doc.footnotes.length > 0;
  const hasEndnotes = doc.endnotes && doc.endnotes.length > 0;
  if (hasFootnotes) {
    addRelationship(documentRels, RelType.Footnotes, "footnotes.xml");
  }
  if (hasEndnotes) {
    addRelationship(documentRels, RelType.Endnotes, "endnotes.xml");
  }

  // Comments
  const hasComments = doc.comments && doc.comments.length > 0;
  if (hasComments) {
    addRelationship(documentRels, RelType.Comments, "comments.xml");
  }

  // Images
  const imageExtensions = new Set<string>();
  const imageRidMap = new Map<string, string>(); // old rId -> new rId
  const imageByRid = new Map<string, ImageDef>(); // O(1) lookup by rId (original or new)
  const svgRidMap = new Map<string, string>(); // SVG fileName -> svgRId
  const svgFallbacks: { fallbackFileName: string; data: Uint8Array }[] = [];
  if (doc.images) {
    // Pre-build lookup index
    for (const img of doc.images) {
      if (img.rId) {
        imageByRid.set(img.rId, img);
      }
    }
    for (const img of doc.images) {
      const oldRid = img.rId;

      if (img.mediaType === "svg" && img.fallbackData) {
        // SVG with fallback: register PNG fallback as main image, SVG as separate
        const baseName = img.fileName.replace(/\.[^.]+$/, "");
        const fallbackFileName = `${baseName}_fallback.png`;

        // Register PNG fallback as the main rId
        const fbRId = addRelationship(documentRels, RelType.Image, `media/${fallbackFileName}`);
        (img as { rId?: string }).rId = fbRId;
        if (oldRid) {
          imageRidMap.set(oldRid, fbRId);
        }
        imageExtensions.add("png");

        // Register SVG as separate image
        const svgRId = addRelationship(documentRels, RelType.Image, `media/${img.fileName}`);
        svgRidMap.set(img.fileName, svgRId);
        const ext = img.fileName.split(".").pop()?.toLowerCase();
        if (ext) {
          imageExtensions.add(ext);
        }

        svgFallbacks.push({ fallbackFileName, data: img.fallbackData });
      } else {
        const rId = addRelationship(documentRels, RelType.Image, `media/${img.fileName}`);
        (img as { rId?: string }).rId = rId;
        if (oldRid) {
          imageRidMap.set(oldRid, rId);
        }
        const ext = img.fileName.split(".").pop()?.toLowerCase();
        if (ext) {
          imageExtensions.add(ext);
        }
      }
    }
  }

  // Update FloatingImage rIds to match new image rIds
  if (imageRidMap.size > 0) {
    for (const content of doc.body) {
      if (content.type === "floatingImage") {
        const fi = content as any;
        if (fi.rId && imageRidMap.has(fi.rId)) {
          // Already mapped — the image's rId was updated in-place above
          // but floatingImage still has the old rId; update it
        }
        // FloatingImage references images by rId; find the matching new rId
        if (fi.rId) {
          const img = imageByRid.get(fi.rId);
          if (!img) {
            // Old rId — look up in map
            const newRid = imageRidMap.get(fi.rId);
            if (newRid) {
              fi.rId = newRid;
            }
          }
        }
      }
    }
  }

  // Assign svgRId to inline and floating images that reference SVG files
  if (svgRidMap.size > 0 && doc.images) {
    const assignSvgRId = (rId: string, target: any): void => {
      const img = doc.images!.find(i => i.rId === rId || imageRidMap.get(rId) === i.rId);
      if (img && img.mediaType === "svg" && img.fallbackData) {
        const svgRId = svgRidMap.get(img.fileName);
        if (svgRId) {
          target.svgRId = svgRId;
        }
      }
    };
    for (const content of doc.body) {
      if (content.type === "floatingImage") {
        assignSvgRId((content as any).rId, content);
      } else if (content.type === "paragraph") {
        for (const child of (content as any).children ?? []) {
          if ("content" in child && Array.isArray(child.content)) {
            for (const c of child.content) {
              if (c.type === "image" && c.rId) {
                assignSvgRId(c.rId, c);
              }
            }
          }
        }
      }
    }
  }

  // Hyperlinks (external)
  const hyperlinks = collectHyperlinks(doc.body);
  for (const link of hyperlinks) {
    if (link.url) {
      const rId = addRelationship(documentRels, RelType.Hyperlink, link.url, "External");
      (link as { rId?: string }).rId = rId;
    }
  }

  // Process altChunk body items: register relationship and prepare data part
  const altChunks: AltChunk[] = [];
  collectAltChunks(doc.body, altChunks);
  altChunks.forEach((chunk, i) => {
    const fileName = chunk.fileName ?? `afchunk${i + 1}.html`;
    // Register relationship for the alt chunk
    const rId = addRelationship(
      documentRels,
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk",
      fileName
    );
    (chunk as { rId?: string }).rId = rId;
    // Register content type
    const ct = chunk.contentType ?? "text/html";
    addContentTypeOverride(contentTypes, `/word/${fileName}`, ct);
    // Write data to archive
    if (chunk.data) {
      archive.add(`word/${fileName}`, chunk.data);
    }
  });

  // Headers
  let headerIndex = 1;
  const headerRelManagers = new Map<string, RelationshipsState>();
  if (doc.headers) {
    for (const [key, headerDef] of doc.headers) {
      const rId = addRelationship(documentRels, RelType.Header, `header${headerIndex}.xml`);
      (headerDef as { rId?: string }).rId = rId;

      // Create per-header relationship state for images and hyperlinks
      const hRels = createRelationships();
      const imgRids = collectImageRidsFromContent(headerDef.content);
      if (imgRids.size > 0 && doc.images) {
        for (const oldRid of imgRids) {
          // O(1) lookup: direct or via remap
          const img = imageByRid.get(oldRid) ?? imageByRid.get(imageRidMap.get(oldRid) ?? "");
          if (img) {
            addRelationshipWithId(hRels, oldRid, RelType.Image, `media/${img.fileName}`);
          }
        }
      }
      // Hyperlinks in header
      const hLinks = collectHyperlinksFromHeaderFooter(headerDef.content);
      for (const link of hLinks) {
        if (link.url) {
          const linkRId = addRelationship(hRels, RelType.Hyperlink, link.url, "External");
          (link as { rId?: string }).rId = linkRId;
        }
      }
      if (getRelationshipCount(hRels) > 0) {
        headerRelManagers.set(key, hRels);
      }

      headerIndex++;
    }
  }

  // Footers
  let footerIndex = 1;
  const footerRelManagers = new Map<string, RelationshipsState>();
  if (doc.footers) {
    for (const [key, footerDef] of doc.footers) {
      const rId = addRelationship(documentRels, RelType.Footer, `footer${footerIndex}.xml`);
      (footerDef as { rId?: string }).rId = rId;

      // Create per-footer relationship state for images and hyperlinks
      const fRels = createRelationships();
      const imgRids = collectImageRidsFromContent(footerDef.content);
      if (imgRids.size > 0 && doc.images) {
        for (const oldRid of imgRids) {
          const img = imageByRid.get(oldRid) ?? imageByRid.get(imageRidMap.get(oldRid) ?? "");
          if (img) {
            addRelationshipWithId(fRels, oldRid, RelType.Image, `media/${img.fileName}`);
          }
        }
      }
      // Hyperlinks in footer
      const fLinks = collectHyperlinksFromHeaderFooter(footerDef.content);
      for (const link of fLinks) {
        if (link.url) {
          const linkRId = addRelationship(fRels, RelType.Hyperlink, link.url, "External");
          (link as { rId?: string }).rId = linkRId;
        }
      }
      if (getRelationshipCount(fRels) > 0) {
        footerRelManagers.set(key, fRels);
      }

      footerIndex++;
    }
  }

  // Watermark header (auto-generated if doc.watermark is set)
  let watermarkHeaderIndex: number | undefined;
  let watermarkHeaderRels: RelationshipsState | undefined;
  let watermarkHeaderRId: string | undefined;
  if (doc.watermark) {
    // Use next header index
    const wmHdrIdx = headerIndex;
    watermarkHeaderRId = addRelationship(documentRels, RelType.Header, `header${wmHdrIdx}.xml`);
    watermarkHeaderIndex = wmHdrIdx;

    // If image watermark, add image relationship in header .rels
    if (doc.watermark.type === "image") {
      const wmRels = createRelationships();
      const wmRId = doc.watermark.rId;
      const img = imageByRid.get(wmRId) ?? imageByRid.get(imageRidMap.get(wmRId) ?? "");
      if (img) {
        addRelationshipWithId(wmRels, wmRId, RelType.Image, `media/${img.fileName}`);
      }
      watermarkHeaderRels = wmRels;
    }
    headerIndex++;
  }

  // --- Content Types ---
  addImageContentTypeDefaults(contentTypes, imageExtensions);
  addContentTypeOverride(contentTypes, PartPath.Document, ContentType.Document);
  addContentTypeOverride(contentTypes, PartPath.Styles, ContentType.Styles);
  addContentTypeOverride(contentTypes, PartPath.Settings, ContentType.Settings);
  addContentTypeOverride(contentTypes, PartPath.FontTable, ContentType.FontTable);
  addContentTypeOverride(contentTypes, PartPath.Theme, ContentType.Theme);

  if (hasNumbering) {
    addContentTypeOverride(contentTypes, PartPath.Numbering, ContentType.Numbering);
  }
  if (hasFootnotes) {
    addContentTypeOverride(contentTypes, PartPath.Footnotes, ContentType.Footnotes);
  }
  if (hasEndnotes) {
    addContentTypeOverride(contentTypes, PartPath.Endnotes, ContentType.Endnotes);
  }
  if (hasComments) {
    addContentTypeOverride(contentTypes, PartPath.Comments, ContentType.Comments);
  }

  headerIndex = 1;
  if (doc.headers) {
    for (const [,] of doc.headers) {
      addContentTypeOverride(contentTypes, PartPath.header(headerIndex), ContentType.Header);
      headerIndex++;
    }
  }

  footerIndex = 1;
  if (doc.footers) {
    for (const [,] of doc.footers) {
      addContentTypeOverride(contentTypes, PartPath.footer(footerIndex), ContentType.Footer);
      footerIndex++;
    }
  }

  if (watermarkHeaderIndex !== undefined) {
    addContentTypeOverride(contentTypes, PartPath.header(watermarkHeaderIndex), ContentType.Header);
  }

  addContentTypeOverride(contentTypes, PartPath.CoreProps, ContentType.CoreProperties);
  addContentTypeOverride(contentTypes, PartPath.AppProps, ContentType.ExtendedProperties);
  if (hasCustomProps) {
    addContentTypeOverride(contentTypes, PartPath.CustomProps, ContentType.CustomProperties);
  }

  // --- Generate & add parts to archive ---
  // Note: [Content_Types].xml and _rels/.rels are serialized LAST so that any
  // relationships/content types registered during content rendering (e.g.
  // thumbnails, chart parts, alt chunks) are included.

  // word/_rels/document.xml.rels
  archive.add(
    PartPath.DocumentRels,
    renderXml(xml => renderRelationships(documentRels, xml))
  );

  // Build an effective doc that includes the auto-generated watermark header
  // reference in section properties (without mutating the caller's doc).
  let effectiveDoc: DocxDocument = doc;
  if (watermarkHeaderIndex !== undefined && watermarkHeaderRId) {
    const existingHeaders = doc.sectionProperties?.headers
      ? [...doc.sectionProperties.headers]
      : [];
    existingHeaders.push({
      type: "default",
      rId: watermarkHeaderRId
    });
    effectiveDoc = {
      ...doc,
      sectionProperties: {
        ...doc.sectionProperties,
        headers: existingHeaders
      }
    };
  }

  // Collect charts and register relationships BEFORE rendering document.xml
  const charts: ChartContent[] = [];
  collectCharts(doc.body, charts);
  if (doc.headers) {
    for (const [, h] of doc.headers) {
      collectChartsFromHeaderFooter(h.content, charts);
    }
  }
  if (doc.footers) {
    for (const [, f] of doc.footers) {
      collectChartsFromHeaderFooter(f.content, charts);
    }
  }

  const chartRIds: string[] = [];
  charts.forEach(chartContent => {
    const num = chartRIds.length + 1;
    const rId = addRelationship(documentRels, RelType.Chart, `charts/chart${num}.xml`);
    chartRIds.push(rId);
    // Store rId in registry so renderChartDrawing can look it up
    CHART_RID_REGISTRY.set(chartContent, rId);
  });

  // word/document.xml
  archive.add(
    PartPath.Document,
    renderXml(xml => renderDocument(xml, effectiveDoc))
  );

  // word/styles.xml
  archive.add(
    PartPath.Styles,
    renderXml(xml => renderStyles(xml, doc.docDefaults, doc.styles))
  );

  // word/settings.xml
  archive.add(
    PartPath.Settings,
    renderXml(xml => renderSettings(xml, doc.settings))
  );

  // word/fontTable.xml
  archive.add(
    PartPath.FontTable,
    renderXml(xml => renderFontTable(xml, doc.fonts))
  );

  // word/fonts/*.odttf (embedded fonts)
  if (doc.embeddedFonts && doc.embeddedFonts.length > 0) {
    const fontTableRels = createRelationships();
    const usedRIds = new Set<string>();

    // Collect rIds referenced in fontTable
    if (doc.fonts) {
      for (const font of doc.fonts) {
        if (font.embedRegular) {
          usedRIds.add(font.embedRegular);
        }
        if (font.embedBold) {
          usedRIds.add(font.embedBold);
        }
        if (font.embedItalic) {
          usedRIds.add(font.embedItalic);
        }
        if (font.embedBoldItalic) {
          usedRIds.add(font.embedBoldItalic);
        }
      }
    }

    for (const ef of doc.embeddedFonts) {
      const partPath = `word/fonts/${ef.fileName}`;
      archive.add(partPath, ef.data);

      // Register relationship from fontTable.xml
      if (usedRIds.has(ef.rId)) {
        addRelationshipWithId(fontTableRels, ef.rId, RelType.Font, `fonts/${ef.fileName}`);
      } else {
        // Add anyway so the embedded font isn't orphaned
        addRelationshipWithId(fontTableRels, ef.rId, RelType.Font, `fonts/${ef.fileName}`);
      }

      // Register content type for .odttf / .ttf
      const ext = ef.fileName.split(".").pop()?.toLowerCase();
      if (ext === "odttf") {
        addContentTypeDefault(contentTypes, "odttf", ContentType.ObfuscatedFont);
      } else if (ext === "ttf") {
        addContentTypeDefault(contentTypes, "ttf", "application/x-font-ttf");
      } else if (ext === "otf") {
        addContentTypeDefault(contentTypes, "otf", "application/x-font-otf");
      }
    }

    // Write fontTable.xml.rels
    archive.add(
      "word/_rels/fontTable.xml.rels",
      renderXml(xml => renderRelationships(fontTableRels, xml))
    );
  }

  // Custom XML parts (for SDT data binding)
  if (doc.customXmlParts && doc.customXmlParts.length > 0) {
    doc.customXmlParts.forEach((part, i) => {
      const num = i + 1;
      const itemPath = `word/customXml/item${num}.xml`;
      const propsPath = `word/customXml/itemProps${num}.xml`;

      // Write the XML content
      archive.add(itemPath, part.xmlContent);

      // Write itemProps*.xml
      const propsXml = renderXml(xml => {
        xml.openXml(STD_DOC_ATTRIBUTES);
        xml.openNode("ds:datastoreItem", {
          "ds:itemID": `{${part.itemId}}`,
          "xmlns:ds": "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
        });
        if (part.schemaReferences && part.schemaReferences.length > 0) {
          xml.openNode("ds:schemaRefs");
          for (const uri of part.schemaReferences) {
            xml.leafNode("ds:schemaRef", { "ds:uri": uri });
          }
          xml.closeNode();
        } else {
          xml.leafNode("ds:schemaRefs");
        }
        xml.closeNode();
      });
      archive.add(propsPath, propsXml);

      // Write item rels (links itemN.xml → itemPropsN.xml)
      const itemRels = createRelationships();
      addRelationship(itemRels, RelType.CustomXmlProps, `itemProps${num}.xml`);
      archive.add(
        `word/customXml/_rels/item${num}.xml.rels`,
        renderXml(xml => renderRelationships(itemRels, xml))
      );

      // Register content types
      addContentTypeOverride(
        contentTypes,
        `/word/customXml/itemProps${num}.xml`,
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
      );

      // Add to document rels
      addRelationship(documentRels, RelType.CustomXml, `customXml/item${num}.xml`);
    });
  }

  // word/theme/theme1.xml
  archive.add(
    PartPath.Theme,
    renderXml(xml => renderTheme(xml, doc.theme))
  );

  // word/numbering.xml
  if (hasNumbering) {
    archive.add(
      PartPath.Numbering,
      renderXml(xml =>
        renderNumbering(xml, doc.abstractNumberings, doc.numberingInstances, doc.numPicBullets)
      )
    );
  }

  // word/footnotes.xml
  if (hasFootnotes) {
    archive.add(
      PartPath.Footnotes,
      renderXml(xml => renderFootnotes(xml, doc.footnotes!))
    );
  }

  // word/endnotes.xml
  if (hasEndnotes) {
    archive.add(
      PartPath.Endnotes,
      renderXml(xml => renderEndnotes(xml, doc.endnotes!))
    );
  }

  // word/comments.xml
  if (hasComments) {
    archive.add(
      PartPath.Comments,
      renderXml(xml => renderComments(xml, doc.comments!))
    );

    // word/commentsExtended.xml (for done/parentId)
    const hasExtended = doc.comments!.some(c => c.done !== undefined || c.parentId !== undefined);
    if (hasExtended) {
      const extXml = renderXml(xml => renderCommentsExtended(xml, doc.comments!));
      archive.add(PartPath.CommentsExtended, extXml);
      addRelationship(documentRels, RelType.CommentsExtended, "commentsExtended.xml");
      addContentTypeOverride(
        contentTypes,
        `/${PartPath.CommentsExtended}`,
        ContentType.CommentsExtended
      );
    }
  }

  // Headers
  headerIndex = 1;
  if (doc.headers) {
    let hIdx = 0;
    const keys = [...doc.headers.keys()];
    for (const [, headerDef] of doc.headers) {
      archive.add(
        PartPath.header(headerIndex),
        renderXml(xml => renderHeader(xml, headerDef.content))
      );
      // Header .rels file
      const hKey = keys[hIdx];
      const hRels = headerRelManagers.get(hKey);
      if (hRels && getRelationshipCount(hRels) > 0) {
        archive.add(
          `word/_rels/header${headerIndex}.xml.rels`,
          renderXml(xml => renderRelationships(hRels, xml))
        );
      }
      headerIndex++;
      hIdx++;
    }
  }

  // Footers
  footerIndex = 1;
  if (doc.footers) {
    let fIdx = 0;
    const keys = [...doc.footers.keys()];
    for (const [, footerDef] of doc.footers) {
      archive.add(
        PartPath.footer(footerIndex),
        renderXml(xml => renderFooter(xml, footerDef.content))
      );
      // Footer .rels file
      const fKey = keys[fIdx];
      const fRels = footerRelManagers.get(fKey);
      if (fRels && getRelationshipCount(fRels) > 0) {
        archive.add(
          `word/_rels/footer${footerIndex}.xml.rels`,
          renderXml(xml => renderRelationships(fRels, xml))
        );
      }
      footerIndex++;
      fIdx++;
    }
  }

  // Watermark header
  if (watermarkHeaderIndex !== undefined && doc.watermark) {
    archive.add(
      PartPath.header(watermarkHeaderIndex),
      renderXml(xml => renderWatermarkHeader(xml, doc.watermark!))
    );
    if (watermarkHeaderRels && getRelationshipCount(watermarkHeaderRels) > 0) {
      archive.add(
        `word/_rels/header${watermarkHeaderIndex}.xml.rels`,
        renderXml(xml => renderRelationships(watermarkHeaderRels!, xml))
      );
    }
  }

  // Media / images
  if (doc.images) {
    for (const img of doc.images) {
      archive.add(PartPath.media(img.fileName), img.data);
    }
  }

  // SVG fallback PNG files
  for (const fb of svgFallbacks) {
    archive.add(PartPath.media(fb.fallbackFileName), fb.data);
  }

  // docProps/core.xml
  archive.add(
    PartPath.CoreProps,
    renderXml(xml => renderCoreProperties(xml, doc.coreProperties))
  );

  // docProps/app.xml
  archive.add(
    PartPath.AppProps,
    renderXml(xml => renderAppProperties(xml, doc.appProperties))
  );

  // docProps/custom.xml
  if (hasCustomProps) {
    archive.add(
      PartPath.CustomProps,
      renderXml(xml => renderCustomProperties(xml, doc.customProperties!))
    );
  }

  // word/webSettings.xml
  if (doc.webSettings) {
    archive.add(
      PartPath.WebSettings,
      renderXml(xml => renderWebSettings(xml, doc.webSettings))
    );
    addRelationship(documentRels, RelType.WebSettings, "webSettings.xml");
    addContentTypeOverride(contentTypes, `/${PartPath.WebSettings}`, ContentType.WebSettings);
  }

  // word/people.xml
  if (doc.people && doc.people.length > 0) {
    archive.add(
      PartPath.People,
      renderXml(xml => renderPeople(xml, doc.people!))
    );
    addRelationship(documentRels, RelType.People, "people.xml");
    addContentTypeOverride(contentTypes, `/${PartPath.People}`, ContentType.People);
  }

  // docProps/thumbnail
  if (doc.thumbnail) {
    const ext =
      doc.thumbnail.contentType === "image/jpeg"
        ? "jpeg"
        : doc.thumbnail.contentType === "image/png"
          ? "png"
          : "wmf";
    const thumbPath = `docProps/thumbnail.${ext}`;
    archive.add(thumbPath, doc.thumbnail.data);
    // Package rels: target is relative to package root (docProps/thumbnail.jpeg)
    addRelationship(
      packageRels,
      "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
      thumbPath
    );
    addContentTypeDefault(contentTypes, ext, doc.thumbnail.contentType);
  }

  // Write opaque (unrecognized) parts for round-trip preservation
  if (doc.opaqueParts) {
    for (const part of doc.opaqueParts) {
      archive.add(part.path, part.data);
      // Register content type: explicit > inferred by extension > skip
      const ext = part.path.split(".").pop()?.toLowerCase();
      if (part.contentType) {
        addContentTypeOverride(contentTypes, `/${part.path}`, part.contentType);
      } else if (ext) {
        // Infer from common extensions so [Content_Types].xml isn't incomplete
        const inferred = inferContentType(ext);
        if (inferred) {
          addContentTypeOverride(contentTypes, `/${part.path}`, inferred);
        }
      }
      // Write part relationships if any
      if (part.relationships && part.relationships.length > 0) {
        const partRels = createRelationships();
        for (const rel of part.relationships) {
          addRelationshipWithId(partRels, rel.id, rel.type, rel.target, rel.targetMode);
        }
        const relsPath = getPartRelsPath(part.path);
        archive.add(
          relsPath,
          renderXml(xml => renderRelationships(partRels, xml))
        );
      }
    }
  }

  // Write chart parts (rIds already registered earlier)
  charts.forEach((chartContent, i) => {
    const num = i + 1;
    const chartPath = `word/charts/chart${num}.xml`;

    // Register chart part
    archive.add(
      chartPath,
      renderXml(xml => renderChartPart(xml, chartContent.chart))
    );

    // Register content type
    addContentTypeOverride(contentTypes, `/word/charts/chart${num}.xml`, ContentType.Chart);
  });

  // LAST: Write [Content_Types].xml and _rels/.rels after all parts have registered
  // their content types and relationships.
  archive.add(
    PartPath.ContentTypes,
    renderXml(xml => renderContentTypes(contentTypes, xml))
  );
  archive.add(
    PartPath.PackageRels,
    renderXml(xml => renderRelationships(packageRels, xml))
  );

  return archive.bytes();
}

/** Recursively collect altChunks from body content. */
function collectAltChunks(body: readonly BodyContent[], out: AltChunk[]): void {
  for (const item of body) {
    if ("type" in item && item.type === "altChunk") {
      out.push(item);
    } else if ("type" in item && item.type === "table") {
      for (const row of item.rows) {
        for (const cell of row.cells) {
          collectAltChunks(cell.content as readonly BodyContent[], out);
        }
      }
    } else if ("type" in item && item.type === "sdt") {
      collectAltChunks(item.content as readonly BodyContent[], out);
    }
  }
}

/** Recursively collect chart contents from body content. */
function collectCharts(body: readonly BodyContent[], out: ChartContent[]): void {
  for (const item of body) {
    if ("type" in item && item.type === "chart") {
      out.push(item as ChartContent);
    } else if ("type" in item && item.type === "table") {
      for (const row of item.rows) {
        for (const cell of row.cells) {
          collectCharts(cell.content as readonly BodyContent[], out);
        }
      }
    } else if ("type" in item && item.type === "sdt") {
      collectCharts(item.content as readonly BodyContent[], out);
    }
  }
}

/** Collect charts from header/footer content. */
function collectChartsFromHeaderFooter(
  content: { children: readonly BodyContent[] },
  out: ChartContent[]
): void {
  collectCharts(content.children, out);
}
