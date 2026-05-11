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

import { ContentType, RelType, PartPath, STD_DOC_ATTRIBUTES } from "../constants";
import { type Mutable } from "../core/internal-utils";
import { getFileExt, getPartRelsPath } from "../core/opc-package";
import { walkBlocks } from "../core/walker";
import { DocxWriteError } from "../errors";
import type {
  DocxDocument,
  BodyContent,
  Paragraph,
  ParagraphChild,
  Table,
  Hyperlink,
  HeaderFooterContent,
  Run,
  RunContent,
  FloatingImage,
  InlineImageContent,
  ChartContent,
  ChartExContent,
  AltChunk,
  ImageDef
} from "../types";
import { renderChartPart } from "./chart-writer";
import { renderComments, renderCommentsExtended } from "./comment-writer";
import {
  createContentTypes,
  addContentTypeDefault,
  addContentTypeOverride,
  addImageContentTypeDefaults,
  renderContentTypes
} from "./content-types";
import { renderDocument } from "./document-writer";
import { renderFootnotes, renderEndnotes } from "./footnote-writer";
import { renderHeader, renderFooter, renderWatermarkHeader } from "./header-footer-writer";
import { renderNumbering } from "./numbering-writer";
import {
  renderSettings,
  renderFontTable,
  renderCoreProperties,
  renderAppProperties,
  renderCustomProperties,
  renderWebSettings,
  renderPeople,
  renderTheme
} from "./parts-writer";
import {
  createRelationships,
  addRelationship,
  addRelationshipWithId,
  getRelationshipCount,
  renderRelationships
} from "./relationships";
import type { RelationshipsState } from "./relationships";
import { createRenderContext } from "./render-context";
import { renderStyles } from "./styles-writer";

/** Render XML to string using XmlWriter. */
function renderXml(renderFn: (xml: XmlWriter) => void): string {
  const writer = new XmlWriter();
  renderFn(writer);
  return writer.xml;
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
  walkBlocks(body, {
    enterParagraph(p) {
      scanChildrenForHyperlinks(p.children, links);
    }
  });
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
 * tables (including nested tables), SDTs, TOC cached paragraphs, and text
 * boxes. Implemented in terms of {@link walkBlocks} so block dispatch stays
 * consistent with the rest of the module.
 */
function walkParagraphs(blocks: readonly BodyContent[], onParagraph: (p: Paragraph) => void): void {
  walkBlocks(blocks, {
    enterParagraph(p) {
      onParagraph(p);
    }
  });
}

/** Collect all image rIds referenced in header/footer content. */
function collectImageRidsFromContent(content: HeaderFooterContent): Set<string> {
  const rIds = new Set<string>();
  walkParagraphs(content.children as readonly BodyContent[], p => {
    scanChildrenForImages(p.children, rIds);
  });
  return rIds;
}

/** Collect hyperlinks from header/footer content. */
function collectHyperlinksFromHeaderFooter(content: HeaderFooterContent): Hyperlink[] {
  const links: Hyperlink[] = [];
  walkParagraphs(content.children as readonly BodyContent[], p => {
    scanChildrenForHyperlinks(p.children, links);
  });
  return links;
}

/** Resolve the main document part content type based on docType. */
function resolveDocumentContentType(doc: DocxDocument): string {
  switch (doc.docType) {
    case "template":
      return ContentType.Template;
    case "macroEnabledDocument":
      return ContentType.DocumentMacroEnabled;
    case "macroEnabledTemplate":
      return ContentType.TemplateMacroEnabled;
    default:
      return ContentType.Document;
  }
}

/**
 * Package a DocxDocument model into a DOCX ZIP file.
 * Returns the ZIP bytes as Uint8Array.
 *
 * This function does NOT modify the input `doc` object. Internally it creates
 * shallow copies of mutable structures (images, headers, footers, etc.) to
 * assign relationship IDs without polluting the caller's model.
 */
export async function packageDocx(
  doc: DocxDocument,
  compressionLevel?: number
): Promise<Uint8Array> {
  // Create a working copy so we never mutate the caller's doc
  const workDoc = shallowCopyDocForPackaging(doc);
  return _packageDocxInner(workDoc, compressionLevel);
}

/**
 * Create a shallow working copy of the document parts that packageDocx still
 * mutates (header/footer/altChunk/hyperlink rId injection). Binary data
 * (Uint8Array) is shared, not cloned. The image table is intentionally NOT
 * cloned: packaging never rewrites image rIds anymore — collisions are now
 * handled via a render-context remap (see `imageRemap` below).
 */
function shallowCopyDocForPackaging(doc: DocxDocument): DocxDocument {
  // Shallow-copy body so we can wrap floating images / altChunks without
  // touching the caller's array.
  const body = doc.body.map(item => {
    if (item.type === "floatingImage" || item.type === "altChunk") {
      return { ...item };
    }
    return item;
  });

  // Shallow-copy headers/footers maps (their `rId` is still assigned in place
  // for backward compatibility with downstream relationship lookups).
  const headers = doc.headers
    ? new Map(Array.from(doc.headers.entries()).map(([k, v]) => [k, { ...v }]))
    : undefined;
  const footers = doc.footers
    ? new Map(Array.from(doc.footers.entries()).map(([k, v]) => [k, { ...v }]))
    : undefined;

  return {
    ...doc,
    body,
    headers,
    footers
  } as DocxDocument;
}

async function _packageDocxInner(
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
  //
  // The model's `image.rId` is the authoritative reference: it is what every
  // inline/floating drawing in the body uses, so the easiest way to keep
  // r:embed and the .rels file in sync is to register relationships under that
  // exact id. If a model rId happens to clash with a previously-registered
  // relationship in this same .rels file (e.g. styles → "rId1" while a model
  // image also asks for "rId1"), we fall back to allocating a fresh rId and
  // record the substitution in `imageRemap`. Writers consult that map via
  // ctx.imageRIdRemap when emitting r:embed so the body and the .rels file
  // stay consistent.
  //
  // For SVG with raster fallback, the model rId points at the PNG fallback
  // resource (so r:embed on a:blip targets a raster image — the legacy wire
  // shape Word readers expect). We additionally register the SVG itself
  // under a secondary rId. We then surface that secondary rId on the
  // inline/floating image's `svgRId` field — but only on shallow copies that
  // live inside the work doc, so the caller's input model is never mutated.
  const imageExtensions = new Set<string>();
  const imageRemap = new Map<string, string>(); // model rId -> registered rId (only set on collision)
  const imageByRid = new Map<string, ImageDef>(); // model rId -> ImageDef
  const svgFallbacks: { fallbackFileName: string; data: Uint8Array }[] = [];
  // For each main image rId that has a SVG fallback, the secondary rId
  // pointing at the SVG file. Used to populate svgRId on inline/floating
  // image content when the source model didn't already provide one.
  const imageSvgRIdMap = new Map<string, string>();
  if (doc.images) {
    for (const img of doc.images) {
      if (img.rId) {
        imageByRid.set(img.rId, img);
      }
      // Index alias rIds so that header/footer-local references resolve to
      // the same ImageDef. Aliases live in their own .rels id space, so when
      // we register them later we use the alias rId verbatim.
      if (img.aliasRIds) {
        for (const alias of img.aliasRIds) {
          if (!imageByRid.has(alias)) {
            imageByRid.set(alias, img);
          }
        }
      }
    }

    const registerImageRel = (modelRId: string | undefined, target: string): string => {
      if (!modelRId) {
        return addRelationship(documentRels, RelType.Image, target);
      }
      if (documentRels.hasId(modelRId)) {
        const newId = addRelationship(documentRels, RelType.Image, target);
        imageRemap.set(modelRId, newId);
        return newId;
      }
      addRelationshipWithId(documentRels, modelRId, RelType.Image, target);
      return modelRId;
    };

    for (const img of doc.images) {
      const oldRid = img.rId;

      if (img.mediaType === "svg" && img.fallbackData) {
        // Main rId points at the PNG fallback (the raster image consumed by
        // a:blip).
        const baseName = img.fileName.replace(/\.[^.]+$/, "");
        const fallbackFileName = `${baseName}_fallback.png`;
        registerImageRel(oldRid, `media/${fallbackFileName}`);
        imageExtensions.add("png");

        // Secondary rId for the actual SVG; auto-allocated.
        const svgRId = addRelationship(documentRels, RelType.Image, `media/${img.fileName}`);
        if (oldRid) {
          imageSvgRIdMap.set(oldRid, svgRId);
        }
        const ext = getFileExt(img.fileName);
        if (ext) {
          imageExtensions.add(ext);
        }

        svgFallbacks.push({ fallbackFileName, data: img.fallbackData });
      } else {
        registerImageRel(oldRid, `media/${img.fileName}`);
        const ext = getFileExt(img.fileName);
        if (ext) {
          imageExtensions.add(ext);
        }
      }
    }
  }

  // Populate svgRId on inline/floating image content nodes for SVG fallbacks.
  // We only mutate shallow copies inside the working doc — never the caller's
  // input. This is invisible to readers that already supplied svgRId.
  if (imageSvgRIdMap.size > 0) {
    const populateSvg = (target: { rId: string; svgRId?: string }): void => {
      if (target.svgRId) {
        return;
      }
      const svgRId = imageSvgRIdMap.get(target.rId);
      if (svgRId) {
        (target as { svgRId?: string }).svgRId = svgRId;
      }
    };
    // Walk the work-doc body. Floating images and altChunks were shallow-copied
    // by shallowCopyDocForPackaging; for paragraphs we lazily clone the run
    // content nodes that carry inline SVG references.
    const processBody = (blocks: readonly BodyContent[]): BodyContent[] => {
      return blocks.map(block => {
        if (block.type === "floatingImage") {
          if (imageSvgRIdMap.has(block.rId)) {
            populateSvg(block as Mutable<FloatingImage>);
          }
          return block;
        }
        if (block.type === "paragraph") {
          let paragraphCopied: Paragraph | null = null;
          let childrenCopied: ParagraphChild[] | null = null;
          for (let i = 0; i < block.children.length; i++) {
            const child = block.children[i];
            if (!("content" in child) || !Array.isArray((child as Run).content)) {
              continue;
            }
            const run = child as Run;
            let runCopied: Run | null = null;
            let contentCopied: RunContent[] | null = null;
            for (let j = 0; j < run.content.length; j++) {
              const c = run.content[j];
              if (c.type !== "image" || !c.rId) {
                continue;
              }
              if (!imageSvgRIdMap.has(c.rId) || c.svgRId) {
                continue;
              }
              if (!contentCopied) {
                contentCopied = [...run.content];
              }
              const cloned: InlineImageContent = { ...(c as InlineImageContent) };
              populateSvg(cloned as Mutable<InlineImageContent>);
              contentCopied[j] = cloned;
            }
            if (contentCopied) {
              runCopied = { ...run, content: contentCopied };
              if (!childrenCopied) {
                childrenCopied = [...block.children];
              }
              childrenCopied[i] = runCopied as ParagraphChild;
            }
          }
          if (childrenCopied) {
            paragraphCopied = { ...block, children: childrenCopied };
            return paragraphCopied;
          }
          return block;
        }
        if (block.type === "table") {
          const newRows = block.rows.map(row => ({
            ...row,
            cells: row.cells.map(cell => ({
              ...cell,
              content: processBody(cell.content as readonly BodyContent[]) as readonly (
                | Paragraph
                | Table
              )[]
            }))
          }));
          return { ...block, rows: newRows } as Table;
        }
        return block;
      });
    };
    (doc as Mutable<DocxDocument>).body = processBody(doc.body);

    // Same treatment for header/footer content (each is essentially a body
    // fragment). headers/footers maps were shallow-copied by
    // shallowCopyDocForPackaging, so swapping `content.children` is safe.
    const rewriteHeaderFooter = <T extends { content: HeaderFooterContent }>(
      defs: ReadonlyMap<string, T> | undefined
    ): void => {
      if (!defs) {
        return;
      }
      for (const [, def] of defs) {
        const newChildren = processBody(
          def.content.children as readonly BodyContent[]
        ) as readonly (Paragraph | Table)[];
        if (newChildren !== def.content.children) {
          (def as { content: HeaderFooterContent }).content = {
            ...def.content,
            children: newChildren
          };
        }
      }
    };
    rewriteHeaderFooter(doc.headers);
    rewriteHeaderFooter(doc.footers);

    // Footnotes and endnotes carry their own paragraph lists.
    const rewriteNotes = (
      notes: readonly { id: number; content: readonly Paragraph[] }[] | undefined
    ): readonly { id: number; content: readonly Paragraph[] }[] | undefined => {
      if (!notes || notes.length === 0) {
        return notes;
      }
      let changed = false;
      const out = notes.map(note => {
        const newContent = processBody(
          note.content as readonly BodyContent[]
        ) as readonly Paragraph[];
        if (newContent === note.content) {
          return note;
        }
        changed = true;
        return { ...note, content: newContent };
      });
      return changed ? out : notes;
    };
    const newFootnotes = rewriteNotes(
      doc.footnotes as readonly { id: number; content: readonly Paragraph[] }[] | undefined
    );
    if (newFootnotes !== doc.footnotes) {
      (doc as Mutable<DocxDocument>).footnotes = newFootnotes as DocxDocument["footnotes"];
    }
    const newEndnotes = rewriteNotes(
      doc.endnotes as readonly { id: number; content: readonly Paragraph[] }[] | undefined
    );
    if (newEndnotes !== doc.endnotes) {
      (doc as Mutable<DocxDocument>).endnotes = newEndnotes as DocxDocument["endnotes"];
    }
  }

  // Hyperlinks (external).
  // We register relationships under the model's existing rId where possible,
  // and stash newly-allocated rIds on a WeakMap keyed by the Hyperlink object
  // so the renderer can look them up without us writing onto the caller's
  // model. The map lives on the render context.
  //
  // Body, footnotes, endnotes, headers and footers each render against a
  // *different* OPC part, so each part has its own .rels namespace. Using one
  // WeakMap is fine because Hyperlink objects are unique by identity, but the
  // **registration** has to happen against the correct rel manager — we register
  // body hyperlinks into documentRels here, and footnote/endnote/header/footer
  // hyperlinks into their own rel managers below.
  const hyperlinkRIds = new WeakMap<object, string>();
  const hyperlinks = collectHyperlinks(doc.body);
  for (const link of hyperlinks) {
    if (link.url) {
      const rId = addRelationship(documentRels, RelType.Hyperlink, link.url, "External");
      hyperlinkRIds.set(link, rId);
    }
  }

  // Footnotes/endnotes have independent .rels parts. Build them lazily so we
  // only emit a footnotes.xml.rels file when there's something to register.
  const footnoteRels = createRelationships();
  const endnoteRels = createRelationships();
  const registerNoteHyperlinks = (
    notes: readonly { content: readonly Paragraph[] }[] | undefined,
    rels: RelationshipsState
  ): void => {
    if (!notes) {
      return;
    }
    for (const note of notes) {
      const links: Hyperlink[] = [];
      for (const p of note.content) {
        scanChildrenForHyperlinks(p.children, links);
      }
      for (const link of links) {
        if (link.url) {
          const rId = addRelationship(rels, RelType.Hyperlink, link.url, "External");
          hyperlinkRIds.set(link, rId);
        }
      }
    }
  };
  registerNoteHyperlinks(doc.footnotes, footnoteRels);
  registerNoteHyperlinks(doc.endnotes, endnoteRels);

  // Footnotes/endnotes may also reference images. Register those references
  // using the same model rId that the run content emits so footnotes.xml's
  // r:embed resolves locally.
  const registerNoteImages = (
    notes: readonly { content: readonly Paragraph[] }[] | undefined,
    rels: RelationshipsState
  ): void => {
    if (!notes || !doc.images) {
      return;
    }
    const seen = new Set<string>();
    for (const note of notes) {
      for (const p of note.content) {
        const rIds = new Set<string>();
        scanChildrenForImages(p.children, rIds);
        for (const oldRid of rIds) {
          if (seen.has(oldRid)) {
            continue;
          }
          seen.add(oldRid);
          const img = imageByRid.get(oldRid);
          if (img) {
            addRelationshipWithId(rels, oldRid, RelType.Image, `media/${img.fileName}`);
            // SVG fallback: also register the secondary svg rId locally if the
            // model knows it (or if we allocated it). For notes we look up the
            // already-allocated svgRId from imageSvgRIdMap if present.
            const svgRId = imageSvgRIdMap.get(oldRid);
            if (svgRId && !rels.hasId(svgRId)) {
              addRelationshipWithId(rels, svgRId, RelType.Image, `media/${img.fileName}`);
            }
          }
        }
      }
    }
  };
  registerNoteImages(doc.footnotes, footnoteRels);
  registerNoteImages(doc.endnotes, endnoteRels);

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

      // Create per-header relationship state for images and hyperlinks.
      // Header/footer .rels is its own id space, so we register images using
      // the model rId — that's also what the body XML emits via r:embed
      // (header/footer XML never goes through imageRemap because the local
      // .rels is independent from the document .rels).
      const hRels = createRelationships();
      const imgRids = collectImageRidsFromContent(headerDef.content);
      if (imgRids.size > 0 && doc.images) {
        for (const oldRid of imgRids) {
          const img = imageByRid.get(oldRid);
          if (img) {
            addRelationshipWithId(hRels, oldRid, RelType.Image, `media/${img.fileName}`);
            // SVG fallback: register the secondary svg rId locally too so
            // asvg:svgBlip in header XML resolves against this part's .rels.
            const svgRId = imageSvgRIdMap.get(oldRid);
            if (svgRId && !hRels.hasId(svgRId)) {
              addRelationshipWithId(hRels, svgRId, RelType.Image, `media/${img.fileName}`);
            }
          }
        }
      }
      // Hyperlinks in header
      const hLinks = collectHyperlinksFromHeaderFooter(headerDef.content);
      for (const link of hLinks) {
        if (link.url) {
          const linkRId = addRelationship(hRels, RelType.Hyperlink, link.url, "External");
          hyperlinkRIds.set(link, linkRId);
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

      // Create per-footer relationship state for images and hyperlinks.
      const fRels = createRelationships();
      const imgRids = collectImageRidsFromContent(footerDef.content);
      if (imgRids.size > 0 && doc.images) {
        for (const oldRid of imgRids) {
          const img = imageByRid.get(oldRid);
          if (img) {
            addRelationshipWithId(fRels, oldRid, RelType.Image, `media/${img.fileName}`);
            // SVG fallback: register the secondary svg rId locally too.
            const svgRId = imageSvgRIdMap.get(oldRid);
            if (svgRId && !fRels.hasId(svgRId)) {
              addRelationshipWithId(fRels, svgRId, RelType.Image, `media/${img.fileName}`);
            }
          }
        }
      }
      // Hyperlinks in footer
      const fLinks = collectHyperlinksFromHeaderFooter(footerDef.content);
      for (const link of fLinks) {
        if (link.url) {
          const linkRId = addRelationship(fRels, RelType.Hyperlink, link.url, "External");
          hyperlinkRIds.set(link, linkRId);
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
      const img = imageByRid.get(wmRId);
      if (img) {
        addRelationshipWithId(wmRels, wmRId, RelType.Image, `media/${img.fileName}`);
      }
      watermarkHeaderRels = wmRels;
    }
    headerIndex++;
  }

  // --- Content Types ---
  addImageContentTypeDefaults(contentTypes, imageExtensions);

  // Determine main document content type based on docType
  const mainDocContentType = resolveDocumentContentType(doc);
  addContentTypeOverride(contentTypes, PartPath.Document, mainDocContentType);
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
  // word/_rels/document.xml.rels is also deferred to include chart relationships.

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
  const renderCtxChartRIds = new Map<object, string>();
  charts.forEach(chartContent => {
    const num = chartRIds.length + 1;
    const rId = addRelationship(documentRels, RelType.Chart, `charts/chart${num}.xml`);
    chartRIds.push(rId);
    renderCtxChartRIds.set(chartContent, rId);
  });

  // Collect ChartEx items and register relationships
  const chartExItems: ChartExContent[] = [];
  collectChartExItems(doc.body, chartExItems);
  const chartExRIds: string[] = [];
  chartExItems.forEach(cxContent => {
    const num = chartExRIds.length + 1;
    const rId = addRelationship(documentRels, RelType.ChartEx, `charts/chartEx${num}.xml`);
    chartExRIds.push(rId);
    renderCtxChartRIds.set(cxContent, rId);
  });

  // Create render context for document serialization
  const renderCtx = createRenderContext({
    chartRIds: renderCtxChartRIds,
    imageRIdRemap: imageRemap,
    hyperlinkRIds
  });

  // word/_rels/document.xml.rels — deferred until after all document-level
  // relationships have been registered (VBA, webSettings, people, charts, etc.)

  // word/document.xml
  archive.add(
    PartPath.Document,
    renderXml(xml => renderDocument(xml, effectiveDoc, renderCtx))
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
      const ext = getFileExt(ef.fileName);
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

  // word/footnotes.xml + footnotes.xml.rels
  if (hasFootnotes) {
    archive.add(
      PartPath.Footnotes,
      // Footnotes are an independent OPC part — their r:id values must
      // resolve against word/_rels/footnotes.xml.rels, not document.xml.rels.
      renderXml(xml =>
        renderFootnotes(xml, doc.footnotes!, { imageRemap: new Map(), hyperlinkRIds })
      )
    );
    if (getRelationshipCount(footnoteRels) > 0) {
      archive.add(
        "word/_rels/footnotes.xml.rels",
        renderXml(xml => renderRelationships(footnoteRels, xml))
      );
    }
  }

  // word/endnotes.xml + endnotes.xml.rels
  if (hasEndnotes) {
    archive.add(
      PartPath.Endnotes,
      renderXml(xml => renderEndnotes(xml, doc.endnotes!, { imageRemap: new Map(), hyperlinkRIds }))
    );
    if (getRelationshipCount(endnoteRels) > 0) {
      archive.add(
        "word/_rels/endnotes.xml.rels",
        renderXml(xml => renderRelationships(endnoteRels, xml))
      );
    }
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
        // Header .rels is independent from document .rels — pass an empty
        // imageRemap so we never rewrite r:embed against the document remap.
        renderXml(xml =>
          renderHeader(xml, headerDef.content, { imageRemap: new Map(), hyperlinkRIds })
        )
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
        // Footer .rels is independent from document .rels.
        renderXml(xml =>
          renderFooter(xml, footerDef.content, { imageRemap: new Map(), hyperlinkRIds })
        )
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

  // word/vbaProject.bin (macro-enabled documents)
  if (doc.vbaProject) {
    archive.add("word/vbaProject.bin", doc.vbaProject);
    addRelationship(documentRels, RelType.VbaProject, "vbaProject.bin");
    addContentTypeOverride(contentTypes, "/word/vbaProject.bin", ContentType.VbaProject);
  }

  // Write opaque (unrecognized) parts for round-trip preservation
  if (doc.opaqueParts) {
    for (const part of doc.opaqueParts) {
      archive.add(part.path, part.data);
      // Register content type: explicit > inferred by extension > skip
      const ext = getFileExt(part.path);
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
  const chartEmbedPromises: Promise<void>[] = [];
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

    // Embed xlsx workbook if requested
    if (chartContent.chart.embedWorkbook) {
      const xlsxPath = `word/embeddings/Microsoft_Excel_Worksheet${num}.xlsx`;

      // Create chart .rels file referencing the embedded xlsx
      const chartRels = createRelationships();
      addRelationship(
        chartRels,
        RelType.Package,
        `../embeddings/Microsoft_Excel_Worksheet${num}.xlsx`
      );
      archive.add(
        `word/charts/_rels/chart${num}.xml.rels`,
        renderXml(xml => renderRelationships(chartRels, xml))
      );

      // Register xlsx content type
      addContentTypeOverride(contentTypes, `/${xlsxPath}`, ContentType.Xlsx);

      // Generate and add xlsx asynchronously via Excel module (dynamic import)
      const promise = import("../bridge/excel-bridge").then(
        async ({ generateChartEmbeddedXlsx }) => {
          const xlsxData = await generateChartEmbeddedXlsx(chartContent.chart.series);
          archive.add(xlsxPath, xlsxData);
        }
      );
      chartEmbedPromises.push(promise);
    }
  });

  // Wait for all embedded xlsx files to be generated
  if (chartEmbedPromises.length > 0) {
    await Promise.all(chartEmbedPromises);
  }

  // Write ChartEx parts (cx:chartSpace XML)
  chartExItems.forEach((cxContent, i) => {
    const num = i + 1;
    const cxPath = `word/charts/chartEx${num}.xml`;
    archive.add(cxPath, cxContent.chartExXml);
    addContentTypeOverride(contentTypes, `/${cxPath}`, ContentType.ChartEx);
  });

  // LAST: Write [Content_Types].xml, _rels/.rels, and word/_rels/document.xml.rels
  // after all parts have registered their content types and relationships.

  // Validate relationships before serializing (catch duplicate IDs, missing TargetMode, etc.)
  const docRelErrors = documentRels.validate();
  const pkgRelErrors = packageRels.validate();
  if (docRelErrors.length > 0 || pkgRelErrors.length > 0) {
    const allErrors = [...pkgRelErrors, ...docRelErrors];
    throw new DocxWriteError(`OPC relationship validation failed:\n${allErrors.join("\n")}`);
  }

  archive.add(
    PartPath.DocumentRels,
    renderXml(xml => renderRelationships(documentRels, xml))
  );
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

/** Recursively collect ChartEx contents from body content. */
function collectChartExItems(body: readonly BodyContent[], out: ChartExContent[]): void {
  for (const item of body) {
    if ("type" in item && item.type === "chartEx") {
      out.push(item as ChartExContent);
    } else if ("type" in item && item.type === "table") {
      for (const row of item.rows) {
        for (const cell of row.cells) {
          collectChartExItems(cell.content as readonly BodyContent[], out);
        }
      }
    } else if ("type" in item && item.type === "sdt") {
      collectChartExItems(item.content as readonly BodyContent[], out);
    }
  }
}
