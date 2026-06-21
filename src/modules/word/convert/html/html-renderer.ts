/**
 * DOCX Module - HTML Renderer
 *
 * Converts a DocxDocument model to semantic HTML5.
 * Supports paragraphs, runs with formatting, tables, lists, images, hyperlinks,
 * headings, comments, footnotes/endnotes, and more.
 */

import { resolveThemeColor } from "@word/core/color-utils";
import { bytesToBase64, sanitizeUrl } from "@word/core/internal-utils";
import { extractMathText, isRun } from "@word/core/text-utils";
import type {
  DocxDocument,
  Paragraph,
  Run,
  RunContent,
  ParagraphChild,
  Table,
  TableCell,
  TableRow,
  BodyContent,
  Hyperlink,
  FloatingImage,
  InlineImageContent,
  ImageDef,
  ColorSpec,
  Border,
  TextBox,
  DrawingShape,
  StructuredDocumentTag,
  DocumentTheme,
  MathBlock,
  CommentDef,
  Chart,
  ChartExContent
} from "@word/types";
import { EMU_PER_INCH } from "@word/units";
import { stripXmlIllegalChars } from "@xml/encode";

/** Options for HTML rendering. */
export interface HtmlRenderOptions {
  /** Include CSS inline in a <style> tag. Default: true. */
  readonly includeStyles?: boolean;
  /** Wrap output in complete HTML document. Default: true. */
  readonly fullDocument?: boolean;
  /** Page title for full document. Default: core property "title" or "Document". */
  readonly title?: string;
  /** Base URL for image data: URLs vs file references. "dataUrl" embeds images as base64. */
  readonly imageMode?: "dataUrl" | "filename" | "none";
  /** Convert comments to HTML. Default: false. */
  readonly includeComments?: boolean;
  /** Render revision marks (insertions/deletions) as <ins>/<del>. Default: false. */
  readonly includeRevisions?: boolean;
  /** Render footnotes/endnotes as anchored references. Default: true. */
  readonly includeNotes?: boolean;
  /** Custom CSS class prefix (default: "docx-"). */
  readonly classPrefix?: string;
  /**
   * Style map: Word style name → HTML tag/class mapping.
   * Values use the format "tag.className" (e.g. "h1.doc-heading", "blockquote.quote").
   * If only a tag is provided (no dot), no class is added.
   * If only a class is provided (starts with "."), the default tag is used with that class.
   */
  readonly styleMap?: ReadonlyMap<string, string> | Record<string, string>;
  /** Optional chart renderer callback. When provided, charts are rendered as inline SVG. */
  readonly chartRenderer?: (chart: Chart) => string;
  /** Optional ChartEx renderer callback. When provided, chartEx blocks are rendered as inline SVG. */
  readonly chartExRenderer?: (chartEx: ChartExContent) => string;
}

/** The result of HTML rendering. */
export interface HtmlRenderResult {
  /** Generated HTML string. */
  readonly html: string;
  /** Warnings encountered during rendering. */
  readonly warnings: readonly string[];
  /** Image map: file name → data URL or reference. */
  readonly images: ReadonlyMap<string, string>;
}

/** Resolved options with all required fields except optional callbacks. */
type ResolvedHtmlRenderOptions = Required<
  Omit<HtmlRenderOptions, "chartRenderer" | "chartExRenderer">
> &
  Pick<HtmlRenderOptions, "chartRenderer" | "chartExRenderer">;

/** Internal rendering state. */
interface RenderState {
  options: ResolvedHtmlRenderOptions;
  doc: DocxDocument;
  imageMap: Map<string, string>;
  rIdToImage: Map<string, ImageDef>;
  warnings: string[];
  /** Current list state per numId. */
  listStack: Array<{ numId: number; level: number; format: string }>;
  /** HTML output buffer. */
  html: string[];
  /** Footnote numbering state. */
  footnoteRefs: Map<number, number>;
  endnoteRefs: Map<number, number>;
  /** Resolved style map (normalized from options). */
  styleMap: Map<string, string>;
  /** Comment definitions indexed by ID. */
  commentMap: Map<number, CommentDef>;
  /** Track active comment ranges (IDs of comments whose range is currently open). */
  activeCommentIds: Set<number>;
  /** Collected aside comments to render at the end of the document. */
  asideComments: Array<{ id: number; def: CommentDef }>;
}

/**
 * Convert a DocxDocument to HTML.
 */
export function renderToHtml(doc: DocxDocument, options?: HtmlRenderOptions): HtmlRenderResult {
  const opts: ResolvedHtmlRenderOptions = {
    includeStyles: options?.includeStyles ?? true,
    fullDocument: options?.fullDocument ?? true,
    title: options?.title ?? doc.coreProperties?.title ?? "Document",
    imageMode: options?.imageMode ?? "dataUrl",
    includeComments: options?.includeComments ?? false,
    includeRevisions: options?.includeRevisions ?? false,
    includeNotes: options?.includeNotes ?? true,
    classPrefix: options?.classPrefix ?? "docx-",
    styleMap: options?.styleMap ?? new Map(),
    chartRenderer: options?.chartRenderer,
    chartExRenderer: options?.chartExRenderer
  };

  // Normalize styleMap to a Map
  const resolvedStyleMap = new Map<string, string>();
  if (opts.styleMap) {
    if (opts.styleMap instanceof Map) {
      for (const [k, v] of opts.styleMap) {
        resolvedStyleMap.set(k, v);
      }
    } else {
      for (const [k, v] of Object.entries(opts.styleMap)) {
        resolvedStyleMap.set(k, v);
      }
    }
  }

  // Build comment map
  const commentMap = new Map<number, CommentDef>();
  if (doc.comments) {
    for (const c of doc.comments) {
      commentMap.set(c.id, c);
    }
  }

  const state: RenderState = {
    options: opts,
    doc,
    imageMap: new Map(),
    rIdToImage: new Map(),
    warnings: [],
    listStack: [],
    html: [],
    footnoteRefs: new Map(),
    endnoteRefs: new Map(),
    styleMap: resolvedStyleMap,
    commentMap,
    activeCommentIds: new Set(),
    asideComments: []
  };

  // Build image map
  if (doc.images) {
    for (const img of doc.images) {
      if (img.rId) {
        state.rIdToImage.set(img.rId, img);
      }
      if (opts.imageMode === "dataUrl") {
        state.imageMap.set(img.fileName, imageToDataUrl(img));
      } else if (opts.imageMode === "filename") {
        state.imageMap.set(img.fileName, img.fileName);
      }
    }
  }

  if (opts.fullDocument) {
    state.html.push("<!DOCTYPE html>");
    state.html.push(`<html lang="en">`);
    state.html.push("<head>");
    state.html.push(`<meta charset="UTF-8">`);
    state.html.push(`<title>${escapeHtml(opts.title)}</title>`);
    if (opts.includeStyles) {
      state.html.push(`<style>${generateCss(opts.classPrefix)}</style>`);
    }
    state.html.push("</head>");
    state.html.push("<body>");
  }

  state.html.push(`<div class="${opts.classPrefix}document">`);

  // Render default header
  if (doc.headers && doc.headers.size > 0) {
    const defaultHeader = doc.headers.get("default");
    if (defaultHeader) {
      state.html.push(`<header class="${opts.classPrefix}header">`);
      for (const child of defaultHeader.content.children) {
        renderBodyContent(state, child as BodyContent);
      }
      closeOpenLists(state);
      state.html.push("</header>");
    }
  }

  // Render body content
  for (const item of doc.body) {
    renderBodyContent(state, item);
  }
  closeOpenLists(state);

  // Render default footer
  if (doc.footers && doc.footers.size > 0) {
    const defaultFooter = doc.footers.get("default");
    if (defaultFooter) {
      state.html.push(`<footer class="${opts.classPrefix}footer">`);
      for (const child of defaultFooter.content.children) {
        renderBodyContent(state, child as BodyContent);
      }
      closeOpenLists(state);
      state.html.push("</footer>");
    }
  }

  // Footnotes
  if (opts.includeNotes && doc.footnotes && doc.footnotes.length > 0) {
    state.html.push(`<hr class="${opts.classPrefix}footnote-separator"/>`);
    state.html.push(`<aside class="${opts.classPrefix}footnotes">`);
    state.html.push("<h2>Footnotes</h2>");
    state.html.push("<ol>");
    for (const note of doc.footnotes) {
      if (note.id <= 0) {
        continue;
      }
      state.html.push(`<li id="footnote-${note.id}">`);
      for (const p of note.content) {
        renderBodyContent(state, p);
      }
      state.html.push("</li>");
    }
    state.html.push("</ol>");
    state.html.push("</aside>");
  }

  // Endnotes
  if (opts.includeNotes && doc.endnotes && doc.endnotes.length > 0) {
    state.html.push(`<hr class="${opts.classPrefix}endnote-separator"/>`);
    state.html.push(`<aside class="${opts.classPrefix}endnotes">`);
    state.html.push("<h2>Endnotes</h2>");
    state.html.push("<ol>");
    for (const note of doc.endnotes) {
      if (note.id <= 0) {
        continue;
      }
      state.html.push(`<li id="endnote-${note.id}">`);
      for (const p of note.content) {
        renderBodyContent(state, p);
      }
      state.html.push("</li>");
    }
    state.html.push("</ol>");
    state.html.push("</aside>");
  }

  // Comments aside
  if (opts.includeComments && state.asideComments.length > 0) {
    state.html.push(`<aside class="${opts.classPrefix}comments">`);
    state.html.push("<h2>Comments</h2>");
    for (const { id, def } of state.asideComments) {
      state.html.push(`<div class="${opts.classPrefix}comment" id="comment-${id}">`);
      state.html.push(
        `<p class="${opts.classPrefix}comment-meta"><strong>${escapeHtml(def.author)}</strong>${def.date ? ` <time>${escapeHtml(def.date)}</time>` : ""}</p>`
      );
      for (const p of def.content) {
        renderParagraph(state, p);
      }
      state.html.push("</div>");
    }
    state.html.push("</aside>");
  }

  state.html.push("</div>");

  if (opts.fullDocument) {
    state.html.push("</body>");
    state.html.push("</html>");
  }

  return {
    html: state.html.join("\n"),
    warnings: state.warnings,
    images: state.imageMap
  };
}

// =============================================================================
// Body Content
// =============================================================================

function renderBodyContent(state: RenderState, item: BodyContent): void {
  if (!("type" in item)) {
    return;
  }
  switch (item.type) {
    case "paragraph":
      renderParagraph(state, item);
      break;
    case "table":
      closeOpenLists(state);
      renderTable(state, item);
      break;
    case "floatingImage":
      closeOpenLists(state);
      renderFloatingImageHtml(state, item);
      break;
    case "textBox":
      closeOpenLists(state);
      renderTextBoxHtml(state, item);
      break;
    case "sdt":
      renderSdtHtml(state, item);
      break;
    case "math":
      closeOpenLists(state);
      renderMathBlockHtml(state, item);
      break;
    case "tableOfContents":
      // Render TOC as placeholder - user would regenerate
      closeOpenLists(state);
      state.html.push(`<nav class="${state.options.classPrefix}toc">`);
      if (item.cachedParagraphs) {
        for (const p of item.cachedParagraphs) {
          renderBodyContent(state, p);
        }
      }
      state.html.push("</nav>");
      break;
    case "drawingShape":
      closeOpenLists(state);
      renderDrawingShapeHtml(state, item);
      break;
    case "chart":
      closeOpenLists(state);
      if (state.options.chartRenderer) {
        state.html.push(`<figure class="${state.options.classPrefix}chart">`);
        if (item.chart.title) {
          state.html.push(`<figcaption>${escapeHtml(item.chart.title)}</figcaption>`);
        }
        state.html.push(state.options.chartRenderer(item.chart));
        state.html.push("</figure>");
      } else {
        state.html.push(`<figure class="${state.options.classPrefix}chart">`);
        if (item.chart.title) {
          state.html.push(`<figcaption>${escapeHtml(item.chart.title)}</figcaption>`);
        }
        state.html.push(
          `<div class="${state.options.classPrefix}chart-placeholder">[${item.chart.type} chart: ${item.chart.series.length} series]</div>`
        );
        state.html.push("</figure>");
      }
      break;
    case "chartEx":
      closeOpenLists(state);
      if (state.options.chartExRenderer) {
        state.html.push(`<figure class="${state.options.classPrefix}chart">`);
        if (item.altText) {
          state.html.push(`<figcaption>${escapeHtml(item.altText)}</figcaption>`);
        }
        state.html.push(state.options.chartExRenderer(item));
        state.html.push("</figure>");
      } else {
        state.html.push(`<figure class="${state.options.classPrefix}chart">`);
        state.html.push(
          `<div class="${state.options.classPrefix}chart-placeholder">[ChartEx${item.altText ? ": " + escapeHtml(item.altText) : ""}]</div>`
        );
        state.html.push("</figure>");
      }
      break;
    case "checkBox":
      closeOpenLists(state);
      state.html.push(`<input type="checkbox" ${item.checked ? "checked" : ""} disabled />`);
      break;
    case "opaqueDrawing":
      // Skip opaque drawings in HTML
      state.warnings.push(`Unsupported element type "opaqueDrawing" skipped`);
      break;
    case "altChunk":
      state.warnings.push(`Unsupported element type "altChunk" skipped`);
      break;
    default: {
      const _type = (item as { type?: string }).type;
      if (_type) {
        state.warnings.push(`Unknown element type "${_type}" skipped`);
      }
      break;
    }
  }
}

// =============================================================================
// Paragraph
// =============================================================================

function renderParagraph(state: RenderState, para: Paragraph): void {
  const props = para.properties;
  const prefix = state.options.classPrefix;

  // Check if this is a list item
  if (props?.numbering) {
    const listInfo = getListInfo(state.doc, props.numbering.numId, props.numbering.level);
    if (listInfo) {
      openListIfNeeded(state, props.numbering.numId, props.numbering.level, listInfo.format);
      state.html.push(`<li>`);
      renderParagraphInline(state, para);
      state.html.push("</li>");
      return;
    }
  }

  closeOpenLists(state);

  // Check styleMap first
  let tag = "p";
  let mappedClass = "";
  let styleMapHit = false;

  if (props?.style && state.styleMap.size > 0) {
    const mapping = state.styleMap.get(props.style);
    if (mapping) {
      styleMapHit = true;
      const parsed = parseStyleMapping(mapping);
      if (parsed.tag) {
        tag = parsed.tag;
      }
      if (parsed.className) {
        mappedClass = parsed.className;
      }
    }
  }

  // If no styleMap hit, determine tag based on style/outline (default logic)
  if (!styleMapHit) {
    if (props?.style) {
      const styleId = props.style.toLowerCase();
      if (styleId === "heading1" || styleId === "heading 1" || styleId === "title") {
        tag = "h1";
      } else if (styleId === "heading2" || styleId === "heading 2") {
        tag = "h2";
      } else if (styleId === "heading3" || styleId === "heading 3") {
        tag = "h3";
      } else if (styleId === "heading4" || styleId === "heading 4") {
        tag = "h4";
      } else if (styleId === "heading5" || styleId === "heading 5") {
        tag = "h5";
      } else if (styleId === "heading6" || styleId === "heading 6") {
        tag = "h6";
      }
    }
    if (props?.outlineLevel !== undefined && props.outlineLevel >= 0 && props.outlineLevel < 6) {
      tag = `h${props.outlineLevel + 1}`;
    }
  }

  // Build inline style
  const styles: string[] = [];
  if (props?.alignment) {
    const alignMap: Record<string, string> = {
      left: "left",
      right: "right",
      center: "center",
      both: "justify",
      justify: "justify",
      distribute: "justify"
    };
    const ta = alignMap[props.alignment];
    if (ta) {
      styles.push(`text-align:${ta}`);
    }
  }
  if (props?.indent) {
    if (props.indent.left !== undefined) {
      styles.push(`margin-left:${twipsToPx(props.indent.left)}px`);
    }
    if (props.indent.right !== undefined) {
      styles.push(`margin-right:${twipsToPx(props.indent.right)}px`);
    }
    if (props.indent.firstLine !== undefined) {
      styles.push(`text-indent:${twipsToPx(props.indent.firstLine)}px`);
    }
    if (props.indent.hanging !== undefined) {
      styles.push(`text-indent:-${twipsToPx(props.indent.hanging)}px`);
    }
  }
  if (props?.spacing) {
    if (props.spacing.before !== undefined) {
      styles.push(`margin-top:${twipsToPx(props.spacing.before)}px`);
    }
    if (props.spacing.after !== undefined) {
      styles.push(`margin-bottom:${twipsToPx(props.spacing.after)}px`);
    }
    if (props.spacing.line !== undefined) {
      styles.push(`line-height:${props.spacing.line / 240}`);
    }
  }
  if (props?.shading?.fill) {
    styles.push(`background-color:#${props.shading.fill}`);
  }
  if (props?.bidi) {
    styles.push(`direction:rtl`);
  }
  if (props?.pageBreakBefore) {
    styles.push(`page-break-before:always`);
  }
  if (props?.keepNext) {
    styles.push(`page-break-after:avoid`);
  }
  if (props?.keepLines) {
    styles.push(`page-break-inside:avoid`);
  }
  if (props?.borders) {
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const border = props.borders[side];
      if (border) {
        const bstyle = borderToCss(border);
        if (bstyle) {
          styles.push(`border-${side}:${bstyle}`);
        }
      }
    }
  }

  const styleAttr = styles.length > 0 ? ` style="${styles.join(";")}"` : "";

  // Build class attribute
  let classAttr = "";
  if (styleMapHit && mappedClass) {
    classAttr = ` class="${mappedClass}"`;
  } else if (props?.style) {
    classAttr = ` class="${prefix}style-${escapeClassName(props.style)}"`;
  }

  state.html.push(`<${tag}${classAttr}${styleAttr}>`);
  renderParagraphInline(state, para);
  state.html.push(`</${tag}>`);
}

function renderParagraphInline(state: RenderState, para: Paragraph): void {
  for (const child of para.children) {
    renderParagraphChild(state, child);
  }
}

function renderParagraphChild(state: RenderState, child: ParagraphChild): void {
  if ("type" in child) {
    switch (child.type) {
      case "hyperlink":
        renderHyperlinkHtml(state, child);
        return;
      case "bookmarkStart":
        state.html.push(`<a id="bookmark-${child.id}" data-name="${escapeHtml(child.name)}"></a>`);
        return;
      case "bookmarkEnd":
        return;
      case "commentRangeStart":
        if (state.options.includeComments) {
          state.activeCommentIds.add(child.id);
          state.html.push(
            `<mark class="${state.options.classPrefix}comment-highlight" data-comment-id="${child.id}">`
          );
        }
        return;
      case "commentRangeEnd":
        if (state.options.includeComments) {
          state.activeCommentIds.delete(child.id);
          state.html.push("</mark>");
          // Register this comment for aside rendering
          const def = state.commentMap.get(child.id);
          if (def) {
            state.asideComments.push({ id: child.id, def });
          }
        }
        return;
      case "commentReference":
        if (state.options.includeComments) {
          state.html.push(
            `<a href="#comment-${child.id}" class="${state.options.classPrefix}comment-ref"><sup>[${child.id}]</sup></a>`
          );
        }
        return;
      case "insertedRun":
        if (state.options.includeRevisions || state.options.includeComments) {
          state.html.push(
            `<ins data-author="${escapeHtml(child.revision.author)}" data-date="${escapeHtml(child.revision.date ?? "")}">`
          );
          renderRun(state, child.run);
          state.html.push("</ins>");
        } else {
          // When revisions are not shown, render the inserted content normally
          renderRun(state, child.run);
        }
        return;
      case "deletedRun":
        if (state.options.includeRevisions || state.options.includeComments) {
          state.html.push(
            `<del data-author="${escapeHtml(child.revision.author)}" data-date="${escapeHtml(child.revision.date ?? "")}">`
          );
          renderRun(state, child.run);
          state.html.push("</del>");
        }
        // When revisions are not shown, deleted content is simply omitted
        return;
      case "movedFromRun":
        if (state.options.includeRevisions || state.options.includeComments) {
          state.html.push(`<del class="${state.options.classPrefix}move-from">`);
          renderRun(state, child.run);
          state.html.push("</del>");
        }
        return;
      case "movedToRun":
        if (state.options.includeRevisions || state.options.includeComments) {
          state.html.push(`<ins class="${state.options.classPrefix}move-to">`);
          renderRun(state, child.run);
          state.html.push("</ins>");
        } else {
          renderRun(state, child.run);
        }
        return;
    }
  }
  // Run (no type discriminator)
  if (isRun(child)) {
    renderRun(state, child);
  }
}

function renderHyperlinkHtml(state: RenderState, link: Hyperlink): void {
  let href = "";
  if (link.url) {
    // Reject javascript:/vbscript:/etc. URLs that could XSS the renderer
    // output. Falls back to no href so the anchor still preserves the visible
    // text.
    href = sanitizeUrl(link.url) ?? "";
  } else if (link.anchor) {
    href = `#bookmark-${link.anchor}`;
  }
  const attrs: string[] = [];
  if (href) {
    attrs.push(`href="${escapeHtml(href)}"`);
  }
  if (link.tooltip) {
    attrs.push(`title="${escapeHtml(link.tooltip)}"`);
  }
  state.html.push(`<a ${attrs.join(" ")}>`);
  for (const ch of link.children) {
    renderParagraphChild(state, ch);
  }
  state.html.push("</a>");
}

// =============================================================================
// Run
// =============================================================================

function renderRun(state: RenderState, run: Run): void {
  const rPr = run.properties;

  // Build tag stack based on formatting
  const tags: string[] = [];
  const styles: string[] = [];

  if (rPr) {
    if (rPr.bold) {
      tags.push("strong");
    }
    if (rPr.italic) {
      tags.push("em");
    }
    if (rPr.underline) {
      const uStyle = typeof rPr.underline === "object" ? rPr.underline.style : rPr.underline;
      if (uStyle !== "none") {
        tags.push("u");
      }
    }
    if (rPr.strike) {
      tags.push("s");
    }
    if (rPr.vertAlign === "superscript") {
      tags.push("sup");
    } else if (rPr.vertAlign === "subscript") {
      tags.push("sub");
    }
    if (rPr.smallCaps) {
      styles.push("font-variant:small-caps");
    }
    if (rPr.caps) {
      styles.push("text-transform:uppercase");
    }
    if (rPr.color) {
      const color = resolveColor(rPr.color, state.doc.theme);
      if (color) {
        styles.push(`color:#${color}`);
      }
    }
    if (rPr.highlight && rPr.highlight !== "none") {
      styles.push(`background-color:${highlightToColor(rPr.highlight)}`);
    }
    if (rPr.shading?.fill) {
      styles.push(`background-color:#${rPr.shading.fill}`);
    }
    if (rPr.size !== undefined) {
      // size is half-points
      styles.push(`font-size:${rPr.size / 2}pt`);
    }
    if (rPr.font) {
      const fontName =
        typeof rPr.font === "string" ? rPr.font : (rPr.font.ascii ?? rPr.font.eastAsia);
      if (fontName) {
        // The value is interpolated into a `style="..."` attribute, so
        // wrap font names containing whitespace or punctuation in
        // single quotes — embedding double quotes here would close the
        // surrounding HTML attribute and produce invalid markup.
        styles.push(`font-family:'${fontName.replace(/'/g, "")}'`);
      }
    }
    if (rPr.doubleStrike) {
      styles.push("text-decoration:line-through;text-decoration-style:double");
    }
    if (rPr.shadow) {
      styles.push("text-shadow:1px 1px 2px rgba(0,0,0,0.3)");
    }
    if (rPr.outline) {
      styles.push("-webkit-text-stroke:1px currentColor;color:transparent");
    }
    if (rPr.emboss) {
      styles.push("text-shadow:-1px -1px 0 rgba(255,255,255,0.6), 1px 1px 0 rgba(0,0,0,0.3)");
    }
    if (rPr.imprint) {
      styles.push("text-shadow:1px 1px 0 rgba(255,255,255,0.6), -1px -1px 0 rgba(0,0,0,0.3)");
    }
    if (rPr.vanish) {
      styles.push("display:none");
    }
    if (rPr.spacing !== undefined) {
      styles.push(`letter-spacing:${rPr.spacing / 20}pt`);
    }
    if (rPr.emphasisMark && rPr.emphasisMark !== "none") {
      const markMap: Record<string, string> = {
        dot: "filled",
        comma: "'\\u3001'",
        circle: "open",
        dotBelow: "filled under"
      };
      const em = markMap[rPr.emphasisMark] ?? "filled";
      styles.push(`text-emphasis:${em}`);
    }
    if (rPr.border) {
      const bstyle = borderToCss(rPr.border);
      if (bstyle) {
        styles.push(`border:${bstyle}`);
      }
    }
  }

  // Outermost <span> if styles, inner tags for semantic markup
  if (styles.length > 0) {
    state.html.push(`<span style="${styles.join(";")}">`);
  }
  for (const t of tags) {
    state.html.push(`<${t}>`);
  }

  // Render run content
  for (const content of run.content) {
    renderRunContentHtml(state, content);
  }

  // Close tags in reverse
  for (let i = tags.length - 1; i >= 0; i--) {
    state.html.push(`</${tags[i]}>`);
  }
  if (styles.length > 0) {
    state.html.push("</span>");
  }
}

function renderRunContentHtml(state: RenderState, content: RunContent): void {
  switch (content.type) {
    case "text":
      state.html.push(escapeHtml(content.text));
      break;
    case "break":
      if (content.breakType === "page") {
        state.html.push(
          `<span class="${state.options.classPrefix}page-break" style="page-break-before:always"></span>`
        );
      } else if (content.breakType === "column") {
        state.html.push(`<span class="${state.options.classPrefix}column-break"></span>`);
      } else {
        state.html.push("<br>");
      }
      break;
    case "tab":
      state.html.push('<span class="docx-tab" style="display:inline-block;min-width:2em"></span>');
      break;
    case "ptab":
      state.html.push('<span class="docx-ptab"></span>');
      break;
    case "carriageReturn":
      state.html.push("<br>");
      break;
    case "noBreakHyphen":
      state.html.push("\u2011");
      break;
    case "softHyphen":
      state.html.push("\u00AD");
      break;
    case "symbol":
      // Convert hex char code to Unicode
      try {
        const code = parseInt(content.char, 16);
        state.html.push(
          `<span style="font-family:'${escapeHtml(content.font).replace(/&#39;/g, "")}'">${String.fromCodePoint(code)}</span>`
        );
      } catch {
        state.html.push(escapeHtml(content.char));
      }
      break;
    case "footnoteRef": {
      const num = state.footnoteRefs.size + 1;
      state.footnoteRefs.set(content.id, num);
      state.html.push(
        `<sup><a href="#footnote-${content.id}" id="footnote-ref-${content.id}">${num}</a></sup>`
      );
      break;
    }
    case "endnoteRef": {
      const num = state.endnoteRefs.size + 1;
      state.endnoteRefs.set(content.id, num);
      state.html.push(
        `<sup><a href="#endnote-${content.id}" id="endnote-ref-${content.id}">${num}</a></sup>`
      );
      break;
    }
    case "image":
      renderInlineImageHtml(state, content);
      break;
    case "field":
      // Use cached value if present
      if (content.cachedValue) {
        state.html.push(escapeHtml(content.cachedValue));
      }
      break;
    case "ruby":
      state.html.push("<ruby>");
      for (const r of content.baseText) {
        renderRun(state, r);
      }
      state.html.push("<rt>");
      for (const r of content.rubyText) {
        renderRun(state, r);
      }
      state.html.push("</rt></ruby>");
      break;
    case "lastRenderedPageBreak":
    case "annotationReference":
      break;
  }
}

function renderInlineImageHtml(state: RenderState, img: InlineImageContent): void {
  const imgDef = state.rIdToImage.get(img.rId);
  if (!imgDef) {
    state.warnings.push(`Image rId ${img.rId} not found`);
    return;
  }
  const src = state.imageMap.get(imgDef.fileName) ?? imgDef.fileName;
  const w = emuToPx(img.width);
  const h = emuToPx(img.height);
  state.html.push(
    `<img src="${escapeHtml(src)}" width="${w}" height="${h}"${img.altText ? ` alt="${escapeHtml(img.altText)}"` : ""}/>`
  );
}

function renderFloatingImageHtml(state: RenderState, img: FloatingImage): void {
  const imgDef = state.rIdToImage.get(img.rId);
  if (!imgDef) {
    return;
  }
  const src = state.imageMap.get(imgDef.fileName) ?? imgDef.fileName;
  const w = emuToPx(img.width);
  const h = emuToPx(img.height);
  const style =
    img.wrap?.style === "square" ? "float:right;margin:10px" : "display:block;margin:10px 0";
  state.html.push(
    `<img src="${escapeHtml(src)}" width="${w}" height="${h}" style="${style}"${img.altText ? ` alt="${escapeHtml(img.altText)}"` : ""}/>`
  );
}

// =============================================================================
// Table
// =============================================================================

function renderTable(state: RenderState, table: Table): void {
  const prefix = state.options.classPrefix;
  state.html.push(
    `<table class="${prefix}table" border="1" style="border-collapse:collapse;width:100%">`
  );

  // Pre-compute rowspan map: key = "rowIdx,colIdx" -> rowspan count
  const rowspanMap = new Map<string, number>();
  const skipCells = new Set<string>();

  for (let rowIdx = 0; rowIdx < table.rows.length; rowIdx++) {
    const row = table.rows[rowIdx];
    let colIdx = 0;
    for (const cell of row.cells) {
      // Skip columns that are occupied by a previous rowspan
      while (skipCells.has(`${rowIdx},${colIdx}`)) {
        colIdx++;
      }
      const gridSpan = cell.properties?.gridSpan ?? 1;
      if (cell.properties?.verticalMerge === "restart") {
        // Count how many subsequent rows continue this merge at the same column
        let span = 1;
        for (let r = rowIdx + 1; r < table.rows.length; r++) {
          const targetCell = getCellAtGridCol(table.rows[r], colIdx);
          if (targetCell?.properties?.verticalMerge === "continue") {
            span++;
            // Mark those cells as skip
            for (let c = colIdx; c < colIdx + gridSpan; c++) {
              skipCells.add(`${r},${c}`);
            }
          } else {
            break;
          }
        }
        if (span > 1) {
          rowspanMap.set(`${rowIdx},${colIdx}`, span);
        }
      }
      colIdx += gridSpan;
    }
  }

  for (let rowIdx = 0; rowIdx < table.rows.length; rowIdx++) {
    renderTableRowWithMerge(state, table.rows[rowIdx], rowIdx, rowspanMap, skipCells);
  }
  state.html.push("</table>");
}

/** Get the cell at a specific grid column index within a row. */
function getCellAtGridCol(row: TableRow, targetCol: number): TableCell | undefined {
  let colIdx = 0;
  for (const cell of row.cells) {
    if (colIdx === targetCol) {
      return cell;
    }
    colIdx += cell.properties?.gridSpan ?? 1;
  }
  return undefined;
}

function renderTableRowWithMerge(
  state: RenderState,
  row: TableRow,
  rowIdx: number,
  rowspanMap: Map<string, number>,
  skipCells: Set<string>
): void {
  const isHeader = row.properties?.tableHeader;
  state.html.push(`<tr>`);
  let colIdx = 0;
  for (const cell of row.cells) {
    // Skip columns occupied by previous rowspan
    while (skipCells.has(`${rowIdx},${colIdx}`)) {
      colIdx++;
    }
    const gridSpan = cell.properties?.gridSpan ?? 1;
    // If this cell is a "continue" cell, skip it (already merged into a rowspan above)
    if (cell.properties?.verticalMerge === "continue") {
      colIdx += gridSpan;
      continue;
    }
    const rowspan = rowspanMap.get(`${rowIdx},${colIdx}`);
    renderTableCellWithRowspan(state, cell, isHeader ?? false, rowspan);
    colIdx += gridSpan;
  }
  state.html.push("</tr>");
}

function renderTableCellWithRowspan(
  state: RenderState,
  cell: TableCell,
  isHeader: boolean,
  rowspan?: number
): void {
  const tag = isHeader ? "th" : "td";
  const attrs: string[] = [];
  const props = cell.properties;

  if (props?.gridSpan && props.gridSpan > 1) {
    attrs.push(`colspan="${props.gridSpan}"`);
  }
  if (rowspan && rowspan > 1) {
    attrs.push(`rowspan="${rowspan}"`);
  }

  const styles: string[] = [];
  if (props?.shading?.fill) {
    styles.push(`background-color:#${props.shading.fill}`);
  }
  if (props?.width) {
    if (props.width.type === "dxa") {
      styles.push(`width:${twipsToPx(props.width.value)}px`);
    } else if (props.width.type === "pct") {
      styles.push(`width:${props.width.value / 50}%`);
    }
  }
  if (props?.borders) {
    for (const side of ["top", "right", "bottom", "left"] as const) {
      const border = props.borders[side];
      if (border) {
        const bstyle = borderToCss(border);
        if (bstyle) {
          styles.push(`border-${side}:${bstyle}`);
        }
      }
    }
  }
  if (props?.verticalAlign) {
    const vaMap: Record<string, string> = {
      top: "top",
      center: "middle",
      bottom: "bottom"
    };
    const va = vaMap[props.verticalAlign];
    if (va) {
      styles.push(`vertical-align:${va}`);
    }
  }

  if (styles.length > 0) {
    attrs.push(`style="${styles.join(";")}"`);
  }

  state.html.push(`<${tag}${attrs.length > 0 ? " " + attrs.join(" ") : ""}>`);
  for (const item of cell.content) {
    renderBodyContent(state, item as BodyContent);
  }
  closeOpenLists(state);
  state.html.push(`</${tag}>`);
}

// =============================================================================
// TextBox
// =============================================================================

function renderTextBoxHtml(state: RenderState, textBox: TextBox): void {
  state.html.push(`<div class="textbox">`);
  for (const p of textBox.content) {
    renderParagraph(state, p);
  }
  state.html.push("</div>");
}

// =============================================================================
// Drawing Shape
// =============================================================================

function renderDrawingShapeHtml(
  state: RenderState,
  item: BodyContent & { type: "drawingShape" }
): void {
  const shape = item as DrawingShape;
  const styles: string[] = [];
  styles.push(`display:inline-block`);
  if (shape.width) {
    styles.push(`width:${emuToPx(shape.width)}px`);
  }
  if (shape.height) {
    styles.push(`height:${emuToPx(shape.height)}px`);
  }
  if (shape.fillColor && !shape.noFill) {
    styles.push(`background-color:#${shape.fillColor}`);
  }
  if (!shape.noOutline) {
    const outlineColor = shape.outlineColor ? `#${shape.outlineColor}` : "#000";
    const outlineWidth = shape.outlineWidth ? `${emuToPx(shape.outlineWidth)}px` : "1px";
    styles.push(`border:${outlineWidth} solid ${outlineColor}`);
  }

  const prefix = state.options.classPrefix;
  state.html.push(`<div class="${prefix}shape" style="${styles.join(";")}">`);
  if (shape.textContent && shape.textContent.length > 0) {
    for (const p of shape.textContent) {
      renderParagraph(state, p);
    }
  }
  state.html.push("</div>");
}

// =============================================================================
// SDT
// =============================================================================

function renderSdtHtml(state: RenderState, sdt: StructuredDocumentTag): void {
  // Check if this is a checkbox SDT
  if (sdt.properties?.checkbox) {
    const checked = sdt.properties.checkbox.checked;
    state.html.push(`<input type="checkbox"${checked ? " checked" : ""} disabled />`);
    return;
  }

  for (const child of sdt.content) {
    if ("type" in child) {
      if (child.type === "paragraph") {
        renderParagraph(state, child);
      } else if (child.type === "table") {
        renderTable(state, child);
      }
    } else if (isRun(child)) {
      // Run
      state.html.push(`<span class="${state.options.classPrefix}sdt">`);
      renderRun(state, child);
      state.html.push("</span>");
    }
  }
}

// =============================================================================
// List Management
// =============================================================================

function getListInfo(
  doc: DocxDocument,
  numId: number,
  level: number
): { format: string } | undefined {
  const instance = doc.numberingInstances?.find(n => n.numId === numId);
  if (!instance) {
    return undefined;
  }
  const abstractNum = doc.abstractNumberings?.find(a => a.abstractNumId === instance.abstractNumId);
  if (!abstractNum) {
    return undefined;
  }
  const levelDef = abstractNum.levels.find(l => l.level === level);
  if (!levelDef) {
    return undefined;
  }
  return { format: levelDef.format };
}

function openListIfNeeded(state: RenderState, numId: number, level: number, format: string): void {
  // Close lists that are too deep
  while (state.listStack.length > 0) {
    const top = state.listStack[state.listStack.length - 1];
    if (top.level > level || (top.level === level && top.numId !== numId)) {
      state.html.push(top.format === "bullet" ? "</ul>" : "</ol>");
      state.listStack.pop();
    } else {
      break;
    }
  }
  // Open new lists if needed
  const currentLevel =
    state.listStack.length > 0 ? state.listStack[state.listStack.length - 1].level : -1;
  if (currentLevel < level || state.listStack.length === 0) {
    while (
      state.listStack.length === 0 ||
      state.listStack[state.listStack.length - 1].level < level
    ) {
      const newLevel =
        state.listStack.length === 0 ? 0 : state.listStack[state.listStack.length - 1].level + 1;
      const tag = format === "bullet" ? "ul" : "ol";
      state.html.push(`<${tag}>`);
      state.listStack.push({ numId, level: newLevel, format });
    }
  }
}

function closeOpenLists(state: RenderState): void {
  while (state.listStack.length > 0) {
    const top = state.listStack.pop()!;
    state.html.push(top.format === "bullet" ? "</ul>" : "</ol>");
  }
}

// =============================================================================
// Math
// =============================================================================

function renderMathBlockHtml(state: RenderState, block: MathBlock): void {
  const text = extractMathText(block.content);
  state.html.push(`<span class="${state.options.classPrefix}math">${escapeHtml(text)}</span>`);
}

// =============================================================================
// Helpers
// =============================================================================

function escapeHtml(s: string): string {
  // Strip characters that are illegal in XML 1.0 / serialised markup —
  // forbidden C0 controls (`0x00`-`0x08`, `0x0B`, `0x0C`, `0x0E`-`0x1F`),
  // DEL (`0x7F`), lone UTF-16 surrogate halves, and the `0xFFFE` / `0xFFFF`
  // noncharacters — before escaping entities. The input is arbitrary docx
  // content (titles, authors, alt text, body runs), so a corrupt or hostile
  // document could otherwise inject control chars / lone surrogates straight
  // into HTML text and attribute values, producing invalid UTF-8 on encode
  // and U+FFFD / parser hiccups in the browser.
  //
  // We keep the HTML-flavoured entity set (`'` → `&#39;`, NOT `&apos;`) and
  // do not route through `xmlEncode`: `&apos;` is not a predefined HTML
  // entity and is unreliable in legacy HTML contexts, so HTML output must
  // use the numeric reference. Only the illegal-char stripping is shared.
  return stripXmlIllegalChars(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeClassName(s: string): string {
  return s.replace(/[^a-zA-Z0-9_-]/g, "-");
}

function twipsToPx(twips: number): number {
  // 1 twip = 1/1440 inch, assuming 96 DPI => 1 inch = 96 px
  return Math.round((twips / 1440) * 96);
}

function emuToPx(emu: number): number {
  // 96 DPI assumed
  return Math.round((emu / EMU_PER_INCH) * 96);
}

function imageToDataUrl(img: ImageDef): string {
  const mime: Record<string, string> = {
    png: "image/png",
    jpeg: "image/jpeg",
    jpg: "image/jpeg",
    gif: "image/gif",
    bmp: "image/bmp",
    svg: "image/svg+xml",
    webp: "image/webp",
    tiff: "image/tiff"
  };
  const mediaType = mime[img.mediaType] ?? "application/octet-stream";
  const base64 = bytesToBase64(img.data);
  return `data:${mediaType};base64,${base64}`;
}

function resolveColor(
  color: string | ColorSpec | undefined,
  theme: DocumentTheme | undefined
): string | undefined {
  if (typeof color === "string") {
    return color === "auto" ? undefined : color;
  }
  if (typeof color === "object" && color !== null) {
    if (typeof color.val === "string") {
      return color.val === "auto" ? undefined : color.val;
    }
    if (color.themeColor && theme) {
      return resolveThemeColor(color, theme);
    }
  }
  return undefined;
}

function highlightToColor(highlight: string): string {
  const map: Record<string, string> = {
    black: "#000000",
    blue: "#0000FF",
    cyan: "#00FFFF",
    green: "#00FF00",
    magenta: "#FF00FF",
    red: "#FF0000",
    yellow: "#FFFF00",
    white: "#FFFFFF",
    darkBlue: "#000080",
    darkCyan: "#008080",
    darkGreen: "#008000",
    darkMagenta: "#800080",
    darkRed: "#800000",
    darkYellow: "#808000",
    darkGray: "#808080",
    lightGray: "#C0C0C0"
  };
  return map[highlight] ?? "transparent";
}

function borderToCss(border: Border): string | undefined {
  if (border.style === "none" || border.style === "nil") {
    return "none";
  }
  const sizePt = border.size !== undefined ? border.size / 8 : 0.5;
  const styleMap: Record<string, string> = {
    single: "solid",
    double: "double",
    dashed: "dashed",
    dotted: "dotted",
    thick: "solid"
  };
  const cssStyle = styleMap[border.style] ?? "solid";
  const color = border.color === "auto" || !border.color ? "#000" : `#${border.color}`;
  return `${sizePt}pt ${cssStyle} ${color}`;
}

/**
 * Parse a style mapping value like "h1.doc-heading" into tag and className.
 * Supported formats:
 *   "h1"           → { tag: "h1", className: "" }
 *   "h1.cls"       → { tag: "h1", className: "cls" }
 *   ".cls"         → { tag: "", className: "cls" } (uses default tag)
 *   "blockquote"   → { tag: "blockquote", className: "" }
 */
function parseStyleMapping(mapping: string): { tag: string; className: string } {
  const dotIdx = mapping.indexOf(".");
  if (dotIdx === -1) {
    return { tag: mapping, className: "" };
  }
  if (dotIdx === 0) {
    return { tag: "", className: mapping.slice(1) };
  }
  return { tag: mapping.slice(0, dotIdx), className: mapping.slice(dotIdx + 1) };
}

function generateCss(prefix: string): string {
  return `
.${prefix}document { font-family: 'Calibri', 'Segoe UI', sans-serif; line-height: 1.15; max-width: 8.5in; margin: 1in auto; padding: 0 1in; }
.${prefix}document p { margin: 0 0 8pt 0; }
.${prefix}document h1, .${prefix}document h2, .${prefix}document h3 { margin-top: 12pt; margin-bottom: 8pt; font-weight: 600; }
.${prefix}table { border-collapse: collapse; margin: 8pt 0; }
.${prefix}table th, .${prefix}table td { padding: 4pt 8pt; }
.${prefix}textbox { border: 1px solid #ccc; padding: 8pt; margin: 8pt 0; }
.${prefix}footnotes, .${prefix}endnotes { font-size: 0.9em; margin-top: 2em; }
.${prefix}footnote-separator, .${prefix}endnote-separator { width: 30%; margin-left: 0; border: none; border-top: 1px solid #000; margin-top: 2em; }
.${prefix}tab { display: inline-block; min-width: 36pt; }
.${prefix}chart { margin: 12pt 0; }
.${prefix}move-from { text-decoration: line-through; color: #a00; }
.${prefix}move-to { text-decoration: underline; color: #0a0; }
.${prefix}comment-highlight { background-color: #fff3cd; }
.${prefix}comment-ref { text-decoration: none; color: #b86800; }
.${prefix}comments { margin-top: 2em; border-top: 1px solid #ccc; padding-top: 1em; }
.${prefix}comment { margin-bottom: 1em; padding: 0.5em; border-left: 3px solid #b86800; padding-left: 1em; }
.${prefix}comment-meta { font-size: 0.85em; color: #666; margin-bottom: 0.25em; }
ins { text-decoration: underline; color: #060; }
del { text-decoration: line-through; color: #600; }
ruby rt { font-size: 0.5em; }
  `.trim();
}
