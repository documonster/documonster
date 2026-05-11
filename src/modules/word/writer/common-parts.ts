/**
 * DOCX Module - Common Auxiliary Parts Builder
 *
 * Shared logic for generating auxiliary XML parts (styles, settings, numbering,
 * fontTable, theme, core properties, app properties, custom properties).
 *
 * Both the buffered packager and streaming writer use this to avoid duplicating
 * the same rendering logic.
 */

import { XmlWriter } from "@xml/writer";

import { PartPath } from "../constants";
import type {
  AbstractNumbering,
  AppProperties,
  CoreProperties,
  CustomProperty,
  DocDefaults,
  DocumentSettings,
  DocumentTheme,
  EndnoteDef,
  FontDef,
  FootnoteDef,
  NumPicBullet,
  NumberingInstance,
  StyleDef
} from "../types";
import { renderFootnotes, renderEndnotes } from "./footnote-writer";
import { renderNumbering } from "./numbering-writer";
import {
  renderSettings,
  renderFontTable,
  renderCoreProperties,
  renderAppProperties,
  renderCustomProperties,
  renderTheme
} from "./parts-writer";
import { renderStyles } from "./styles-writer";

// =============================================================================
// Types
// =============================================================================

/** Input options for building common auxiliary parts. */
export interface AuxiliaryPartsInput {
  readonly docDefaults?: DocDefaults;
  readonly styles?: readonly StyleDef[];
  readonly settings?: DocumentSettings;
  readonly fonts?: readonly FontDef[];
  readonly theme?: DocumentTheme;
  readonly abstractNumberings?: readonly AbstractNumbering[];
  readonly numberingInstances?: readonly NumberingInstance[];
  readonly numPicBullets?: readonly NumPicBullet[];
  readonly coreProperties?: CoreProperties;
  readonly appProperties?: AppProperties;
  readonly customProperties?: readonly CustomProperty[];
  readonly footnotes?: readonly FootnoteDef[];
  readonly endnotes?: readonly EndnoteDef[];
}

/** A rendered XML part ready to be written to ZIP. */
export interface RenderedPart {
  readonly path: string;
  readonly content: string;
}

// =============================================================================
// Public API
// =============================================================================

/**
 * Build all common auxiliary XML parts that are shared between
 * the buffered packager and streaming writer.
 *
 * Returns an array of rendered parts (path + XML string).
 * Does NOT include: document.xml, images, headers/footers, comments,
 * charts, custom XML, embedded fonts, opaque parts, watermarks.
 * Those are handled by the respective packager/streaming-writer because
 * they require relationship ID coordination.
 */
export function buildCommonAuxiliaryParts(input: AuxiliaryPartsInput): RenderedPart[] {
  const parts: RenderedPart[] = [];

  // styles.xml (always written)
  parts.push({
    path: PartPath.Styles,
    content: renderXml(xml => renderStyles(xml, input.docDefaults, input.styles))
  });

  // settings.xml
  parts.push({
    path: PartPath.Settings,
    content: renderXml(xml => renderSettings(xml, input.settings))
  });

  // fontTable.xml
  parts.push({
    path: PartPath.FontTable,
    content: renderXml(xml => renderFontTable(xml, input.fonts))
  });

  // theme/theme1.xml
  parts.push({
    path: PartPath.Theme,
    content: renderXml(xml => renderTheme(xml, input.theme))
  });

  // numbering.xml (conditional)
  const hasNumbering =
    (input.abstractNumberings && input.abstractNumberings.length > 0) ||
    (input.numberingInstances && input.numberingInstances.length > 0);
  if (hasNumbering) {
    parts.push({
      path: PartPath.Numbering,
      content: renderXml(xml =>
        renderNumbering(
          xml,
          input.abstractNumberings,
          input.numberingInstances,
          input.numPicBullets
        )
      )
    });
  }

  // footnotes.xml (conditional)
  if (input.footnotes && input.footnotes.length > 0) {
    parts.push({
      path: PartPath.Footnotes,
      content: renderXml(xml => renderFootnotes(xml, input.footnotes!))
    });
  }

  // endnotes.xml (conditional)
  if (input.endnotes && input.endnotes.length > 0) {
    parts.push({
      path: PartPath.Endnotes,
      content: renderXml(xml => renderEndnotes(xml, input.endnotes!))
    });
  }

  // docProps/core.xml
  parts.push({
    path: PartPath.CoreProps,
    content: renderXml(xml => renderCoreProperties(xml, input.coreProperties))
  });

  // docProps/app.xml
  parts.push({
    path: PartPath.AppProps,
    content: renderXml(xml => renderAppProperties(xml, input.appProperties))
  });

  // docProps/custom.xml (conditional)
  if (input.customProperties && input.customProperties.length > 0) {
    parts.push({
      path: PartPath.CustomProps,
      content: renderXml(xml => renderCustomProperties(xml, input.customProperties!))
    });
  }

  return parts;
}

// =============================================================================
// Internal
// =============================================================================

/** Render XML to string using XmlWriter. */
function renderXml(renderFn: (xml: XmlWriter) => void): string {
  const writer = new XmlWriter();
  renderFn(writer);
  return writer.xml;
}
