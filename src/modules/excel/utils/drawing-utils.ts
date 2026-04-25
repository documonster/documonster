/**
 * Shared utilities for building drawing models (anchors + relationships)
 * used by both the streaming WorksheetWriter and the non-streaming WorkSheetXform.
 *
 * This eliminates the duplicated anchor/rel building logic and provides
 * a single, correct image-rel deduplication strategy.
 */

import { mediaRelTargetFromRels } from "@excel/utils/ooxml-paths";
import { RelType } from "@excel/xlsx/rel-type";

// =============================================================================
// Types
// =============================================================================

export interface DrawingAnchor {
  picture: {
    rId: string;
    hyperlinks?: { tooltip?: string; rId: string };
    /** Alpha modulation for transparency (OOXML percentage, e.g. 15000 = 15%). */
    alphaModFix?: number;
  };
  range: any;
}

export interface DrawingRel {
  Id: string;
  Type: string;
  Target: string;
  TargetMode?: string;
}

export interface DrawingModel {
  anchors: DrawingAnchor[];
  rels: DrawingRel[];
}

interface ImageMedium {
  imageId: string | number;
  range: any;
  hyperlinks?: { hyperlink?: string; tooltip?: string };
  /** Opacity 0-1 for watermark overlay mode. */
  opacity?: number;
}

/**
 * Resolves a media filename into the drawing-level relative target path.
 *
 * In the non-streaming path, media entries have separate `name` and `extension`
 * fields (e.g. name="image0", extension="png").
 * In the streaming path, `name` already includes the extension (e.g. "image0.png").
 *
 * This function accepts both forms and returns e.g. `"../media/image0.png"`.
 */
export function resolveMediaTarget(medium: { name?: string; extension?: string }): string {
  // When name already contains the extension (streaming path), use it directly.
  // Otherwise concatenate name + extension (non-streaming path).
  // Note: name may be undefined in the non-streaming path; we preserve the legacy
  // behavior of `${undefined}.${ext}` = "undefined.ext" to match addMedia().
  const filename =
    medium.name && medium.extension && medium.name.endsWith(`.${medium.extension}`)
      ? medium.name
      : `${medium.name}.${medium.extension}`;
  return mediaRelTargetFromRels(filename);
}

// =============================================================================
// Anchor / Rel Building
// =============================================================================

/** Options for {@link buildDrawingAnchorsAndRels}. */
interface BuildDrawingOptions {
  /** Look up a book-level image by its id. Return `undefined` if not found. */
  getBookImage: (imageId: string | number) => { name?: string; extension?: string } | undefined;

  /** Generate the next unique rId string for the drawing rels. */
  nextRId: (rels: DrawingRel[]) => string;
}

/**
 * Build the drawing anchors and relationships from a list of image media entries.
 *
 * This is the core logic shared between:
 * - `WorksheetWriter._writeDrawing()` (streaming)
 * - `WorkSheetXform.prepare()` (non-streaming)
 *
 * It correctly deduplicates image rels: if the same `imageId` is used for
 * multiple anchors, only one image relationship is created and shared.
 */
export function buildDrawingAnchorsAndRels(
  media: ImageMedium[],
  existingRels: DrawingRel[],
  options: BuildDrawingOptions
): DrawingModel {
  const anchors: DrawingAnchor[] = [];
  const rels: DrawingRel[] = [...existingRels];

  // Map imageId → rId for deduplication (handles non-consecutive duplicates correctly)
  const imageRIdMap: Record<string, string> = {};

  for (const medium of media) {
    const imageId = String(medium.imageId);
    const bookImage = options.getBookImage(medium.imageId);
    if (!bookImage) {
      continue;
    }

    // Deduplicate: reuse rId if same imageId already has a drawing rel
    let rIdImage = imageRIdMap[imageId];
    if (!rIdImage) {
      rIdImage = options.nextRId(rels);
      imageRIdMap[imageId] = rIdImage;
      rels.push({
        Id: rIdImage,
        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        Target: resolveMediaTarget(bookImage)
      });
    }

    const anchor: DrawingAnchor = {
      picture: {
        rId: rIdImage
      },
      range: medium.range
    };

    // Pass through watermark opacity as alphaModFix
    if (medium.opacity !== undefined) {
      const clamped = Math.max(0, Math.min(1, medium.opacity));
      anchor.picture.alphaModFix = Math.round(clamped * 100000);
    }

    // Handle image hyperlinks
    if (medium.hyperlinks && medium.hyperlinks.hyperlink) {
      const rIdHyperlink = options.nextRId(rels);
      anchor.picture.hyperlinks = {
        tooltip: medium.hyperlinks.tooltip,
        rId: rIdHyperlink
      };
      rels.push({
        Id: rIdHyperlink,
        Type: RelType.Hyperlink,
        Target: medium.hyperlinks.hyperlink,
        TargetMode: "External"
      });
    }

    anchors.push(anchor);
  }

  return { anchors, rels };
}

// =============================================================================
// Anchor Filtering
// =============================================================================

/**
 * Filter drawing anchors to remove invalid entries before XML generation.
 *
 * Shared between streaming `WorkbookWriterBase.addDrawings()` and
 * non-streaming `XLSX.addDrawings()`.
 */
export function filterDrawingAnchors(anchors: any[]): any[] {
  return anchors.filter(a => {
    if (a == null) {
      return false;
    }
    // Absolute anchors need a valid picture
    if (a.range?.pos !== undefined) {
      return !!a.picture;
    }
    // Form controls have range.br and shape properties
    if (a.range?.br && a.shape) {
      return true;
    }
    // One-cell anchors need a valid picture or graphicFrame (charts)
    if (!a.range?.br && !a.picture && !a.graphicFrame) {
      return false;
    }
    // Two-cell anchors need either picture, shape, or graphicFrame (charts)
    if (a.range?.br && !a.picture && !a.shape && !a.graphicFrame) {
      return false;
    }
    return true;
  });
}
