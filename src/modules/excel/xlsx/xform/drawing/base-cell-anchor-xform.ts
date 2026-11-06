import { inferExternalImageExtension } from "@excel/utils/drawing-utils";
import { BaseXform } from "@excel/xlsx/xform/base-xform";
import { GenericEchoXform } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import type { EchoNode } from "@excel/xlsx/xform/drawing/generic-echo-xform";
import type { ParseOpenTag, XmlSink } from "@xml/types";

/**
 * The drawing objects an anchor carries that this library does not model: a group, a shape — a text box is one — and
 * a connector. Each is captured whole and written back as read (see {@link GenericEchoXform}), relationship ids and
 * all. Without an entry for the tag, the anchor's dispatch fell through to the generic per-tag lookup: a shape parsed
 * into nothing and was deleted on write, and a picture nested in a group was promoted to the anchor's own.
 */
const PRESERVED_TAGS = ["xdr:sp", "xdr:grpSp", "xdr:cxnSp"] as const;

/** Parsers for {@link PRESERVED_TAGS}, to spread into an anchor's `map`. */
export function preservedObjectXforms(): Record<string, BaseXform> {
  return Object.fromEntries(PRESERVED_TAGS.map(tag => [tag, new GenericEchoXform(tag)]));
}

/** The picture fields the anchor reconciler reads off `model.picture`. */
interface ReconcilePicture {
  rId?: string;
  external?: boolean;
  svgRId?: string;
  alphaModFix?: number;
}

/** A media entry resolved from the package. Extensible (sourced from varied
 * relationship/media records), so kept structurally open rather than over-typed. */
type MediaEntry = Record<string, unknown>;

/** The reconcile context threaded from the drawing reader. */
interface ReconcileOptions {
  rels: Record<string, { Target?: string; TargetMode?: string } | undefined>;
  media: MediaEntry[];
  mediaIndex: Record<string, number>;
}

abstract class BaseCellAnchorXform extends BaseXform {
  declare public map: Record<string, BaseXform>;
  declare public parser?: BaseXform;
  declare public model: { range: { editAs?: string } & Record<string, unknown> } & Record<
    string,
    unknown
  >;

  abstract get tag(): string;

  parseOpen(node: ParseOpenTag): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          range: {
            editAs: node.attributes.editAs
          }
        };
        break;
      default:
        this.parser = this.map[node.name];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  /**
   * The unmodelled object this anchor parsed, if any, with its parser reset. The parsers are shared across every
   * anchor of a drawing, so a model left behind would be inherited by the next anchor.
   */
  protected takePreserved(): EchoNode | undefined {
    let preserved: EchoNode | undefined;
    for (const tag of PRESERVED_TAGS) {
      const xform = this.map[tag];
      if (xform?.model) {
        preserved ??= xform.model as EchoNode;
        xform.reset();
      }
    }
    return preserved;
  }

  /** Write an object captured by {@link takePreserved} back as it was read. */
  protected renderPreserved(xmlStream: XmlSink, preserved: EchoNode): void {
    this.map[preserved.tag]?.render(xmlStream, preserved);
  }

  reconcilePicture(
    model: ReconcilePicture | undefined,
    options: ReconcileOptions
  ): MediaEntry | undefined {
    if (model && model.rId) {
      const rel = options.rels[model.rId];
      if (!rel) {
        return undefined;
      }

      // External (linked) image: the relationship uses TargetMode="External"
      // and there is no media part in the package. Synthesize a media entry
      // carrying the `link` target so it round-trips and surfaces on the
      // worksheet as an external image. Entries are deduplicated by link.
      if (rel.TargetMode === "External" || model.external) {
        return this.reconcileExternalPicture(rel.Target, options);
      }

      const match = rel.Target?.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
      if (match) {
        const name = match[1];
        const mediaId = options.mediaIndex[name];
        const medium = options.media[mediaId];
        if (!medium) {
          return undefined;
        }

        // Resolve an SVG companion (asvg:svgBlip extension) back to its media
        // index and record it on the raster media entry itself, so callers that
        // look the image up by id (e.g. Workbook.getImage) surface the vector
        // companion alongside the raster fallback.
        if (model.svgRId) {
          const svgRel = options.rels[model.svgRId];
          const svgMatch = svgRel && svgRel.Target?.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
          if (svgMatch) {
            const svgMediaId = options.mediaIndex[svgMatch[1]];
            if (svgMediaId !== undefined) {
              medium.svgMediaId = svgMediaId;
            }
          }
        }

        // Preserve alphaModFix (transparency) from the picture model if present
        if (model.alphaModFix !== undefined) {
          return { ...medium, alphaModFix: model.alphaModFix };
        }
        return medium;
      }
    }
    return undefined;
  }

  /**
   * Resolve (or create) the media entry for an external linked image. The
   * synthesized entry is appended to `options.media` and indexed by its link
   * so repeated references to the same external image share one entry.
   */
  private reconcileExternalPicture(
    link: string | undefined,
    options: ReconcileOptions
  ): MediaEntry | undefined {
    if (!link) {
      return undefined;
    }
    const indexKey = `external:${link}`;
    let mediaId = options.mediaIndex[indexKey];
    if (mediaId === undefined) {
      mediaId = options.media.length;
      const medium: MediaEntry = {
        type: "image",
        extension: inferExternalImageExtension(link),
        link,
        index: mediaId
      };
      options.media.push(medium);
      options.mediaIndex[indexKey] = mediaId;
    }
    return options.media[mediaId];
  }
}

export { BaseCellAnchorXform };
