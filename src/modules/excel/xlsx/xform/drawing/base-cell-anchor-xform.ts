import { inferExternalImageExtension } from "@excel/utils/drawing-utils";
import { BaseXform } from "@excel/xlsx/xform/base-xform";

abstract class BaseCellAnchorXform extends BaseXform {
  declare public map: { [key: string]: any };
  declare public parser: any;
  declare public model: any;

  abstract get tag(): string;

  parseOpen(node: any): boolean {
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

  reconcilePicture(model: any, options: any): any {
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

      const match = rel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
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
          const svgMatch = svgRel && svgRel.Target.match(/.*\/media\/(.+[.][a-zA-Z]{3,4})/);
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
  private reconcileExternalPicture(link: string, options: any): any {
    if (!link) {
      return undefined;
    }
    const indexKey = `external:${link}`;
    let mediaId = options.mediaIndex[indexKey];
    if (mediaId === undefined) {
      mediaId = options.media.length;
      const medium = {
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
