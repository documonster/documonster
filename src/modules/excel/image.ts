import { Anchor, type AnchorModel } from "@excel/anchor";
import { ImageError } from "@excel/errors";
import { colCache } from "@excel/utils/col-cache";
import type { Worksheet } from "@excel/worksheet";

interface ImageHyperlinks {
  hyperlink?: string;
  tooltip?: string;
}

interface ImageExt {
  width?: number;
  height?: number;
}

/** Absolute position in pixels (for absoluteAnchor). */
interface ImagePos {
  x: number;
  y: number;
}

interface ImageRange {
  tl: Anchor;
  br?: Anchor;
  ext?: ImageExt;
  editAs?: string;
  hyperlinks?: ImageHyperlinks;
  /** Absolute position — mutually exclusive with tl/br cell anchors. */
  pos?: ImagePos;
}

interface BackgroundModel {
  type: "background";
  imageId: string;
}

interface WatermarkModel {
  type: "watermark";
  imageId: string;
  /** Opacity 0-1 for overlay mode */
  opacity?: number;
}

interface HeaderImageModel {
  type: "headerImage";
  imageId: string;
  headerWidth?: number;
  headerHeight?: number;
  applyTo?: "all" | "odd" | "even" | "first";
}

interface ImageRangeModel {
  tl: AnchorModel;
  br?: AnchorModel;
  ext?: ImageExt;
  editAs?: string;
  /** Absolute position — when present, tl/br are ignored. */
  pos?: ImagePos;
}

interface ImageModel {
  type: "image";
  imageId: string;
  hyperlinks?: ImageHyperlinks;
  range: ImageRangeModel;
}

type Model = BackgroundModel | ImageModel | WatermarkModel | HeaderImageModel;
type ImageModelInput = ModelInput;

interface RangeInput {
  tl?: AnchorModel | { col: number; row: number } | string;
  br?: AnchorModel | { col: number; row: number } | string;
  ext?: ImageExt;
  editAs?: string;
  hyperlinks?: ImageHyperlinks;
  /** Absolute position — when present, tl/br are ignored. */
  pos?: ImagePos;
}

interface ModelInput {
  type: string;
  imageId: string;
  range?: string | RangeInput | ImageRangeModel;
  hyperlinks?: ImageHyperlinks;
  opacity?: number;
  headerWidth?: number;
  headerHeight?: number;
  applyTo?: "all" | "odd" | "even" | "first";
}

class Image {
  readonly worksheet: Worksheet;
  type?: string;
  imageId?: string;
  range?: ImageRange;
  /** Opacity for watermark overlay mode (0-1). */
  opacity?: number;
  /** Header image width in points. */
  headerWidth?: number;
  /** Header image height in points. */
  headerHeight?: number;
  /** Header watermark applyTo setting. */
  applyTo?: "all" | "odd" | "even" | "first";

  constructor(worksheet: Worksheet, model?: ModelInput) {
    this.worksheet = worksheet;
    if (model) {
      this.model = model;
    }
  }

  get model(): Model {
    switch (this.type) {
      case "background":
        return {
          type: this.type,
          imageId: this.imageId ?? ""
        };
      case "watermark":
        return {
          type: this.type,
          imageId: this.imageId ?? "",
          opacity: this.opacity
        };
      case "headerImage":
        return {
          type: this.type,
          imageId: this.imageId ?? "",
          headerWidth: this.headerWidth,
          headerHeight: this.headerHeight,
          applyTo: this.applyTo
        };
      case "image": {
        const range = this.range;
        if (!range) {
          throw new ImageError("Image has no range");
        }
        // Absolute positioning — no cell anchors
        if (range.pos) {
          return {
            type: this.type,
            imageId: this.imageId ?? "",
            hyperlinks: range.hyperlinks,
            range: {
              tl: { nativeCol: 0, nativeColOff: 0, nativeRow: 0, nativeRowOff: 0 },
              ext: range.ext,
              pos: range.pos
            }
          };
        }
        return {
          type: this.type,
          imageId: this.imageId ?? "",
          hyperlinks: range.hyperlinks,
          range: {
            tl: range.tl.model,
            br: range.br?.model,
            ext: range.ext,
            editAs: range.editAs
          }
        };
      }
      default:
        throw new ImageError("Invalid Image Type");
    }
  }

  set model({
    type,
    imageId,
    range,
    hyperlinks,
    opacity,
    headerWidth,
    headerHeight,
    applyTo
  }: ModelInput) {
    this.type = type;
    this.imageId = imageId;
    this.opacity = opacity;
    this.headerWidth = headerWidth;
    this.headerHeight = headerHeight;
    this.applyTo = applyTo;

    if (type === "image") {
      if (typeof range === "string") {
        const decoded = colCache.decode(range);
        if ("top" in decoded) {
          // It's a Location (range like "A1:C3")
          this.range = {
            tl: new Anchor(this.worksheet, { col: decoded.left, row: decoded.top }, -1),
            br: new Anchor(this.worksheet, { col: decoded.right, row: decoded.bottom }, 0),
            editAs: "oneCell"
          };
        }
      } else if (range && "pos" in range && range.pos) {
        // Absolute positioning — preserve pos/ext, use dummy tl anchor
        this.range = {
          tl: new Anchor(this.worksheet, null, 0),
          ext: range.ext,
          hyperlinks: hyperlinks || ("hyperlinks" in range ? range.hyperlinks : undefined),
          pos: range.pos
        };
      } else if (range) {
        // Anchor inputs:
        //   - string addresses ("A1") are 1-based; the internal Anchor uses
        //     0-based nativeCol/nativeRow, so the top-left address is shifted
        //     by -1. The bottom-right address is intentionally NOT shifted —
        //     this matches the string-range path (e.g. "A1:B2") where br
        //     anchors the cell *past* the address (i.e. its right/bottom edge).
        //   - object inputs ({ col, row } / AnchorModel) already use the
        //     0-based convention and are passed through as-is.
        const tlInput = range.tl;
        const brInput = range.br;
        this.range = {
          tl: new Anchor(this.worksheet, tlInput, typeof tlInput === "string" ? -1 : 0),
          br: brInput ? new Anchor(this.worksheet, brInput, 0) : undefined,
          ext: range.ext,
          editAs: range.editAs,
          hyperlinks: hyperlinks || ("hyperlinks" in range ? range.hyperlinks : undefined)
        };
      }
    }
  }

  clone(worksheet?: Worksheet): Image {
    const target = worksheet ?? this.worksheet;
    const cloned = new Image(target);
    cloned.type = this.type;
    cloned.imageId = this.imageId;
    cloned.opacity = this.opacity;
    cloned.headerWidth = this.headerWidth;
    cloned.headerHeight = this.headerHeight;
    cloned.applyTo = this.applyTo;

    if (this.range) {
      cloned.range = {
        tl: this.range.tl.clone(target),
        br: this.range.br ? this.range.br.clone(target) : undefined,
        ext: this.range.ext ? { ...this.range.ext } : undefined,
        editAs: this.range.editAs,
        hyperlinks: this.range.hyperlinks ? { ...this.range.hyperlinks } : undefined,
        pos: this.range.pos ? { ...this.range.pos } : undefined
      };
    }

    return cloned;
  }
}

export { Image, type Model as ImageModel, type ImageModelInput };
