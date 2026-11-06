/**
 * The image media types a document accepts, and the checks that keep an image
 * part writable: its name must carry an extension `[Content_Types].xml` can
 * type. Shared by the public entry points (which reject a bad media type where
 * it is supplied), the writers (which refuse to emit an untyped part) and
 * `Advanced.validateDocument` (which reports the same conditions as issues).
 */

import { IMAGE_CONTENT_TYPES } from "@word/constants";
import { messagePrefix } from "@word/core/argument-checks";
import { DocxInvalidImageError } from "@word/errors";
import type { ImageMediaType } from "@word/types";

/** Every {@link ImageMediaType}, in declaration order. */
export const IMAGE_MEDIA_TYPES: readonly ImageMediaType[] = [
  "png",
  "jpeg",
  "gif",
  "bmp",
  "tiff",
  "svg",
  "webp",
  "emf",
  "wmf"
];

/**
 * Media types an existing `ImageDef` may carry. Wider than
 * {@link IMAGE_MEDIA_TYPES} because the reader names a picture's type after its
 * file extension, so a read document holds `"jpg"` and `"tif"`.
 */
export const KNOWN_IMAGE_MEDIA_TYPES: readonly string[] = [...IMAGE_MEDIA_TYPES, "jpg", "tif"];

/** File extension → the media types a picture with that extension may declare. */
export const EXTENSION_MEDIA_TYPE_MAP: ReadonlyMap<string, readonly ImageMediaType[]> = new Map<
  string,
  readonly ImageMediaType[]
>([
  ["png", ["png"]],
  ["jpg", ["jpeg"]],
  ["jpeg", ["jpeg"]],
  ["gif", ["gif"]],
  ["bmp", ["bmp"]],
  ["tiff", ["tiff"]],
  ["tif", ["tiff"]],
  ["svg", ["svg"]],
  ["webp", ["webp"]],
  ["emf", ["emf"]],
  ["wmf", ["wmf"]]
]);

/** The media type a picture with this file extension has, or `undefined` for an unknown extension. */
export function mediaTypeForExtension(ext: string): ImageMediaType | undefined {
  return EXTENSION_MEDIA_TYPE_MAP.get(ext.toLowerCase())?.[0];
}

/** Whether `value` is an {@link ImageMediaType}. */
export function isImageMediaType(value: unknown): value is ImageMediaType {
  return typeof value === "string" && (IMAGE_MEDIA_TYPES as readonly string[]).includes(value);
}

/**
 * Throw {@link DocxInvalidImageError} unless `value` is an {@link ImageMediaType}.
 * `where` names the entry point, for the message.
 */
export function assertImageMediaType(
  value: unknown,
  where: string,
  field = "mediaType"
): asserts value is ImageMediaType {
  if (!isImageMediaType(value)) {
    throw new DocxInvalidImageError(
      field,
      `${where}: invalid image media type ${JSON.stringify(value)}; expected one of ${IMAGE_MEDIA_TYPES.join(", ")} (a file-extension name, not a MIME type)`
    );
  }
}

/** The extension of a part or file name, lower-cased, or `undefined` if it has none. */
function extensionOf(fileName: string): string | undefined {
  const dot = fileName.lastIndexOf(".");
  return dot > 0 && dot < fileName.length - 1 ? fileName.slice(dot + 1).toLowerCase() : undefined;
}

/**
 * Why an image with this file name and media type cannot be written as a
 * package part, or `undefined` if it can: the name must be a single path
 * segment whose extension maps to an image content type, and the media type
 * must be one the writer knows — unless the picture carries the content type
 * its source package declared, which types the part whatever its extension.
 */
export function describeImagePartProblem(
  fileName: string | undefined,
  mediaType: string | undefined,
  contentType?: string
): { readonly field: "fileName" | "mediaType"; readonly problem: string } | undefined {
  if (!fileName) {
    return { field: "fileName", problem: "has no fileName" };
  }
  if (/[/\\]/.test(fileName)) {
    return {
      field: "fileName",
      problem: `fileName ${JSON.stringify(fileName)} is not a single path segment`
    };
  }
  // A read picture typed by its own package is written with that type, by Override.
  if (contentType) {
    return undefined;
  }
  const ext = extensionOf(fileName);
  if (!ext || !(ext in IMAGE_CONTENT_TYPES)) {
    return {
      field: "fileName",
      problem: `fileName ${JSON.stringify(fileName)} has no image extension with a content type (expected one of ${Object.keys(IMAGE_CONTENT_TYPES).join(", ")})`
    };
  }
  if (!mediaType || !KNOWN_IMAGE_MEDIA_TYPES.includes(mediaType)) {
    return {
      field: "mediaType",
      problem: `invalid mediaType ${JSON.stringify(mediaType)}; expected one of ${IMAGE_MEDIA_TYPES.join(", ")}`
    };
  }
  return undefined;
}

/**
 * The content type a picture's part is written with when its extension has no
 * image `Default`: the one its source package declared. `undefined` when the
 * extension's `Default` covers it.
 */
export function imagePartOverride(image: {
  readonly fileName: string;
  readonly contentType?: string;
}): string | undefined {
  const ext = extensionOf(image.fileName);
  return image.contentType && !(ext && ext in IMAGE_CONTENT_TYPES) ? image.contentType : undefined;
}

/** Throw {@link DocxInvalidImageError} if an image cannot be written as a typed package part. */
export function assertWritableImage(
  image: { readonly fileName?: string; readonly mediaType?: string; readonly contentType?: string },
  where: string,
  at: string
): void {
  const found = describeImagePartProblem(image.fileName, image.mediaType, image.contentType);
  if (found) {
    throw new DocxInvalidImageError(
      `${at}.${found.field}`,
      `${messagePrefix(where)}image ${found.problem}`
    );
  }
}
