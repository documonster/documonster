/**
 * OPC Path Utilities
 *
 * Small helpers for working with Open Packaging Conventions part paths.
 * These are used throughout the reader/writer when resolving the location of
 * a relationships file or extracting an extension to choose a content type.
 */

/**
 * Get the `.rels` part path for a given part path.
 *
 * Per OPC convention, the relationships for `dir/name.ext` live at
 * `dir/_rels/name.ext.rels`; for a top-level `name.ext` they live at
 * `_rels/name.ext.rels`.
 */
export function getPartRelsPath(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash) : "";
  const name = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
  return dir ? `${dir}/_rels/${name}.rels` : `_rels/${name}.rels`;
}

/**
 * Extract the file name (last path segment) from a part path.
 *
 * @example
 * getFileName("word/media/image1.png")  // "image1.png"
 * getFileName("standalone.xml")         // "standalone.xml"
 */
export function getFileName(partPath: string): string {
  const lastSlash = partPath.lastIndexOf("/");
  return lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
}

/**
 * Extract the lower-cased file extension (without the dot) from a part path
 * or file name. Returns an empty string when there is no extension.
 *
 * @example
 * getFileExt("image1.PNG")        // "png"
 * getFileExt("word/media/x.tiff") // "tiff"
 * getFileExt("noext")             // ""
 */
export function getFileExt(partPath: string): string {
  const fileName = getFileName(partPath);
  const dot = fileName.lastIndexOf(".");
  if (dot < 0 || dot === fileName.length - 1) {
    return "";
  }
  return fileName.substring(dot + 1).toLowerCase();
}
