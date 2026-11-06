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

/**
 * Resolve an OPC relationship target against its source part.
 *
 * - Leading `/` → package root absolute
 * - `../` / `./` → resolved relative to the source part's directory
 * - Plain paths → resolved relative to the source part's directory
 *
 * Returns an empty string when the resolved path would escape the package
 * root (i.e. more `..` segments than there are directories above the
 * source). The previous behaviour silently treated extra `..` as a no-op,
 * which let a malicious or malformed `Target` value point at an entirely
 * different part than the directory traversal would suggest.
 */
export function resolvePartPath(sourcePart: string, target: string): string {
  if (!target) {
    return "";
  }
  if (target.startsWith("/")) {
    return target.slice(1);
  }
  const lastSlash = sourcePart.lastIndexOf("/");
  const baseDir = lastSlash >= 0 ? sourcePart.substring(0, lastSlash).split("/") : [];
  const segs = target.split("/");
  for (const seg of segs) {
    if (seg === "..") {
      if (baseDir.length === 0) {
        // The target asks to step above the package root. Treat as
        // unresolvable instead of silently swallowing the segment.
        return "";
      }
      baseDir.pop();
    } else if (seg !== "." && seg !== "") {
      baseDir.push(seg);
    }
  }
  return baseDir.join("/");
}
