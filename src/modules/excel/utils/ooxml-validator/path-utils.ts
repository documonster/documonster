/**
 * Path helpers shared across all validator checkers.
 *
 * All paths are handled as **forward-slash POSIX strings** because that is
 * how OPC part names are defined. We never touch the host's path separator.
 */

/**
 * Strip a single leading `/` if present. OPC `PartName` values start with
 * `/` by spec, but zip entries do not.
 */
export function stripLeadingSlash(p: string): string {
  return p.startsWith("/") ? p.slice(1) : p;
}

/** `.xml`, `.rels`, or `.vml` — the payloads we parse as XML. */
export function isXmlLike(pathName: string): boolean {
  return pathName.endsWith(".xml") || pathName.endsWith(".rels") || pathName.endsWith(".vml");
}

/**
 * Directory in which a relationships part's **source** part lives. OPC
 * resolves rel Targets relative to the source part's directory, NOT the
 * `.rels` file's directory. Examples:
 *
 * | .rels path                              | source dir          |
 * |-----------------------------------------|---------------------|
 * | `_rels/.rels`                           | ``                  |
 * | `xl/_rels/workbook.xml.rels`            | `xl`                |
 * | `xl/worksheets/_rels/sheet1.xml.rels`   | `xl/worksheets`     |
 */
export function getRelsSourceDir(relsPath: string): string {
  if (relsPath === "_rels/.rels") {
    return "";
  }
  const marker = "/_rels/";
  const idx = relsPath.indexOf(marker);
  if (idx === -1) {
    return relsPath.includes("/") ? relsPath.slice(0, relsPath.lastIndexOf("/")) : "";
  }
  return relsPath.slice(0, idx);
}

/**
 * Basename without directory, e.g. `xl/foo/bar.xml` → `bar.xml`.
 */
export function posixBasename(p: string): string {
  const idx = p.lastIndexOf("/");
  return idx === -1 ? p : p.slice(idx + 1);
}

/** Extension without the leading dot, lowercase. `""` when missing. */
export function getExtension(p: string): string {
  const base = posixBasename(p);
  const idx = base.lastIndexOf(".");
  return idx === -1 ? "" : base.slice(idx + 1).toLowerCase();
}

/**
 * Resolve a relationship Target against a .rels path. Handles absolute
 * targets starting with `/` (relative to package root), `../` traversal,
 * `./`, and plain relative segments. Never returns a leading slash.
 *
 * We do NOT use `node:path` so this module is usable in browsers.
 */
export function resolveRelTarget(relsPath: string, target: string): string {
  // Absolute-in-package targets start with "/".
  if (target.startsWith("/")) {
    return normalizeSegments(target.split("/"));
  }
  const baseDir = getRelsSourceDir(relsPath);
  const segments: string[] = [];
  if (baseDir) {
    segments.push(...baseDir.split("/"));
  }
  segments.push(...target.split("/"));
  return normalizeSegments(segments);
}

function normalizeSegments(segments: string[]): string {
  const out: string[] = [];
  let escaped = false;
  for (const seg of segments) {
    if (seg === "" || seg === ".") {
      continue;
    }
    if (seg === "..") {
      if (out.length === 0) {
        // Record traversal past the package root — caller's job to flag it.
        escaped = true;
      } else {
        out.pop();
      }
      continue;
    }
    out.push(seg);
  }
  const joined = out.join("/");
  return escaped ? `../${joined}` : joined;
}

/**
 * `false` when a resolved rel target escapes the package root. Targets
 * that escape are malicious or accidentally produced by broken
 * serialisers and always indicate a corrupt package.
 */
export function isSafeResolvedPath(resolved: string): boolean {
  return !(resolved === ".." || resolved.startsWith("../") || resolved.includes("/../"));
}

/**
 * Derive the .rels path for a given source part. Symmetric with
 * `getRelsSourceDir`. Examples:
 *
 * | source part                    | rels path                              |
 * |--------------------------------|----------------------------------------|
 * | `xl/workbook.xml`              | `xl/_rels/workbook.xml.rels`           |
 * | `xl/worksheets/sheet1.xml`     | `xl/worksheets/_rels/sheet1.xml.rels`  |
 */
export function relsPathForPart(partPath: string): string {
  const base = posixBasename(partPath);
  const dir =
    partPath.length > base.length ? partPath.slice(0, partPath.length - base.length - 1) : "";
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

/**
 * Inverse of {@link relsPathForPart}: source part for a given rels file.
 * Returns undefined for the root rels (which has no conventional source).
 */
export function sourcePartForRels(relsPath: string): string | undefined {
  if (relsPath === "_rels/.rels") {
    return undefined;
  }
  const base = posixBasename(relsPath);
  if (!base.endsWith(".rels")) {
    return undefined;
  }
  const srcDir = getRelsSourceDir(relsPath);
  const sourceName = base.slice(0, -".rels".length);
  return srcDir ? `${srcDir}/${sourceName}` : sourceName;
}

/**
 * Part name satisfies OPC PartName grammar — not empty, uses only safe
 * characters, does not end with `/` and does not contain `.` or `..`
 * segments. We do not validate the full grammar but cover the cases that
 * break Excel in practice.
 */
export function isLegalPartName(partName: string): boolean {
  if (!partName) {
    return false;
  }
  if (partName.endsWith("/")) {
    return false;
  }
  if (/[\\:*?"<>|]/.test(partName)) {
    return false;
  }
  const segments = partName.split("/");
  for (const seg of segments) {
    if (seg === "" || seg === "." || seg === "..") {
      return false;
    }
  }
  return true;
}
