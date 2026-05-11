/**
 * OPC Package Model — format-agnostic representation of an Open Packaging package.
 *
 * This abstraction layer sits between:
 * - Reader: ZIP bytes → OpcPackage → DocxDocument
 * - Writer: DocxDocument → PackagePlan → OpcPackage → ZIP bytes / Flat OPC XML
 * - Streaming: DocxDocument → PackagePlan → incremental OPC parts
 */

/** A fully-qualified part name within the package (e.g. "/word/document.xml"). */
export type PartName = string;

/** Relationship target mode. */
export type TargetMode = "Internal" | "External";

/** A single relationship entry. */
export interface OpcRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: TargetMode;
}

/** Set of relationships for a source part (or the package root). */
export interface OpcRelationshipSet {
  readonly source: string;
  readonly relationships: readonly OpcRelationship[];
}

/** A single part within the package. */
export interface OpcPart {
  /** Part name (absolute path starting with /). */
  readonly name: PartName;
  /** Content type (MIME type). */
  readonly contentType: string;
  /** Part data. String for XML parts, Uint8Array for binary parts. */
  readonly data: Uint8Array | string;
  /** Relationships originating from this part. */
  readonly relationships?: readonly OpcRelationship[];
}

/** Content type definition (extension default or per-part override). */
export interface ContentTypeEntry {
  readonly partName?: PartName; // If set, it's an override
  readonly extension?: string; // If set, it's a default
  readonly contentType: string;
}

/** Complete OPC package representation. */
export interface OpcPackage {
  /** All parts indexed by part name. */
  readonly parts: ReadonlyMap<PartName, OpcPart>;
  /** Package-level relationships (_rels/.rels). */
  readonly packageRelationships: readonly OpcRelationship[];
  /** Content type catalog. */
  readonly contentTypes: readonly ContentTypeEntry[];
}

/** Options for writing an OPC package. */
export interface OpcWriteOptions {
  /** Compression level (0-9). Default: 6. */
  readonly compressionLevel?: number;
}

/** A "package plan" — the serialization plan for a Word document before final output. */
export interface WordPackagePlan {
  /** Main document part. */
  readonly documentPart: OpcPart;
  /** All auxiliary parts (styles, settings, numbering, headers, footers, etc.). */
  readonly auxiliaryParts: readonly OpcPart[];
  /** Media parts (images, embedded xlsx, etc.). */
  readonly mediaParts: readonly OpcPart[];
  /** Package-level relationships. */
  readonly packageRelationships: readonly OpcRelationship[];
  /** Content type entries. */
  readonly contentTypes: readonly ContentTypeEntry[];
}

/** Convert a WordPackagePlan to an OpcPackage (ready for serialization). */
export function planToPackage(plan: WordPackagePlan): OpcPackage {
  const parts = new Map<PartName, OpcPart>();

  // Add the main document part
  parts.set(normalizePartName(plan.documentPart.name), plan.documentPart);

  // Add auxiliary parts
  for (const part of plan.auxiliaryParts) {
    parts.set(normalizePartName(part.name), part);
  }

  // Add media parts
  for (const part of plan.mediaParts) {
    parts.set(normalizePartName(part.name), part);
  }

  return {
    parts,
    packageRelationships: plan.packageRelationships,
    contentTypes: plan.contentTypes
  };
}

/** Normalize a part name (ensure leading /, normalize separators). */
export function normalizePartName(name: string): PartName {
  // Replace backslashes with forward slashes
  let normalized = name.replace(/\\/g, "/");

  // Collapse consecutive slashes
  normalized = normalized.replace(/\/+/g, "/");

  // Ensure leading slash
  if (!normalized.startsWith("/")) {
    normalized = "/" + normalized;
  }

  // Remove trailing slash (part names are never directories)
  if (normalized.length > 1 && normalized.endsWith("/")) {
    normalized = normalized.slice(0, -1);
  }

  return normalized;
}

/** Resolve a relative target path against a source part name. */
export function resolveTarget(sourcePart: PartName, relativeTarget: string): PartName {
  // If the target is already absolute, just normalize it
  if (relativeTarget.startsWith("/")) {
    return normalizePartName(relativeTarget);
  }

  // Get the directory of the source part
  const lastSlash = sourcePart.lastIndexOf("/");
  const sourceDir = lastSlash > 0 ? sourcePart.slice(0, lastSlash) : "";

  // Combine source directory with relative target
  const combined = sourceDir + "/" + relativeTarget;

  // Resolve . and .. segments
  const segments = combined.split("/");
  const resolved: string[] = [];

  for (const segment of segments) {
    if (segment === "" || segment === ".") {
      continue;
    }
    if (segment === "..") {
      resolved.pop();
    } else {
      resolved.push(segment);
    }
  }

  return "/" + resolved.join("/");
}

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
