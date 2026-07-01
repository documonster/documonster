import { ArchiveError } from "@archive/core/errors";

export type ZipPathMode = "legacy" | "posix" | "safe";

export interface ZipPathOptions {
  /**
   * - "legacy": replace backslashes with '/', strip leading '/'
   * - "posix": also normalizes '.' and '..' segments using POSIX rules
   * - "safe": like "posix" but rejects absolute paths and any remaining '..' traversal
   */
  mode?: ZipPathMode;

  /** If true, ensures a single leading '/'. */
  prependSlash?: boolean;

  /** If true, strips a Windows drive prefix like `C:`. Default: true for posix/safe. */
  stripDrive?: boolean;
}

function replaceBackslashes(p: string): string {
  return p.replace(/\\/g, "/");
}

function stripLeadingSlashes(p: string): string {
  return p.replace(/^\/+/, "");
}

function stripWindowsDrive(p: string): string {
  // C:/path or C:path
  return p.replace(/^[a-zA-Z]:\/?/, "");
}

function normalizePosix(p: string): string {
  const parts = p.split("/");
  const out: string[] = [];
  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }
    if (part === "..") {
      if (out.length > 0) {
        out.pop();
      } else {
        // Preserve leading '..' to be handled by caller (safe mode may reject).
        out.push("..");
      }
      continue;
    }
    out.push(part);
  }
  return out.join("/");
}

function isAbsoluteZipPath(p: string): boolean {
  // Backslashes are normalized earlier, so this only needs to check '/'.
  return p.startsWith("/");
}

function hasWindowsDrivePrefix(p: string): boolean {
  // C:/path or C:path
  return /^[a-zA-Z]:/.test(p);
}

export function normalizeZipPath(pathName: string, options: ZipPathOptions = {}): string {
  const mode: ZipPathMode = options.mode ?? "legacy";

  let p = pathName;
  p = replaceBackslashes(p);

  const hadWindowsDrive = hasWindowsDrivePrefix(p);

  if (mode === "legacy") {
    p = stripLeadingSlashes(p);
    return options.prependSlash ? "/" + p : p;
  }

  const stripDrive = options.stripDrive ?? true;
  if (mode === "safe" && hadWindowsDrive && !stripDrive) {
    throw new ArchiveError(`Unsafe ZIP path (drive): ${pathName}`);
  }
  if (stripDrive) {
    p = stripWindowsDrive(p);
  }

  const hadLeadingSlash = isAbsoluteZipPath(p);
  p = stripLeadingSlashes(p);
  p = normalizePosix(p);

  if (mode === "safe") {
    if (hadLeadingSlash) {
      throw new ArchiveError(`Unsafe ZIP path (absolute): ${pathName}`);
    }
    if (p === ".." || p.startsWith("../")) {
      throw new ArchiveError(`Unsafe ZIP path (traversal): ${pathName}`);
    }
  }

  if (options.prependSlash) {
    return "/" + p;
  }
  return p;
}

export function joinZipPath(options: ZipPathOptions, ...parts: string[]): string {
  const normalizeOptions: ZipPathOptions = { ...options, prependSlash: false };
  const normalized: string[] = [];

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    if (!part) {
      continue;
    }
    const next = normalizeZipPath(part, normalizeOptions);
    if (next) {
      normalized.push(next);
    }
  }

  const joined = normalized.join("/");
  return options.prependSlash ? "/" + stripLeadingSlashes(joined) : joined;
}
