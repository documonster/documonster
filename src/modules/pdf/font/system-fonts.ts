/**
 * System font discovery for PDF generation.
 *
 * When no embedded font is provided and the document contains non-WinAnsi
 * characters, this module searches standard system font directories for a
 * TrueType font (.ttf or .ttc) with broad Unicode coverage.
 *
 * This is a Node.js-only feature — browser environments do not have
 * file system access and must always provide fonts explicitly.
 *
 * .ttc (TrueType Collection) files are supported — parseTtf() extracts
 * the first font from the collection automatically.
 *
 * Results are cached: the filesystem search runs only once per process.
 */

import { fileExistsSync, readFileBytesSync, traverseDirectorySync } from "@utils/fs";

// =============================================================================
// Platform Font Directories
// =============================================================================

function getSystemFontDirs(): string[] {
  const platform = typeof process !== "undefined" ? process.platform : "";
  const home =
    typeof process !== "undefined" ? (process.env.HOME ?? process.env.USERPROFILE ?? "") : "";

  const dirs: string[] = [];

  switch (platform) {
    case "darwin":
      dirs.push(
        "/System/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
        "/Library/Fonts",
        `${home}/Library/Fonts`
      );
      break;
    case "win32": {
      const winDir = process.env.WINDIR ?? process.env.SystemRoot ?? "C:\\Windows";
      dirs.push(`${winDir}\\Fonts`, `${process.env.LOCALAPPDATA ?? ""}\\Microsoft\\Windows\\Fonts`);
      break;
    }
    case "linux":
    default:
      dirs.push(
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        "/usr/share/fonts/truetype",
        "/usr/share/fonts/opentype",
        "/usr/share/fonts/TTF",
        "/usr/share/fonts/noto",
        "/usr/share/fonts/noto-cjk",
        "/usr/share/fonts/google-noto",
        "/usr/share/fonts/google-noto-cjk",
        "/usr/share/fonts/truetype/noto",
        "/usr/share/fonts/truetype/dejavu",
        "/usr/share/fonts/truetype/liberation",
        "/usr/share/fonts/truetype/droid",
        "/usr/share/fonts/wqy",
        `${home}/.local/share/fonts`,
        `${home}/.fonts`
      );
      break;
  }

  return dirs;
}

// =============================================================================
// Preferred Font Names (ordered by preference — first match wins)
// =============================================================================

const PREFERRED_FONTS = [
  // Noto Sans CJK — broadest open-source coverage
  "NotoSansCJKsc-Regular.ttf",
  "NotoSansCJK-Regular.ttc",
  "NotoSansCJKSC-Regular.otf",
  "NotoSansSC-Regular.ttf",
  "NotoSansTC-Regular.ttf",
  "NotoSansJP-Regular.ttf",
  "NotoSansKR-Regular.ttf",
  "NotoSans-Regular.ttf",
  // Arial Unicode MS
  "Arial Unicode.ttf",
  "Arial Unicode MS.ttf",
  "ArialUnicode.ttf",
  "arialuni.ttf",
  // macOS
  "PingFang.ttc",
  "Hiragino Sans GB.ttc",
  "STHeiti Light.ttc",
  "STHeiti Medium.ttc",
  "Songti.ttc",
  "AppleSDGothicNeo.ttc",
  // Windows
  "msyh.ttc",
  "msyhbd.ttc",
  "msjh.ttc",
  "simsun.ttc",
  "simhei.ttf",
  "malgun.ttf",
  "meiryo.ttc",
  "yugothic.ttf",
  "segoeui.ttf",
  "arial.ttf",
  // Linux
  "DejaVuSans.ttf",
  "LiberationSans-Regular.ttf",
  "FreeSans.ttf",
  "DroidSansFallbackFull.ttf",
  "DroidSansFallback.ttf",
  "wqy-microhei.ttc",
  "wqy-zenhei.ttc",
  "uming.ttc",
  "NanumGothic.ttf",
  "IPAexGothic.ttf"
];

// =============================================================================
// Font Discovery
// =============================================================================

let _cachedCandidates: Uint8Array[] | undefined;

/**
 * Return all discoverable system font candidates, ordered by preference.
 *
 * Each entry is the raw font file bytes of a `.ttf` or `.ttc` file.
 * The caller decides which candidate to use (e.g. by checking cmap coverage).
 *
 * Results are cached — the filesystem scan runs only once per process.
 */
export function discoverSystemFontCandidates(): Uint8Array[] {
  if (_cachedCandidates !== undefined) {
    return _cachedCandidates;
  }

  if (typeof process === "undefined" || !process.platform) {
    _cachedCandidates = [];
    return _cachedCandidates;
  }

  const candidates: Uint8Array[] = [];
  const seen = new Set<string>(); // dedupe by path

  const dirs = getSystemFontDirs();

  // Strategy 1: Check preferred font filenames (in order)
  for (const fontName of PREFERRED_FONTS) {
    for (const dir of dirs) {
      const fontPath = `${dir}/${fontName}`;
      if (seen.has(fontPath)) {
        continue;
      }
      if (fileExistsSync(fontPath)) {
        const data = tryReadFont(fontPath);
        if (data) {
          candidates.push(data);
          seen.add(fontPath);
        }
      }
    }
  }

  // Strategy 2: Scan directories for any .ttf/.ttc not already found
  const broadRe =
    /noto|unicode|cjk|yahei|heiti|gothic|sans|serif|ming|song|dejavu|liberation|droid|wqy/i;

  for (const dir of dirs) {
    try {
      const entries = traverseDirectorySync(dir, { recursive: true, filter: e => !e.isDirectory });
      const fonts = entries.filter(
        e => /\.tt[cf]$/i.test(e.absolutePath) && !seen.has(e.absolutePath)
      );

      // Broad-coverage names first, then large files
      const broad = fonts.filter(e => broadRe.test(e.absolutePath));
      const rest = fonts.filter(e => !broadRe.test(e.absolutePath) && e.size > 50000);

      for (const entry of [...broad, ...rest]) {
        if (seen.has(entry.absolutePath)) {
          continue;
        }
        const data = tryReadFont(entry.absolutePath);
        if (data) {
          candidates.push(data);
          seen.add(entry.absolutePath);
        }
      }
    } catch {
      // Directory doesn't exist or not readable
    }
  }

  _cachedCandidates = candidates;
  return candidates;
}

/**
 * Search for a system font suitable for Unicode rendering.
 *
 * Returns the raw font file bytes of the highest-priority candidate,
 * or `null` if no font was found. This is a convenience wrapper around
 * {@link discoverSystemFontCandidates}.
 */
export function discoverSystemFont(): Uint8Array | null {
  const candidates = discoverSystemFontCandidates();
  return candidates.length > 0 ? candidates[0] : null;
}

/**
 * Reset the cached font discovery result (for testing).
 */
export function resetFontDiscoveryCache(): void {
  _cachedCandidates = undefined;
}

/**
 * Override the cached candidates with a custom list (for testing).
 * Call {@link resetFontDiscoveryCache} to clear the override.
 */
export function _setCandidatesForTest(candidates: Uint8Array[]): void {
  _cachedCandidates = candidates;
}

// =============================================================================
// Internal
// =============================================================================

function tryReadFont(fontPath: string): Uint8Array | null {
  try {
    return readFileBytesSync(fontPath);
  } catch {
    return null;
  }
}
