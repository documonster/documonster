/**
 * Glob pattern matching utilities.
 *
 * This module provides glob pattern matching functionality that works
 * in both Node.js and browser environments. It's used by the file system
 * utilities but can also be used standalone for pattern matching.
 *
 * @module
 */

// =============================================================================
// Pattern Cache
// =============================================================================

/**
 * Cache for compiled glob regex patterns.
 * Key format: `${pattern}:${dot ? '1' : '0'}`
 */
const regexCache = new Map<string, RegExp>();

/**
 * Maximum number of patterns to cache.
 * Prevents unbounded memory growth in long-running processes.
 */
const MAX_CACHE_SIZE = 1000;

/**
 * Get cache key for a pattern and options.
 */
function getCacheKey(pattern: string, dot: boolean): string {
  return `${pattern}:${dot ? "1" : "0"}`;
}

// =============================================================================
// Glob Pattern Conversion
// =============================================================================

/**
 * Escape special regex characters.
 */
function escapeRegexChar(char: string): string {
  return char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Convert a glob pattern to a RegExp.
 *
 * Supports:
 * - `*` - Match any characters except path separator
 * - `**` - Match any characters including path separator
 * - `?` - Match single character
 * - `[abc]` - Character class
 * - `{a,b,c}` - Alternation
 * - `!pattern` - Negation (when used in ignore list)
 *
 * Results are cached for performance.
 *
 * @param pattern - Glob pattern to convert
 * @param options - Conversion options
 * @returns Compiled RegExp
 */
export function globToRegex(pattern: string, options: { dot?: boolean } = {}): RegExp {
  const { dot = false } = options;

  // Check cache first
  const cacheKey = getCacheKey(pattern, dot);
  const cached = regexCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  // Compile the pattern
  let regex = "";
  let i = 0;
  let inBracket = false;
  let inBrace = false;

  while (i < pattern.length) {
    const char = pattern[i]!;

    if (inBracket) {
      if (char === "]") {
        regex += "]";
        inBracket = false;
      } else if (char === "\\") {
        regex += "\\";
        if (i + 1 < pattern.length) {
          regex += pattern[++i];
        }
      } else {
        regex += char;
      }
      i++;
      continue;
    }

    if (inBrace) {
      if (char === "}") {
        regex += ")";
        inBrace = false;
      } else if (char === ",") {
        regex += "|";
      } else {
        regex += escapeRegexChar(char);
      }
      i++;
      continue;
    }

    switch (char) {
      case "*":
        if (pattern[i + 1] === "*") {
          // ** - match anything including path separators
          if (pattern[i + 2] === "/" || pattern[i + 2] === "\\") {
            regex += "(?:.*[/\\\\])?";
            i += 3;
          } else {
            regex += ".*";
            i += 2;
          }
        } else {
          // * - match anything except path separators
          regex += "[^/\\\\]*";
          i++;
        }
        break;

      case "?":
        regex += "[^/\\\\]";
        i++;
        break;

      case "[":
        regex += "[";
        inBracket = true;
        i++;
        break;

      case "{":
        regex += "(?:";
        inBrace = true;
        i++;
        break;

      case "/":
      case "\\":
        regex += "[/\\\\]";
        i++;
        break;

      default:
        regex += escapeRegexChar(char);
        i++;
    }
  }

  // Handle dot files - insert negative lookahead at the start if needed
  if (!dot && regex.startsWith("[^/\\\\]*")) {
    regex = "(?!\\.)[^/\\\\]*" + regex.slice("[^/\\\\]*".length);
  }

  const result = new RegExp("^" + regex + "$", "i");

  // Cache the result (with size limit to prevent memory leaks)
  if (regexCache.size >= MAX_CACHE_SIZE) {
    // Clear oldest entries (simple approach: clear half the cache)
    const keys = Array.from(regexCache.keys());
    for (let j = 0; j < keys.length / 2; j++) {
      regexCache.delete(keys[j]!);
    }
  }
  regexCache.set(cacheKey, result);

  return result;
}

// =============================================================================
// Path Normalization
// =============================================================================

/**
 * Normalize path separators to forward slashes.
 *
 * This is used to ensure consistent path matching regardless of platform.
 *
 * @param filePath - Path to normalize
 * @returns Path with forward slashes
 */
export function normalizePath(filePath: string): string {
  return filePath.replace(/\\/g, "/");
}

// =============================================================================
// Pattern Matching
// =============================================================================

/**
 * Test if a path matches a glob pattern.
 *
 * @param filePath - File path to test
 * @param pattern - Glob pattern
 * @param options - Match options
 * @returns True if path matches pattern
 */
export function matchGlob(filePath: string, pattern: string, options?: { dot?: boolean }): boolean {
  const regex = globToRegex(pattern, options);
  const normalizedPath = normalizePath(filePath);
  return regex.test(normalizedPath);
}

/**
 * Test if a path matches any of the given patterns.
 *
 * This is more efficient than calling matchGlob multiple times
 * because it avoids normalizing the path repeatedly.
 *
 * @param filePath - File path to test
 * @param patterns - Array of glob patterns
 * @param options - Match options
 * @returns True if path matches any pattern
 */
export function matchGlobAny(
  filePath: string,
  patterns: readonly string[],
  options?: { dot?: boolean }
): boolean {
  if (patterns.length === 0) {
    return false;
  }

  // Normalize path once
  const normalizedPath = normalizePath(filePath);

  // Test against all patterns
  for (const pattern of patterns) {
    const regex = globToRegex(pattern, options);
    if (regex.test(normalizedPath)) {
      return true;
    }
  }

  return false;
}

/**
 * Create a pre-compiled matcher for multiple patterns.
 *
 * This is useful when you need to match many files against the same
 * set of patterns - it avoids repeated regex compilation.
 *
 * @param patterns - Array of glob patterns
 * @param options - Match options
 * @returns A function that tests if a path matches any pattern
 */
export function createGlobMatcher(
  patterns: string[],
  options?: { dot?: boolean }
): (filePath: string) => boolean {
  if (patterns.length === 0) {
    return () => false;
  }

  // Pre-compile all patterns
  const regexes = patterns.map(p => globToRegex(p, options));

  return (filePath: string) => {
    const normalizedPath = normalizePath(filePath);
    return regexes.some(r => r.test(normalizedPath));
  };
}

/**
 * Clear the regex cache.
 *
 * This is mainly useful for testing or memory-sensitive applications.
 */
export function clearGlobCache(): void {
  regexCache.clear();
}
