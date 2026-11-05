/**
 * Markdown Module - Public API
 *
 * Pure Markdown table parsing/formatting functionality with no Excel dependencies.
 * For Markdown-Worksheet integration, use `ExcelMarkdown` from `documonster/excel/markdown` instead.
 *
 * Design principles:
 * - Only export types and functions that are part of the PUBLIC API
 * - Internal utilities are used internally but not exported
 * - This reduces bundle size and simplifies the public interface
 */

import type { MarkdownOptions as MarkdownOptionsImpl } from "@markdown/types";

// =============================================================================
// Core Types
// =============================================================================

export type {
  // Alignment
  MarkdownAlignment,

  // Parse types
  MarkdownParseResult,
  MarkdownParseOptions,

  // Format types
  MarkdownColumnConfig,
  MarkdownFormatOptions
} from "@markdown/types";

/**
 * @deprecated Worksheet options do not belong to the standalone Markdown module. Use
 * `ExcelMarkdownReadOptions` / `ExcelMarkdownWriteOptions` from `documonster/excel/markdown`, which the
 * `ExcelMarkdown` functions take. Removed in 1.0.
 */
export type MarkdownOptions = MarkdownOptionsImpl;

// =============================================================================
// Core Functions — `Markdown` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Markdown from "@markdown/surface/markdown";

// =============================================================================
// Errors
// =============================================================================

export {
  MarkdownError,
  MarkdownArgumentError,
  MarkdownParseError,
  isMarkdownError
} from "@markdown/errors";

// The options every error constructor of this entry takes (`{ cause }`), named so a subclass can pass them on.
export type { BaseErrorOptions } from "@utils/errors";
