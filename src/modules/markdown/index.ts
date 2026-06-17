/**
 * Markdown Module - Public API
 *
 * Pure Markdown table parsing/formatting functionality with no Excel dependencies.
 * For Markdown-Worksheet integration, use Workbook.readMarkdown/writeMarkdown methods instead.
 *
 * Design principles:
 * - Only export types and functions that are part of the PUBLIC API
 * - Internal utilities are used internally but not exported
 * - This reduces bundle size and simplifies the public interface
 */

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
  MarkdownFormatOptions,

  // Workbook integration types
  MarkdownOptions
} from "./types";

// =============================================================================
// Core Functions — `Markdown` domain namespace (tree-shaken via `export * as`)
// =============================================================================

export * as Markdown from "./surface/markdown";

// =============================================================================
// Errors
// =============================================================================

export { MarkdownError, MarkdownParseError } from "./errors";
