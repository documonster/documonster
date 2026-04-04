/**
 * Markdown Module - Public API
 *
 * Pure Markdown table parsing/formatting functionality with no Excel dependencies.
 * For Markdown-Worksheet integration, use Workbook.readMd/writeMd methods instead.
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
  MdAlignment,

  // Parse types
  MdParseResult,
  MdParseOptions,

  // Format types
  MdColumnConfig,
  MdFormatOptions,

  // Workbook integration types
  MdOptions
} from "./types";

// =============================================================================
// Core Functions
// =============================================================================

// Parser
export { parseMd, parseMdAll } from "./parse/index";

// Formatter
export { formatMd } from "./format/index";

// =============================================================================
// Errors
// =============================================================================

export { MdError, MdParseError } from "./errors";
