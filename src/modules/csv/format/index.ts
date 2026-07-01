/**
 * CSV Format Module - Public Exports
 *
 * Provides all CSV formatting functionality:
 * - formatCsv: Main batch formatting function
 * - Low-level utilities for streaming formatters
 */

// =============================================================================
// Configuration
// =============================================================================

export type { FormatConfig, FormatRowOptions } from "@csv/format/config";
export { createFormatConfig } from "@csv/format/config";

// =============================================================================
// Formatting
// =============================================================================

export { formatCsv, formatRowWithLookup } from "@csv/format/formatter";
