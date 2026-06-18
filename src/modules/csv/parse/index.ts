/**
 * CSV Parse Module - Public Exports
 *
 * Provides all CSV parsing functionality:
 * - parseCsv: Main synchronous parsing function
 * - parseFastMode, parseWithScanner: Low-level parsing generators
 * - Configuration and state types/factories
 */

// =============================================================================
// Synchronous Parsing
// =============================================================================

export { parseCsv, parseFastMode, parseWithScanner } from "@csv/parse/sync";

// =============================================================================
// Configuration
// =============================================================================

export type { ParseConfig, CreateParseConfigOptions, ParseConfigResult } from "@csv/parse/config";
export { createParseConfig, resolveParseConfig, makeTrimField } from "@csv/parse/config";

// =============================================================================
// State Management
// =============================================================================

export type { ParseState } from "@csv/parse/state";
export { createParseState, resetInfoState } from "@csv/parse/state";

// =============================================================================
// Row Processing
// =============================================================================

export type { RowProcessResult } from "@csv/parse/row-processor";
export { processCompletedRow, shouldSkipRow } from "@csv/parse/row-processor";

// =============================================================================
// Asynchronous Parsing
// =============================================================================

export { parseCsvAsync, parseCsvRows, parseCsvWithProgress } from "@csv/parse/async";
