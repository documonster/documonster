/**
 * CSV Scanner Module
 *
 * High-performance CSV field scanner using indexOf-based batch scanning.
 * Re-exports all scanner functionality from scanner.ts.
 */

// Types
export type { ScannerConfig, RowScanResult, Scanner } from "@csv/parse/scanner/scanner";

// Constants and state factory
export { DEFAULT_SCANNER_CONFIG } from "@csv/parse/scanner/scanner";

// Core scanning functions
export { scanQuotedField, scanUnquotedField, scanRow } from "@csv/parse/scanner/scanner";

// Scanner factory and utilities
export { createScanner, scanAllRows, scanRowsAsync } from "@csv/parse/scanner/scanner";
