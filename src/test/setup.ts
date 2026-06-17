/**
 * Vitest global setup — installs the default formula syntax probe so the
 * general suite classifies defined-name text (formula vs. opaque) during
 * XLSX load without each test file wiring it up.
 *
 * Formula evaluation itself no longer needs any install step: call the
 * `calculateFormulas(workbook)` free function from `@formula/...` directly.
 *
 * Chart support likewise needs no install step — the chart implementation is
 * imported statically by the high-level chart APIs and tree-shaken out of
 * builds that never use them.
 */

import { installFormulaEngine } from "@formula/install";

installFormulaEngine();
