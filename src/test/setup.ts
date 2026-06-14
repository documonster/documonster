/**
 * Vitest global setup — installs the default formula syntax probe so the
 * general suite classifies defined-name text (formula vs. opaque) during
 * XLSX load without each test file wiring it up.
 *
 * Formula evaluation itself no longer needs any install step: call the
 * `calculateFormulas(workbook)` free function from `@formula/...` directly.
 *
 * Chart support is NOT installed globally — only the chart test files
 * install it locally via `beforeAll(installChartSupport)`. This keeps the
 * non-chart test files from paying the ~30k-line chart module
 * transform/evaluate cost on every run.
 */

import { installFormulaEngine } from "@formula/install";

installFormulaEngine();
