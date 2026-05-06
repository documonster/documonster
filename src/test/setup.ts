/**
 * Vitest global setup — installs the optional formula engine so the
 * general suite can call `Workbook.calculateFormulas()` without each
 * test file wiring it up.
 *
 * Chart support is NOT installed globally — only the 16 chart test
 * files install it locally via `beforeAll(installChartSupport)`. This
 * keeps the 324 non-chart test files from paying the ~30k-line chart
 * module transform/evaluate cost on every run.
 */

import { installFormulaEngine } from "@formula/install";

installFormulaEngine();
