/**
 * Vitest global setup — installs the formula engine so the general
 * suite can call `Workbook.calculateFormulas()` without each test
 * file wiring it up.
 *
 * The "no engine installed" error path is still exercised, in
 * `src/modules/formula/__tests__/registry-contract.test.ts`, which
 * temporarily unregisters the engine via `registerFormulaEngine(null)`.
 *
 * Imports from `@formula/install` directly so this setup doesn't
 * unnecessarily pull in the `calculateFormulas`/`tokenize`/`parse`
 * exports from the public barrel.
 */

import { installFormulaEngine } from "@formula/install";

installFormulaEngine();
