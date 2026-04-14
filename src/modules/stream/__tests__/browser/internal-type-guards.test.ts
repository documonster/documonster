/**
 * Internal Type Guards Browser Tests
 *
 * Runs the shared internal type guards test suite against the Browser implementation.
 */

import { runInternalTypeGuardsTests } from "@stream/__tests__/internal-type-guards.shared";
import { describe } from "vitest";

describe("internal/type-guards (Browser)", () => {
  runInternalTypeGuardsTests();
});
