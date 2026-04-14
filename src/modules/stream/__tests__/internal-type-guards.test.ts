/**
 * Tests for internal/type-guards.ts (Node.js runner)
 *
 * Tests for Web Streams API type guard functions.
 * Runs the shared test suite against Node.js runtime.
 */

import { describe } from "vitest";

import { runInternalTypeGuardsTests } from "./internal-type-guards.shared";

describe("internal/type-guards (Node)", () => {
  runInternalTypeGuardsTests();
});
