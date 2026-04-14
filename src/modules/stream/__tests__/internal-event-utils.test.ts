/**
 * Tests for internal/event-utils.ts (Node.js runner)
 *
 * Tests for onceEvent() promise-based event listener.
 * Runs the shared test suite against Node.js runtime.
 */

import { describe } from "vitest";

import { runInternalEventUtilsTests } from "./internal-event-utils.shared";

describe("internal/event-utils (Node)", () => {
  runInternalEventUtilsTests();
});
