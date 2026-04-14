/**
 * Tests for stream/errors.ts (Node.js runner)
 *
 * Tests for stream-specific error classes.
 * Runs the shared test suite against Node.js runtime.
 */

import { describe } from "vitest";

import { runStreamErrorsTests } from "./stream-errors.shared";

describe("stream/errors (Node)", () => {
  runStreamErrorsTests();
});
