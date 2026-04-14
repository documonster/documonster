/**
 * Tests for stream/utils.ts (Node.js runner)
 *
 * Tests for cross-platform stream utility functions.
 * Runs the shared test suite against Node.js runtime.
 */

import { describe } from "vitest";

import { runStreamUtilsTests } from "./stream-utils.shared";

describe("stream/utils (Node)", () => {
  runStreamUtilsTests();
});
