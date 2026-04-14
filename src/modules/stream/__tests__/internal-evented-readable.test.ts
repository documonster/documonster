/**
 * Tests for internal/evented-readable-to-async-iterable.ts (Node.js runner)
 *
 * Tests for eventedReadableToAsyncIterableNoDestroy().
 * Runs the shared test suite against Node.js runtime.
 */

import { describe } from "vitest";

import { runInternalEventedReadableTests } from "./internal-evented-readable.shared";

describe("internal/evented-readable-to-async-iterable (Node)", () => {
  runInternalEventedReadableTests();
});
