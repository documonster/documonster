/**
 * Internal Evented Readable Browser Tests
 *
 * Runs the shared internal evented readable test suite against the Browser implementation.
 */

import { runInternalEventedReadableTests } from "@stream/__tests__/internal-evented-readable.shared";
import { describe } from "vitest";

describe("internal/evented-readable-to-async-iterable (Browser)", () => {
  runInternalEventedReadableTests();
});
