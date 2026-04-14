/**
 * Stream Errors Browser Tests
 *
 * Runs the shared stream errors test suite against the Browser implementation.
 */

import { runStreamErrorsTests } from "@stream/__tests__/stream-errors.shared";
import { describe } from "vitest";

describe("stream/errors (Browser)", () => {
  runStreamErrorsTests();
});
