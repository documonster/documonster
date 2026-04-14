/**
 * Stream Utils Browser Tests
 *
 * Runs the shared stream utils test suite against the Browser implementation.
 */

import { runStreamUtilsTests } from "@stream/__tests__/stream-utils.shared";
import { describe } from "vitest";

describe("stream/utils (Browser)", () => {
  runStreamUtilsTests();
});
