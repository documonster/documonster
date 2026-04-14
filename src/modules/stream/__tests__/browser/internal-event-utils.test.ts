/**
 * Internal Event Utils Browser Tests
 *
 * Runs the shared internal event utils test suite against the Browser implementation.
 */

import { runInternalEventUtilsTests } from "@stream/__tests__/internal-event-utils.shared";
import { describe } from "vitest";

describe("internal/event-utils (Browser)", () => {
  runInternalEventUtilsTests();
});
