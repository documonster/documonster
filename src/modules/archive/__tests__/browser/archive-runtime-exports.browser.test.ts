import * as archive from "@archive";
import {
  ARCHIVE_BROWSER_EXPORTS,
  ARCHIVE_NAMESPACE_EXPORTS,
  getRuntimeExportKeys
} from "@archive/__tests__/runtime/archive-runtime-exports";
import { describe, expect, it } from "vitest";

describe("archive/index runtime exports (browser)", () => {
  it("should match the export contract", () => {
    const actual = getRuntimeExportKeys(archive);
    const expected = [...ARCHIVE_BROWSER_EXPORTS].sort();
    expect(actual).toEqual(expected);
  });

  it("should expose the expected `Archive` namespace members", () => {
    const actual = getRuntimeExportKeys(archive.Archive);
    const expected = [...ARCHIVE_NAMESPACE_EXPORTS].sort();
    expect(actual).toEqual(expected);
  });
});
