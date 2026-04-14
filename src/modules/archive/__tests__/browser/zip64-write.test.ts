import { runZip64WriteTests } from "@archive/__tests__/zip/zip64-write.shared";
import { describe } from "vitest";

describe("ZIP64 write - Browser", () => {
  runZip64WriteTests();
});
