import { zip, unzip } from "@archive";
import { runZipE2ETests, type ZipE2EModuleImports } from "@archive/__tests__/zip/zip-e2e.v2.shared";
import { describe } from "vitest";

describe("Archive - ZIP E2E (Browser)", () => {
  const imports: ZipE2EModuleImports = {
    zip: zip as ZipE2EModuleImports["zip"],
    unzip: unzip as ZipE2EModuleImports["unzip"]
  };

  runZipE2ETests(imports);
});
