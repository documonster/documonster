import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { SharedStringsXform } from "@excel/xlsx/xform/strings/shared-strings-xform";
import { describe } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Import JSON data
const sharedStringsData = JSON.parse(
  fs.readFileSync(path.join(__dirname, "data", "sharedStrings.json"), "utf8")
);

const expectations = [
  {
    title: "Shared Strings",
    create() {
      return new SharedStringsXform();
    },
    preparedModel: sharedStringsData,
    xml: fs.readFileSync(path.join(__dirname, "data", "sharedStrings.xml")).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("SharedStringsXform", () => {
  testXformHelper(expectations);
});
