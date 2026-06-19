import { readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { AppXform } from "@excel/xlsx/xform/core/app-xform";
import { describe } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const expectations = [
  {
    title: "app.01",
    create() {
      return new AppXform();
    },
    preparedModel: { worksheets: [{ name: "Sheet1" }] },
    xml: readFileSync(join(__dirname, "./data/app.01.xml")).toString().replace(/\r\n/g, "\n"),
    tests: ["render", "renderIn"]
  },
  {
    title: "app.02",
    create() {
      return new AppXform();
    },
    preparedModel: {
      worksheets: [{ name: "Sheet1" }, { name: "Sheet2" }],
      company: "Cyber Sapiens, Ltd.",
      manager: "Test Manager"
    },
    xml: readFileSync(join(__dirname, "./data/app.02.xml")).toString().replace(/\r\n/g, "\n"),
    tests: ["render", "renderIn"]
  }
];

describe("AppXform", () => {
  testXformHelper(expectations);
});
