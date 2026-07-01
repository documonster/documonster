import fs from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { TableXform } from "@excel/xlsx/xform/table/table-xform";
import { describe } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const expectations = [
  {
    title: "showing filter",
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: JSON.parse(fs.readFileSync(join(__dirname, "data/table.1.1.json")).toString()),
    xml: fs.readFileSync(join(__dirname, "data/table.1.2.xml")).toString(),
    parsedModel: JSON.parse(fs.readFileSync(join(__dirname, "data/table.1.3.json")).toString()),
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "table with calculatedColumnFormula child elements",
    create() {
      return new TableXform();
    },
    xml: fs.readFileSync(join(__dirname, "data/table.2.2.xml")).toString(),
    parsedModel: JSON.parse(fs.readFileSync(join(__dirname, "data/table.2.3.json")).toString()),
    tests: ["parse"]
  }
];

describe("TableXform", () => {
  testXformHelper(expectations);
});
