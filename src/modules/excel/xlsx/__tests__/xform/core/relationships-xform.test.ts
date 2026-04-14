import { readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { RelationshipsXform } from "@excel/xlsx/xform/core/relationships-xform";
import { describe } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

import { worksheetRels1 } from "@excel/xlsx/__tests__/xform/core/data/worksheet.rels.1";

const expectations: any[] = [
  {
    title: "worksheet.rels",
    create() {
      return new RelationshipsXform();
    },
    preparedModel: worksheetRels1,
    xml: readFileSync(join(__dirname, "./data/worksheet.rels.xml"))
      .toString()
      .replace(/\r\n/g, "\n"),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("RelationshipsXform", () => {
  testXformHelper(expectations);
});
