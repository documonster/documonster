import { readFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { CoreXform } from "@excel/xlsx/xform/core/core-xform";
import { describe } from "vitest";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const expectations: any[] = [
  {
    title: "core.xml",
    create: () => new CoreXform(),
    preparedModel: {
      creator: "Test Creator",
      lastModifiedBy: "Test Creator",
      created: new Date("2016-04-20T16:26:46Z"),
      modified: new Date("2016-05-12T06:52:49Z")
    },
    xml: readFileSync(join(__dirname, "./data/core.01.xml")).toString().replace(/\r\n/g, "\n"),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "core.xml - with metadata",
    create: () => new CoreXform(),
    preparedModel: {
      creator: "Test Creator",
      title: "My Little Xlsx",
      subject: "An Xlsx about Xlsxs",
      description: "A lot of stuff",
      keywords: "xlsx,test",
      category: "unit test material",
      lastModifiedBy: "Test Creator",
      revision: 1,
      created: new Date("2016-04-20T16:26:46Z"),
      modified: new Date("2016-05-12T06:52:49Z")
    },
    xml: readFileSync(join(__dirname, "./data/core.02.xml")).toString().replace(/\r\n/g, "\n"),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "core.xml - with cp:lastPrinted",
    create: () => new CoreXform(),
    preparedModel: {
      creator: "Test Creator",
      lastModifiedBy: "Test Creator",
      lastPrinted: new Date("2016-08-16T19:56:07Z"),
      created: new Date("2016-04-20T16:26:46Z"),
      modified: new Date("2016-05-12T06:52:49Z")
    },
    xml: readFileSync(join(__dirname, "./data/core.03.xml")).toString().replace(/\r\n/g, "\n"),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "core.xml - with cp:contentStatus",
    create: () => new CoreXform(),
    preparedModel: {
      creator: "Test Creator",
      lastModifiedBy: "Test Creator",
      contentStatus: "Final",
      created: new Date("2016-04-20T16:26:46Z"),
      modified: new Date("2016-05-12T06:52:49Z")
    },
    xml: readFileSync(join(__dirname, "./data/core.04.xml")).toString().replace(/\r\n/g, "\n"),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ["render", "renderIn", "parse"]
  },
  {
    title: "core.xml - with empty cp:version",
    create: () => new CoreXform(),
    preparedModel: {},
    xml: readFileSync(join(__dirname, "./data/core.05.xml")).toString().replace(/\r\n/g, "\n"),
    parsedModel: {
      title: "...",
      creator: "...",
      lastModifiedBy: "...",
      lastPrinted: new Date("2017-05-15T16:17:00Z"),
      created: new Date("2015-07-15T16:27:34Z"),
      modified: new Date("2017-09-06T15:39:12Z")
    },
    tests: ["parse"]
  },
  {
    title: "core.xml - without namespace for coreProperties node",
    create: () => new CoreXform(),
    preparedModel: {},
    xml: readFileSync(join(__dirname, "./data/core.06.xml")).toString().replace(/\r\n/g, "\n"),
    parsedModel: {
      creator: "Apache POI",
      created: new Date("2018-05-08T14:56:50Z")
    },
    tests: ["parse"]
  }
];

describe("CoreXform", () => {
  testXformHelper(expectations);
});
