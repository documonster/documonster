import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { AppTitlesOfPartsXform } from "@excel/xlsx/xform/core/app-titles-of-parts-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "app.01",
    create() {
      return new AppTitlesOfPartsXform();
    },
    preparedModel: [{ name: "Sheet1" }],
    xml: '<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>',
    tests: ["render", "renderIn"]
  },
  {
    title: "app.02",
    create() {
      return new AppTitlesOfPartsXform();
    },
    preparedModel: [{ name: "Sheet1" }, { name: "Sheet2" }],
    xml: '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr></vt:vector></TitlesOfParts>',
    tests: ["render", "renderIn"]
  }
];

describe("AppTitlesOfPartsXform", () => {
  testXformHelper(expectations);
});
