import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { PageSetupPropertiesXform } from "@excel/xlsx/xform/sheet/page-setup-properties-xform";
import { describe } from "vitest";

const expectations = [
  {
    title: "fitToPage",
    create() {
      return new PageSetupPropertiesXform();
    },
    preparedModel: { fitToPage: true },
    xml: '<pageSetUpPr fitToPage="1"/>',
    parsedModel: { fitToPage: true },
    tests: ["render", "renderIn", "parse"]
  }
];

describe("PageSetupPropertiesXform", () => {
  testXformHelper(expectations);
});
