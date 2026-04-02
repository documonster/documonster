import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { describe } from "vitest";
import { testXformHelper } from "@excel/xlsx/__tests__/xform/test-xform-helper";
import { DrawingXform } from "@excel/xlsx/xform/drawing/drawing-xform";
import { drawing10 } from "@excel/xlsx/__tests__/xform/drawing/data/drawing.1.0";
import { drawing11 } from "@excel/xlsx/__tests__/xform/drawing/data/drawing.1.1";
import { drawing13 } from "@excel/xlsx/__tests__/xform/drawing/data/drawing.1.3";
import { drawing14 } from "@excel/xlsx/__tests__/xform/drawing/data/drawing.1.4";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const options = {
  rels: {
    rId1: { Target: "../media/image1.jpg" },
    rId2: { Target: "../media/image2.jpg" }
  },
  mediaIndex: { image1: 0, image2: 1 },
  media: [{}, {}]
};

const expectations = [
  {
    title: "Drawing 1",
    create() {
      return new DrawingXform();
    },
    initialModel: drawing10,
    preparedModel: drawing11,
    xml: fs.readFileSync(path.join(__dirname, "data", "drawing.1.2.xml")).toString(),
    parsedModel: drawing13,
    reconciledModel: drawing14,
    // XML comparison can handle element ordering differences
    tests: ["prepare", "render", "renderIn", "parse", "reconcile"],
    options
  }
];

describe("DrawingXform", () => {
  testXformHelper(expectations);
});
