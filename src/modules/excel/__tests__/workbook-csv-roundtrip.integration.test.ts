import { testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook, ValueType } from "../../../index";

describe("Workbook", () => {
  describe("CSV", () => {
    it("roundtrips numbers/strings/formulas/dates (from legacy examples)", async () => {
      const csvFile = testFilePath("csv-roundtrip-from-examples", ".csv");

      // Write a CSV (mirrors former src/examples/testCsvOut.ts)
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      ws.columns = [
        { header: "Col 1", key: "key", width: 25 },
        { header: "Col 2", key: "name", width: 25 },
        { header: "Col 3", key: "age", width: 21 },
        { header: "Col 4", key: "addr1", width: 18 },
        { header: "Col 5", key: "addr2", width: 8 }
      ];

      ws.getCell("A2").value = 7;
      ws.getCell("B2").value = "Hello, World!";
      ws.getCell("C2").value = -5.55;
      ws.getCell("D2").value = new Date(2015, 2, 10, 7, 8, 9);
      ws.getCell("E2").value = "Hello, World!";

      ws.getCell("A3").value = { text: "www.google.com", hyperlink: "http://www.google.com" };
      ws.getCell("A4").value = "Boo!";
      ws.getCell("C4").value = "Hoo!";

      ws.getCell("A5").value = 1;
      ws.getCell("B5").value = 2;
      ws.getCell("C5").value = { formula: "A5+B5", result: 3 };

      ws.getCell("A6").value = "Hello";
      ws.getCell("B6").value = "World";
      ws.getCell("C6").value = {
        formula: "CONCATENATE(A6,', ',B6,'!')",
        result: "Hello, World!"
      };

      ws.getCell("A7").value = 1;
      ws.getCell("B7").value = 2;
      // C7 intentionally left blank
      ws.getCell("D7").value = 4;

      ws.getCell("A10").value = "<";
      ws.getCell("B10").value = ">";
      ws.getCell("C10").value = "<a>";
      ws.getCell("D10").value = "><";

      await wb.writeCsvFile(csvFile, { dateFormat: "DD/MM/YYYY HH:mm:ss" });

      // Read it back (mirrors former src/examples/testCsvIn.ts)
      const wb2 = new Workbook();
      await wb2.readCsvFile(csvFile, { dateFormats: ["DD/MM/YYYY HH:mm:ss"] });

      const ws2 = wb2.getWorksheet()!;
      expect(ws2).toBeTruthy();

      expect(ws2.getCell("A2").type).toBe(ValueType.Number);
      expect(ws2.getCell("A2").value).toBe(7);

      expect(ws2.getCell("B2").type).toBe(ValueType.String);
      expect(ws2.getCell("B2").value).toBe("Hello, World!");

      expect(ws2.getCell("C2").type).toBe(ValueType.Number);
      expect(Math.abs((ws2.getCell("C2").value as number) + 5.55)).toBeLessThan(0.000001);

      const d2 = ws2.getCell("D2");
      expect(d2.type).toBe(ValueType.Date);
      expect(d2.value).toBeInstanceOf(Date);
      const date = d2.value as Date;
      expect(date.getFullYear()).toBe(2015);
      expect(date.getMonth()).toBe(2);
      expect(date.getDate()).toBe(10);
      expect(date.getHours()).toBe(7);
      expect(date.getMinutes()).toBe(8);
      expect(date.getSeconds()).toBe(9);

      expect(ws2.getCell("C5").type).toBe(ValueType.Number);
      expect(ws2.getCell("C5").value).toBe(3);

      expect(ws2.getCell("A7").type).toBe(ValueType.Number);
      expect(ws2.getCell("A7").value).toBe(1);

      expect(ws2.getCell("B7").type).toBe(ValueType.Number);
      expect(ws2.getCell("B7").value).toBe(2);

      expect(ws2.getCell("C7").type).toBe(ValueType.Null);
      expect(ws2.getCell("C7").value).toBe(null);

      expect(ws2.getCell("D7").type).toBe(ValueType.Number);
      expect(ws2.getCell("D7").value).toBe(4);

      expect(ws2.getCell("A10").value).toBe("<");
      expect(ws2.getCell("B10").value).toBe(">");
      expect(ws2.getCell("C10").value).toBe("<a>");
      expect(ws2.getCell("D10").value).toBe("><");
    }, 6000);
  });
});
