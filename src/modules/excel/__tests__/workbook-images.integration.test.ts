import fs from "fs";
import { promisify } from "util";

import { makeTestDataPath, testFilePath } from "@test/utils";
import { describe, it, expect } from "vitest";

import { Workbook } from "../../../index";

const excelTestDataPath = makeTestDataPath(import.meta.url, "./data");

const IMAGE_FILENAME = excelTestDataPath("image.png");

const TEST_XLSX_FILE_NAME = testFilePath("workbook-images.test");
const fsReadFileAsync = promisify(fs.readFile);

// =============================================================================
// Tests

describe("Workbook", () => {
  describe("Images", () => {
    it("stores background image", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");
      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.getCell("A1").value = "Hello, World!";
      ws.addBackgroundImage(imageId);

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const backgroundId2 = ws2!.getBackgroundImageId();
      const image = wb2.getImage(backgroundId2!);

      expect(Buffer.compare(imageData, image!.buffer as Uint8Array)).toBe(0);
    });

    it("stores embedded image and hyperlink", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.getCell("A1").value = "Hello, World!";
      ws.getCell("A2").value = {
        hyperlink: "http://www.somewhere.com",
        text: "www.somewhere.com"
      };
      ws.addImage(imageId, "C3:E6");

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      expect(ws.getCell("A1").value).toBe("Hello, World!");
      expect(ws.getCell("A2").value).toEqual({
        hyperlink: "http://www.somewhere.com",
        text: "www.somewhere.com"
      });

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const images = ws2!.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.tl.col).toBe(2);
      expect(imageDesc.range!.tl.row).toBe(2);
      expect(imageDesc.range!.br!.col).toBe(5);
      expect(imageDesc.range!.br!.row).toBe(6);

      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer as Uint8Array)).toBe(0);
    });

    it("stores embedded image with oneCell", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        br: { col: 2.101046875, row: 3.4 },
        editAs: "oneCell"
      });

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const images = ws2!.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("oneCell");

      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer as Uint8Array)).toBe(0);
    });

    it("stores embedded image with one-cell-anchor", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 },
        editAs: "oneCell"
      });

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const images = ws2!.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("oneCell");
      expect(imageDesc.range!.ext!.width).toBe(100);
      expect(imageDesc.range!.ext!.height).toBe(100);

      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer as Uint8Array)).toBe(0);
    });

    it("stores embedded image with hyperlinks", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const imageId = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.addImage(imageId, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 },
        editAs: "absolute",
        hyperlinks: {
          hyperlink: "http://www.somewhere.com",
          tooltip: "www.somewhere.com"
        }
      });

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const images = ws2!.getImages();
      expect(images.length).toBe(1);

      const imageDesc = images[0];
      expect(imageDesc.range!.editAs).toBe("absolute");
      expect(imageDesc.range!.ext!.width).toBe(100);
      expect(imageDesc.range!.ext!.height).toBe(100);

      expect(imageDesc.range!.hyperlinks).toEqual({
        hyperlink: "http://www.somewhere.com",
        tooltip: "www.somewhere.com"
      });

      const image = wb2.getImage(imageDesc.imageId!);
      expect(Buffer.compare(imageData, image!.buffer as Uint8Array)).toBe(0);
    });

    it("image extensions should not be case sensitive", async () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet("blort");

      const imageId1 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "png"
      });

      const imageId2 = wb.addImage({
        filename: IMAGE_FILENAME,
        extension: "jpeg"
      });

      ws.addImage(imageId1, {
        tl: { col: 0.1125, row: 0.4 },
        ext: { width: 100, height: 100 }
      });

      ws.addImage(imageId2, {
        tl: { col: 0.1125, row: 0.4 },
        br: { col: 2.101046875, row: 3.4 },
        editAs: "oneCell"
      });

      await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

      const wb2 = new Workbook();
      await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);

      const ws2 = wb2.getWorksheet("blort");
      expect(ws2).toBeDefined();

      const imageData = await fsReadFileAsync(IMAGE_FILENAME);

      const images = ws2!.getImages();
      expect(images.length).toBe(2);

      const imageDesc1 = images[0];
      expect(imageDesc1.range!.ext!.width).toBe(100);
      expect(imageDesc1.range!.ext!.height).toBe(100);
      const image1 = wb2.getImage(imageDesc1.imageId!);

      const imageDesc2 = images[1];
      expect(imageDesc2.range!.editAs).toBe("oneCell");

      const image2 = wb2.getImage(imageDesc1.imageId!);

      expect(Buffer.compare(imageData, image1!.buffer!)).toBe(0);
      expect(Buffer.compare(imageData, image2!.buffer!)).toBe(0);
    });

    describe("read-write round-trip (issue #58)", () => {
      it("does not duplicate images after read-write cycles", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, { tl: { col: 1, row: 0 }, br: { col: 2, row: 1 } });

        // First write
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        // Read back and write again
        await wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        // Read the final file and verify images are not duplicated
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1");
        expect(ws2).toBeDefined();

        const images = ws2!.getImages();
        expect(images.length).toBe(1);
      });

      it("does not duplicate images after multiple read-write cycles", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "B2:D4");

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        // Perform 3 read-write cycles on the same workbook
        for (let i = 0; i < 3; i++) {
          await wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
          await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        }

        // Read the final file with a fresh workbook
        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1");
        expect(ws2).toBeDefined();

        const images = ws2!.getImages();
        expect(images.length).toBe(1);
      });

      it("does not duplicate when multiple images exist", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId1, "A1:B2");
        ws.addImage(imgId2, "C3:D4");

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        // Read-write cycle
        await wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1");
        expect(ws2).toBeDefined();

        const images = ws2!.getImages();
        expect(images.length).toBe(2);
      });

      it("preserves image data through read-write cycle", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "C3:E6");

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);
        await wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1");
        const images = ws2!.getImages();
        expect(images.length).toBe(1);

        const imageData = await fsReadFileAsync(IMAGE_FILENAME);
        const image = wb2.getImage(images[0].imageId!);
        expect(Buffer.compare(imageData, image!.buffer!)).toBe(0);
      });
    });

    describe("image range updates on row/column splice", () => {
      it("updates image range after insertRow", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "B2:D4");

        // Insert a row before the image
        ws.insertRow(1, []);

        const images = ws.getImages();
        expect(images.length).toBe(1);
        const img = images[0];
        // Image should shift down by 1 row (B2:D4 -> B3:D5)
        // nativeRow is 0-based: row 2 -> nativeRow 1, after insert -> nativeRow 2
        expect(img.range!.tl.nativeRow).toBe(2);
        expect(img.range!.tl.nativeCol).toBe(1);
        expect(img.range!.br!.nativeRow).toBe(5);
        expect(img.range!.br!.nativeCol).toBe(4);
      });

      it("does not update image range when inserting row after the image", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "A1:B2");

        // Insert a row after the image
        ws.insertRow(5, []);

        const images = ws.getImages();
        const img = images[0];
        // Image should not move (A1:B2 stays the same)
        // nativeRow for A1 with string range uses offset -1: row=1 -> nativeRow=0
        expect(img.range!.tl.nativeRow).toBe(0);
        expect(img.range!.tl.nativeCol).toBe(0);
        expect(img.range!.br!.nativeRow).toBe(2);
        expect(img.range!.br!.nativeCol).toBe(2);
      });

      it("updates image range after spliceRows with remove", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "A3:B4");

        // Remove 1 row at row 1
        ws.spliceRows(1, 1);

        const images = ws.getImages();
        const img = images[0];
        // Image should shift up by 1 row (A3:B4 -> A2:B3)
        // nativeRow: row 3 -> nativeRow 2, after remove -> nativeRow 1
        expect(img.range!.tl.nativeRow).toBe(1);
        expect(img.range!.br!.nativeRow).toBe(3);
      });

      it("updates image range after spliceColumns with insert", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, "B1:C2");

        // Insert a column before column B
        ws.spliceColumns(1, 0, []);

        const images = ws.getImages();
        const img = images[0];
        // Image should shift right by 1 column (B1:C2 -> C1:D2)
        // tl: col=2 with offset -1 -> nativeCol=1, after insert -> nativeCol=2
        // br: col=3 with offset 0 -> nativeCol=3, after insert -> nativeCol=4
        expect(img.range!.tl.nativeCol).toBe(2);
        expect(img.range!.br!.nativeCol).toBe(4);
      });

      it("handles multiple images correctly during row splice", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId1, "A1:A1");
        ws.addImage(imgId2, "A3:B4");

        // Insert 2 rows at row 2
        ws.spliceRows(2, 0, [], []);

        const images = ws.getImages();
        // First image at A1 should not move (nativeRow 0 < start-1 = 1)
        expect(images[0].range!.tl.nativeRow).toBe(0);
        // Second image at A3 should shift down by 2 (nativeRow 2 >= start-1 = 1)
        expect(images[1].range!.tl.nativeRow).toBe(4);
        expect(images[1].range!.br!.nativeRow).toBe(6);
      });

      it("does not update background images during splice", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addBackgroundImage(imgId);

        // Should not throw
        ws.insertRow(1, []);

        // Background image should still exist
        expect(ws.getBackgroundImageId()).toBeDefined();
      });
    });

    describe("image duplication during row duplication (issue #57)", () => {
      it("duplicates images anchored to the source row", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, { tl: { col: 1, row: 0 }, br: { col: 2, row: 1 } });
        ws.getCell("A1").value = "Row 1";

        // Duplicate row 1 twice (creates rows 2 and 3 as copies)
        ws.duplicateRow(1, 2);

        const images = ws.getImages();
        // Original + 2 clones = 3 images
        expect(images.length).toBe(3);

        // Original stays at row 0 (0-based)
        expect(images[0].range!.tl.nativeRow).toBe(0);
        // Clone 1 at row 1 (0-based)
        expect(images[1].range!.tl.nativeRow).toBe(1);
        // Clone 2 at row 2 (0-based)
        expect(images[2].range!.tl.nativeRow).toBe(2);
      });

      it("preserves two-cell anchor span when duplicating", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        // Image spans 2 rows: tl at row 0, br at row 2
        ws.addImage(imgId, {
          tl: { col: 0, row: 0 },
          br: { col: 2, row: 2 }
        });

        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        expect(images.length).toBe(2);

        // Original: tl row 0, br row 2
        expect(images[0].range!.tl.nativeRow).toBe(0);
        expect(images[0].range!.br!.nativeRow).toBe(2);

        // Clone: tl row 1, br row 3 (same 2-row span)
        expect(images[1].range!.tl.nativeRow).toBe(1);
        expect(images[1].range!.br!.nativeRow).toBe(3);
      });

      it("duplicates one-cell anchor images (ext-based, no br)", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, {
          tl: { col: 0, row: 0 },
          ext: { width: 200, height: 150 }
        });

        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        expect(images.length).toBe(2);

        // Clone should have the same ext and no br
        const cloned = images[1];
        expect(cloned.range!.tl.nativeRow).toBe(1);
        expect(cloned.range!.br).toBeUndefined();
        expect(cloned.range!.ext).toEqual({ width: 200, height: 150 });
      });

      it("duplicates multiple images on the same source row", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        // Two images on row 1 (0-based row 0), different columns
        ws.addImage(imgId1, {
          tl: { col: 0, row: 0 },
          br: { col: 1, row: 1 }
        });
        ws.addImage(imgId2, {
          tl: { col: 3, row: 0 },
          br: { col: 4, row: 1 }
        });

        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        // 2 originals + 2 clones = 4
        expect(images.length).toBe(4);

        // Clones should be at row 1 (0-based), same columns as originals
        expect(images[2].range!.tl.nativeRow).toBe(1);
        expect(images[2].range!.tl.nativeCol).toBe(0);
        expect(images[3].range!.tl.nativeRow).toBe(1);
        expect(images[3].range!.tl.nativeCol).toBe(3);
      });

      it("does not clone images from other rows", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        // Image on row 1 (0-based row 0)
        ws.addImage(imgId1, {
          tl: { col: 0, row: 0 },
          br: { col: 1, row: 1 }
        });
        // Image on row 5 (0-based row 4)
        ws.addImage(imgId2, {
          tl: { col: 0, row: 4 },
          br: { col: 1, row: 5 }
        });

        // insert=true so spliceRows inserts a new row, shifting images below
        ws.duplicateRow(1, 1, true);

        const images = ws.getImages();
        // 2 originals + 1 clone (from row 1 only) = 3
        expect(images.length).toBe(3);

        // The row-5 image should be shifted down by 1 (spliceRows effect)
        // but not duplicated. Original was nativeRow 4, after splice it becomes 5
        const row5Image = images[1]; // the second original
        expect(row5Image.range!.tl.nativeRow).toBe(5);
      });

      it("preserves hyperlinks on cloned images", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.addImage(imgId, {
          tl: { col: 0, row: 0 },
          ext: { width: 100, height: 100 },
          hyperlinks: {
            hyperlink: "http://example.com",
            tooltip: "Example"
          }
        });

        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        expect(images.length).toBe(2);

        const cloned = images[1];
        expect(cloned.range!.hyperlinks).toEqual({
          hyperlink: "http://example.com",
          tooltip: "Example"
        });
      });

      it("round-trips duplicated images through write/read", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.getCell("A1").value = "Hello";
        ws.addImage(imgId, {
          tl: { col: 1, row: 0 },
          br: { col: 2, row: 1 }
        });

        ws.duplicateRow(1, 2);

        // Write and read back
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1");
        expect(ws2).toBeDefined();

        const images = ws2!.getImages();
        expect(images.length).toBe(3);

        // Verify positions survived round-trip
        expect(images[0].range!.tl.nativeRow).toBe(0);
        expect(images[1].range!.tl.nativeRow).toBe(1);
        expect(images[2].range!.tl.nativeRow).toBe(2);

        // Verify image data is intact
        const imageData = await fsReadFileAsync(IMAGE_FILENAME);
        for (const img of images) {
          const imgBuffer = wb2.getImage(img.imageId!);
          expect(Buffer.compare(imageData, imgBuffer!.buffer!)).toBe(0);
        }
      });

      it("works with insert mode", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        ws.getCell("A1").value = "Row 1";
        ws.getCell("A2").value = "Row 2";
        ws.addImage(imgId, {
          tl: { col: 1, row: 0 },
          br: { col: 2, row: 1 }
        });

        // insert=true: inserts new rows instead of overwriting
        ws.duplicateRow(1, 1, true);

        const images = ws.getImages();
        expect(images.length).toBe(2);

        // Original at row 0
        expect(images[0].range!.tl.nativeRow).toBe(0);
        // Clone at row 1
        expect(images[1].range!.tl.nativeRow).toBe(1);
      });

      it("loads test file and duplicates row with images", async () => {
        const wb = new Workbook();
        await wb.xlsx.readFile(excelTestDataPath("duplicate-row-images.xlsx"));

        // The test file has a drawing on sheet3 with an image at row 12
        const ws = wb.getWorksheet("Sheet3");
        expect(ws).toBeDefined();

        const imagesBefore = ws!.getImages();
        const countBefore = imagesBefore.length;
        expect(countBefore).toBeGreaterThan(0);

        // Find which row has the image (0-based nativeRow -> 1-based rowNum)
        const srcRow0 = imagesBefore[0].range!.tl.nativeRow;
        const srcRowNum = srcRow0 + 1;

        // Duplicate that row once (insert mode to push existing rows down)
        ws!.duplicateRow(srcRowNum, 1, true);

        const imagesAfter = ws!.getImages();
        // Each image on the source row should be cloned once
        const imagesOnSrcRow = imagesBefore.filter(
          img => img.range && img.range.tl.nativeRow === srcRow0
        );
        expect(imagesAfter.length).toBe(countBefore + imagesOnSrcRow.length);

        // Write and read back to verify integrity
        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet3");
        const finalImages = ws2!.getImages();
        expect(finalImages.length).toBe(countBefore + imagesOnSrcRow.length);
      });

      it("overwrite mode removes images on target rows before cloning", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });

        // Image on row 1 (0-based row 0)
        ws.addImage(imgId1, {
          tl: { col: 0, row: 0 },
          br: { col: 1, row: 1 }
        });
        // Image on row 2 (0-based row 1) — this will be overwritten
        ws.addImage(imgId2, {
          tl: { col: 2, row: 1 },
          br: { col: 3, row: 2 }
        });

        // Overwrite mode (default): duplicate row 1 once, overwriting row 2
        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        // Original on row 1 + clone on row 2 = 2 (old row-2 image removed)
        expect(images.length).toBe(2);

        // Original stays at row 0
        expect(images[0].range!.tl.nativeRow).toBe(0);
        expect(images[0].range!.tl.nativeCol).toBe(0);

        // Clone at row 1 (0-based), same column as the source
        expect(images[1].range!.tl.nativeRow).toBe(1);
        expect(images[1].range!.tl.nativeCol).toBe(0);
      });

      it("overwrite mode with no source images still removes target row images", () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");
        const imgId = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });

        ws.getCell("A1").value = "Row 1 (no images)";
        // Image on row 2 (0-based row 1)
        ws.addImage(imgId, {
          tl: { col: 0, row: 1 },
          br: { col: 1, row: 2 }
        });

        // Overwrite mode: duplicate row 1 once, overwriting row 2
        ws.duplicateRow(1, 1);

        const images = ws.getImages();
        // Old row-2 image should be removed since the row was overwritten
        expect(images.length).toBe(0);
      });
    });

    describe("image deduplication", () => {
      it("deduplicates drawing rels for non-consecutive same imageId", async () => {
        const wb = new Workbook();
        const ws = wb.addWorksheet("Sheet1");

        // Add two different images and use the first one again (non-consecutive)
        const imgId1 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });
        const imgId2 = wb.addImage({
          filename: IMAGE_FILENAME,
          extension: "png"
        });

        // Pattern: imgId1, imgId2, imgId1 — tests non-consecutive dedup
        ws.addImage(imgId1, "A1:B2");
        ws.addImage(imgId2, "C3:D4");
        ws.addImage(imgId1, "E5:F6");

        await wb.xlsx.writeFile(TEST_XLSX_FILE_NAME);

        const wb2 = new Workbook();
        await wb2.xlsx.readFile(TEST_XLSX_FILE_NAME);
        const ws2 = wb2.getWorksheet("Sheet1")!;

        const images = ws2.getImages();
        expect(images.length).toBe(3);

        // All 3 images should be valid and readable
        const imageData = await fsReadFileAsync(IMAGE_FILENAME);
        for (const img of images) {
          const imgBuffer = wb2.getImage(img.imageId!);
          expect(imgBuffer).toBeDefined();
          expect(Buffer.compare(imageData, imgBuffer!.buffer!)).toBe(0);
        }
      });
    });
  });
});
