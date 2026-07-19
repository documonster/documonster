import { Workbook, Worksheet } from "@excel/index";
import { buildSheetProtection, verifySheetPassword } from "@excel/utils/sheet-protection";
import { describe, expect, it } from "vitest";

describe("sheet password verification", () => {
  describe("verifySheetPassword", () => {
    it("accepts the correct password and rejects a wrong one", async () => {
      const protection = await buildSheetProtection("s3cret");
      expect(await verifySheetPassword(protection, "s3cret")).toBe(true);
      expect(await verifySheetPassword(protection, "nope")).toBe(false);
    });

    it("returns false when there is no stored hash to check against", async () => {
      // Protected without a password — nothing to verify.
      const noPassword = await buildSheetProtection();
      expect(await verifySheetPassword(noPassword, "anything")).toBe(false);
      expect(await verifySheetPassword(null, "anything")).toBe(false);
      expect(await verifySheetPassword(undefined, "anything")).toBe(false);
    });

    it("honours a custom spin count", async () => {
      const protection = await buildSheetProtection("pw", { spinCount: 5000 });
      expect(await verifySheetPassword(protection, "pw")).toBe(true);
      expect(await verifySheetPassword(protection, "PW")).toBe(false);
    });

    it("supports an explicitly empty password", async () => {
      const protection = await buildSheetProtection("");
      expect(protection.hashValue).toBeDefined();
      expect(await verifySheetPassword(protection, "")).toBe(true);
      expect(await verifySheetPassword(protection, "not-empty")).toBe(false);
    });
  });

  describe("Worksheet.verifyPassword", () => {
    it("verifies against a sheet protected with a password", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      await Worksheet.protect(ws, "letmein");

      expect(await Worksheet.verifyPassword(ws, "letmein")).toBe(true);
      expect(await Worksheet.verifyPassword(ws, "wrong")).toBe(false);
    });

    it("returns false for an unprotected sheet", async () => {
      const wb = Workbook.create();
      const ws = Workbook.addWorksheet(wb, "Sheet1");
      expect(await Worksheet.verifyPassword(ws, "anything")).toBe(false);
    });

    it("verifies a password against protection loaded back from a saved file", async () => {
      // The hash must survive a real read/write round-trip so verifyPassword
      // works on a workbook opened from disk, not just one protected in-memory.
      const wb = Workbook.create();
      await Worksheet.protect(Workbook.addWorksheet(wb, "Sheet1"), "diskpass");
      const buffer = await Workbook.toBuffer(wb);

      const reopened = Workbook.create();
      await Workbook.read(reopened, buffer);
      const ws = Workbook.getWorksheet(reopened, "Sheet1")!;
      expect(await Worksheet.verifyPassword(ws, "diskpass")).toBe(true);
      expect(await Worksheet.verifyPassword(ws, "wrong")).toBe(false);
    });
  });
});
