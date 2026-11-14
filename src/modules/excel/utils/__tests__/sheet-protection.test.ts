import { ExcelUnsupportedFeatureError } from "@excel/errors";
import { Workbook, Worksheet } from "@excel/index";
import { Encryptor } from "@excel/utils/encryptor";
import { buildSheetProtection, verifySheetPassword } from "@excel/utils/sheet-protection";
import { scheduleMacrotask } from "@utils/yield";
import { describe, expect, it, vi } from "vitest";

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

    // A file may name any algorithm ECMA-376 lists; one the platform has no digest for failed with a bare `Error`.
    it("refuses an algorithm the platform does not have as the module's error", async () => {
      const protection = {
        ...(await buildSheetProtection("pw", { spinCount: 1 })),
        algorithmName: "WHIRLPOOL"
      };
      await expect(verifySheetPassword(protection, "pw")).rejects.toThrow(
        ExcelUnsupportedFeatureError
      );
      await expect(verifySheetPassword(protection, "pw")).rejects.toThrow(
        'the algorithm "WHIRLPOOL"'
      );
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
      await Worksheet.protect(ws, { password: "letmein" });

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
      await Worksheet.protect(Workbook.addWorksheet(wb, "Sheet1"), { password: "diskpass" });
      const buffer = await Workbook.toBuffer(wb);

      const reopened = Workbook.create();
      await Workbook.read(reopened, buffer);
      const ws = Workbook.getWorksheet(reopened, "Sheet1")!;
      expect(await Worksheet.verifyPassword(ws, "diskpass")).toBe(true);
      expect(await Worksheet.verifyPassword(ws, "wrong")).toBe(false);
    });
  });
});

describe("the protection password hash", () => {
  // ECMA-376 Part 1 §18.2.29: H0 = SHA-512(salt + UTF-16LE password), Hn = SHA-512(Hn-1 + n as a 32-bit little-endian
  // integer), 100,000 rounds — the value computed independently, with Python's hashlib.
  it("is the hash the format defines", async () => {
    expect(
      await Encryptor.convertPasswordToHash(
        "password",
        "SHA-512",
        "AAAAAAAAAAAAAAAAAAAAAA==",
        100_000
      )
    ).toBe(
      "qd8sMcwp/cRqmrVv2Az3F/5D2ZnRl6tGb3f/80S1TxiEVka5FGqNpUoZD6F0SrQUEzFvXJBzuEAQEFaPzvU4mg=="
    );
  });

  // The rounds ran as one synchronous loop in Node, and as one awaited Web Crypto digest per round in a browser —
  // four seconds in Safari. They run synchronously everywhere now, giving the event loop a turn once the budget is
  // spent.
  it("lets the event loop run while it spins", async () => {
    let clock = 0;
    vi.spyOn(performance, "now").mockImplementation(() => (clock += 1));
    let turns = 0;
    let running = true;
    const tick = (): void => {
      if (running) {
        turns++;
        scheduleMacrotask(tick);
      }
    };
    scheduleMacrotask(tick);
    try {
      await Encryptor.convertPasswordToHash("pw", "SHA-512", "AAAAAAAAAAAAAAAAAAAAAA==", 100_000);
    } finally {
      running = false;
      vi.restoreAllMocks();
    }
    expect(turns).toBeGreaterThan(5);
  });
});
