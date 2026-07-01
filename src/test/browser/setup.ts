import { loadIife } from "@test/browser/load-iife";
import { beforeAll } from "vitest";

// Eagerly load the Excel IIFE so the legacy `Documonster.Excel` smoke test
// (which reads the global without loading it itself) keeps working. Other
// per-bundle smoke tests load their own bundle via `loadIife(...)`.
beforeAll(async () => {
  const Excel = await loadIife("excel", "Excel");
  // eslint-disable-next-line no-console
  console.log("Documonster.Excel loaded:", typeof Excel);
}, 60000);
