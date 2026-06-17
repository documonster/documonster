import { beforeAll } from "vitest";

beforeAll(async () => {
  const script = document.createElement("script");
  script.src = "/dist/iife/documonster.excel.iife.min.js";

  await new Promise((resolve, reject) => {
    script.onload = resolve;
    script.onerror = e => {
      console.error("Failed to load Documonster.Excel:", e);
      reject(e);
    };
    document.head.appendChild(script);
  });

  console.log("Documonster.Excel loaded:", typeof (globalThis as any).Documonster?.Excel);
}, 60000);
