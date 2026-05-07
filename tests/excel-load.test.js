import { describe, it, expect } from "vitest";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");

// excel-load expects global XLSX
globalThis.XLSX = XLSX;
await import("../js/excel-load.js");
const { parseArrayBuffer } = globalThis.TorSExcel || {};

function makeWorkbookAOA(aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Bars");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return buf;
}

describe("TorSExcel.parseArrayBuffer", () => {
  it("parses a simple single-day workbook with explicit levels", () => {
    const buf = makeWorkbookAOA([
      ["Time", "Open", "High", "Low", "Close", "PD_High", "PD_Low", "ON_High", "ON_Low", "OR_High", "OR_Low"],
      ["2026-05-01 09:30", 200, 201, 199, 200.5, 210, 190, 205, 195, 202, 198],
      ["2026-05-01 09:31", 200.5, 202, 200, 201.0, "", "", "", "", "", ""],
    ]);

    const res = parseArrayBuffer(buf, 30);
    expect(res.ok).toBe(true);
    expect(res.bars.length).toBeGreaterThanOrEqual(2);
    expect(res.levels.pdHigh).toBe(210);
    expect(res.levels.pdLow).toBe(190);
    expect(res.levels.onHigh).toBe(205);
    expect(res.levels.onLow).toBe(195);
    expect(res.levels.orHigh).toBe(202);
    expect(res.levels.orLow).toBe(198);
  });

  it("computes opening range from timestamps when OR columns absent", () => {
    const buf = makeWorkbookAOA([
      ["Time", "Open", "High", "Low", "Close"],
      ["2026-05-01 09:30", 200, 201, 199, 200.5],
      ["2026-05-01 09:31", 200.5, 203, 200, 202.0],
      ["2026-05-01 10:01", 202.0, 202.2, 201.2, 201.5],
    ]);
    const res = parseArrayBuffer(buf, 30);
    expect(res.ok).toBe(true);
    expect(res.levels.orHigh).toBe(203);
    expect(res.levels.orLow).toBe(199);
  });

  it("returns a helpful error when required columns missing", () => {
    const buf = makeWorkbookAOA([
      ["Foo", "Bar"],
      [1, 2],
      [3, 4],
    ]);
    const res = parseArrayBuffer(buf, 30);
    expect(res.ok).toBe(false);
    expect(String(res.error || "")).toMatch(/Need at least 2 rows/i);
  });
});

