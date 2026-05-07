/**
 * Excel → bars + levels for TorS preview (browser only).
 * Depends on global XLSX (SheetJS).
 */
(function (global) {
  "use strict";

  const RTH_START = 9 * 60 + 30;
  const RTH_END = 16 * 60;

  function normalizeHeader(h) {
    return String(h ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/gu, "_")
      .replace(/[^a-z0-9_]/gu, "");
  }

  function pickNorm(row, aliases) {
    for (const a of aliases) {
      const v = row[a];
      if (v !== undefined && v !== null && v !== "") return v;
    }
    return undefined;
  }

  function num(x) {
    if (typeof x === "number" && Number.isFinite(x)) return x;
    const n = parseFloat(String(x).replace(/,/gu, ""));
    return Number.isFinite(n) ? n : NaN;
  }

  function parseTimeCell(v) {
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    if (typeof v === "number" && Number.isFinite(v)) {
      const ms = (v - 25569) * 86400 * 1000;
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d;
    }
    if (typeof v === "string") {
      const t = Date.parse(v);
      if (!isNaN(t)) return new Date(t);
    }
    return null;
  }

  function getETDayKey(d) {
    return new Intl.DateTimeFormat("en-CA", {
      timeZone: "America/New_York",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).format(d);
  }

  function getETMinuteOfDay(d) {
    const parts = new Intl.DateTimeFormat("en-US", {
      timeZone: "America/New_York",
      hour: "numeric",
      minute: "numeric",
      hour12: false,
    }).formatToParts(d);
    let hour = 0;
    let minute = 0;
    for (const p of parts) {
      if (p.type === "hour") hour = parseInt(p.value, 10) || 0;
      if (p.type === "minute") minute = parseInt(p.value, 10) || 0;
    }
    return hour * 60 + minute;
  }

  function isRTH(etMins) {
    return etMins >= RTH_START && etMins < RTH_END;
  }

  const ALIAS = {
    time: ["time", "datetime", "timestamp", "date", "dt", "bar_time", "bartime"],
    open: ["open", "o"],
    high: ["high", "h"],
    low: ["low", "l"],
    close: ["close", "c", "last"],
    pdHigh: ["pd_high", "pdhigh", "previous_day_high", "prev_day_high", "priordayhigh"],
    pdLow: ["pd_low", "pdlow", "previous_day_low", "prev_day_low", "priordaylow"],
    onHigh: ["on_high", "onhigh", "overnight_high", "overnighthigh"],
    onLow: ["on_low", "onlow", "overnight_low", "overnightlow"],
    orHigh: ["or_high", "orhigh", "opening_range_high", "openingrangehigh"],
    orLow: ["or_low", "orlow", "opening_range_low", "openingrangelow"],
  };

  function firstNumericInColumn(rows, aliases) {
    for (const raw of rows) {
      const row = raw.__norm || raw;
      const v = num(pickNorm(row, aliases));
      if (Number.isFinite(v)) return v;
    }
    return null;
  }

  function buildNormRows(rows) {
    return rows.map((r) => {
      const o = { __norm: {} };
      for (const [k, v] of Object.entries(r)) {
        o.__norm[normalizeHeader(k)] = v;
      }
      return o;
    });
  }

  function parseArrayBuffer(arrayBuffer, orMinutes) {
    const warnings = [];
    if (typeof XLSX === "undefined") {
      return { ok: false, error: "Excel library not loaded.", warnings };
    }
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheetName = wb.SheetNames[0];
    if (!sheetName) {
      return { ok: false, error: "Workbook has no sheets.", warnings };
    }
    const sheet = wb.Sheets[sheetName];
    const objects = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
    if (!objects.length) {
      return { ok: false, error: "First sheet is empty.", warnings };
    }

    const rows = buildNormRows(objects);
    const barsRaw = [];

    for (const wrap of rows) {
      const row = wrap.__norm;
      const tRaw = pickNorm(row, ALIAS.time);
      const t = parseTimeCell(tRaw);
      const o = num(pickNorm(row, ALIAS.open));
      const h = num(pickNorm(row, ALIAS.high));
      const l = num(pickNorm(row, ALIAS.low));
      const c = num(pickNorm(row, ALIAS.close));
      if (!t || !Number.isFinite(h) || !Number.isFinite(l)) continue;
      const oF = Number.isFinite(o) ? o : c;
      const cF = Number.isFinite(c) ? c : oF;
      if (!Number.isFinite(oF) || !Number.isFinite(cF)) continue;
      const dayKey = getETDayKey(t);
      const etMins = getETMinuteOfDay(t);
      barsRaw.push({ t, o: oF, h, l, c: cF, dayKey, etMins });
    }

    if (barsRaw.length < 2) {
      return {
        ok: false,
        error:
          "Need at least 2 rows with Time + High + Low (+ Open/Close). Check headers (Time, Open, High, Low, Close).",
        warnings,
      };
    }

    barsRaw.sort((a, b) => a.t - b.t);

    const fromCols = {
      pdHigh: firstNumericInColumn(rows, ALIAS.pdHigh),
      pdLow: firstNumericInColumn(rows, ALIAS.pdLow),
      onHigh: firstNumericInColumn(rows, ALIAS.onHigh),
      onLow: firstNumericInColumn(rows, ALIAS.onLow),
      orHigh: firstNumericInColumn(rows, ALIAS.orHigh),
      orLow: firstNumericInColumn(rows, ALIAS.orLow),
    };

    const dayKeys = [...new Set(barsRaw.map((b) => b.dayKey))].sort();
    const lastDay = dayKeys[dayKeys.length - 1];
    const prevDay = dayKeys.length > 1 ? dayKeys[dayKeys.length - 2] : null;

    const rthLast = barsRaw.filter((b) => b.dayKey === lastDay && isRTH(b.etMins));
    if (!rthLast.length) {
      warnings.push("No RTH bars (9:30–16:00 ET) on the latest date — showing all times.");
    }

    const displayBars = rthLast.length ? rthLast : barsRaw.filter((b) => b.dayKey === lastDay);
    const chartBars =
      displayBars.length >= 2
        ? displayBars
        : barsRaw.filter((b) => b.dayKey === lastDay).length >= 2
          ? barsRaw.filter((b) => b.dayKey === lastDay)
          : barsRaw;

    let pdHigh = fromCols.pdHigh;
    let pdLow = fromCols.pdLow;
    let onHigh = fromCols.onHigh;
    let onLow = fromCols.onLow;

    if (prevDay) {
      const rthPrev = barsRaw.filter((b) => b.dayKey === prevDay && isRTH(b.etMins));
      if (rthPrev.length) {
        if (pdHigh == null) pdHigh = Math.max(...rthPrev.map((b) => b.h));
        if (pdLow == null) pdLow = Math.min(...rthPrev.map((b) => b.l));
      } else {
        warnings.push("Prior calendar day has no RTH bars; PD high/low need columns or more data.");
      }

      const onBars = barsRaw.filter((b) => {
        if (b.dayKey === prevDay && b.etMins >= RTH_END) return true;
        if (b.dayKey === lastDay && b.etMins < RTH_START) return true;
        return false;
      });
      if (onBars.length) {
        if (onHigh == null) onHigh = Math.max(...onBars.map((b) => b.h));
        if (onLow == null) onLow = Math.min(...onBars.map((b) => b.l));
      } else {
        warnings.push("No overnight window rows (16:00 prior → 9:30 current); add rows or ON_* columns.");
      }
    } else {
      if (pdHigh == null || pdLow == null) {
        warnings.push("Single-day file: set PD_High / PD_Low columns or include the prior session day.");
      }
      if (onHigh == null || onLow == null) {
        warnings.push("Single-day file: set ON_High / ON_Low columns or include overnight rows.");
      }
    }

    const orM = Math.max(5, Math.min(120, orMinutes || 30));
    const rthToday = barsRaw.filter((b) => b.dayKey === lastDay && isRTH(b.etMins));
    const orSegment = rthToday.filter((b) => b.etMins < RTH_START + orM);
    let orHigh = fromCols.orHigh;
    let orLow = fromCols.orLow;
    if (orSegment.length) {
      if (orHigh == null) orHigh = Math.max(...orSegment.map((b) => b.h));
      if (orLow == null) orLow = Math.min(...orSegment.map((b) => b.l));
    } else {
      warnings.push("Could not compute opening range from timestamps; set OR_High / OR_Low columns.");
    }

    const levels = { pdHigh, pdLow, onHigh, onLow, orHigh, orLow };
    const simpleBars = chartBars.map((b) => ({ o: b.o, h: b.h, l: b.l, c: b.c, t: b.t, etMins: b.etMins }));

    return {
      ok: true,
      bars: simpleBars,
      levels,
      sheetName,
      rowCount: barsRaw.length,
      lastDay,
      warnings,
    };
  }

  function downloadTemplate() {
    if (typeof XLSX === "undefined") return;
    const headers = [
      "Time",
      "Open",
      "High",
      "Low",
      "Close",
      "PD_High",
      "PD_Low",
      "ON_High",
      "ON_Low",
      "OR_High",
      "OR_Low",
    ];
    const sample = [
      "2026-05-01 09:30",
      100.0,
      100.15,
      99.95,
      100.05,
      101.2,
      99.1,
      100.9,
      99.3,
      "",
      "",
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers, sample]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Bars");
    XLSX.writeFile(wb, "tors-excel-template.xlsx");
  }

  const api = { parseArrayBuffer: parseArrayBuffer, downloadTemplate: downloadTemplate };
  global.TorSExcel = api;
  // Node/unit-test support
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }
})(typeof window !== "undefined" ? window : globalThis);
