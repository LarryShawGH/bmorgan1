/**
 * TorS demo: synthetic intraday series + level drawing (browser preview only).
 * Times are modeled in ET minutes from midnight for labeling.
 */

const RTH_START = 9 * 60 + 30; // 9:30
const RTH_END = 16 * 60; // 16:00
const BAR_MINUTES = 1;
const BARS_PER_DAY = (RTH_END - RTH_START) / BAR_MINUTES;

/** Fallbacks keyed by HTML id (see getConfig). */
const defaults = {
  cPdHigh: "#ef4444",
  cPdLow: "#22c55e",
  cOnHigh: "#86efac",
  cOnLow: "#fca5a5",
  cOrHigh: "#4ade80",
  cOrLow: "#f87171",
  cBoxFill: "#6366f1",
  cCenter: "#f8fafc",
  cHalfHour: "#e2e8f0",
  sPdHigh: true,
  sPdLow: true,
  sOnHigh: true,
  sOnLow: true,
  sOrHigh: true,
  sOrLow: true,
  sBox: true,
  sHalfHour: true,
  orMinutes: 30,
  boxOpacity: 0.12,
};

const PRESET_VERSION = 1;
const STORAGE_KEY = "tors_presets_v1";

const PRESET_FIELDS = {
  text: ["clientLabel", "clientNotes", "seed", "orMinutes"],
  range: ["boxOpacity"],
  color: [
    "cPdHigh",
    "cPdLow",
    "cOnHigh",
    "cOnLow",
    "cOrHigh",
    "cOrLow",
    "cBoxFill",
    "cCenter",
    "cHalfHour",
  ],
  check: [
    "sPdHigh",
    "sPdLow",
    "sOnHigh",
    "sOnLow",
    "sOrHigh",
    "sOrLow",
    "sBox",
    "sHalfHour",
  ],
};

function hexToRgba(hex, alpha) {
  const h = (hex || "").replace("#", "");
  if (h.length !== 6) return `rgba(99, 102, 241, ${alpha})`;
  const n = parseInt(h, 16);
  if (!Number.isFinite(n)) return `rgba(99, 102, 241, ${alpha})`;
  return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${alpha})`;
}

function applyQueryToForm() {
  const p = new URLSearchParams(window.location.search);
  if (p.has("or")) {
    const v = Math.min(120, Math.max(5, Number(p.get("or")) || 30));
    const el = document.getElementById("orMinutes");
    if (el) el.value = String(v);
  }
  if (p.has("seed")) {
    const v = Number(p.get("seed"));
    const el = document.getElementById("seed");
    if (el && Number.isFinite(v)) el.value = String(v >>> 0);
  }
}

function presetToBase64Url(preset) {
  const json = JSON.stringify(preset);
  const b64 = btoa(unescape(encodeURIComponent(json)));
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/u, "");
}

function presetFromBase64Url(s) {
  if (!s || typeof s !== "string") return null;
  let b64 = s.replace(/-/g, "+").replace(/_/g, "/");
  const pad = b64.length % 4 ? "=".repeat(4 - (b64.length % 4)) : "";
  b64 += pad;
  const json = decodeURIComponent(escape(atob(b64)));
  return JSON.parse(json);
}

function collectPreset() {
  const p = { v: PRESET_VERSION };
  for (const id of PRESET_FIELDS.text) {
    const el = document.getElementById(id);
    p[id] = el ? String(el.value) : "";
  }
  for (const id of PRESET_FIELDS.range) {
    const el = document.getElementById(id);
    p[id] = el ? String(el.value) : String(defaults.boxOpacity);
  }
  for (const id of PRESET_FIELDS.color) {
    const el = document.getElementById(id);
    p[id] = el ? el.value : defaults[id];
  }
  for (const id of PRESET_FIELDS.check) {
    const el = document.getElementById(id);
    p[id] = Boolean(el?.checked);
  }
  return p;
}

function applyPreset(raw) {
  if (!raw || raw.v !== PRESET_VERSION) return false;
  for (const id of PRESET_FIELDS.text) {
    if (raw[id] === undefined) continue;
    const el = document.getElementById(id);
    if (el) el.value = String(raw[id]);
  }
  for (const id of PRESET_FIELDS.range) {
    if (raw[id] === undefined) continue;
    const el = document.getElementById(id);
    if (el) el.value = String(raw[id]);
  }
  for (const id of PRESET_FIELDS.color) {
    if (raw[id] === undefined) continue;
    const el = document.getElementById(id);
    if (el) el.value = String(raw[id]);
  }
  for (const id of PRESET_FIELDS.check) {
    if (raw[id] === undefined) continue;
    const el = document.getElementById(id);
    if (el) el.checked = Boolean(raw[id]);
  }
  return true;
}

function applyHashPreset() {
  const h = window.location.hash.replace(/^#/u, "");
  if (!h.startsWith("p=")) return;
  const encoded = h.slice(2);
  try {
    const data = presetFromBase64Url(encoded);
    if (applyPreset(data)) draw();
  } catch (e) {
    console.warn("Invalid preset hash", e);
  }
}

function listPresets() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch {
    return [];
  }
}

function savePresetsList(arr) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(arr));
}

function populatePresetSelect() {
  const sel = document.getElementById("presetSelect");
  if (!sel) return;
  const current = sel.value;
  const items = listPresets();
  sel.innerHTML = '<option value="">— Load a saved preset —</option>';
  for (const { name } of items) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  }
  if (current && items.some((x) => x.name === current)) sel.value = current;
}

function showCopyToast(msg) {
  const el = document.getElementById("copyToast");
  if (!el) return;
  el.textContent = msg;
  clearTimeout(showCopyToast._t);
  showCopyToast._t = setTimeout(() => {
    el.textContent = "";
  }, 2800);
}

/** In-memory workbook for re-parse when opening-range minutes change. */
const excelCache = {
  buffer: null,
  fileName: "",
  parsedOr: NaN,
  parsed: null,
};

function setExcelStatus(text) {
  const el = document.getElementById("excelStatus");
  if (el) el.textContent = text || "";
}

function clearExcelData() {
  excelCache.buffer = null;
  excelCache.fileName = "";
  excelCache.parsedOr = NaN;
  excelCache.parsed = null;
  setExcelStatus("");
}

function getExcelParsed() {
  if (!excelCache.buffer || typeof TorSExcel === "undefined") return null;
  const orM = Math.max(
    5,
    Math.min(120, parseInt(document.getElementById("orMinutes")?.value || "30", 10) || 30)
  );
  if (excelCache.parsed && excelCache.parsedOr === orM) return excelCache.parsed;
  const res = TorSExcel.parseArrayBuffer(excelCache.buffer, orM);
  excelCache.parsed = res;
  excelCache.parsedOr = orM;
  return res;
}

function getSeriesForDraw() {
  const ex = getExcelParsed();
  if (ex && ex.ok && ex.bars.length >= 2) {
    return {
      source: "excel",
      bars: ex.bars,
      levels: ex.levels,
      fileName: excelCache.fileName,
      sheetName: ex.sheetName,
    };
  }
  const seed = parseInt(document.getElementById("seed")?.value || "42", 10) || 42;
  const g = generateSession(seed);
  return { source: "demo", bars: g.bars, levels: g.levels, fileName: "", sheetName: "" };
}

function barIndexForHalfHour(series, etMinute) {
  if (series.source === "excel" && series.bars[0]?.etMins != null) {
    const bars = series.bars;
    for (let i = 0; i < bars.length; i++) {
      if (bars[i].etMins >= etMinute) return i;
    }
    return Math.max(0, bars.length - 1);
  }
  const idx = Math.round(((etMinute - RTH_START) / BAR_MINUTES) * (series.bars.length / BARS_PER_DAY));
  return Math.min(series.bars.length - 1, Math.max(0, idx));
}

function levelOk(v) {
  return v != null && Number.isFinite(v);
}

function mulberry32(a) {
  return function () {
    let t = (a += 0x6d2b79f5);
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function generateSession(seed) {
  const rnd = mulberry32(seed >>> 0);
  const base = 100 + rnd() * 20;

  const pdHigh = base + 0.4 + rnd() * 0.5;
  const pdLow = base - 0.6 - rnd() * 0.5;
  const onHigh = pdHigh - 0.05 + rnd() * 0.3;
  const onLow = pdLow + 0.05 - rnd() * 0.3;

  const bars = [];
  let price = base + (rnd() - 0.5) * 0.3;
  for (let i = 0; i < BARS_PER_DAY; i++) {
    const drift = (rnd() - 0.48) * 0.08;
    const o = price;
    const c = o + drift + (rnd() - 0.5) * 0.12;
    const h = Math.max(o, c) + rnd() * 0.06;
    const l = Math.min(o, c) - rnd() * 0.06;
    bars.push({ o, h, l, c });
    price = c;
  }

  const orMinutes = Math.max(
    5,
    Math.min(120, parseInt(document.getElementById("orMinutes")?.value || "30", 10) || 30)
  );
  const orBars = Math.round(orMinutes / BAR_MINUTES);
  let orH = -Infinity;
  let orL = Infinity;
  for (let i = 0; i < Math.min(orBars, bars.length); i++) {
    orH = Math.max(orH, bars[i].h);
    orL = Math.min(orL, bars[i].l);
  }

  return { bars, levels: { pdHigh, pdLow, onHigh, onLow, orHigh: orH, orLow: orL }, base };
}

function etMinuteToLabel(mins) {
  const h = Math.floor(mins / 60);
  const m = mins % 60;
  const am = h < 12;
  const hh = h % 12 || 12;
  const mm = m.toString().padStart(2, "0");
  return `${hh}:${mm} ${am ? "AM" : "PM"}`;
}

function getConfig() {
  const g = (id) => document.getElementById(id);
  const c = (id) => g(id)?.value || defaults[id];
  const chk = (id) => g(id)?.checked ?? defaults[id];
  return {
    pdHigh: c("cPdHigh"),
    pdLow: c("cPdLow"),
    onHigh: c("cOnHigh"),
    onLow: c("cOnLow"),
    orHigh: c("cOrHigh"),
    orLow: c("cOrLow"),
    boxFill: c("cBoxFill"),
    centerLine: c("cCenter"),
    halfHour: c("cHalfHour"),
    showPdH: chk("sPdHigh"),
    showPdL: chk("sPdLow"),
    showOnH: chk("sOnHigh"),
    showOnL: chk("sOnLow"),
    showOrH: chk("sOrHigh"),
    showOrL: chk("sOrLow"),
    showBox: chk("sBox"),
    showHalfHour: chk("sHalfHour"),
    boxOpacity: parseFloat(g("boxOpacity")?.value || String(defaults.boxOpacity)),
  };
}

function draw() {
  const canvas = document.getElementById("chart");
  if (!canvas) return;
  const ctx = canvas.getContext("2d");
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  canvas.width = Math.floor(rect.width * dpr);
  canvas.height = Math.floor(rect.height * dpr);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);

  const W = rect.width;
  const H = rect.height;
  const pad = { l: 56, r: 12, t: 28, b: 36 };
  const plotW = W - pad.l - pad.r;
  const plotH = H - pad.t - pad.b;

  const series = getSeriesForDraw();
  const bars = series.bars;
  const levels = series.levels;
  const cfg = getConfig();

  let ymin = Infinity;
  let ymax = -Infinity;
  for (const b of bars) {
    ymin = Math.min(ymin, b.l);
    ymax = Math.max(ymax, b.h);
  }
  for (const v of Object.values(levels)) {
    if (levelOk(v)) {
      ymin = Math.min(ymin, v);
      ymax = Math.max(ymax, v);
    }
  }
  const padY = (ymax - ymin) * 0.08 || 0.1;
  ymin -= padY;
  ymax += padY;

  const yScale = (p) => pad.t + plotH - ((p - ymin) / (ymax - ymin)) * plotH;
  const xScale = (i) => pad.l + (i / (bars.length - 1 || 1)) * plotW;

  ctx.fillStyle = "#010409";
  ctx.fillRect(0, 0, W, H);

  ctx.strokeStyle = "#21262d";
  ctx.lineWidth = 1;
  for (let i = 0; i <= 5; i++) {
    const y = pad.t + (plotH * i) / 5;
    ctx.beginPath();
    ctx.moveTo(pad.l, y);
    ctx.lineTo(pad.l + plotW, y);
    ctx.stroke();
    const price = ymax - (i / 5) * (ymax - ymin);
    ctx.fillStyle = "#6e7681";
    ctx.font = "11px " + getComputedStyle(document.body).fontFamily;
    ctx.textAlign = "right";
    ctx.fillText(price.toFixed(2), pad.l - 8, y + 4);
  }

  ctx.fillStyle = "#8b949e";
  ctx.font = "12px " + getComputedStyle(document.body).fontFamily;
  ctx.textAlign = "center";
  const subtitle =
    series.source === "excel"
      ? `EXCEL MODE — ${series.fileName} (${bars.length} bars)`
      : "Eastern Time (concept demo)";
  ctx.fillText(subtitle, W / 2, 18);

  function hLine(price, color, dash) {
    const y = yScale(price);
    ctx.beginPath();
    if (dash) ctx.setLineDash(dash);
    ctx.strokeStyle = color;
    ctx.lineWidth = 1.5;
    ctx.moveTo(pad.l, y);
    ctx.lineTo(pad.l + plotW, y);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  if (cfg.showBox && levelOk(levels.pdHigh) && levelOk(levels.pdLow)) {
    const y1 = yScale(levels.pdHigh);
    const y2 = yScale(levels.pdLow);
    const top = Math.min(y1, y2);
    const h = Math.abs(y2 - y1);
    ctx.fillStyle = hexToRgba(cfg.boxFill, cfg.boxOpacity);
    ctx.fillRect(pad.l, top, plotW, h);
    const mid = (levels.pdHigh + levels.pdLow) / 2;
    hLine(mid, cfg.centerLine, [4, 4]);
    ctx.fillStyle = cfg.centerLine;
    ctx.font = "10px " + getComputedStyle(document.body).fontFamily;
    ctx.textAlign = "left";
    ctx.fillText(mid.toFixed(2), pad.l + plotW - 48, yScale(mid) - 4);
  }

  if (cfg.showPdH && levelOk(levels.pdHigh)) hLine(levels.pdHigh, cfg.pdHigh);
  if (cfg.showPdL && levelOk(levels.pdLow)) hLine(levels.pdLow, cfg.pdLow);
  if (cfg.showOnH && levelOk(levels.onHigh)) hLine(levels.onHigh, cfg.onHigh, [6, 4]);
  if (cfg.showOnL && levelOk(levels.onLow)) hLine(levels.onLow, cfg.onLow, [6, 4]);
  if (cfg.showOrH && levelOk(levels.orHigh)) hLine(levels.orHigh, cfg.orHigh);
  if (cfg.showOrL && levelOk(levels.orLow)) hLine(levels.orLow, cfg.orLow);

  const halfHourStarts = [];
  for (let m = RTH_START; m <= RTH_END; m += 30) {
    halfHourStarts.push(m);
  }

  if (cfg.showHalfHour) {
    ctx.font = "9px " + getComputedStyle(document.body).fontFamily;
    halfHourStarts.forEach((etMin, idx) => {
      const i = barIndexForHalfHour(series, etMin);
      const x = xScale(i);
      const close = bars[i].c;
      ctx.beginPath();
      ctx.strokeStyle = cfg.halfHour;
      ctx.lineWidth = 1;
      ctx.setLineDash([3, 3]);
      ctx.moveTo(x, pad.t);
      ctx.lineTo(x, pad.t + plotH);
      ctx.stroke();
      ctx.setLineDash([]);
      ctx.fillStyle = cfg.halfHour;
      ctx.textAlign = "center";
      ctx.fillText(etMinuteToLabel(etMin), x, pad.t + plotH + 14);
      ctx.textAlign = "left";
      ctx.fillText(close.toFixed(2), x + 3, pad.t + 12 + (idx % 3) * 11);
    });
  }

  const candleW = Math.max(1, (plotW / bars.length) * 0.6);
  const upColor = series.source === "excel" ? "#38bdf8" : "#3fb950";
  const downColor = series.source === "excel" ? "#fb7185" : "#f85149";
  bars.forEach((b, i) => {
    const x = xScale(i);
    const yO = yScale(b.o);
    const yC = yScale(b.c);
    const yH = yScale(b.h);
    const yL = yScale(b.l);
    const up = b.c >= b.o;
    ctx.strokeStyle = up ? upColor : downColor;
    ctx.fillStyle = up ? upColor : downColor;
    ctx.beginPath();
    ctx.moveTo(x, yH);
    ctx.lineTo(x, yL);
    ctx.stroke();
    const top = Math.min(yO, yC);
    const h = Math.max(1, Math.abs(yC - yO));
    ctx.fillRect(x - candleW / 2, top, candleW, h);
  });

}

function wire() {
  document.querySelectorAll("[data-refresh]").forEach((el) => {
    el.addEventListener("input", draw);
    el.addEventListener("change", draw);
  });
  document.getElementById("btnRegen")?.addEventListener("click", () => {
    const s = document.getElementById("seed");
    if (s) s.value = String((Math.random() * 1e9) | 0);
    draw();
  });
  document.getElementById("btnRedraw")?.addEventListener("click", draw);

  document.getElementById("btnLoadExcel")?.addEventListener("click", () => {
    document.getElementById("excelFile")?.click();
  });

  document.getElementById("btnExcelTemplate")?.addEventListener("click", () => {
    if (typeof TorSExcel === "undefined" || typeof XLSX === "undefined") {
      showCopyToast("Excel library failed to load. Check network / CDN.");
      return;
    }
    TorSExcel.downloadTemplate();
    showCopyToast("Template downloaded (tors-excel-template.xlsx).");
  });

  document.getElementById("btnUseDemo")?.addEventListener("click", () => {
    clearExcelData();
    const inp = document.getElementById("excelFile");
    if (inp) inp.value = "";
    draw();
    showCopyToast("Using synthetic demo data.");
  });

  document.getElementById("excelFile")?.addEventListener("change", (ev) => {
    const file = ev.target.files?.[0];
    if (!file) return;
    if (typeof TorSExcel === "undefined" || typeof XLSX === "undefined") {
      showCopyToast("Excel library not loaded.");
      ev.target.value = "";
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const buf = reader.result;
      if (!(buf instanceof ArrayBuffer)) {
        showCopyToast("Could not read file.");
        ev.target.value = "";
        return;
      }
      excelCache.buffer = buf;
      excelCache.fileName = file.name;
      excelCache.parsedOr = NaN;
      excelCache.parsed = null;
      const res = TorSExcel.parseArrayBuffer(buf, parseInt(document.getElementById("orMinutes")?.value || "30", 10) || 30);
      excelCache.parsed = res;
      excelCache.parsedOr = parseInt(document.getElementById("orMinutes")?.value || "30", 10) || 30;
      if (!res.ok) {
        clearExcelData();
        const inp2 = document.getElementById("excelFile");
        if (inp2) inp2.value = "";
        showCopyToast(res.error || "Excel parse failed.");
        setExcelStatus(res.error || "");
        draw();
        ev.target.value = "";
        return;
      }
      const modified = new Date(file.lastModified);
      const msg = [
        `Loaded ${file.name} (${Math.round(file.size / 1024)} KB, modified ${modified.toLocaleString()}) — sheet “${res.sheetName}”: ${res.rowCount} rows → ${res.bars.length} chart bars (latest ET day).`,
        ...(res.warnings || []),
      ].join(" ");
      setExcelStatus(msg);
      showCopyToast("Excel loaded.");
      draw();
      // Important: reset file input so re-selecting same filename triggers change event.
      ev.target.value = "";
    };
    reader.readAsArrayBuffer(file);
  });

  document.getElementById("btnCopyLink")?.addEventListener("click", () => {
    const preset = collectPreset();
    const base = `${window.location.origin}${window.location.pathname}${window.location.search}`;
    const url = `${base}#p=${presetToBase64Url(preset)}`;
    if (url.length > 8000) {
      showCopyToast("Link too long for some browsers — use Download JSON instead.");
      return;
    }
    const excelHint = excelCache.buffer
      ? " Send the .xlsx separately (not embedded in the link)."
      : "";
    navigator.clipboard.writeText(url).then(
      () => showCopyToast(`Share link copied.${excelHint}`),
      () => {
        window.prompt("Copy this link:", url);
        showCopyToast("Copy from the dialog if clipboard was blocked.");
      }
    );
  });

  document.getElementById("btnDownloadJson")?.addEventListener("click", () => {
    const preset = collectPreset();
    const slug = (preset.clientLabel || "tors-preset")
      .replace(/[^\w\-]+/gu, "_")
      .replace(/_+/gu, "_")
      .replace(/^_|_$/gu, "")
      .slice(0, 48) || "tors-preset";
    const blob = new Blob([JSON.stringify(preset, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${slug}.json`;
    a.click();
    URL.revokeObjectURL(a.href);
    showCopyToast("JSON file downloaded.");
  });

  document.getElementById("importFile")?.addEventListener("change", (ev) => {
    const file = ev.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(String(reader.result));
        if (applyPreset(data)) {
          draw();
          showCopyToast("Imported preset from file.");
        } else showCopyToast("Invalid preset file (wrong version).");
      } catch (e) {
        showCopyToast("Could not read JSON.");
        console.warn(e);
      }
      ev.target.value = "";
    };
    reader.readAsText(file);
  });

  document.getElementById("btnApplyPaste")?.addEventListener("click", () => {
    const ta = document.getElementById("importPaste");
    if (!ta) return;
    try {
      const data = JSON.parse(ta.value.trim());
      if (applyPreset(data)) {
        draw();
        showCopyToast("Applied pasted JSON.");
      } else showCopyToast("Invalid JSON preset (need v:1).");
    } catch (e) {
      showCopyToast("Invalid JSON.");
      console.warn(e);
    }
  });

  document.getElementById("presetSelect")?.addEventListener("change", (ev) => {
    const name = ev.target.value;
    if (!name) return;
    const item = listPresets().find((x) => x.name === name);
    if (item?.preset && applyPreset(item.preset)) draw();
  });

  document.getElementById("btnSavePreset")?.addEventListener("click", () => {
    const label = document.getElementById("clientLabel")?.value?.trim() || "";
    const name = window.prompt("Name for this preset (saved in this browser only)", label);
    if (!name || !String(name).trim()) return;
    const key = String(name).trim();
    const preset = collectPreset();
    const next = listPresets().filter((x) => x.name !== key);
    next.push({ name: key, preset });
    savePresetsList(next);
    populatePresetSelect();
    const sel = document.getElementById("presetSelect");
    if (sel) sel.value = key;
    showCopyToast(`Saved "${key}" on this device.`);
  });

  document.getElementById("btnDeletePreset")?.addEventListener("click", () => {
    const sel = document.getElementById("presetSelect");
    const name = sel?.value;
    if (!name) {
      showCopyToast("Select a preset to delete.");
      return;
    }
    if (!window.confirm(`Delete preset "${name}" from this browser?`)) return;
    savePresetsList(listPresets().filter((x) => x.name !== name));
    populatePresetSelect();
    showCopyToast("Preset deleted.");
  });

  let resizeT;
  window.addEventListener("resize", () => {
    clearTimeout(resizeT);
    resizeT = setTimeout(draw, 120);
  });

  window.addEventListener("hashchange", () => {
    applyHashPreset();
  });
}

document.addEventListener("DOMContentLoaded", () => {
  applyQueryToForm();
  applyHashPreset();
  populatePresetSelect();
  wire();
  draw();
});
