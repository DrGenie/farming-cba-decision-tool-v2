// Farming CBA Tool - Newcastle Business School
// Fully upgraded script with working tabs, robust import pipeline (upload + paste TSV/CSV + dictionary parsing),
// replicate-specific control baselines, plot-level deltas, treatment summaries with missing-safe stats,
// control-centric Results (leaderboard + comparison-to-control grid + filters + narrative),
// discounted CBA engine + sensitivity grid (price, discount, persistence, recurrence),
// scenario save/load to localStorage, exports (cleaned TSV, summaries CSV, sensitivity CSV, Excel workbook if available),
// AI Briefing (copy-ready narrative prompt with no bullets, no em dash, no abbreviations) + Copy Results JSON,
// and bottom-right toasts for every major action.

(() => {
  "use strict";

  // =========================
  // 0) CONSTANTS + DEFAULTS
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  const STORAGE_KEYS = {
    scenarios: "farming_cba_scenarios_v1",
    activeScenario: "farming_cba_active_scenario_v1"
  };

  // Default sensitivity grids (can be overridden via UI if present)
  const DEFAULT_SENS_PRICE = [300, 350, 400, 450, 500, 550, 600];
  const DEFAULT_SENS_DISC = [2, 4, 7, 10, 12];
  const DEFAULT_SENS_PERSIST = [1, 2, 3, 5, 7, 10];
  const DEFAULT_SENS_RECURRENCE = [1, 2, 3, 4, 5, 7, 10, 0]; // 0 = once only at year 0

  // =========================
  // 1) ID + UTIL
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

  const fmt = n =>
    isFinite(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 2 })
      : "n/a";

  const money = n => (isFinite(n) ? "$" + fmt(n) : "n/a");
  const percent = n => (isFinite(n) ? fmt(n) + "%" : "n/a");
  const slug = s =>
    (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");

  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const s = String(value).trim();
    if (!s || s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === "string" && v.trim() === "") || v === "?";
  }

  function median(arr) {
    const a = arr.filter(v => Number.isFinite(v)).slice().sort((x, y) => x - y);
    if (!a.length) return NaN;
    const mid = Math.floor(a.length / 2);
    return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
  }

  function mean(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (!a.length) return NaN;
    return a.reduce((s, v) => s + v, 0) / a.length;
  }

  function sd(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (a.length < 2) return NaN;
    const m = mean(a);
    const v = a.reduce((s, x) => s + (x - m) * (x - m), 0) / (a.length - 1);
    return Math.sqrt(v);
  }

  function quantile(arr, q) {
    const a = arr.filter(v => Number.isFinite(v)).slice().sort((x, y) => x - y);
    if (!a.length) return NaN;
    const pos = (a.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (a[base + 1] === undefined) return a[base];
    return a[base] + rest * (a[base + 1] - a[base]);
  }

  function iqrOutlierFlags(arr) {
    const a = arr.filter(v => Number.isFinite(v));
    if (a.length < 8) return { low: NaN, high: NaN, outliers: 0 };
    const q1 = quantile(a, 0.25);
    const q3 = quantile(a, 0.75);
    const iqr = q3 - q1;
    const low = q1 - 1.5 * iqr;
    const high = q3 + 1.5 * iqr;
    const outliers = a.filter(v => v < low || v > high).length;
    return { low, high, outliers };
  }

  function annuityFactor(N, rPct) {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  }

  function rng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6d2b79f5;
      let x = t;
      x = Math.imul(x ^ (x >>> 15), 1 | x);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }

  function triangular(r, a, c, b) {
    const F = (c - a) / (b - a);
    if (r < F) return a + Math.sqrt(r * (b - a) * (c - a));
    return b - Math.sqrt((1 - r) * (b - a) * (b - c));
  }

  function ensureToastRoot() {
    if (document.getElementById("toast-root")) return;
    const div = document.createElement("div");
    div.id = "toast-root";
    div.setAttribute("aria-live", "polite");
    div.setAttribute("aria-atomic", "true");
    document.body.appendChild(div);
  }

  function showToast(message) {
    ensureToastRoot();
    const root = document.getElementById("toast-root") || document.body;
    const toast = document.createElement("div");
    toast.className = "toast";
    toast.textContent = message;
    root.appendChild(toast);
    void toast.offsetWidth;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
      setTimeout(() => toast.remove(), 200);
    }, 3500);
  }

  // =========================
  // 2) MODEL (kept, extended)
  // =========================
  const model = {
    project: {
      name: "Faba bean soil amendment trial",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum and fertiliser treatments against a control.",
      objectives: "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities: "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal:
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Faba bean growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
      withoutProject:
        "Growers continue with baseline practice and do not access detailed economic evidence on soil amendments."
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      // Value for Grain yield is treated as grain price ($ per tonne) when unit is t/ha deltas.
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (baseline)",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        notes: "Control definition is taken from the uploaded dataset where available.",
        recurrenceYears: 0 // 0 = once only at year 0; control is baseline so unused
      }
    ],
    benefits: [],
    otherCosts: [],
    adoption: { base: 1.0, low: 0.6, high: 1.0 },
    risk: { base: 0.15, low: 0.05, high: 0.3, tech: 0.05, nonCoop: 0.04, socio: 0.02, fin: 0.03, man: 0.02 },
    sim: {
      n: 1000,
      targetBCR: 2,
      bcrMode: "all",
      seed: null,
      results: { npv: [], bcr: [] },
      details: [],
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false
    }
  };

  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!t.deltas) t.deltas = {};
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (typeof t.labourCost === "undefined") t.labourCost = 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0;
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      if (typeof t.recurrenceYears !== "number" || isNaN(t.recurrenceYears)) t.recurrenceYears = 0;
    });
  }
  initTreatmentDeltas();

  // =========================
  // 3) STATE FOR DATASET + SCENARIOS
  // =========================
  const state = {
    dataset: {
      sourceName: "",
      rawText: "",
      rows: [],
      dictionary: null,
      schema: null,
      derived: {
        cleanedRows: [],
        checks: [],
        replicateBaselines: new Map(), // repKey -> { yieldMean, costMeansByCol }
        plotDeltas: [], // one per row
        treatmentSummary: [], // derived summaries
        controlKey: null
      },
      committedAt: null
    },
    config: {
      // Persistence is how many years yield benefits persist after application.
      // Recurrence is how often costs are incurred (year 0 then every k years), 0 = once only.
      persistenceYears: 5,
      // Sensitivity grids
      sensPrice: DEFAULT_SENS_PRICE.slice(),
      sensDiscount: DEFAULT_SENS_DISC.slice(),
      sensPersistence: DEFAULT_SENS_PERSIST.slice(),
      sensRecurrence: DEFAULT_SENS_RECURRENCE.slice()
    },
    results: {
      perTreatmentBaseCase: [], // computed vs control
      sensitivityGrid: [], // computed per treatment
      lastComputedAt: null
    }
  };

  // =========================
  // 4) CSV/TSV + DICTIONARY PARSING
  // =========================
  function parseDelimited(text, delimiter) {
    const rows = [];
    let i = 0;
    const len = text.length;
    let field = "";
    let row = [];
    let inQuotes = false;
    while (i < len) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          const next = text[i + 1];
          if (next === '"') {
            field += '"';
            i += 2;
            continue;
          } else {
            inQuotes = false;
            i++;
            continue;
          }
        } else {
          field += ch;
          i++;
          continue;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
          i++;
          continue;
        }
        if (ch === delimiter) {
          row.push(field);
          field = "";
          i++;
          continue;
        }
        if (ch === "\r") {
          i++;
          continue;
        }
        if (ch === "\n") {
          row.push(field);
          rows.push(row);
          row = [];
          field = "";
          i++;
          continue;
        }
        field += ch;
        i++;
      }
    }
    row.push(field);
    rows.push(row);
    while (rows.length && rows[rows.length - 1].every(c => String(c ?? "").trim() === "")) rows.pop();
    return rows;
  }

  function detectDelimiter(text) {
    const firstLine = (text || "").split(/\n/).find(l => l.trim().length > 0) || "";
    const tabCount = (firstLine.match(/\t/g) || []).length;
    const commaCount = (firstLine.match(/,/g) || []).length;
    if (tabCount >= commaCount && tabCount > 0) return "\t";
    if (commaCount > 0) return ",";
    const tabs = (text.match(/\t/g) || []).length;
    const commas = (text.match(/,/g) || []).length;
    if (tabs >= commas && tabs > 0) return "\t";
    return ",";
  }

  function normaliseHeader(h) {
    return String(h ?? "")
      .trim()
      .replace(/\s+/g, " ")
      .replace(/[^\S\r\n]+/g, " ");
  }

  function headersToObjects(table) {
    if (!table.length) return [];
    const header = table[0].map(normaliseHeader);
    const out = [];
    for (let r = 1; r < table.length; r++) {
      const row = table[r];
      const obj = {};
      for (let c = 0; c < header.length; c++) {
        const key = header[c] || `col_${c + 1}`;
        obj[key] = row[c] ?? "";
      }
      out.push(obj);
    }
    return out;
  }

  function looksLikeDictionaryHeader(cols) {
    const lower = cols.map(c => String(c || "").toLowerCase());
    const hasVar = lower.some(c => c.includes("variable") || c.includes("field") || c === "name" || c.includes("column"));
    const hasDesc = lower.some(c => c.includes("description") || c.includes("label") || c.includes("notes") || c.includes("definition"));
    return hasVar && hasDesc;
  }

  function splitDictionaryAndDataFromText(rawText) {
    const text = String(rawText || "");
    const chunks = text.split(/\n{2,}/g).map(s => s.trim()).filter(Boolean);
    if (chunks.length < 2) return { dictText: null, dataText: text };

    let dictIdx = -1;
    let dataIdx = -1;

    for (let i = 0; i < chunks.length; i++) {
      const del = detectDelimiter(chunks[i]);
      const tbl = parseDelimited(chunks[i], del);
      if (!tbl.length) continue;
      const head = tbl[0] || [];
      if (looksLikeDictionaryHeader(head) && tbl.length >= 2) {
        dictIdx = i;
        break;
      }
    }

    if (dictIdx >= 0) {
      for (let j = dictIdx + 1; j < chunks.length; j++) {
        const del = detectDelimiter(chunks[j]);
        const tbl = parseDelimited(chunks[j], del);
        if (!tbl.length) continue;
        const head = tbl[0].map(h => String(h || "").toLowerCase());
        const wide = (tbl[0] || []).length >= 6;
        const hasYield = head.some(h => h.includes("yield"));
        const hasTreat = head.some(h => h.includes("treatment") || h.includes("amend") || h.includes("variant"));
        const hasRep = head.some(h => h.includes("rep") || h.includes("block") || h.includes("replicate"));
        if (wide || (hasYield && hasTreat) || (hasYield && hasRep)) {
          dataIdx = j;
          break;
        }
      }
    }

    if (dictIdx >= 0 && dataIdx >= 0) {
      return { dictText: chunks[dictIdx], dataText: chunks[dataIdx] };
    }
    return { dictText: null, dataText: text };
  }

  function parseDictionaryText(dictText) {
    if (!dictText) return null;
    const del = detectDelimiter(dictText);
    const tbl = parseDelimited(dictText, del);
    if (!tbl.length) return null;
    const objs = headersToObjects(tbl);

    const keys = Object.keys(objs[0] || {});
    const lowerKeys = keys.map(k => k.toLowerCase());

    const varKey =
      keys[lowerKeys.findIndex(k => k.includes("variable") || k.includes("field") || k === "name" || k.includes("column"))] ||
      keys[0];
    const descKey =
      keys[lowerKeys.findIndex(k => k.includes("description") || k.includes("label") || k.includes("definition") || k.includes("notes"))] ||
      keys[Math.min(1, keys.length - 1)] ||
      keys[0];

    const roleKey =
      keys[lowerKeys.findIndex(k => k.includes("role") || k.includes("type") || k.includes("category") || k.includes("domain"))] ||
      null;

    const unitKey = keys[lowerKeys.findIndex(k => k.includes("unit"))] || null;

    const dict = new Map();
    objs.forEach(r => {
      const v = String(r[varKey] ?? "").trim();
      if (!v) return;
      dict.set(v, {
        variable: v,
        description: String(r[descKey] ?? "").trim(),
        role: roleKey ? String(r[roleKey] ?? "").trim() : "",
        unit: unitKey ? String(r[unitKey] ?? "").trim() : ""
      });
    });

    return { rows: objs, map: dict };
  }

  function inferSchema(rows, dictionary) {
    const headers = rows.length ? Object.keys(rows[0]) : [];

    function bestHeader(cands) {
      const lower = headers.map(h => h.toLowerCase());
      for (const c of cands) {
        const idx = lower.findIndex(h => h === c || h.includes(c));
        if (idx >= 0) return headers[idx];
      }
      return null;
    }

    const dictRoleMatch = role => {
      if (!dictionary || !dictionary.map) return null;
      for (const [k, meta] of dictionary.map.entries()) {
        const r = String(meta.role || "").toLowerCase();
        if (r.includes(role)) {
          const idx = headers.findIndex(h => h.trim() === k.trim());
          if (idx >= 0) return headers[idx];
        }
      }
      return null;
    };

    const treatmentCol =
      dictRoleMatch("treatment") ||
      bestHeader(["amendment", "treatment", "variant", "package", "option", "arm"]);
    const replicateCol =
      dictRoleMatch("replicate") ||
      bestHeader(["replicate", "rep", "block", "trial block", "replication"]);
    const plotCol = dictRoleMatch("plot") || bestHeader(["plot", "plot id", "plotid", "plot_no", "plot number"]);
    const controlFlagCol = dictRoleMatch("control") || bestHeader(["is_control", "control", "baseline"]);
    const yieldCol = dictRoleMatch("yield") || bestHeader(["yield t/ha", "yield", "grain yield", "yield_tha", "yield (t/ha)"]);

    const costCols = headers.filter(h => {
      const s = h.toLowerCase();
      const isCosty =
        s.includes("cost") || s.includes("labour") || s.includes("labor") || s.includes("input") || s.includes("fert") ||
        s.includes("herb") || s.includes("fung") || s.includes("insect") || s.includes("fuel") || s.includes("machinery") ||
        s.includes("spray") || s.includes("seed");
      const isClearlyNotCost =
        s.includes("yield") || s.includes("protein") || s.includes("screen") || s.includes("moist") || s.includes("rep") ||
        s.includes("plot") || s.includes("treatment") || s.includes("amend");
      return isCosty && !isClearlyNotCost;
    });

    const plotAreaCol = bestHeader(["plot area", "plot_area", "area (ha)", "area_ha", "plot_ha", "ha"]);

    return {
      headers,
      treatmentCol,
      replicateCol,
      plotCol,
      controlFlagCol,
      yieldCol,
      costCols,
      plotAreaCol
    };
  }

  function normaliseTreatmentKey(v) {
    return String(v ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ")
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function detectControlKey(rows, schema) {
    if (!rows.length) return null;
    const tCol = schema.treatmentCol;
    const cCol = schema.controlFlagCol;

    if (cCol) {
      for (const r of rows) {
        const v = r[cCol];
        const s = String(v ?? "").trim().toLowerCase();
        const truthy = s === "1" || s === "true" || s === "yes" || s === "y";
        if (truthy && tCol && !isBlank(r[tCol])) return normaliseTreatmentKey(r[tCol]);
      }
    }

    if (tCol) {
      const counts = new Map();
      for (const r of rows) {
        const tv = String(r[tCol] ?? "");
        if (!tv.trim()) continue;
        const key = normaliseTreatmentKey(tv);
        const low = tv.toLowerCase();
        const isCtrl = low.includes("control") || low.includes("baseline") || low.includes("check");
        if (isCtrl) counts.set(key, (counts.get(key) || 0) + 1);
      }
      if (counts.size) {
        let best = null;
        let bestN = -1;
        for (const [k, n] of counts.entries()) {
          if (n > bestN) {
            best = k;
            bestN = n;
          }
        }
        return best;
      }
    }
    return null;
  }

  function costPerHaFromRow(row, schema, col) {
    const raw = parseNumber(row[col]);
    if (!Number.isFinite(raw)) return NaN;

    const h = String(col || "").toLowerCase();

    const looksPerHa = h.includes("/ha") || h.includes("per ha") || h.includes("per_ha") || h.includes("ha)");
    if (looksPerHa) return raw;

    if (schema.plotAreaCol) {
      const a = parseNumber(row[schema.plotAreaCol]);
      if (Number.isFinite(a) && a > 0) return raw / a;
    }
    return raw;
  }

  function computeDerivedFromDataset(rows, schema) {
    const checks = [];
    const derived = {
      cleanedRows: [],
      checks,
      replicateBaselines: new Map(),
      plotDeltas: [],
      treatmentSummary: [],
      controlKey: null
    };

    if (!rows.length) {
      checks.push({ code: "NO_ROWS", severity: "error", message: "No data rows found after parsing.", count: 0, detail: "" });
      return derived;
    }

    if (!schema.treatmentCol) checks.push({ code: "NO_TREATMENT_COL", severity: "error", message: "Treatment column not found.", count: 0, detail: "" });
    if (!schema.replicateCol) checks.push({ code: "NO_REPLICATE_COL", severity: "warn", message: "Replicate column not found. Replicate-specific baselines will fall back to overall control mean.", count: 0, detail: "" });
    if (!schema.yieldCol) checks.push({ code: "NO_YIELD_COL", severity: "error", message: "Yield column not found.", count: 0, detail: "" });

    const controlKey = detectControlKey(rows, schema);
    derived.controlKey = controlKey;

    if (!controlKey) checks.push({ code: "NO_CONTROL_DETECTED", severity: "error", message: "Control treatment could not be detected. Provide an is_control column or ensure the control label includes the word control.", count: 0, detail: "" });

    const cleaned = rows.map((r, idx) => {
      const treatVal = schema.treatmentCol ? r[schema.treatmentCol] : "";
      const repVal = schema.replicateCol ? r[schema.replicateCol] : "";
      const plotVal = schema.plotCol ? r[schema.plotCol] : "";
      const y = schema.yieldCol ? parseNumber(r[schema.yieldCol]) : NaN;

      const tKey = normaliseTreatmentKey(treatVal);
      const repKey = schema.replicateCol ? String(repVal ?? "").trim() : "";
      const pKey = schema.plotCol ? String(plotVal ?? "").trim() : String(idx + 1);

      let isControl = false;
      if (schema.controlFlagCol) {
        const v = String(r[schema.controlFlagCol] ?? "").trim().toLowerCase();
        isControl = v === "1" || v === "true" || v === "yes" || v === "y";
      }
      if (!isControl && controlKey && tKey === controlKey) isControl = true;

      const costByCol = {};
      (schema.costCols || []).forEach(c => {
        costByCol[c] = costPerHaFromRow(r, schema, c);
      });

      return {
        __rowIndex: idx,
        treatment: String(treatVal ?? "").trim(),
        treatmentKey: tKey,
        replicate: repKey,
        plot: pKey,
        isControl,
        yield: y,
        costsPerHa: costByCol,
        original: r
      };
    });

    derived.cleanedRows = cleaned;

    const missingYield = cleaned.filter(r => !Number.isFinite(r.yield)).length;
    if (missingYield) checks.push({ code: "MISSING_YIELD", severity: "warn", message: "Some rows have missing yield values. These are excluded from yield summaries.", count: missingYield, detail: "" });

    const negYield = cleaned.filter(r => Number.isFinite(r.yield) && r.yield < 0).length;
    if (negYield) checks.push({ code: "NEGATIVE_YIELD", severity: "warn", message: "Some rows have negative yield values. Check units or data entry.", count: negYield, detail: "" });

    const reps = new Map();
    const overallCtrlY = [];
    const overallCtrlCosts = new Map();
    (schema.costCols || []).forEach(c => overallCtrlCosts.set(c, []));

    cleaned.forEach(r => {
      if (!r.isControl) return;
      if (Number.isFinite(r.yield)) overallCtrlY.push(r.yield);
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        if (Number.isFinite(v)) overallCtrlCosts.get(c).push(v);
      });

      const repKey = schema.replicateCol ? (r.replicate || "__NO_REP__") : "__NO_REP__";
      if (!reps.has(repKey)) {
        const m = new Map();
        (schema.costCols || []).forEach(c => m.set(c, []));
        reps.set(repKey, { ctrlY: [], ctrlCostsByCol: m });
      }
      const entry = reps.get(repKey);
      if (Number.isFinite(r.yield)) entry.ctrlY.push(r.yield);
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        if (Number.isFinite(v)) entry.ctrlCostsByCol.get(c).push(v);
      });
    });

    const overallCtrlMeanYield = mean(overallCtrlY);
    if (!Number.isFinite(overallCtrlMeanYield)) {
      checks.push({ code: "CONTROL_YIELD_MISSING", severity: "error", message: "Control yields are missing. Cannot compute deltas.", count: 0, detail: "" });
    }

    const replicateBaselines = new Map();
    for (const [repKey, entry] of reps.entries()) {
      const yMean = mean(entry.ctrlY);
      const costsMean = {};
      (schema.costCols || []).forEach(c => {
        costsMean[c] = mean(entry.ctrlCostsByCol.get(c) || []);
      });
      replicateBaselines.set(repKey, {
        yieldMean: Number.isFinite(yMean) ? yMean : overallCtrlMeanYield,
        costsMeanByCol: costsMean
      });
    }
    derived.replicateBaselines = replicateBaselines;

    if (schema.replicateCol) {
      const allRepKeys = new Set(cleaned.map(r => r.replicate || "__MISSING_REP__"));
      let repsNoCtrl = 0;
      allRepKeys.forEach(k => {
        const has = replicateBaselines.has(k);
        if (!has) repsNoCtrl++;
      });
      if (repsNoCtrl) {
        checks.push({
          code: "REPS_WITHOUT_CONTROL",
          severity: "warn",
          message: "Some replicates have no control rows. Their baselines fall back to overall control mean.",
          count: repsNoCtrl,
          detail: ""
        });
      }
    }

    const plotDeltas = cleaned.map(r => {
      const repKey = schema.replicateCol ? (r.replicate || "__NO_REP__") : "__NO_REP__";
      const base = replicateBaselines.get(repKey) || { yieldMean: overallCtrlMeanYield, costsMeanByCol: {} };
      const dy = Number.isFinite(r.yield) && Number.isFinite(base.yieldMean) ? r.yield - base.yieldMean : NaN;
      const dCosts = {};
      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        const b = base.costsMeanByCol ? base.costsMeanByCol[c] : NaN;
        dCosts[c] = Number.isFinite(v) && Number.isFinite(b) ? v - b : NaN;
      });
      return {
        ...r,
        controlYieldMeanRep: base.yieldMean,
        deltaYield: dy,
        deltaCostsPerHa: dCosts
      };
    });
    derived.plotDeltas = plotDeltas;

    const byTreat = new Map();
    plotDeltas.forEach(r => {
      if (!r.treatmentKey) return;
      if (!byTreat.has(r.treatmentKey)) {
        byTreat.set(r.treatmentKey, {
          treatmentKey: r.treatmentKey,
          treatmentLabel: r.treatment || r.treatmentKey,
          isControl: !!r.isControl,
          n: 0,
          yield: [],
          deltaYield: [],
          costsByCol: {},
          deltaCostsByCol: {}
        });
        (schema.costCols || []).forEach(c => {
          byTreat.get(r.treatmentKey).costsByCol[c] = [];
          byTreat.get(r.treatmentKey).deltaCostsByCol[c] = [];
        });
      }
      const g = byTreat.get(r.treatmentKey);
      g.n++;
      if (Number.isFinite(r.yield)) g.yield.push(r.yield);
      if (Number.isFinite(r.deltaYield)) g.deltaYield.push(r.deltaYield);

      (schema.costCols || []).forEach(c => {
        const v = r.costsPerHa[c];
        const dv = r.deltaCostsPerHa[c];
        if (Number.isFinite(v)) g.costsByCol[c].push(v);
        if (Number.isFinite(dv)) g.deltaCostsByCol[c].push(dv);
      });
    });

    const summaries = [];
    for (const [, g] of byTreat.entries()) {
      const dy = g.deltaYield;
      const y = g.yield;
      const out = {
        treatmentKey: g.treatmentKey,
        treatmentLabel: g.treatmentLabel,
        isControl: g.isControl,
        nRows: g.n,
        nYield: y.filter(Number.isFinite).length,
        yieldMean: mean(y),
        yieldSD: sd(y),
        deltaYieldMean: mean(dy),
        deltaYieldSD: sd(dy),
        deltaYieldMedian: median(dy),
        costs: {},
        deltaCosts: {}
      };
      (schema.costCols || []).forEach(c => {
        out.costs[c] = { mean: mean(g.costsByCol[c] || []), sd: sd(g.costsByCol[c] || []) };
        out.deltaCosts[c] = { mean: mean(g.deltaCostsByCol[c] || []), sd: sd(g.deltaCostsByCol[c] || []) };
      });
      summaries.push(out);
    }

    const dyAll = plotDeltas.map(r => r.deltaYield).filter(v => Number.isFinite(v));
    const outFlags = iqrOutlierFlags(dyAll);
    if (Number.isFinite(outFlags.outliers) && outFlags.outliers > 0) {
      checks.push({
        code: "YIELD_DELTA_OUTLIERS",
        severity: "warn",
        message: "Some yield deltas are outliers by IQR rule. Check plots or consider robustness.",
        count: outFlags.outliers,
        detail: Number.isFinite(outFlags.low) && Number.isFinite(outFlags.high) ? `IQR bounds are ${fmt(outFlags.low)} to ${fmt(outFlags.high)} t/ha.` : ""
      });
    }

    const lowN = summaries.filter(s => !s.isControl && (s.nYield || 0) < 2).length;
    if (lowN) {
      checks.push({
        code: "LOW_REPLICATION",
        severity: "warn",
        message: "Some treatments have fewer than 2 yield observations. Means are unstable.",
        count: lowN,
        detail: ""
      });
    }

    let missingCostCells = 0;
    (schema.costCols || []).forEach(c => {
      plotDeltas.forEach(r => {
        if (!Number.isFinite(r.costsPerHa[c]) && !isBlank(r.original[c])) missingCostCells += 1;
      });
    });
    if (missingCostCells) {
      checks.push({
        code: "NON_NUMERIC_COSTS",
        severity: "warn",
        message: "Some cost cells are non-numeric. They are treated as missing and excluded from cost summaries.",
        count: missingCostCells,
        detail: ""
      });
    }

    derived.treatmentSummary = summaries;
    return derived;
  }

  // =========================
  // 5) CALIBRATE MODEL FROM DATASET SUMMARY
  // =========================
  function ensureYieldOutput() {
    let out = model.outputs.find(o => String(o.name || "").toLowerCase().includes("yield"));
    if (!out) {
      out = { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" };
      model.outputs.unshift(out);
    }
    return out;
  }

  function extractIncrementalCostsFromSummary(summary, schema) {
    let labour = 0;
    let services = 0;
    let materials = 0;

    const cols = schema.costCols || [];
    cols.forEach(c => {
      const dv = summary.deltaCosts && summary.deltaCosts[c] ? summary.deltaCosts[c].mean : NaN;
      if (!Number.isFinite(dv)) return;
      const h = String(c || "").toLowerCase();
      const isLab = h.includes("labour") || h.includes("labor") || h.includes("hours") || h.includes("wage");
      const isServ = h.includes("contract") || h.includes("service") || h.includes("hire") || h.includes("machinery") || h.includes("fuel") || h.includes("spray");
      if (isLab) labour += dv;
      else if (isServ) services += dv;
      else materials += dv;
    });

    return { labour, services, materials };
  }

  function applyDatasetToModel() {
    const derived = state.dataset.derived;
    const schema = state.dataset.schema;

    if (!derived || !derived.treatmentSummary || !derived.treatmentSummary.length) {
      showToast("No derived treatment summary available to apply.");
      return;
    }

    const yieldOut = ensureYieldOutput();
    const yieldId = yieldOut.id;

    const controlSummary =
      derived.treatmentSummary.find(s => s.isControl) ||
      (derived.controlKey ? derived.treatmentSummary.find(s => s.treatmentKey === derived.controlKey) : null);

    const controlName = controlSummary ? controlSummary.treatmentLabel : "Control (baseline)";

    const currentControl = model.treatments.find(t => t.isControl) || model.treatments[0];
    const farmArea = currentControl ? Number(currentControl.area) || 100 : 100;

    const newTreatments = [];
    newTreatments.push({
      id: uid(),
      name: controlName || "Control (baseline)",
      area: farmArea,
      adoption: 1,
      deltas: { [yieldId]: 0 },
      labourCost: 0,
      materialsCost: 0,
      servicesCost: 0,
      capitalCost: 0,
      constrained: true,
      source: "Imported dataset",
      isControl: true,
      notes: "Control is defined by the dataset control flag or by the treatment label.",
      recurrenceYears: 0
    });

    derived.treatmentSummary
      .filter(s => !s.isControl)
      .sort((a, b) => {
        const A = Number.isFinite(a.deltaYieldMean) ? a.deltaYieldMean : -Infinity;
        const B = Number.isFinite(b.deltaYieldMean) ? b.deltaYieldMean : -Infinity;
        return B - A;
      })
      .forEach(s => {
        const incCosts = extractIncrementalCostsFromSummary(s, schema);
        const t = {
          id: uid(),
          name: s.treatmentLabel || s.treatmentKey,
          area: farmArea,
          adoption: 1,
          deltas: { [yieldId]: Number.isFinite(s.deltaYieldMean) ? s.deltaYieldMean : 0 },
          labourCost: Number.isFinite(incCosts.labour) ? incCosts.labour : 0,
          materialsCost: Number.isFinite(incCosts.materials) ? incCosts.materials : 0,
          servicesCost: Number.isFinite(incCosts.services) ? incCosts.services : 0,
          capitalCost: 0,
          constrained: true,
          source: "Imported dataset",
          isControl: false,
          notes: "Incremental values are computed as replicate-specific deltas relative to the control mean within each replicate.",
          recurrenceYears: 0
        };
        newTreatments.push(t);
      });

    model.treatments = newTreatments;
    initTreatmentDeltas();
    showToast("Dataset applied. Treatments calibrated from replicate-specific deltas versus control.");
  }

  // =========================
  // 6) CBA ENGINE: PER TREATMENT VS CONTROL
  // =========================
  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;
    let lo = -0.99;
    let hi = 5.0;
    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);
    let nLo = npvAt(lo);
    let nHi = npvAt(hi);
    if (nLo * nHi > 0) {
      for (let k = 0; k < 20 && nLo * nHi > 0; k++) {
        hi *= 1.5;
        nHi = npvAt(hi);
      }
      if (nLo * nHi > 0) return NaN;
    }
    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2;
      const nMid = npvAt(mid);
      if (Math.abs(nMid) < 1e-8) return mid * 100;
      if (nLo * nMid <= 0) {
        hi = mid;
        nHi = nMid;
      } else {
        lo = mid;
        nLo = nMid;
      }
    }
    return ((lo + hi) / 2) * 100;
  }

  function mirr(cf, financeRatePct, reinvestRatePct) {
    const n = cf.length - 1;
    const fr = financeRatePct / 100;
    const rr = reinvestRatePct / 100;
    let pvNeg = 0;
    let fvPos = 0;
    for (let t = 0; t <= n; t++) {
      const v = cf[t];
      if (v < 0) pvNeg += v / Math.pow(1 + fr, t);
      if (v > 0) fvPos += v * Math.pow(1 + rr, n - t);
    }
    if (pvNeg === 0) return NaN;
    const mirrVal = Math.pow(-fvPos / pvNeg, 1 / n) - 1;
    return mirrVal * 100;
  }

  function payback(cf, ratePct) {
    let cum = 0;
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + ratePct / 100, t);
      if (cum >= 0) return t;
    }
    return null;
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + ratePct / 100, t);
    }
    return pv;
  }

  function getGrainPrice() {
    const el = document.getElementById("grainPrice");
    const v = el ? parseNumber(el.value) : NaN;
    if (Number.isFinite(v)) return v;
    const yieldOut = ensureYieldOutput();
    const p = Number(yieldOut.value) || 0;
    return Number.isFinite(p) ? p : 0;
  }

  function getPersistenceYears() {
    const el = document.getElementById("persistenceYears");
    const v = el ? parseNumber(el.value) : NaN;
    if (Number.isFinite(v) && v >= 0) return Math.floor(v);
    return Math.max(0, Math.floor(state.config.persistenceYears || 0));
  }

  function getRecurrenceYears(t) {
    const v = Number(t.recurrenceYears);
    return Number.isFinite(v) ? Math.max(0, Math.floor(v)) : 0;
  }

  function buildTreatmentCashflowsVsControl(t, opts) {
    const years = Math.max(0, Math.floor(opts.years));
    const disc = Number(opts.discountRatePct) || 0;
    const price = Number(opts.pricePerTonne) || 0;
    const persistence = Math.max(0, Math.floor(opts.persistenceYears));
    const recurrence = Math.max(0, Math.floor(opts.recurrenceYears));
    const adoption = clamp(Number(opts.adoptionMultiplier) || 1, 0, 1);
    const risk = clamp(Number(opts.riskMultiplier) || 0, 0, 1);

    const area = Number(t.area) || 0;
    const yieldOut = ensureYieldOutput();
    const yieldDelta = Number(t.deltas && t.deltas[yieldOut.id]) || 0;

    const benefitByYear = new Array(years + 1).fill(0);
    for (let y = 1; y <= years; y++) {
      if (persistence === 0 || y > persistence) {
        benefitByYear[y] = 0;
      } else {
        const annual = yieldDelta * price * area * adoption * (1 - risk);
        benefitByYear[y] = annual;
      }
    }

    const costByYear = new Array(years + 1).fill(0);
    const perHaApplicationCost =
      (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);

    const cap0 = Number(t.capitalCost) || 0;
    costByYear[0] += cap0;

    if (!t.isControl) {
      costByYear[0] += perHaApplicationCost * area;
      if (recurrence > 0) {
        for (let y = recurrence; y <= years; y += recurrence) {
          costByYear[y] += perHaApplicationCost * area;
        }
      }
    }

    const cf = new Array(years + 1).fill(0).map((_, i) => benefitByYear[i] - costByYear[i]);

    const pvBenefits = presentValue(benefitByYear, disc);
    const pvCosts = presentValue(costByYear, disc);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const pb = payback(cf, disc);

    return {
      treatmentId: t.id,
      treatmentName: t.name,
      isControl: !!t.isControl,
      areaHa: area,
      pricePerTonne: price,
      discountRatePct: disc,
      persistenceYears: persistence,
      recurrenceYears: recurrence,
      adoptionMultiplier: adoption,
      riskMultiplier: risk,
      perHaApplicationCost,
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roiPct: roi,
      irrPct: irrVal,
      mirrPct: mirrVal,
      paybackYears: pb,
      benefitByYear,
      costByYear,
      cf
    };
  }

  function computeBaseCaseResultsVsControl() {
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const years = Math.max(0, Math.floor(model.time.years || 0));
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const control = model.treatments.find(x => x.isControl) || null;

    const results = model.treatments.map(t =>
      buildTreatmentCashflowsVsControl(t, {
        pricePerTonne: price,
        discountRatePct: disc,
        years,
        persistenceYears: persistence,
        recurrenceYears: getRecurrenceYears(t),
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      })
    );

    const ranked = results
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => {
        const A = Number.isFinite(a.npv) ? a.npv : -Infinity;
        const B = Number.isFinite(b.npv) ? b.npv : -Infinity;
        return B - A;
      })
      .map((r, i) => ({ ...r, rankByNpv: i + 1 }));

    const out = results.map(r => {
      if (r.isControl) return { ...r, rankByNpv: null };
      const rr = ranked.find(x => x.treatmentId === r.treatmentId);
      return rr ? rr : { ...r, rankByNpv: null };
    });

    state.results.perTreatmentBaseCase = out;
    state.results.lastComputedAt = new Date().toISOString();

    const totalBenefit = new Array(years + 1).fill(0);
    const totalCost = new Array(years + 1).fill(0);
    out.forEach(r => {
      if (r.isControl) return;
      for (let y = 0; y <= years; y++) {
        totalBenefit[y] += r.benefitByYear[y] || 0;
        totalCost[y] += r.costByYear[y] || 0;
      }
    });
    const pvB = presentValue(totalBenefit, disc);
    const pvC = presentValue(totalCost, disc);
    const total = {
      pvBenefits: pvB,
      pvCosts: pvC,
      npv: pvB - pvC,
      bcr: pvC > 0 ? pvB / pvC : NaN
    };

    return { control, perTreatment: out, projectTotal: total };
  }

  function computeSensitivityGrid() {
    const priceGrid = (state.config.sensPrice || DEFAULT_SENS_PRICE).slice();
    const discGrid = (state.config.sensDiscount || DEFAULT_SENS_DISC).slice();
    const persistGrid = (state.config.sensPersistence || DEFAULT_SENS_PERSIST).slice();
    const recurGrid = (state.config.sensRecurrence || DEFAULT_SENS_RECURRENCE).slice();

    const years = Math.max(0, Math.floor(model.time.years || 0));
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const treatments = model.treatments.filter(t => !t.isControl);

    const grid = [];
    treatments.forEach(t => {
      const baseRec = getRecurrenceYears(t);
      recurGrid.forEach(rec => {
        const recurrenceYears = Number.isFinite(rec) ? Math.max(0, Math.floor(rec)) : baseRec;
        persistGrid.forEach(persistenceYears => {
          discGrid.forEach(discountRatePct => {
            priceGrid.forEach(pricePerTonne => {
              const r = buildTreatmentCashflowsVsControl(t, {
                pricePerTonne,
                discountRatePct,
                years,
                persistenceYears,
                recurrenceYears,
                adoptionMultiplier: adopt,
                riskMultiplier: risk
              });
              grid.push({
                treatment: t.name,
                treatmentId: t.id,
                pricePerTonne,
                discountRatePct,
                persistenceYears,
                recurrenceYears,
                pvBenefits: r.pvBenefits,
                pvCosts: r.pvCosts,
                npv: r.npv,
                bcr: r.bcr,
                roiPct: r.roiPct
              });
            });
          });
        });
      });
    });

    state.results.sensitivityGrid = grid;
    showToast("Sensitivity grid computed.");
    return grid;
  }

  // =========================
  // 7) RESULTS RENDERING (CONTROL-CENTRIC)
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const setVal = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

  function classifyDelta(val) {
    if (!Number.isFinite(val)) return "";
    if (val > 0) return "pos";
    if (val < 0) return "neg";
    return "zero";
  }

  function filterTreatments(results, mode) {
    const list = results.filter(r => !r.isControl);
    if (!mode || mode === "all") return list;

    if (mode === "top5_npv") {
      return list
        .slice()
        .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity))
        .slice(0, 5);
    }
    if (mode === "top5_bcr") {
      return list
        .slice()
        .sort((a, b) => (Number.isFinite(b.bcr) ? b.bcr : -Infinity) - (Number.isFinite(a.bcr) ? a.bcr : -Infinity))
        .slice(0, 5);
    }
    if (mode === "improve_only") {
      return list.filter(r => Number.isFinite(r.npv) && r.npv > 0);
    }
    return list;
  }

  function renderLeaderboard(perTreatment, filterMode) {
    const root = document.getElementById("resultsLeaderboard") || document.getElementById("leaderboard") || document.getElementById("treatmentLeaderboard");
    if (!root) return;

    const list = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    root.innerHTML = "";

    const table = document.createElement("table");
    table.className = "summary-table leaderboard-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Rank</th>
          <th>Treatment</th>
          <th>NPV</th>
          <th>BCR</th>
          <th>PV benefits</th>
          <th>PV costs</th>
        </tr>
      </thead>
      <tbody>
        ${list
          .map((r, i) => {
            const rank = i + 1;
            const npvCls = classifyDelta(r.npv);
            const bcrText = Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a";
            return `
              <tr>
                <td>${rank}</td>
                <td>${esc(r.treatmentName)}</td>
                <td class="${npvCls}">${money(r.npv)}</td>
                <td>${bcrText}</td>
                <td>${money(r.pvBenefits)}</td>
                <td>${money(r.pvCosts)}</td>
              </tr>
            `;
          })
          .join("")}
      </tbody>
    `;
    root.appendChild(table);
  }

  function renderComparisonToControl(perTreatment, filterMode) {
    const root =
      document.getElementById("comparisonToControl") ||
      document.getElementById("comparisonTable") ||
      document.getElementById("comparisonGrid") ||
      document.getElementById("resultsComparison");

    if (!root) return;

    const control = perTreatment.find(r => r.isControl) || null;
    const treatments = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const indicators = [
      { key: "pvBenefits", label: "PV benefits" },
      { key: "pvCosts", label: "PV costs" },
      { key: "npv", label: "NPV" },
      { key: "bcr", label: "BCR" },
      { key: "roiPct", label: "ROI" },
      { key: "rankByNpv", label: "Rank by NPV" },
      { key: "deltaNpv", label: "Δ NPV vs control" },
      { key: "deltaPvCost", label: "Δ PV cost vs control" }
    ];

    const colHeaders = [];
    colHeaders.push({ type: "control", name: control ? control.treatmentName : "Control (baseline)" });
    treatments.forEach(t => {
      colHeaders.push({ type: "treatment", name: t.treatmentName, id: t.treatmentId });
      colHeaders.push({ type: "delta", name: "Δ", id: t.treatmentId });
    });

    const table = document.createElement("table");
    table.className = "comparison-table summary-table";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th class="sticky-col">Indicator</th>
        ${colHeaders
          .map(h => {
            if (h.type === "control") return `<th class="sticky-head control-col">${esc(h.name)} (baseline)</th>`;
            if (h.type === "treatment") return `<th class="sticky-head">${esc(h.name)}</th>`;
            return `<th class="sticky-head delta-col">Δ vs control</th>`;
          })
          .join("")}
      </tr>
    `;

    const tbody = document.createElement("tbody");

    indicators.forEach(ind => {
      const tr = document.createElement("tr");
      const first = document.createElement("td");
      first.className = "sticky-col";
      first.textContent = ind.label;
      tr.appendChild(first);

      let controlVal = "";
      if (!control) controlVal = "n/a";
      else controlVal = formatIndicatorValue(ind.key, control, true);
      const tdControl = document.createElement("td");
      tdControl.className = "control-col";
      tdControl.textContent = controlVal;
      tr.appendChild(tdControl);

      treatments.forEach(t => {
        const td = document.createElement("td");
        td.textContent = formatIndicatorValue(ind.key, t, false);
        td.className = classifyIndicatorCell(ind.key, t, "value");
        tr.appendChild(td);

        const tdD = document.createElement("td");
        tdD.textContent = formatDeltaValue(ind.key, t);
        tdD.className = classifyIndicatorCell(ind.key, t, "delta");
        tr.appendChild(tdD);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);

    root.innerHTML = "";
    const wrap = document.createElement("div");
    wrap.className = "comparison-wrap";
    wrap.appendChild(table);
    root.appendChild(wrap);
  }

  function formatIndicatorValue(key, r, isControl) {
    if (key === "pvBenefits") return money(isControl ? 0 : r.pvBenefits);
    if (key === "pvCosts") return money(isControl ? 0 : r.pvCosts);
    if (key === "npv") return money(isControl ? 0 : r.npv);
    if (key === "bcr") return isControl ? "n/a" : (Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a");
    if (key === "roiPct") return isControl ? "n/a" : (Number.isFinite(r.roiPct) ? percent(r.roiPct) : "n/a");
    if (key === "rankByNpv") return isControl ? "" : (r.rankByNpv != null ? String(r.rankByNpv) : "");
    if (key === "deltaNpv") return isControl ? "" : money(r.npv);
    if (key === "deltaPvCost") return isControl ? "" : money(r.pvCosts);
    return "";
  }

  function formatDeltaValue(key, r) {
    if (key === "pvBenefits") return money(r.pvBenefits);
    if (key === "pvCosts") return money(r.pvCosts);
    if (key === "npv") return money(r.npv);
    if (key === "bcr") return Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a";
    if (key === "roiPct") return Number.isFinite(r.roiPct) ? percent(r.roiPct) : "n/a";
    if (key === "rankByNpv") return r.rankByNpv != null ? String(r.rankByNpv) : "";
    if (key === "deltaNpv") return money(r.npv);
    if (key === "deltaPvCost") return money(r.pvCosts);
    return "";
  }

  function classifyIndicatorCell(key, r, mode) {
    if (key === "pvCosts" || key === "deltaPvCost") {
      const v = r.pvCosts;
      if (!Number.isFinite(v)) return mode === "delta" ? "delta-cell" : "value-cell";
      const cls = v > 0 ? "neg" : v < 0 ? "pos" : "zero";
      return `${mode}-cell ${cls}`;
    }
    if (key === "npv" || key === "deltaNpv") {
      const cls = classifyDelta(r.npv);
      return `${mode}-cell ${cls}`;
    }
    if (key === "pvBenefits") {
      const cls = classifyDelta(r.pvBenefits);
      return `${mode}-cell ${cls}`;
    }
    if (key === "bcr") {
      const cls = Number.isFinite(r.bcr) ? (r.bcr >= 1 ? "pos" : "neg") : "";
      return `${mode}-cell ${cls}`;
    }
    return `${mode}-cell`;
  }

  function renderResultsNarrative(perTreatment, filterMode) {
    const root =
      document.getElementById("resultsNarrative") ||
      document.getElementById("whatThisMeans") ||
      document.getElementById("resultsWhatThisMeans");

    if (!root) return;

    const treatments = filterTreatments(perTreatment, filterMode)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const top = treatments[0] || null;
    const worst = treatments.slice().reverse().find(x => Number.isFinite(x.npv)) || null;

    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const years = Math.floor(model.time.years || 0);
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const parts = [];
    parts.push(
      `This results panel compares each treatment against the control baseline using discounted cashflows over ${years} years. The grain price used in the base case is ${money(price)} per tonne, the discount rate is ${fmt(disc)} percent per year, the assumed persistence of yield effects is ${persistence} years, the adoption multiplier is ${fmt(adopt)}, and the risk multiplier reduces benefits by ${fmt(risk)} as a proportion.`
    );

    if (top) {
      parts.push(
        `The strongest base case result by net present value is ${top.treatmentName}. Its present value of benefits is ${money(top.pvBenefits)}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(top.npv)}. This result is driven by the combination of yield uplift against the control and the incremental costs applied under the recurrence setting for that treatment.`
      );
    }

    if (worst && top && worst.treatmentId !== top.treatmentId) {
      parts.push(
        `A weaker base case result is ${worst.treatmentName}. Its present value of benefits is ${money(worst.pvBenefits)}, its present value of costs is ${money(worst.pvCosts)}, and its net present value is ${money(worst.npv)}. This pattern usually reflects either a smaller yield uplift compared with the control, higher incremental costs, or both.`
      );
    }

    const improves = treatments.filter(r => Number.isFinite(r.npv) && r.npv > 0).length;
    const total = treatments.length;
    if (total) {
      parts.push(
        `Under the current assumptions, ${improves} of ${total} treatments have a positive net present value relative to the control. This does not decide anything by itself, but it highlights which packages are more sensitive to costs and grain price assumptions.`
      );
    }

    root.textContent = parts.join("\n\n");
  }

  // =========================
  // 8) DATA CHECKS PANEL RENDERING
  // =========================
  function renderDataChecks() {
    const root = document.getElementById("dataChecks") || document.getElementById("dataChecksList") || document.getElementById("checksPanel");
    if (!root) return;

    const checks = (state.dataset.derived && state.dataset.derived.checks) ? state.dataset.derived.checks : [];
    if (!checks.length) {
      root.innerHTML = `<p class="small muted">No data checks triggered. If you have imported a dataset, this means required columns were found and core summaries could be computed.</p>`;
      return;
    }

    const rows = checks.slice().sort((a, b) => {
      const sevRank = s => (s === "error" ? 0 : s === "warn" ? 1 : 2);
      return sevRank(a.severity) - sevRank(b.severity);
    });

    root.innerHTML = `
      <table class="summary-table checks-table">
        <thead>
          <tr>
            <th>Severity</th>
            <th>Check</th>
            <th>Count</th>
            <th>Summary</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(r => {
              const sev = String(r.severity || "info").toUpperCase();
              const cls = r.severity === "error" ? "neg" : r.severity === "warn" ? "warn" : "zero";
              return `
                <tr>
                  <td class="${cls}">${esc(sev)}</td>
                  <td><code>${esc(r.code || "")}</code></td>
                  <td>${Number.isFinite(r.count) ? fmt(r.count) : ""}</td>
                  <td>${esc(r.message || "")}${r.detail ? " " + esc(r.detail) : ""}</td>
                </tr>
              `;
            })
            .join("")}
        </tbody>
      </table>
    `;
  }

  // =========================
  // 9) EXPORTS (TSV, CSV, XLSX if available)
  // =========================
  function downloadFile(filename, content, mime) {
    const blob = new Blob([content], { type: mime || "text/plain" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  }

  function toCsv(rows) {
    return rows
      .map(r =>
        r
          .map(x => {
            const s = x == null ? "" : String(x);
            const needs = s.includes(",") || s.includes('"') || s.includes("\n") || s.includes("\r");
            const safe = s.replace(/"/g, '""');
            return needs ? `"${safe}"` : safe;
          })
          .join(",")
      )
      .join("\r\n");
  }

  function exportCleanedDatasetTsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.plotDeltas || !derived.plotDeltas.length) {
      showToast("No cleaned dataset to export.");
      return;
    }
    const schema = state.dataset.schema;
    const rows = derived.plotDeltas;

    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
    const header = [
      "treatment",
      "treatment_key",
      "replicate",
      "plot",
      "is_control",
      "yield",
      "control_yield_mean_replicate",
      "delta_yield",
      ...costCols.map(c => `cost_per_ha:${c}`),
      ...costCols.map(c => `delta_cost_per_ha:${c}`)
    ];

    const lines = [header.join("\t")];
    rows.forEach(r => {
      const out = [
        r.treatment || "",
        r.treatmentKey || "",
        r.replicate || "",
        r.plot || "",
        r.isControl ? "1" : "0",
        Number.isFinite(r.yield) ? r.yield : "",
        Number.isFinite(r.controlYieldMeanRep) ? r.controlYieldMeanRep : "",
        Number.isFinite(r.deltaYield) ? r.deltaYield : ""
      ];

      costCols.forEach(c => out.push(Number.isFinite(r.costsPerHa[c]) ? r.costsPerHa[c] : ""));
      costCols.forEach(c => out.push(Number.isFinite(r.deltaCostsPerHa[c]) ? r.deltaCostsPerHa[c] : ""));

      lines.push(out.join("\t"));
    });

    const name = slug(model.project.name || "project");
    downloadFile(`${name}_cleaned_dataset.tsv`, lines.join("\n"), "text/tab-separated-values");
    showToast("Cleaned dataset TSV downloaded.");
  }

  function exportTreatmentSummaryCsv() {
    const derived = state.dataset.derived;
    if (!derived || !derived.treatmentSummary || !derived.treatmentSummary.length) {
      showToast("No treatment summary to export.");
      return;
    }

    const schema = state.dataset.schema;
    const costCols = schema && schema.costCols ? schema.costCols.slice() : [];

    const rows = [];
    rows.push([
      "treatment",
      "is_control",
      "n_yield",
      "yield_mean",
      "yield_sd",
      "delta_yield_mean",
      "delta_yield_sd",
      "delta_yield_median",
      ...costCols.map(c => `delta_cost_mean:${c}`),
      ...costCols.map(c => `delta_cost_sd:${c}`)
    ]);

    derived.treatmentSummary
      .slice()
      .sort((a, b) => (b.isControl ? 1 : 0) - (a.isControl ? 1 : 0))
      .forEach(s => {
        const row = [
          s.treatmentLabel || s.treatmentKey,
          s.isControl ? "1" : "0",
          Number.isFinite(s.nYield) ? s.nYield : "",
          Number.isFinite(s.yieldMean) ? s.yieldMean : "",
          Number.isFinite(s.yieldSD) ? s.yieldSD : "",
          Number.isFinite(s.deltaYieldMean) ? s.deltaYieldMean : "",
          Number.isFinite(s.deltaYieldSD) ? s.deltaYieldSD : "",
          Number.isFinite(s.deltaYieldMedian) ? s.deltaYieldMedian : ""
        ];
        costCols.forEach(c => row.push(Number.isFinite(s.deltaCosts?.[c]?.mean) ? s.deltaCosts[c].mean : ""));
        costCols.forEach(c => row.push(Number.isFinite(s.deltaCosts?.[c]?.sd) ? s.deltaCosts[c].sd : ""));
        rows.push(row);
      });

    const csv = toCsv(rows);
    const name = slug(model.project.name || "project");
    downloadFile(`${name}_treatment_summary.csv`, csv, "text/csv");
    showToast("Treatment summary CSV downloaded.");
  }

  function exportSensitivityGridCsv() {
    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      showToast("No sensitivity grid to export. Run sensitivity first.");
      return;
    }

    const rows = [];
    rows.push([
      "treatment",
      "price_per_tonne",
      "discount_rate_pct",
      "persistence_years",
      "recurrence_years",
      "pv_benefits",
      "pv_costs",
      "npv",
      "bcr",
      "roi_pct"
    ]);

    grid.forEach(g => {
      rows.push([
        g.treatment,
        g.pricePerTonne,
        g.discountRatePct,
        g.persistenceYears,
        g.recurrenceYears,
        g.pvBenefits,
        g.pvCosts,
        g.npv,
        g.bcr,
        g.roiPct
      ]);
    });

    const csv = toCsv(rows);
    const name = slug(model.project.name || "project");
    downloadFile(`${name}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Sensitivity grid CSV downloaded.");
  }

  function exportWorkbookIfAvailable() {
    if (typeof XLSX === "undefined") {
      showToast("Excel export requires the XLSX library.");
      return;
    }

    const wb = XLSX.utils.book_new();
    const name = slug(model.project.name || "project");

    const settingsAoA = [
      ["Project name", model.project.name],
      ["Analysis years", model.time.years],
      ["Discount rate base percent", model.time.discBase],
      ["Grain price per tonne", getGrainPrice()],
      ["Persistence years", getPersistenceYears()],
      ["Adoption multiplier", model.adoption.base],
      ["Risk multiplier", model.risk.base],
      ["Dataset source", state.dataset.sourceName || ""],
      ["Dataset committed at", state.dataset.committedAt || ""]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(settingsAoA), "Settings");

    const derived = state.dataset.derived;
    if (derived && derived.plotDeltas && derived.plotDeltas.length) {
      const schema = state.dataset.schema;
      const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
      const rows = derived.plotDeltas.map(r => {
        const obj = {
          treatment: r.treatment,
          treatment_key: r.treatmentKey,
          replicate: r.replicate,
          plot: r.plot,
          is_control: r.isControl ? 1 : 0,
          yield: Number.isFinite(r.yield) ? r.yield : null,
          control_yield_mean_replicate: Number.isFinite(r.controlYieldMeanRep) ? r.controlYieldMeanRep : null,
          delta_yield: Number.isFinite(r.deltaYield) ? r.deltaYield : null
        };
        costCols.forEach(c => { obj[`cost_per_ha:${c}`] = Number.isFinite(r.costsPerHa[c]) ? r.costsPerHa[c] : null; });
        costCols.forEach(c => { obj[`delta_cost_per_ha:${c}`] = Number.isFinite(r.deltaCostsPerHa[c]) ? r.deltaCostsPerHa[c] : null; });
        return obj;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "CleanedData");
    }

    if (derived && derived.treatmentSummary && derived.treatmentSummary.length) {
      const schema = state.dataset.schema;
      const costCols = schema && schema.costCols ? schema.costCols.slice() : [];
      const rows = derived.treatmentSummary.map(s => {
        const obj = {
          treatment: s.treatmentLabel || s.treatmentKey,
          is_control: s.isControl ? 1 : 0,
          n_yield: s.nYield,
          yield_mean: s.yieldMean,
          yield_sd: s.yieldSD,
          delta_yield_mean: s.deltaYieldMean,
          delta_yield_sd: s.deltaYieldSD,
          delta_yield_median: s.deltaYieldMedian
        };
        costCols.forEach(c => {
          obj[`delta_cost_mean:${c}`] = s.deltaCosts?.[c]?.mean ?? null;
          obj[`delta_cost_sd:${c}`] = s.deltaCosts?.[c]?.sd ?? null;
        });
        return obj;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "TreatmentSummary");
    }

    const base = state.results.perTreatmentBaseCase || [];
    if (base.length) {
      const rows = base.map(r => ({
        treatment: r.treatmentName,
        is_control: r.isControl ? 1 : 0,
        area_ha: r.areaHa,
        recurrence_years: r.recurrenceYears,
        price_per_tonne: r.pricePerTonne,
        discount_rate_pct: r.discountRatePct,
        persistence_years: r.persistenceYears,
        pv_benefits: r.pvBenefits,
        pv_costs: r.pvCosts,
        npv: r.npv,
        bcr: r.bcr,
        roi_pct: r.roiPct,
        irr_pct: r.irrPct,
        mirr_pct: r.mirrPct,
        payback_years: r.paybackYears
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "BaseCaseResults");
    }

    const grid = state.results.sensitivityGrid || [];
    if (grid.length) {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(grid), "SensitivityGrid");
    }

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${name}_workbook.xlsx`, wbout, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Excel workbook downloaded.");
  }

  // =========================
  // 10) AI BRIEFING (NO BULLETS, NO EM DASH, NO ABBREVIATIONS)
  // =========================
  function buildAiBriefingText() {
    const base = state.results.perTreatmentBaseCase || [];
    const treatments = base.filter(r => !r.isControl).slice().sort((a, b) => (b.npv || -Infinity) - (a.npv || -Infinity));

    const years = Math.floor(model.time.years || 0);
    const price = getGrainPrice();
    const disc = Number(model.time.discBase) || 0;
    const persistence = getPersistenceYears();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const derived = state.dataset.derived;
    const checks = derived && derived.checks ? derived.checks : [];

    const top = treatments[0] || null;

    const p = [];
    p.push(
      `Write a decision brief in clear plain language for an on farm manager. Use full sentences and paragraphs only. Do not use bullet points. Do not use an em dash. Do not use abbreviations.`
    );

    p.push(
      `The analysis compares each soil amendment treatment against the control baseline using discounted cashflows over ${years} years. The grain price used in the base case is ${money(price)} per tonne and the discount rate is ${fmt(disc)} percent per year. Yield effects are assumed to persist for ${persistence} years after application. The adoption multiplier is ${fmt(adopt)} and the risk multiplier reduces benefits by ${fmt(risk)} as a proportion.`
    );

    if (derived && derived.treatmentSummary && derived.treatmentSummary.length) {
      const nTreat = derived.treatmentSummary.filter(s => !s.isControl).length;
      const nRows = derived.plotDeltas ? derived.plotDeltas.length : 0;
      p.push(
        `The underlying dataset includes ${nRows} plot level records and ${nTreat} non control treatments. Treatment effects are computed using replicate specific control baselines, meaning each plot is compared with the control mean within the same replicate before averaging.`
      );
    }

    if (checks.length) {
      const errs = checks.filter(c => c.severity === "error").length;
      const warns = checks.filter(c => c.severity === "warn").length;
      p.push(
        `Data checks were run after import. There are ${errs} error level checks and ${warns} warning level checks. Summarise the most important implications for interpretation and sensitivity, without recommending a single option.`
      );
    } else {
      p.push(`Data checks were run after import and no triggers were recorded in the checks panel.`);
    }

    if (top) {
      p.push(
        `In the base case, the strongest treatment by net present value is ${top.treatmentName}. Its present value of benefits is ${money(top.pvBenefits)}, its present value of costs is ${money(top.pvCosts)}, and its net present value is ${money(top.npv)}. Explain in practical terms what is driving this result, focusing on yield uplift against the control and incremental costs under the recurrence assumption for that treatment.`
      );
    }

    if (treatments.length) {
      const pos = treatments.filter(t => Number.isFinite(t.npv) && t.npv > 0).length;
      p.push(
        `Across all treatments, ${pos} have a positive net present value relative to the control under the base case. Explain what this means in terms of trade offs. Do not instruct the user to choose anything.`
      );
    }

    p.push(`Include a section that explains the meaning of net present value, present value of benefits, present value of costs, benefit cost ratio, and return on investment in farmer facing terms.`);
    p.push(`Include a section that explains sensitivity, focusing on grain price, discount rate, persistence of yield effects, and recurrence of costs. Explain which inputs are likely to change and how that could move results.`);
    p.push(`Include a section that lists practical options to improve weaker treatments without giving a directive, such as reducing cost items, changing timing, improving establishment, or verifying the yield response with additional seasons or sites.`);

    const resultsBlock = {
      project: {
        name: model.project.name,
        years,
        discountRatePct: disc,
        grainPricePerTonne: price,
        persistenceYears: persistence,
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      },
      treatments: treatments.slice(0, 12).map(t => ({
        name: t.treatmentName,
        recurrenceYears: t.recurrenceYears,
        pvBenefits: t.pvBenefits,
        pvCosts: t.pvCosts,
        npv: t.npv,
        bcr: t.bcr,
        roiPct: t.roiPct,
        irrPct: t.irrPct,
        paybackYears: t.paybackYears
      })),
      dataChecks: checks.slice(0, 12)
    };

    p.push(`Use the following computed results as the only quantitative basis for the brief.`);
    p.push(JSON.stringify(resultsBlock, null, 2));
    return p.join("\n\n");
  }

  function renderAiBriefing() {
    const text = buildAiBriefingText();
    const box = document.getElementById("aiBriefingText") || document.getElementById("copilotPreview");
    if (box && "value" in box) box.value = text;

    const jsonBox = document.getElementById("resultsJson") || document.getElementById("aiResultsJson");
    if (jsonBox && "value" in jsonBox) {
      const payload = buildResultsJsonPayload();
      jsonBox.value = JSON.stringify(payload, null, 2);
    }
  }

  function buildResultsJsonPayload() {
    const derived = state.dataset.derived || null;
    const schema = state.dataset.schema || null;

    return {
      toolName: "Farming CBA Decision Tool",
      project: model.project,
      time: model.time,
      config: {
        grainPricePerTonne: getGrainPrice(),
        persistenceYears: getPersistenceYears(),
        adoptionMultiplier: model.adoption.base,
        riskMultiplier: model.risk.base
      },
      dataset: {
        sourceName: state.dataset.sourceName,
        committedAt: state.dataset.committedAt,
        schema,
        checks: derived ? derived.checks : [],
        treatmentSummary: derived ? derived.treatmentSummary : []
      },
      results: {
        baseCasePerTreatment: state.results.perTreatmentBaseCase,
        sensitivityGridCount: (state.results.sensitivityGrid || []).length
      }
    };
  }

  async function copyToClipboard(text, successMsg, failMsg) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        showToast(successMsg);
      } else {
        throw new Error("Clipboard API unavailable");
      }
    } catch {
      showToast(failMsg);
    }
  }

  // =========================
  // 11) IMPORT PIPELINE
  // =========================
  function renderImportSummary() {
    const el = document.getElementById("importSummary") || document.getElementById("dataImportSummary") || document.getElementById("importStatus");
    if (!el) return;

    const rows = state.dataset.rows || [];
    const schema = state.dataset.schema;
    const derived = state.dataset.derived;

    if (!rows.length || !schema) {
      el.textContent = "No dataset parsed yet.";
      return;
    }

    const parts = [];
    parts.push(`Rows parsed: ${rows.length.toLocaleString()}.`);
    parts.push(`Treatment column: ${schema.treatmentCol || "not found"}.`);
    parts.push(`Replicate column: ${schema.replicateCol || "not found"}.`);
    parts.push(`Yield column: ${schema.yieldCol || "not found"}.`);
    parts.push(`Cost columns: ${(schema.costCols || []).length.toLocaleString()}.`);
    if (derived && derived.controlKey) parts.push(`Detected control key: ${derived.controlKey}.`);
    if (state.dataset.committedAt) parts.push(`Committed at: ${state.dataset.committedAt}.`);
    el.textContent = parts.join(" ");
  }

  function parseAndStageFromText(rawText, sourceName) {
    const split = splitDictionaryAndDataFromText(rawText);
    const dict = parseDictionaryText(split.dictText);
    const dataText = split.dataText;

    const del = detectDelimiter(dataText);
    const tbl = parseDelimited(dataText, del);
    const rows = headersToObjects(tbl);

    state.dataset.sourceName = sourceName || "";
    state.dataset.rawText = rawText || "";
    state.dataset.dictionary = dict;
    state.dataset.rows = rows;
    state.dataset.schema = inferSchema(rows, dict);
    state.dataset.derived = computeDerivedFromDataset(rows, state.dataset.schema, dict);
    state.dataset.committedAt = null;

    renderImportSummary();
    renderDataChecks();
    showToast("Dataset parsed and staged. Review Data Checks, then commit.");
  }

  function commitStagedDataset() {
    const derived = state.dataset.derived;
    const schema = state.dataset.schema;

    const errors = (derived && derived.checks ? derived.checks.filter(c => c.severity === "error") : []);
    if (errors.length) {
      showToast("Cannot commit. Fix the error level data checks first.");
      renderDataChecks();
      return;
    }
    if (!schema || !schema.treatmentCol || !schema.yieldCol) {
      showToast("Cannot commit. Missing required columns.");
      renderDataChecks();
      return;
    }

    state.dataset.committedAt = new Date().toISOString();
    applyDatasetToModel();

    renderAll();
    setBasicsFieldsFromModel();
    calcAndRender();
    renderControlCentricResults();
    renderAiBriefing();

    showToast("Dataset committed. Results updated.");
  }

  function initImportBindings() {
    const fileInput =
      document.getElementById("dataFile") ||
      document.getElementById("datasetFile") ||
      document.getElementById("uploadData") ||
      document.getElementById("uploadFile") ||
      document.getElementById("trialFile");

    const pasteBox =
      document.getElementById("dataPaste") ||
      document.getElementById("datasetPaste") ||
      document.getElementById("pasteData") ||
      document.getElementById("pasteBox");

    const parseBtn =
      document.getElementById("parseData") ||
      document.getElementById("parseImport") ||
      document.getElementById("parseDataset");

    const commitBtn =
      document.getElementById("commitData") ||
      document.getElementById("commitImport") ||
      document.getElementById("importCommit");

    if (fileInput) {
      fileInput.addEventListener("change", async e => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        const text = await f.text();
        parseAndStageFromText(text, f.name);
        showToast("File loaded and parsed.");
        e.target.value = "";
      });
    }

    if (parseBtn && pasteBox) {
      parseBtn.addEventListener("click", e => {
        e.preventDefault();
        const text = String(pasteBox.value || "");
        if (!text.trim()) {
          showToast("Paste data is empty.");
          return;
        }
        parseAndStageFromText(text, "pasted_text");
      });
    }

    if (commitBtn) {
      commitBtn.addEventListener("click", e => {
        e.preventDefault();
        commitStagedDataset();
      });
    }

    document.addEventListener("click", e => {
      const btn = e.target.closest("[data-action]");
      if (!btn) return;
      const act = btn.getAttribute("data-action");
      if (!act) return;

      if (act === "parse-import") {
        e.preventDefault();
        const box =
          document.getElementById("dataPaste") ||
          document.getElementById("datasetPaste") ||
          document.getElementById("pasteData") ||
          document.getElementById("pasteBox");
        const text = box ? String(box.value || "") : "";
        if (!text.trim()) {
          showToast("Paste data is empty.");
          return;
        }
        parseAndStageFromText(text, "pasted_text");
        return;
      }

      if (act === "commit-import") {
        e.preventDefault();
        commitStagedDataset();
        return;
      }

      if (act === "export-cleaned-tsv") {
        e.preventDefault();
        exportCleanedDatasetTsv();
        return;
      }

      if (act === "export-treatment-summary-csv") {
        e.preventDefault();
        exportTreatmentSummaryCsv();
        return;
      }

      if (act === "run-sensitivity") {
        e.preventDefault();
        computeSensitivityGrid();
        renderSensitivitySummary();
        return;
      }

      if (act === "export-sensitivity-csv") {
        e.preventDefault();
        exportSensitivityGridCsv();
        return;
      }

      if (act === "export-workbook") {
        e.preventDefault();
        exportWorkbookIfAvailable();
        return;
      }

      if (act === "copy-ai-briefing") {
        e.preventDefault();
        const box = document.getElementById("aiBriefingText") || document.getElementById("copilotPreview");
        const txt = box && "value" in box ? String(box.value || "") : "";
        if (!txt.trim()) {
          showToast("AI briefing text is empty.");
          return;
        }
        copyToClipboard(txt, "AI briefing text copied.", "Unable to copy. Please copy manually.");
        return;
      }

      if (act === "copy-results-json") {
        e.preventDefault();
        const payload = buildResultsJsonPayload();
        const txt = JSON.stringify(payload, null, 2);
        copyToClipboard(txt, "Results JSON copied.", "Unable to copy. Please copy manually.");
        return;
      }

      if (act === "save-scenario") {
        e.preventDefault();
        saveScenarioFromUi();
        return;
      }

      if (act === "load-scenario") {
        e.preventDefault();
        loadScenarioFromUi();
        return;
      }

      if (act === "add-output") {
        e.preventDefault();
        addOutput();
        return;
      }

      if (act === "add-treatment") {
        e.preventDefault();
        addTreatment();
        return;
      }
    });
  }

  // =========================
  // 12) SCENARIOS: SAVE/LOAD TO LOCALSTORAGE
  // =========================
  function getScenarioStore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.scenarios);
      if (!raw) return {};
      const obj = JSON.parse(raw);
      return obj && typeof obj === "object" ? obj : {};
    } catch {
      return {};
    }
  }

  function setScenarioStore(obj) {
    try {
      localStorage.setItem(STORAGE_KEYS.scenarios, JSON.stringify(obj));
    } catch {
      // ignore
    }
  }

  function currentScenarioObject() {
    return {
      name: model.project.name,
      time: {
        years: model.time.years,
        discBase: model.time.discBase,
        discLow: model.time.discLow,
        discHigh: model.time.discHigh
      },
      config: {
        grainPricePerTonne: getGrainPrice(),
        persistenceYears: getPersistenceYears(),
        sensPrice: state.config.sensPrice,
        sensDiscount: state.config.sensDiscount,
        sensPersistence: state.config.sensPersistence,
        sensRecurrence: state.config.sensRecurrence
      },
      adoption: model.adoption,
      risk: model.risk,
      treatments: model.treatments.map(t => ({
        id: t.id,
        name: t.name,
        area: t.area,
        recurrenceYears: getRecurrenceYears(t),
        labourCost: t.labourCost,
        materialsCost: t.materialsCost,
        servicesCost: t.servicesCost,
        capitalCost: t.capitalCost,
        deltas: t.deltas,
        isControl: !!t.isControl
      })),
      dataset: {
        sourceName: state.dataset.sourceName,
        committedAt: state.dataset.committedAt
      }
    };
  }

  function applyScenarioObject(scn) {
    if (!scn || typeof scn !== "object") return;

    if (scn.time) {
      if (Number.isFinite(scn.time.years)) model.time.years = Math.floor(scn.time.years);
      if (Number.isFinite(scn.time.discBase)) model.time.discBase = scn.time.discBase;
      if (Number.isFinite(scn.time.discLow)) model.time.discLow = scn.time.discLow;
      if (Number.isFinite(scn.time.discHigh)) model.time.discHigh = scn.time.discHigh;
    }

    if (scn.adoption) {
      model.adoption.base = Number.isFinite(scn.adoption.base) ? scn.adoption.base : model.adoption.base;
      model.adoption.low = Number.isFinite(scn.adoption.low) ? scn.adoption.low : model.adoption.low;
      model.adoption.high = Number.isFinite(scn.adoption.high) ? scn.adoption.high : model.adoption.high;
    }

    if (scn.risk) {
      model.risk.base = Number.isFinite(scn.risk.base) ? scn.risk.base : model.risk.base;
      model.risk.low = Number.isFinite(scn.risk.low) ? scn.risk.low : model.risk.low;
      model.risk.high = Number.isFinite(scn.risk.high) ? scn.risk.high : model.risk.high;
      model.risk.tech = Number.isFinite(scn.risk.tech) ? scn.risk.tech : model.risk.tech;
      model.risk.nonCoop = Number.isFinite(scn.risk.nonCoop) ? scn.risk.nonCoop : model.risk.nonCoop;
      model.risk.socio = Number.isFinite(scn.risk.socio) ? scn.risk.socio : model.risk.socio;
      model.risk.fin = Number.isFinite(scn.risk.fin) ? scn.risk.fin : model.risk.fin;
      model.risk.man = Number.isFinite(scn.risk.man) ? scn.risk.man : model.risk.man;
    }

    if (scn.config) {
      if (Number.isFinite(scn.config.persistenceYears)) state.config.persistenceYears = Math.floor(scn.config.persistenceYears);
      if (Array.isArray(scn.config.sensPrice)) state.config.sensPrice = scn.config.sensPrice.slice();
      if (Array.isArray(scn.config.sensDiscount)) state.config.sensDiscount = scn.config.sensDiscount.slice();
      if (Array.isArray(scn.config.sensPersistence)) state.config.sensPersistence = scn.config.sensPersistence.slice();
      if (Array.isArray(scn.config.sensRecurrence)) state.config.sensRecurrence = scn.config.sensRecurrence.slice();

      const gp = document.getElementById("grainPrice");
      if (gp && Number.isFinite(scn.config.grainPricePerTonne)) gp.value = scn.config.grainPricePerTonne;
      const py = document.getElementById("persistenceYears");
      if (py && Number.isFinite(scn.config.persistenceYears)) py.value = scn.config.persistenceYears;
    }

    if (Array.isArray(scn.treatments) && scn.treatments.length) {
      const existing = model.treatments.slice();
      scn.treatments.forEach(st => {
        const match =
          existing.find(t => t.id === st.id) ||
          existing.find(t => String(t.name || "").trim().toLowerCase() === String(st.name || "").trim().toLowerCase());
        if (!match) return;
        if (Number.isFinite(st.area)) match.area = st.area;
        if (Number.isFinite(st.recurrenceYears)) match.recurrenceYears = Math.floor(st.recurrenceYears);
        if (Number.isFinite(st.labourCost)) match.labourCost = st.labourCost;
        if (Number.isFinite(st.materialsCost)) match.materialsCost = st.materialsCost;
        if (Number.isFinite(st.servicesCost)) match.servicesCost = st.servicesCost;
        if (Number.isFinite(st.capitalCost)) match.capitalCost = st.capitalCost;
        if (st.deltas && typeof st.deltas === "object") match.deltas = { ...match.deltas, ...st.deltas };
      });
      initTreatmentDeltas();
    }

    setBasicsFieldsFromModel();
    renderAll();
    calcAndRender();
    renderControlCentricResults();
    renderAiBriefing();
    showToast("Scenario applied.");
  }

  function saveScenarioFromUi() {
    const nameEl = document.getElementById("scenarioName") || document.getElementById("scenarioTitle");
    const nameRaw = nameEl ? String(nameEl.value || "").trim() : "";
    const name = nameRaw || `${slug(model.project.name || "scenario")}_${new Date().toISOString().slice(0, 10)}`;

    const store = getScenarioStore();
    store[name] = currentScenarioObject();
    setScenarioStore(store);

    try {
      localStorage.setItem(STORAGE_KEYS.activeScenario, name);
    } catch {
      // ignore
    }

    refreshScenarioSelect();
    showToast("Scenario saved.");
  }

  function loadScenarioFromUi() {
    const sel = document.getElementById("scenarioSelect") || document.getElementById("scenarioList");
    const chosen = sel ? String(sel.value || "").trim() : "";
    const store = getScenarioStore();
    if (!chosen || !store[chosen]) {
      showToast("No saved scenario selected.");
      return;
    }
    applyScenarioObject(store[chosen]);
    try {
      localStorage.setItem(STORAGE_KEYS.activeScenario, chosen);
    } catch {
      // ignore
    }
    showToast("Scenario loaded.");
  }

  function refreshScenarioSelect() {
    const sel = document.getElementById("scenarioSelect") || document.getElementById("scenarioList");
    if (!sel) return;

    const store = getScenarioStore();
    const names = Object.keys(store).sort((a, b) => a.localeCompare(b));

    const active = (() => {
      try {
        return localStorage.getItem(STORAGE_KEYS.activeScenario) || "";
      } catch {
        return "";
      }
    })();

    sel.innerHTML = `<option value="">Select a saved scenario</option>` + names.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
    if (active && names.includes(active)) sel.value = active;
  }

  function tryLoadActiveScenario() {
    let active = "";
    try {
      active = localStorage.getItem(STORAGE_KEYS.activeScenario) || "";
    } catch {
      active = "";
    }
    if (!active) return;
    const store = getScenarioStore();
    if (store[active]) applyScenarioObject(store[active]);
  }

  // =========================
  // 13) DOM: TABS
  // =========================
  function switchTab(target) {
    if (!target) return;

    const navEls = $$("[data-tab],[data-tab-target],[data-tab-jump]");
    navEls.forEach(el => {
      const key = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      const isActive = key === target;
      el.classList.toggle("active", isActive);
      el.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    const panels = $$(".tab-panel");
    panels.forEach(p => {
      const key = p.dataset.tabPanel || (p.id ? p.id.replace(/^tab-/, "") : "");
      const match = key === target || p.id === target || p.id === "tab-" + target;
      const show = !!match;
      p.classList.toggle("active", show);
      p.classList.toggle("show", show);
      p.hidden = !show;
      p.setAttribute("aria-hidden", show ? "false" : "true");
      p.style.display = show ? "" : "none";
    });

    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const el = e.target.closest("[data-tab],[data-tab-target],[data-tab-jump]");
      if (!el) return;
      const target = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      if (!target) return;
      e.preventDefault();
      switchTab(target);
      showToast(`Switched to ${target} tab.`);
    });

    const activeNav =
      document.querySelector("[data-tab].active, [data-tab-target].active, [data-tab-jump].active") ||
      document.querySelector("[data-tab], [data-tab-target], [data-tab-jump]");
    if (activeNav) {
      const target = activeNav.dataset.tab || activeNav.dataset.tabTarget || activeNav.dataset.tabJump;
      if (target) {
        switchTab(target);
        return;
      }
    }

    const firstPanel = document.querySelector(".tab-panel");
    if (firstPanel) {
      const key = firstPanel.dataset.tabPanel || (firstPanel.id ? firstPanel.id.replace(/^tab-/, "") : "");
      if (key) switchTab(key);
    }
  }

  // =========================
  // 14) FORMS: PROJECT + TIME + RISK
  // =========================
  function setBasicsFieldsFromModel() {
    if ($("#projectName")) $("#projectName").value = model.project.name || "";
    if ($("#projectLead")) $("#projectLead").value = model.project.lead || "";
    if ($("#analystNames")) $("#analystNames").value = model.project.analysts || "";
    if ($("#projectTeam")) $("#projectTeam").value = model.project.team || "";
    if ($("#projectSummary")) $("#projectSummary").value = model.project.summary || "";
    if ($("#projectObjectives")) $("#projectObjectives").value = model.project.objectives || "";
    if ($("#projectActivities")) $("#projectActivities").value = model.project.activities || "";
    if ($("#stakeholderGroups")) $("#stakeholderGroups").value = model.project.stakeholders || "";
    if ($("#lastUpdated")) $("#lastUpdated").value = model.project.lastUpdated || "";
    if ($("#projectGoal")) $("#projectGoal").value = model.project.goal || "";
    if ($("#withProject")) $("#withProject").value = model.project.withProject || "";
    if ($("#withoutProject")) $("#withoutProject").value = model.project.withoutProject || "";
    if ($("#organisation")) $("#organisation").value = model.project.organisation || "";
    if ($("#contactEmail")) $("#contactEmail").value = model.project.contactEmail || "";
    if ($("#contactPhone")) $("#contactPhone").value = model.project.contactPhone || "";

    if ($("#startYear")) $("#startYear").value = model.time.startYear;
    if ($("#projectStartYear")) $("#projectStartYear").value = model.time.projectStartYear || model.time.startYear;
    if ($("#years")) $("#years").value = model.time.years;
    if ($("#discBase")) $("#discBase").value = model.time.discBase;
    if ($("#discLow")) $("#discLow").value = model.time.discLow;
    if ($("#discHigh")) $("#discHigh").value = model.time.discHigh;
    if ($("#mirrFinance")) $("#mirrFinance").value = model.time.mirrFinance;
    if ($("#mirrReinvest")) $("#mirrReinvest").value = model.time.mirrReinvest;

    if ($("#adoptBase")) $("#adoptBase").value = model.adoption.base;
    if ($("#adoptLow")) $("#adoptLow").value = model.adoption.low;
    if ($("#adoptHigh")) $("#adoptHigh").value = model.adoption.high;

    if ($("#riskBase")) $("#riskBase").value = model.risk.base;
    if ($("#riskLow")) $("#riskLow").value = model.risk.low;
    if ($("#riskHigh")) $("#riskHigh").value = model.risk.high;
    if ($("#rTech")) $("#rTech").value = model.risk.tech;
    if ($("#rNonCoop")) $("#rNonCoop").value = model.risk.nonCoop;
    if ($("#rSocio")) $("#rSocio").value = model.risk.socio;
    if ($("#rFin")) $("#rFin").value = model.risk.fin;
    if ($("#rMan")) $("#rMan").value = model.risk.man;

    if ($("#persistenceYears")) $("#persistenceYears").value = getPersistenceYears();
    if ($("#grainPrice")) $("#grainPrice").value = getGrainPrice();

    const sched = model.time.discountSchedule || DEFAULT_DISCOUNT_SCHEDULE;
    $$("input[data-disc-period]").forEach(inp => {
      const idx = +inp.dataset.discPeriod;
      const scenario = inp.dataset.scenario;
      const row = sched[idx];
      if (!row) return;
      let v = "";
      if (scenario === "low") v = row.low;
      else if (scenario === "base") v = row.base;
      else if (scenario === "high") v = row.high;
      inp.value = v ?? "";
    });

    refreshScenarioSelect();
    renderImportSummary();
  }

  let debTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(debTimer);
    debTimer = setTimeout(calcAndRender, 120);
  }

  function bindBasics() {
    setBasicsFieldsFromModel();

    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", e => {
        e.stopPropagation();
        const r =
          1 -
          (1 - parseNumber($("#rTech")?.value)) *
            (1 - parseNumber($("#rNonCoop")?.value)) *
            (1 - parseNumber($("#rSocio")?.value)) *
            (1 - parseNumber($("#rFin")?.value)) *
            (1 - parseNumber($("#rMan")?.value));
        if ($("#combinedRiskOut")) $("#combinedRiskOut").textContent = "Combined: " + (r * 100).toFixed(2) + "%";
        if ($("#riskBase")) $("#riskBase").value = r.toFixed(3);
        model.risk.base = r;
        calcAndRender();
        showToast("Combined risk updated from component risks.");
      });
    }

    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const val = +t.value;
          if (scenario === "low") row.low = val;
          else if (scenario === "base") row.base = val;
          else if (scenario === "high") row.high = val;
          calcAndRenderDebounced();
        }
        return;
      }

      const id = t.id;
      if (!id) return;

      switch (id) {
        case "projectName": model.project.name = t.value; break;
        case "projectLead": model.project.lead = t.value; break;
        case "analystNames": model.project.analysts = t.value; break;
        case "projectTeam": model.project.team = t.value; break;
        case "projectSummary": model.project.summary = t.value; break;
        case "projectObjectives": model.project.objectives = t.value; break;
        case "projectActivities": model.project.activities = t.value; break;
        case "stakeholderGroups": model.project.stakeholders = t.value; break;
        case "lastUpdated": model.project.lastUpdated = t.value; break;
        case "projectGoal": model.project.goal = t.value; break;
        case "withProject": model.project.withProject = t.value; break;
        case "withoutProject": model.project.withoutProject = t.value; break;
        case "organisation": model.project.organisation = t.value; break;
        case "contactEmail": model.project.contactEmail = t.value; break;
        case "contactPhone": model.project.contactPhone = t.value; break;

        case "startYear": model.time.startYear = +t.value; break;
        case "projectStartYear": model.time.projectStartYear = +t.value; break;
        case "years": model.time.years = +t.value; break;
        case "discBase": model.time.discBase = +t.value; break;
        case "discLow": model.time.discLow = +t.value; break;
        case "discHigh": model.time.discHigh = +t.value; break;
        case "mirrFinance": model.time.mirrFinance = +t.value; break;
        case "mirrReinvest": model.time.mirrReinvest = +t.value; break;

        case "adoptBase": model.adoption.base = +t.value; break;
        case "adoptLow": model.adoption.low = +t.value; break;
        case "adoptHigh": model.adoption.high = +t.value; break;

        case "riskBase": model.risk.base = +t.value; break;
        case "riskLow": model.risk.low = +t.value; break;
        case "riskHigh": model.risk.high = +t.value; break;
        case "rTech": model.risk.tech = +t.value; break;
        case "rNonCoop": model.risk.nonCoop = +t.value; break;
        case "rSocio": model.risk.socio = +t.value; break;
        case "rFin": model.risk.fin = +t.value; break;
        case "rMan": model.risk.man = +t.value; break;

        case "persistenceYears":
          state.config.persistenceYears = Math.max(0, Math.floor(+t.value || 0));
          break;

        case "grainPrice": {
          const yieldOut = ensureYieldOutput();
          yieldOut.value = +t.value || yieldOut.value;
          break;
        }
      }

      calcAndRenderDebounced();
    });

    const saveScn = $("#saveScenario") || $("#saveScenarioBtn");
    if (saveScn) saveScn.addEventListener("click", e => { e.preventDefault(); saveScenarioFromUi(); });

    const loadScn = $("#loadScenario") || $("#loadScenarioBtn");
    if (loadScn) loadScn.addEventListener("click", e => { e.preventDefault(); loadScenarioFromUi(); });

    const scenarioSelect = $("#scenarioSelect") || $("#scenarioList");
    if (scenarioSelect) {
      scenarioSelect.addEventListener("change", () => {
        const val = String(scenarioSelect.value || "").trim();
        if (!val) return;
        const store = getScenarioStore();
        if (store[val]) {
          applyScenarioObject(store[val]);
          try { localStorage.setItem(STORAGE_KEYS.activeScenario, val); } catch { /* ignore */ }
          showToast("Scenario loaded.");
        }
      });
    }
  }

  // =========================
  // 15) RENDER OUTPUTS + TREATMENTS
  // =========================
  function addOutput() {
    const o = { id: uid(), name: "New output", unit: "units", value: 0, source: "Input Directly" };
    model.outputs.push(o);
    model.treatments.forEach(t => { if (!t.deltas) t.deltas = {}; t.deltas[o.id] = 0; });
    renderOutputs();
    renderTreatments();
    calcAndRenderDebounced();
    showToast("Output metric added.");
  }

  function addTreatment() {
    const control = model.treatments.find(t => t.isControl) || model.treatments[0];
    const area = control ? Number(control.area) || 100 : 100;
    const t = {
      id: uid(),
      name: "New treatment",
      area,
      adoption: 1,
      deltas: {},
      labourCost: 0,
      materialsCost: 0,
      servicesCost: 0,
      capitalCost: 0,
      constrained: false,
      source: "Input Directly",
      isControl: false,
      notes: "",
      recurrenceYears: 0
    };
    model.outputs.forEach(o => { t.deltas[o.id] = 0; });
    model.treatments.push(t);
    renderTreatments();
    calcAndRenderDebounced();
    showToast("Treatment added.");
  }

  function renderOutputs() {
    const root = $("#outputsList");
    if (!root) return;
    root.innerHTML = "";
    model.outputs.forEach(o => {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Output: ${esc(o.name)}</h4>
        <div class="row-6">
          <div class="field"><label>Name</label><input value="${esc(o.name)}" data-k="name" data-id="${o.id}" /></div>
          <div class="field"><label>Unit</label><input value="${esc(o.unit)}" data-k="unit" data-id="${o.id}" /></div>
          <div class="field"><label>Value ($ per unit)</label><input type="number" step="0.01" value="${o.value}" data-k="value" data-id="${o.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-k="source" data-id="${o.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                .map(s => `<option ${s === o.source ? "selected" : ""}>${s}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-output="${o.id}">Remove</button></div>
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${o.id}</code></div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const k = e.target.dataset.k;
      const id = e.target.dataset.id;
      if (!k || !id) return;
      const o = model.outputs.find(x => x.id === id);
      if (!o) return;
      if (k === "value") o[k] = +e.target.value;
      else o[k] = e.target.value;
      model.treatments.forEach(t => {
        if (!(id in t.deltas)) t.deltas[id] = 0;
      });
      renderTreatments();
      calcAndRenderDebounced();
    };

    root.onclick = e => {
      const id = e.target.dataset.delOutput;
      if (!id) return;
      if (!confirm("Remove this output metric?")) return;
      model.outputs = model.outputs.filter(o => o.id !== id);
      model.treatments.forEach(t => delete t.deltas[id]);
      renderOutputs();
      renderTreatments();
      calcAndRender();
      showToast("Output metric removed.");
    };

    const addBtn = document.getElementById("addOutput") || document.getElementById("addOutputBtn");
    if (addBtn && !addBtn.__bound) {
      addBtn.__bound = true;
      addBtn.addEventListener("click", e => { e.preventDefault(); addOutput(); });
    }
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";

    model.treatments.forEach(t => {
      const materials = Number(t.materialsCost) || 0;
      const services = Number(t.servicesCost) || 0;
      const labour = Number(t.labourCost) || 0;
      const totalPerHa = materials + services + labour;
      const rec = getRecurrenceYears(t);

      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${t.area}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Imported dataset","Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                .map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`)
                .join("")}
            </select>
          </div>
          <div class="field"><label>Control vs treatment</label>
            <select data-tk="isControl" data-id="${t.id}">
              <option value="treatment" ${!t.isControl ? "selected" : ""}>Treatment</option>
              <option value="control" ${t.isControl ? "selected" : ""}>Control</option>
            </select>
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Materials cost per ha per application</label><input type="number" step="0.01" value="${t.materialsCost || 0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost per ha per application</label><input type="number" step="0.01" value="${t.servicesCost || 0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost per ha per application</label><input type="number" step="0.01" value="${t.labourCost || 0}" data-tk="labourCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total cost per ha per application</label><input type="number" step="0.01" value="${totalPerHa}" readonly data-total-cost="${t.id}" /></div>
          <div class="field"><label>Capital cost (year 0)</label><input type="number" step="0.01" value="${t.capitalCost || 0}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Recurrence (years)</label>
            <select data-tk="recurrenceYears" data-id="${t.id}">
              ${[
                { v: 0, label: "Once only" },
                { v: 1, label: "Every year" },
                { v: 2, label: "Every 2 years" },
                { v: 3, label: "Every 3 years" },
                { v: 4, label: "Every 4 years" },
                { v: 5, label: "Every 5 years" },
                { v: 7, label: "Every 7 years" },
                { v: 10, label: "Every 10 years" }
              ]
                .map(opt => `<option value="${opt.v}" ${Number(rec) === opt.v ? "selected" : ""}>${opt.label}</option>`)
                .join("")}
            </select>
          </div>
        </div>

        <div class="field">
          <label>Notes</label>
          <textarea data-tk="notes" data-id="${t.id}" rows="2">${esc(t.notes || "")}</textarea>
        </div>

        <h5>Output deltas (per ha)</h5>
        <div class="row">
          ${model.outputs
            .map(
              o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${t.deltas[o.id] ?? 0}" data-td="${o.id}" data-id="${t.id}" />
            </div>
          `
            )
            .join("")}
        </div>
        <div class="kv"><small class="muted">id:</small> <code>${t.id}</code></div>
      `;
      root.appendChild(el);
    });

    root.oninput = e => {
      const id = e.target.dataset.id;
      if (!id) return;
      const t = model.treatments.find(x => x.id === id);
      if (!t) return;

      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "name" || tk === "source" || tk === "notes") {
          t[tk] = e.target.value;
        } else if (tk === "isControl") {
          const val = e.target.value === "control";
          model.treatments.forEach(tt => (tt.isControl = false));
          if (val) t.isControl = true;
          renderTreatments();
          calcAndRenderDebounced();
          showToast(`Control treatment set to ${t.name}.`);
          return;
        } else if (tk === "recurrenceYears") {
          t.recurrenceYears = Math.max(0, Math.floor(+e.target.value || 0));
        } else {
          t[tk] = +e.target.value;
        }

        if (tk === "materialsCost" || tk === "servicesCost" || tk === "labourCost") {
          const container = e.target.closest(".item");
          if (container) {
            const mats = parseNumber(container.querySelector(`input[data-tk="materialsCost"][data-id="${id}"]`)?.value) || 0;
            const serv = parseNumber(container.querySelector(`input[data-tk="servicesCost"][data-id="${id}"]`)?.value) || 0;
            const lab = parseNumber(container.querySelector(`input[data-tk="labourCost"][data-id="${id}"]`)?.value) || 0;
            const totalField = container.querySelector(`input[data-total-cost="${id}"]`);
            if (totalField) totalField.value = mats + serv + lab;
          }
        }
      }

      const td = e.target.dataset.td;
      if (td) t.deltas[td] = +e.target.value;

      calcAndRenderDebounced();
    };

    root.addEventListener("click", e => {
      const id = e.target.dataset.delTreatment;
      if (!id) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== id);
      if (!model.treatments.some(t => t.isControl) && model.treatments.length) model.treatments[0].isControl = true;
      renderTreatments();
      calcAndRender();
      renderControlCentricResults();
      showToast("Treatment removed.");
    });

    const addBtn = document.getElementById("addTreatment") || document.getElementById("addTreatmentBtn");
    if (addBtn && !addBtn.__bound) {
      addBtn.__bound = true;
      addBtn.addEventListener("click", e => { e.preventDefault(); addTreatment(); });
    }
  }

  function renderAll() {
    renderOutputs();
    renderTreatments();
    renderDataChecks();
    refreshScenarioSelect();
  }

  // =========================
  // 16) RESULTS + FILTERS
  // =========================
  function renderControlCentricResults(filterMode) {
    const base = computeBaseCaseResultsVsControl();
    renderLeaderboard(base.perTreatment, filterMode || currentResultsFilterMode());
    renderComparisonToControl(base.perTreatment, filterMode || currentResultsFilterMode());
    renderResultsNarrative(base.perTreatment, filterMode || currentResultsFilterMode());
    renderAiBriefing();
  }

  function currentResultsFilterMode() {
    const sel =
      document.getElementById("resultsFilter") ||
      document.getElementById("resultsQuickFilter") ||
      document.getElementById("comparisonFilter");
    const v = sel ? String(sel.value || "").trim() : "";
    return v || "all";
  }

  function bindResultsFilters() {
    const sel =
      document.getElementById("resultsFilter") ||
      document.getElementById("resultsQuickFilter") ||
      document.getElementById("comparisonFilter");
    if (sel && !sel.__bound) {
      sel.__bound = true;
      sel.addEventListener("change", () => {
        renderControlCentricResults(currentResultsFilterMode());
        showToast("Results filter applied.");
      });
    }

    const map = [
      { id: "filterTopNpv", mode: "top5_npv" },
      { id: "filterTopBcr", mode: "top5_bcr" },
      { id: "filterImproveOnly", mode: "improve_only" },
      { id: "filterShowAll", mode: "all" }
    ];
    map.forEach(m => {
      const b = document.getElementById(m.id);
      if (b && !b.__bound) {
        b.__bound = true;
        b.addEventListener("click", e => {
          e.preventDefault();
          renderControlCentricResults(m.mode);
          showToast("Results filter applied.");
        });
      }
    });

    document.addEventListener("click", e => {
      const b = e.target.closest("[data-results-filter]");
      if (!b) return;
      const mode = String(b.getAttribute("data-results-filter") || "").trim();
      if (!mode) return;
      e.preventDefault();
      renderControlCentricResults(mode);
      showToast("Results filter applied.");
    });
  }

  // =========================
  // 17) TOP SUMMARY CARDS (BACKWARD COMPATIBLE)
  // =========================
  function setFirstExistingText(ids, value) {
    for (const id of ids) {
      const el = document.getElementById(id);
      if (el) {
        el.textContent = value;
        return true;
      }
    }
    return false;
  }

  function calcAndRender() {
    const disc = Number(model.time.discBase) || 0;
    const years = Math.max(0, Math.floor(model.time.years || 0));
    const persistence = getPersistenceYears();
    const price = getGrainPrice();
    const adopt = clamp(Number(model.adoption.base) || 1, 0, 1);
    const risk = clamp(Number(model.risk.base) || 0, 0, 1);

    const treatments = model.treatments.filter(t => !t.isControl);
    const benefitByYear = new Array(years + 1).fill(0);
    const costByYear = new Array(years + 1).fill(0);

    treatments.forEach(t => {
      const r = buildTreatmentCashflowsVsControl(t, {
        pricePerTonne: price,
        discountRatePct: disc,
        years,
        persistenceYears: persistence,
        recurrenceYears: getRecurrenceYears(t),
        adoptionMultiplier: adopt,
        riskMultiplier: risk
      });
      for (let y = 0; y <= years; y++) {
        benefitByYear[y] += r.benefitByYear[y] || 0;
        costByYear[y] += r.costByYear[y] || 0;
      }
    });

    const pvB = presentValue(benefitByYear, disc);
    const pvC = presentValue(costByYear, disc);
    const npv = pvB - pvC;
    const bcr = pvC > 0 ? pvB / pvC : NaN;
    const roiPct = pvC > 0 ? (npv / pvC) * 100 : NaN;

    setFirstExistingText(["pvBenefits", "pvBenefitsVal", "kpiPvBenefits", "kpiPVB"], money(pvB));
    setFirstExistingText(["pvCosts", "pvCostsVal", "kpiPvCosts", "kpiPVC"], money(pvC));
    setFirstExistingText(["npv", "npvVal", "kpiNpv", "kpiNPV"], money(npv));
    setFirstExistingText(["bcr", "bcrVal", "kpiBcr", "kpiBCR"], Number.isFinite(bcr) ? fmt(bcr) : "n/a");
    setFirstExistingText(["roi", "roiVal", "kpiRoi", "kpiROI"], Number.isFinite(roiPct) ? percent(roiPct) : "n/a");

    renderControlCentricResults(currentResultsFilterMode());
    renderAiBriefing();
  }

  // =========================
  // 18) SENSITIVITY SUMMARY RENDERING
  // =========================
  function renderSensitivitySummary() {
    const root =
      document.getElementById("sensitivitySummary") ||
      document.getElementById("sensitivityResults") ||
      document.getElementById("sensitivityPanel");

    if (!root) return;

    const grid = state.results.sensitivityGrid || [];
    if (!grid.length) {
      root.innerHTML = `<p class="small muted">No sensitivity results yet. Run sensitivity to generate the grid.</p>`;
      return;
    }

    const treatments = Array.from(new Set(grid.map(g => g.treatment)));
    const rows = grid.length;

    const byTreat = new Map();
    treatments.forEach(t => byTreat.set(t, []));
    grid.forEach(g => { if (byTreat.has(g.treatment)) byTreat.get(g.treatment).push(g); });

    const summary = [];
    for (const [t, arr] of byTreat.entries()) {
      const npvs = arr.map(x => x.npv).filter(Number.isFinite);
      const bcrs = arr.map(x => x.bcr).filter(Number.isFinite);
      summary.push({
        treatment: t,
        scenarios: arr.length,
        npv_min: npvs.length ? Math.min(...npvs) : NaN,
        npv_med: median(npvs),
        npv_max: npvs.length ? Math.max(...npvs) : NaN,
        bcr_med: median(bcrs)
      });
    }

    summary.sort((a, b) => (Number.isFinite(b.npv_med) ? b.npv_med : -Infinity) - (Number.isFinite(a.npv_med) ? a.npv_med : -Infinity));

    root.innerHTML = `
      <div class="small muted">Grid rows: ${rows.toLocaleString()}. Treatments: ${treatments.length.toLocaleString()}.</div>
      <table class="summary-table">
        <thead>
          <tr>
            <th>Treatment</th>
            <th>Scenarios</th>
            <th>NPV min</th>
            <th>NPV median</th>
            <th>NPV max</th>
            <th>BCR median</th>
          </tr>
        </thead>
        <tbody>
          ${summary.slice(0, 20).map(s => `
            <tr>
              <td>${esc(s.treatment)}</td>
              <td>${fmt(s.scenarios)}</td>
              <td class="${classifyDelta(s.npv_min)}">${money(s.npv_min)}</td>
              <td class="${classifyDelta(s.npv_med)}">${money(s.npv_med)}</td>
              <td class="${classifyDelta(s.npv_max)}">${money(s.npv_max)}</td>
              <td>${Number.isFinite(s.bcr_med) ? fmt(s.bcr_med) : "n/a"}</td>
            </tr>
          `).join("")}
        </tbody>
      </table>
      <div class="small muted">Showing up to 20 treatments by median NPV across the grid.</div>
    `;
  }

  function bindSensitivityControls() {
    const runBtn = document.getElementById("runSensitivity") || document.getElementById("runSensitivityBtn");
    if (runBtn && !runBtn.__bound) {
      runBtn.__bound = true;
      runBtn.addEventListener("click", e => {
        e.preventDefault();
        computeSensitivityGrid();
        renderSensitivitySummary();
      });
    }

    const exportBtn = document.getElementById("exportSensitivity") || document.getElementById("exportSensitivityBtn");
    if (exportBtn && !exportBtn.__bound) {
      exportBtn.__bound = true;
      exportBtn.addEventListener("click", e => {
        e.preventDefault();
        exportSensitivityGridCsv();
      });
    }
  }

  // =========================
  // 19) AI COPY BUTTONS (OPTIONAL IDS)
  // =========================
  function bindAiButtons() {
    const copyBrief = document.getElementById("copyAiBriefing") || document.getElementById("copyBriefingBtn");
    if (copyBrief && !copyBrief.__bound) {
      copyBrief.__bound = true;
      copyBrief.addEventListener("click", async e => {
        e.preventDefault();
        const box = document.getElementById("aiBriefingText") || document.getElementById("copilotPreview");
        const txt = box && "value" in box ? String(box.value || "") : "";
        if (!txt.trim()) return showToast("AI briefing text is empty.");
        await copyToClipboard(txt, "AI briefing text copied.", "Unable to copy. Please copy manually.");
      });
    }

    const copyJson = document.getElementById("copyResultsJson") || document.getElementById("copyResultsJsonBtn");
    if (copyJson && !copyJson.__bound) {
      copyJson.__bound = true;
      copyJson.addEventListener("click", async e => {
        e.preventDefault();
        const payload = buildResultsJsonPayload();
        await copyToClipboard(JSON.stringify(payload, null, 2), "Results JSON copied.", "Unable to copy. Please copy manually.");
      });
    }
  }

  // =========================
  // 20) INIT
  // =========================
  function init() {
    ensureToastRoot();
    initTabs();
    initImportBindings();
    bindBasics();
    bindResultsFilters();
    bindSensitivityControls();
    bindAiButtons();

    renderAll();
    renderImportSummary();
    renderDataChecks();
    refreshScenarioSelect();

    tryLoadActiveScenario();

    calcAndRender();
    renderSensitivitySummary();
    renderAiBriefing();

    // If a Results tab exists and nothing is active, prefer it
    const preferResults = document.querySelector('[data-tab="results"],[data-tab-target="results"],[data-tab-jump="results"]');
    if (preferResults && !document.querySelector("[data-tab].active,[data-tab-target].active,[data-tab-jump].active")) {
      switchTab("results");
    }

    showToast("Tool ready.");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
