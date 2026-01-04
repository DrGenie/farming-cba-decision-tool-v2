// Farming CBA Tool - Newcastle Business School
// Fully upgraded script: robust TSV/CSV/TXT+dictionary import (upload + paste),
// replicate-specific control baselines, plot-level deltas, missing-safe summaries,
// cost scaling with per-treatment recurrence, full discounted CBA + sensitivity grid,
// scenario save/load, comparison-to-control Results grid + filters + narrative,
// exports (cleaned TSV, treatment CSV, sensitivity CSV, XLSX workbook),
// AI Briefing (copy-ready narrative prompt: no bullets, no em dash, no abbreviations) + Copy Results JSON,
// and bottom-right toasts for major actions.

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

  const APP_STORAGE_PREFIX = "farming_cba_tool_v3";
  const STORAGE_KEYS = {
    scenarioIndex: `${APP_STORAGE_PREFIX}.scenario_index`,
    lastScenario: `${APP_STORAGE_PREFIX}.last_scenario`,
    lastImport: `${APP_STORAGE_PREFIX}.last_import`
  };

  const DEFAULT_SENSITIVITY = {
    priceMultipliers: [0.8, 0.9, 1.0, 1.1, 1.2],
    discountRatesPct: null, // if null -> [discLow, discBase, discHigh] from model
    persistenceYears: [1, 2, 3, 5, 7, 10],
    recurrenceMultipliers: [1.0] // scaling on recurring costs and recurring benefits if configured
  };

  // =========================
  // 1) ID + UTIL HELPERS
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

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

  function parseNumber(value) {
    if (value === null || value === undefined) return NaN;
    if (typeof value === "number") return Number.isFinite(value) ? value : NaN;
    const str = String(value).trim();
    if (!str) return NaN;
    if (str === "?" || str.toLowerCase() === "na" || str.toLowerCase() === "n/a" || str.toLowerCase() === "null")
      return NaN;
    const cleaned = str.replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function safeMean(arr) {
    const clean = arr.filter(v => Number.isFinite(v));
    if (!clean.length) return NaN;
    return clean.reduce((a, b) => a + b, 0) / clean.length;
  }

  function safeSum(arr) {
    const clean = arr.filter(v => Number.isFinite(v));
    if (!clean.length) return 0;
    return clean.reduce((a, b) => a + b, 0);
  }

  function safeMedian(arr) {
    const clean = arr.filter(v => Number.isFinite(v)).slice().sort((a, b) => a - b);
    if (!clean.length) return NaN;
    const mid = Math.floor(clean.length / 2);
    return clean.length % 2 ? clean[mid] : (clean[mid - 1] + clean[mid]) / 2;
  }

  function safeStd(arr) {
    const clean = arr.filter(v => Number.isFinite(v));
    if (clean.length < 2) return NaN;
    const m = safeMean(clean);
    const v = clean.reduce((a, b) => a + (b - m) * (b - m), 0) / (clean.length - 1);
    return Math.sqrt(v);
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
    const F = (c - a) / (b - a || 1);
    if (r < F) return a + Math.sqrt(r * (b - a) * (c - a));
    return b - Math.sqrt((1 - r) * (b - a) * (b - c));
  }

  function showToast(message) {
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

  // DOM
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));
  const num = sel => +(document.querySelector(sel)?.value || 0);
  const setVal = (sel, text) => {
    const el = document.querySelector(sel);
    if (el) el.textContent = text;
  };

  // Clipboard
  async function copyToClipboard(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
        return true;
      }
    } catch (e) {
      // ignore
    }
    try {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.setAttribute("readonly", "true");
      ta.style.position = "fixed";
      ta.style.top = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand("copy");
      document.body.removeChild(ta);
      return ok;
    } catch (e) {
      return false;
    }
  }

  // =========================
  // 2) CORE MODEL
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
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Input Directly" },
      { id: uid(), name: "Screenings", unit: "percentage point", value: -20, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "percentage point", value: 10, source: "Input Directly" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (no amendment)",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 40,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        notes: "Baseline faba bean practice without deep soil amendment.",
        recurrence: {
          cost: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" },
          benefit: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" }
        }
      },
      {
        id: uid(),
        name: "Deep organic matter CP1",
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 60,
        materialsCost: 16500,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: false,
        notes: "Deep incorporation of organic matter at CP1 rate.",
        recurrence: {
          cost: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" },
          benefit: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" }
        }
      }
    ],
    benefits: [
      {
        id: uid(),
        label: "Reduced recurring costs (energy and water)",
        category: "C4",
        theme: "Cost savings",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        year: new Date().getFullYear(),
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 15000,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 120000,
        notes: "Project wide operating cost saving"
      },
      {
        id: uid(),
        label: "Reduced risk of quality downgrades",
        category: "C7",
        theme: "Risk reduction",
        frequency: "Annual",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 9,
        year: new Date().getFullYear(),
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 0,
        growthPct: 0,
        linkAdoption: true,
        linkRisk: false,
        p0: 0.1,
        p1: 0.07,
        consequence: 120000,
        notes: ""
      },
      {
        id: uid(),
        label: "Soil asset value uplift (carbon and structure)",
        category: "C6",
        theme: "Soil carbon",
        frequency: "Once",
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear(),
        year: new Date().getFullYear() + 5,
        unitValue: 0,
        quantity: 0,
        abatement: 0,
        annualAmount: 50000,
        growthPct: 0,
        linkAdoption: false,
        linkRisk: true,
        p0: 0,
        p1: 0,
        consequence: 0,
        notes: ""
      }
    ],
    otherCosts: [
      {
        id: uid(),
        label: "Project management and monitoring and evaluation",
        type: "annual",
        category: "Capital",
        annual: 20000,
        startYear: new Date().getFullYear(),
        endYear: new Date().getFullYear() + 4,
        capital: 50000,
        year: new Date().getFullYear(),
        constrained: true,
        depMethod: "declining",
        depLife: 5,
        depRate: 30
      }
    ],
    adoption: { base: 0.9, low: 0.6, high: 1.0 },
    risk: {
      base: 0.15,
      low: 0.05,
      high: 0.3,
      tech: 0.05,
      nonCoop: 0.04,
      socio: 0.02,
      fin: 0.03,
      man: 0.02
    },
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
    },
    // Trial and CBA calibration settings
    calibration: {
      mode: "model", // "model" uses model.treatment deltas; "trial" uses imported trial plot data
      pricePerTonne: null, // if null uses outputs "Grain yield" value
      persistenceYears: null, // if null uses model.time.years
      controlNameHint: "control",
      // If provided via import dictionary, these can override heuristics:
      columnMap: {
        treatment: "",
        replicate: "",
        plot: "",
        yield: ""
      }
    },
    sensitivity: JSON.parse(JSON.stringify(DEFAULT_SENSITIVITY))
  };

  // Ensure treatment deltas exist and recurrence blocks exist
  function initTreatmentDeltasAndRecurrence() {
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (typeof t.labourCost === "undefined") t.labourCost = Number(t.annualCost || 0) || 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      if (!t.recurrence) {
        t.recurrence = {
          cost: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" },
          benefit: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" }
        };
      } else {
        if (!t.recurrence.cost) t.recurrence.cost = { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" };
        if (!t.recurrence.benefit) t.recurrence.benefit = { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" };
      }
      delete t.annualCost;
    });
  }
  initTreatmentDeltasAndRecurrence();

  // =========================
  // 3) OPTIONAL EMBEDDED TRIAL DEFAULTS (FALLBACK)
  // =========================
  const FABABEAN_SHEET_NAMES = ["FabaBeanRaw", "FabaBeansRaw", "FabaBean", "FabaBeans"];

  const RAW_PLOTS = [
    { Amendment: "control", "Yield t/ha": 2.4, "Pre sowing Labour": 40, "Treatment Input Cost Only /Ha": 0 },
    { Amendment: "deep_om_cp1", "Yield t/ha": 3.1, "Pre sowing Labour": 55, "Treatment Input Cost Only /Ha": 16500 },
    { Amendment: "deep_om_cp1_plus_liq_gypsum_cht", "Yield t/ha": 3.2, "Pre sowing Labour": 56, "Treatment Input Cost Only /Ha": 16850 },
    { Amendment: "deep_gypsum", "Yield t/ha": 2.9, "Pre sowing Labour": 50, "Treatment Input Cost Only /Ha": 500 },
    { Amendment: "deep_om_cp1_plus_pam", "Yield t/ha": 3.0, "Pre sowing Labour": 57, "Treatment Input Cost Only /Ha": 18000 },
    { Amendment: "deep_om_cp1_plus_ccm", "Yield t/ha": 3.25, "Pre sowing Labour": 58, "Treatment Input Cost Only /Ha": 21225 },
    { Amendment: "deep_ccm_only", "Yield t/ha": 2.95, "Pre sowing Labour": 52, "Treatment Input Cost Only /Ha": 3225 },
    { Amendment: "deep_om_cp2_plus_gypsum", "Yield t/ha": 3.3, "Pre sowing Labour": 60, "Treatment Input Cost Only /Ha": 24000 },
    { Amendment: "deep_liq_gypsum_cht", "Yield t/ha": 2.8, "Pre sowing Labour": 48, "Treatment Input Cost Only /Ha": 350 },
    { Amendment: "surface_silicon", "Yield t/ha": 2.7, "Pre sowing Labour": 45, "Treatment Input Cost Only /Ha": 1000 },
    { Amendment: "deep_liq_npks", "Yield t/ha": 3.0, "Pre sowing Labour": 53, "Treatment Input Cost Only /Ha": 2200 },
    { Amendment: "deep_ripping_only", "Yield t/ha": 2.85, "Pre sowing Labour": 47, "Treatment Input Cost Only /Ha": 0 }
  ];

  const LABOUR_COLUMNS = [
    "Pre sowing Labour",
    "Amendment Labour",
    "Sowing Labour",
    "Herbicide Labour",
    "Herbicide Labour 2",
    "Herbicide Labour 3",
    "Harvesting Labour",
    "Harvesting Labour 2"
  ];

  const OPERATING_COLUMNS = [
    "Treatment Input Cost Only /Ha",
    "Cavalier (Oxyfluofen 240)",
    "Factor",
    "Roundup CT",
    "Roundup Ultra Max",
    "Supercharge Elite Discontinued",
    "Platnium (Clethodim 360)",
    "Mentor",
    "Simazine 900",
    "Veritas Opti",
    "FLUTRIAFOL fungicide",
    "Barrack fungicide discontinued",
    "Talstar"
  ];

  function slugifyTreatmentName(name) {
    return (name || "")
      .toString()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  // =========================
  // 4) IMPORT PIPELINE: UPLOAD + PASTE (TSV/CSV/TXT) + DICTIONARY PARSING
  // =========================
  const trialState = {
    source: "embedded",
    importedAt: null,
    dictionary: null,
    dataRows: null, // array of objects, keys are column names
    detected: {
      treatmentCol: null,
      replicateCol: null,
      plotCol: null,
      yieldCol: null,
      costCols: { labour: [], operating: [], other: [] }
    },
    validation: { issues: [], summary: "" },
    metrics: null, // computed summaries including replicate-specific controls and plot deltas
    cleaned: null // cleaned table with normalised keys and numeric parsing (missing-safe)
  };

  function detectDelimiter(text) {
    const firstLine = (text || "").split(/\r?\n/).find(l => l.trim().length > 0) || "";
    const tabs = (firstLine.match(/\t/g) || []).length;
    const commas = (firstLine.match(/,/g) || []).length;
    const semis = (firstLine.match(/;/g) || []).length;
    if (tabs >= commas && tabs >= semis && tabs > 0) return "\t";
    if (commas >= semis && commas > 0) return ",";
    if (semis > 0) return ";";
    return "\t";
  }

  function parseDelimitedTextToRows(text, delimiter) {
    const d = delimiter || detectDelimiter(text);
    const lines = (text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    const nonEmpty = lines.filter(l => l.trim().length > 0);
    if (!nonEmpty.length) return { header: [], rows: [] };

    // Simple robust CSV/TSV parsing with quotes support
    function splitLine(line, delim) {
      const out = [];
      let cur = "";
      let inQ = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQ && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQ = !inQ;
          }
          continue;
        }
        if (!inQ && ch === delim) {
          out.push(cur);
          cur = "";
          continue;
        }
        cur += ch;
      }
      out.push(cur);
      return out.map(s => s.trim());
    }

    const header = splitLine(nonEmpty[0], d).map(h => h.replace(/^\uFEFF/, "").trim());
    const rows = [];
    for (let i = 1; i < nonEmpty.length; i++) {
      const parts = splitLine(nonEmpty[i], d);
      const obj = {};
      for (let j = 0; j < header.length; j++) {
        obj[header[j]] = parts[j] !== undefined ? parts[j] : "";
      }
      rows.push(obj);
    }
    return { header, rows };
  }

  function parseDictionaryCsv(text) {
    // Accept common dictionary formats. Output: { columns: [{name,label,type,unit,notes}], mapByName: {...} }
    const d = detectDelimiter(text);
    const parsed = parseDelimitedTextToRows(text, d);
    const header = parsed.header.map(h => h.toLowerCase());
    const rows = parsed.rows || [];
    const idx = k => header.indexOf(k);

    const colNameIdx = idx("name") >= 0 ? idx("name") : (idx("variable") >= 0 ? idx("variable") : idx("field"));
    const labelIdx = idx("label") >= 0 ? idx("label") : (idx("description") >= 0 ? idx("description") : idx("question"));
    const typeIdx = idx("type");
    const unitIdx = idx("unit");
    const notesIdx = idx("notes") >= 0 ? idx("notes") : idx("comments");

    const columns = rows
      .map(r => {
        const keys = Object.keys(r);
        const getByIndex = i => (i >= 0 ? r[keys[i]] : "");
        const name = (colNameIdx >= 0 ? getByIndex(colNameIdx) : "").toString().trim();
        if (!name) return null;
        return {
          name,
          label: (labelIdx >= 0 ? getByIndex(labelIdx) : "").toString().trim(),
          type: (typeIdx >= 0 ? getByIndex(typeIdx) : "").toString().trim(),
          unit: (unitIdx >= 0 ? getByIndex(unitIdx) : "").toString().trim(),
          notes: (notesIdx >= 0 ? getByIndex(notesIdx) : "").toString().trim()
        };
      })
      .filter(Boolean);

    const mapByName = {};
    columns.forEach(c => (mapByName[c.name] = c));
    return { columns, mapByName };
  }

  function splitCombinedDictionaryAndData(text) {
    // Handles a combined TXT containing a dictionary section and then a data section.
    // Heuristic: find first line that looks like a data header (contains tabs/commas and at least 3 fields)
    // and that also contains a likely treatment/yield field.
    const raw = (text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const lines = raw.split("\n");

    function looksLikeHeader(line) {
      const d = detectDelimiter(line);
      const parts = line.split(d).map(s => s.trim()).filter(s => s.length);
      if (parts.length < 3) return false;
      const low = parts.map(p => p.toLowerCase());
      const hasTreatment = low.some(p => p.includes("treat") || p.includes("amend") || p === "amendment");
      const hasYield = low.some(p => p.includes("yield"));
      return hasTreatment && hasYield;
    }

    let headerIdx = -1;
    for (let i = 0; i < Math.min(lines.length, 3000); i++) {
      const line = lines[i];
      if (!line || !line.trim()) continue;
      if (looksLikeHeader(line)) {
        headerIdx = i;
        break;
      }
    }

    if (headerIdx <= 0) return { dictText: "", dataText: raw };
    const dictText = lines.slice(0, headerIdx).join("\n").trim();
    const dataText = lines.slice(headerIdx).join("\n").trim();
    return { dictText, dataText };
  }

  function normaliseKey(k) {
    return (k || "")
      .toString()
      .trim()
      .replace(/\u00A0/g, " ")
      .replace(/\s+/g, " ");
  }

  function detectTrialColumns(rows, dictionary) {
    const issues = [];
    const cols = rows && rows.length ? Object.keys(rows[0]).map(normaliseKey) : [];
    const colLower = cols.map(c => c.toLowerCase());

    const preferFromMap = (mapped) => {
      if (!mapped) return null;
      const exact = cols.find(c => c === mapped) || cols.find(c => c.toLowerCase() === mapped.toLowerCase());
      return exact || null;
    };

    const userMap = model.calibration.columnMap || {};
    const userTreatment = preferFromMap(userMap.treatment);
    const userReplicate = preferFromMap(userMap.replicate);
    const userPlot = preferFromMap(userMap.plot);
    const userYield = preferFromMap(userMap.yield);

    const pick = (cands) => {
      for (const cand of cands) {
        const idx = colLower.indexOf(cand);
        if (idx >= 0) return cols[idx];
      }
      return null;
    };

    const treatmentCol =
      userTreatment ||
      pick(["amendment", "treatment", "treat", "treatment_name", "amendment_name", "management", "option"]) ||
      cols.find(c => c.toLowerCase().includes("amend")) ||
      cols.find(c => c.toLowerCase().includes("treat"));

    const replicateCol =
      userReplicate ||
      pick(["replicate", "rep", "block", "trial_block", "replication", "repeat"]) ||
      cols.find(c => /rep/i.test(c) && !/report/i.test(c)) ||
      cols.find(c => /block/i.test(c));

    const plotCol =
      userPlot ||
      pick(["plot", "plot_id", "plotid", "plot number", "plot_number", "sub_plot", "subplot"]) ||
      cols.find(c => c.toLowerCase().includes("plot"));

    const yieldCol =
      userYield ||
      pick(["yield t/ha", "yield", "yield_t_ha", "grain yield", "grain_yield", "yield (t/ha)", "yield_t/ha"]) ||
      cols.find(c => c.toLowerCase().includes("yield"));

    // Identify likely cost columns
    const labourCols = cols.filter(c => /labour|labor/i.test(c));
    const operatingCols = cols.filter(c => /cost|input|chemical|fert|herb|fung|insect|diesel|fuel|mach/i.test(c));
    const otherCols = cols.filter(c => /service|contract|hire|freight|transport/i.test(c));

    if (!treatmentCol) issues.push({ code: "missing_treatment_col", severity: "error", message: "No treatment or amendment column detected." });
    if (!yieldCol) issues.push({ code: "missing_yield_col", severity: "error", message: "No yield column detected." });
    if (!replicateCol) issues.push({ code: "missing_replicate_col", severity: "warn", message: "No replicate or block column detected. Replicate-specific control baselines will not be available." });

    // Remove yield col from operating cost set if it matched
    const opNoYield = operatingCols.filter(c => c !== yieldCol);

    return {
      treatmentCol,
      replicateCol,
      plotCol,
      yieldCol,
      costCols: { labour: labourCols, operating: opNoYield, other: otherCols },
      issues
    };
  }

  function cleanTrialRows(rows, detected) {
    if (!rows || !rows.length) return [];
    const tCol = detected.treatmentCol;
    const rCol = detected.replicateCol;
    const pCol = detected.plotCol;
    const yCol = detected.yieldCol;

    const labourCols = detected.costCols.labour || [];
    const opCols = detected.costCols.operating || [];
    const otherCols = detected.costCols.other || [];

    return rows.map((row, idx) => {
      const obj = { __row: idx + 1 };

      // Raw values (strings)
      obj.treatment = tCol ? String(row[tCol] ?? "").trim() : "";
      obj.replicate = rCol ? String(row[rCol] ?? "").trim() : "";
      obj.plot = pCol ? String(row[pCol] ?? "").trim() : "";

      // Numerics
      obj.yield_t_ha = yCol ? parseNumber(row[yCol]) : NaN;

      let labour = 0;
      let labourAny = false;
      labourCols.forEach(c => {
        const v = parseNumber(row[c]);
        if (Number.isFinite(v)) {
          labour += v;
          labourAny = true;
        }
      });
      obj.labour_cost_per_ha = labourAny ? labour : NaN;

      let op = 0;
      let opAny = false;
      opCols.forEach(c => {
        const v = parseNumber(row[c]);
        if (Number.isFinite(v)) {
          op += v;
          opAny = true;
        }
      });
      obj.operating_cost_per_ha = opAny ? op : NaN;

      let oth = 0;
      let othAny = false;
      otherCols.forEach(c => {
        const v = parseNumber(row[c]);
        if (Number.isFinite(v)) {
          oth += v;
          othAny = true;
        }
      });
      obj.services_cost_per_ha = othAny ? oth : NaN;

      // Preserve original row for traceability
      obj.__raw = row;
      return obj;
    });
  }

  function identifyControlTreatments(treatName, hint) {
    const s = (treatName || "").toString().trim().toLowerCase();
    if (!s) return false;
    const h = (hint || "control").toString().trim().toLowerCase();
    if (h && s.includes(h)) return true;
    // common control markers
    if (s === "control" || s.includes("control")) return true;
    if (s.includes("baseline")) return true;
    if (s.includes("no amendment")) return true;
    if (s.includes("nil")) return true;
    return false;
  }

  function computeTrialMetrics(cleanedRows, detected) {
    const issues = [];
    const hint = model.calibration.controlNameHint || "control";

    if (!cleanedRows || !cleanedRows.length) {
      return {
        summary: { nRows: 0, nTreatments: 0, nReplicates: 0 },
        treatments: [],
        replicates: [],
        plots: [],
        controlByReplicate: new Map(),
        issues: [{ code: "no_rows", severity: "error", message: "No data rows available." }]
      };
    }

    const hasRep = !!detected.replicateCol;

    // Group by replicate then treatment
    const repSet = new Set();
    const treatSet = new Set();

    const repTreat = new Map(); // key: rep|treat => { yields:[], labour:[], op:[], svc:[], plots:[] }
    const treatAll = new Map(); // key: treat => arrays
    const controlByRep = new Map(); // rep => {controlTreatName, meanYield}

    const plots = [];
    cleanedRows.forEach(r => {
      const treat = r.treatment || "";
      const rep = hasRep ? (r.replicate || "") : "__single__";
      repSet.add(rep);
      treatSet.add(treat);

      const key = rep + "||" + treat;
      if (!repTreat.has(key)) repTreat.set(key, { rep, treat, yields: [], labour: [], op: [], svc: [], rows: [] });
      const g = repTreat.get(key);
      g.yields.push(r.yield_t_ha);
      g.labour.push(r.labour_cost_per_ha);
      g.op.push(r.operating_cost_per_ha);
      g.svc.push(r.services_cost_per_ha);
      g.rows.push(r);

      if (!treatAll.has(treat)) treatAll.set(treat, { treat, yields: [], labour: [], op: [], svc: [], reps: new Map() });
      const all = treatAll.get(treat);
      all.yields.push(r.yield_t_ha);
      all.labour.push(r.labour_cost_per_ha);
      all.op.push(r.operating_cost_per_ha);
      all.svc.push(r.services_cost_per_ha);
      if (!all.reps.has(rep)) all.reps.set(rep, { yields: [], labour: [], op: [], svc: [] });
      const rr = all.reps.get(rep);
      rr.yields.push(r.yield_t_ha);
      rr.labour.push(r.labour_cost_per_ha);
      rr.op.push(r.operating_cost_per_ha);
      rr.svc.push(r.services_cost_per_ha);

      plots.push(r);
    });

    // Determine control mean per replicate
    repSet.forEach(rep => {
      const repTreats = [];
      treatSet.forEach(treat => {
        const key = rep + "||" + treat;
        if (repTreat.has(key)) repTreats.push(repTreat.get(key));
      });

      const controls = repTreats.filter(g => identifyControlTreatments(g.treat, hint));
      if (controls.length === 0) {
        issues.push({
          code: "no_control_in_rep",
          severity: "warn",
          message: hasRep
            ? `No control detected in replicate "${rep}". Deltas vs control cannot be computed for this replicate.`
            : "No control detected. Deltas vs control cannot be computed."
        });
        return;
      }
      if (controls.length > 1) {
        issues.push({
          code: "multiple_controls_in_rep",
          severity: "warn",
          message: hasRep
            ? `Multiple control-like treatments detected in replicate "${rep}". The first control-like treatment will be used.`
            : "Multiple control-like treatments detected. The first control-like treatment will be used."
        });
      }
      const chosen = controls[0];
      const meanYield = safeMean(chosen.yields);
      if (!Number.isFinite(meanYield)) {
        issues.push({
          code: "control_yield_missing",
          severity: "warn",
          message: hasRep
            ? `Control yield is missing in replicate "${rep}". Deltas vs control cannot be computed for this replicate.`
            : "Control yield is missing. Deltas vs control cannot be computed."
        });
        return;
      }
      controlByRep.set(rep, { controlTreatName: chosen.treat, meanYield });
    });

    // Compute plot-level deltas where possible
    const plotDeltas = plots.map(p => {
      const rep = hasRep ? (p.replicate || "") : "__single__";
      const ctrl = controlByRep.get(rep);
      const deltaYield = ctrl && Number.isFinite(p.yield_t_ha) ? p.yield_t_ha - ctrl.meanYield : NaN;
      return { ...p, delta_yield_vs_control_t_ha: deltaYield, __control_mean_yield_t_ha: ctrl ? ctrl.meanYield : NaN };
    });

    // Per replicate per treatment stats
    const repTreatStats = [];
    repTreat.forEach(g => {
      const rep = g.rep;
      const treat = g.treat;
      const ctrl = controlByRep.get(rep);
      const meanYield = safeMean(g.yields);
      const meanLabour = safeMean(g.labour);
      const meanOp = safeMean(g.op);
      const meanSvc = safeMean(g.svc);
      const deltaYield = ctrl && Number.isFinite(meanYield) && Number.isFinite(ctrl.meanYield) ? meanYield - ctrl.meanYield : NaN;

      repTreatStats.push({
        replicate: rep,
        treatment: treat,
        nPlots: g.rows.length,
        meanYieldTHa: meanYield,
        stdYieldTHa: safeStd(g.yields),
        meanLabourCostPerHa: meanLabour,
        meanOperatingCostPerHa: meanOp,
        meanServicesCostPerHa: meanSvc,
        deltaYieldVsControlTHa: deltaYield,
        controlMeanYieldTHa: ctrl ? ctrl.meanYield : NaN,
        isControlLike: identifyControlTreatments(treat, hint)
      });
    });

    // Overall treatment stats using replicate-means when available (missing-safe)
    const treatmentStats = [];
    treatAll.forEach(all => {
      const treat = all.treat;
      const repMeansYield = [];
      const repMeansLab = [];
      const repMeansOp = [];
      const repMeansSvc = [];
      const repDeltas = [];

      all.reps.forEach((rr, rep) => {
        const my = safeMean(rr.yields);
        const ml = safeMean(rr.labour);
        const mo = safeMean(rr.op);
        const ms = safeMean(rr.svc);
        if (Number.isFinite(my)) repMeansYield.push(my);
        if (Number.isFinite(ml)) repMeansLab.push(ml);
        if (Number.isFinite(mo)) repMeansOp.push(mo);
        if (Number.isFinite(ms)) repMeansSvc.push(ms);

        const ctrl = controlByRep.get(rep);
        if (ctrl && Number.isFinite(my) && Number.isFinite(ctrl.meanYield)) repDeltas.push(my - ctrl.meanYield);
      });

      const overallMeanYield = safeMean(repMeansYield.length ? repMeansYield : all.yields);
      const overallStdYield = safeStd(all.yields);
      const meanLabour = safeMean(repMeansLab.length ? repMeansLab : all.labour);
      const meanOp = safeMean(repMeansOp.length ? repMeansOp : all.op);
      const meanSvc = safeMean(repMeansSvc.length ? repMeansSvc : all.svc);
      const meanDelta = safeMean(repDeltas);

      treatmentStats.push({
        treatment: treat,
        isControlLike: identifyControlTreatments(treat, hint),
        nRows: all.yields.length,
        nReplicates: all.reps.size,
        meanYieldTHa: overallMeanYield,
        stdYieldTHa: overallStdYield,
        meanLabourCostPerHa: meanLabour,
        meanOperatingCostPerHa: meanOp,
        meanServicesCostPerHa: meanSvc,
        meanDeltaYieldVsControlTHa: meanDelta,
        medianDeltaYieldVsControlTHa: safeMedian(repDeltas),
        stdDeltaYieldVsControlTHa: safeStd(repDeltas)
      });
    });

    // Additional checks
    const missingYield = plotDeltas.filter(p => !Number.isFinite(p.yield_t_ha)).length;
    const missingTreat = plotDeltas.filter(p => !p.treatment).length;
    const missingRep = hasRep ? plotDeltas.filter(p => !p.replicate).length : 0;
    const negYield = plotDeltas.filter(p => Number.isFinite(p.yield_t_ha) && p.yield_t_ha < 0).length;

    if (missingTreat > 0) issues.push({ code: "missing_treatment_values", severity: "error", message: `${missingTreat} row(s) are missing treatment labels.` });
    if (missingYield > 0) issues.push({ code: "missing_yield_values", severity: "warn", message: `${missingYield} row(s) have missing yield values.` });
    if (missingRep > 0) issues.push({ code: "missing_replicate_values", severity: "warn", message: `${missingRep} row(s) have missing replicate values.` });
    if (negYield > 0) issues.push({ code: "negative_yield_values", severity: "warn", message: `${negYield} row(s) have negative yield values.` });

    const controlCount = treatmentStats.filter(t => t.isControlLike).length;
    if (controlCount === 0) issues.push({ code: "no_control_detected", severity: "error", message: "No control-like treatment detected overall. Control comparisons will not be available." });
    if (controlCount > 1) issues.push({ code: "multiple_controls_detected", severity: "warn", message: "Multiple control-like treatments detected overall. The tool will select one control for comparison." });

    // Choose global control name: most common control-like treatment
    const ctrlNameCounts = new Map();
    treatmentStats.filter(t => t.isControlLike).forEach(t => {
      ctrlNameCounts.set(t.treatment, (ctrlNameCounts.get(t.treatment) || 0) + t.nRows);
    });
    let globalControlName = null;
    let best = -1;
    ctrlNameCounts.forEach((count, name) => {
      if (count > best) {
        best = count;
        globalControlName = name;
      }
    });

    return {
      summary: {
        nRows: plotDeltas.length,
        nTreatments: treatSet.size,
        nReplicates: repSet.size,
        hasReplicate: hasRep,
        globalControlName
      },
      treatments: treatmentStats.sort((a, b) => (a.isControlLike === b.isControlLike ? a.treatment.localeCompare(b.treatment) : a.isControlLike ? -1 : 1)),
      replicates: Array.from(repSet),
      repTreatStats,
      plots: plotDeltas,
      controlByReplicate: controlByRep,
      issues
    };
  }

  function buildValidationPanel(issues) {
    const out = [];
    const counts = { error: 0, warn: 0, info: 0 };
    (issues || []).forEach(i => {
      const sev = (i.severity || "info").toLowerCase();
      if (counts[sev] !== undefined) counts[sev] += 1;
      else counts.info += 1;
      out.push({ ...i, severity: sev });
    });
    const summary =
      counts.error > 0
        ? `${counts.error} error(s) and ${counts.warn} warning(s) detected.`
        : counts.warn > 0
          ? `${counts.warn} warning(s) detected.`
          : "No data issues detected.";
    return { items: out, summary };
  }

  function renderDataChecksPanel() {
    const root = $("#dataChecks") || $("#dataChecksList") || $("#dataChecksPanel");
    const listRoot = $("#dataChecksList") || root;
    if (!listRoot) return;

    listRoot.innerHTML = "";

    const pack = buildValidationPanel(trialState.validation.issues || []);
    trialState.validation.summary = pack.summary;

    const headerEl = $("#dataChecksSummary");
    if (headerEl) headerEl.textContent = pack.summary;

    if (!pack.items.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "No data checks to show.";
      listRoot.appendChild(p);
      return;
    }

    const ul = document.createElement("div");
    ul.className = "checks";
    pack.items.forEach(it => {
      const row = document.createElement("div");
      row.className = `check ${esc(it.severity)}`;
      row.innerHTML = `
        <div class="check-badge">${esc(it.severity.toUpperCase())}</div>
        <div class="check-body">
          <div class="check-code"><code>${esc(it.code)}</code></div>
          <div class="check-message">${esc(it.message)}</div>
        </div>
      `;
      ul.appendChild(row);
    });
    listRoot.appendChild(ul);
  }

  // Upload + Paste handlers: all are optional; they activate only if matching DOM exists.
  async function handleImportText(text, opts = {}) {
    const combined = splitCombinedDictionaryAndData(text);
    let dict = null;

    // Try dictionary parsing if combined section looks like delimited content
    const dictText = combined.dictText || "";
    if (dictText.trim().length > 0) {
      try {
        dict = parseDictionaryCsv(dictText);
      } catch (e) {
        dict = null;
      }
    }

    const dataText = combined.dataText || text;
    const delimiter = opts.delimiter || detectDelimiter(dataText);
    const parsed = parseDelimitedTextToRows(dataText, delimiter);

    const rows = parsed.rows || [];
    trialState.source = opts.source || "paste";
    trialState.importedAt = new Date().toISOString();
    trialState.dictionary = dict;
    trialState.dataRows = rows;

    const detected = detectTrialColumns(rows, dict);
    trialState.detected = detected;

    const cleaned = cleanTrialRows(rows, detected);
    trialState.cleaned = cleaned;

    const metrics = computeTrialMetrics(cleaned, detected);
    trialState.metrics = metrics;

    const combinedIssues = [...(detected.issues || []), ...(metrics.issues || [])];
    trialState.validation.issues = combinedIssues;

    // Prefer trial mode if it is usable
    const hasCore = !!detected.treatmentCol && !!detected.yieldCol && cleaned.length > 0;
    if (hasCore) model.calibration.mode = "trial";

    // Persist last import snapshot (small)
    try {
      const snapshot = {
        importedAt: trialState.importedAt,
        source: trialState.source,
        nRows: metrics.summary.nRows,
        nTreatments: metrics.summary.nTreatments,
        hasReplicate: metrics.summary.hasReplicate,
        detected: trialState.detected,
        controlNameHint: model.calibration.controlNameHint,
        columnMap: model.calibration.columnMap
      };
      localStorage.setItem(STORAGE_KEYS.lastImport, JSON.stringify(snapshot));
    } catch (e) {
      // ignore
    }

    // Render optional import status fields
    const statusEl = $("#importStatus") || $("#dataImportStatus");
    if (statusEl) {
      statusEl.textContent = `Imported ${metrics.summary.nRows.toLocaleString()} row(s) with ${metrics.summary.nTreatments.toLocaleString()} treatment(s).`;
    }

    renderDataChecksPanel();
    applyTrialCalibrationToModel();
    renderAll();
    calcAndRender();
    showToast("Data import complete. Trial calibration updated.");
  }

  async function handleUploadFile(file) {
    if (!file) return;
    const name = (file.name || "").toLowerCase();
    const ext = name.split(".").pop();

    // Excel supported if SheetJS exists
    if ((ext === "xlsx" || ext === "xlsm" || ext === "xlsb" || ext === "xls") && typeof XLSX !== "undefined") {
      try {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array" });

        // Try to find a raw data sheet
        const sheetName = wb.SheetNames.find(n => FABABEAN_SHEET_NAMES.includes(n)) || wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        // Convert to delimited text (tab) for a single pipeline
        const cols = rows.length ? Object.keys(rows[0]) : [];
        const lines = [];
        lines.push(cols.join("\t"));
        rows.forEach(r => {
          lines.push(cols.map(c => (r[c] == null ? "" : String(r[c]))).join("\t"));
        });
        const text = lines.join("\n");

        await handleImportText(text, { source: `upload:${file.name}`, delimiter: "\t" });
        showToast(`Excel uploaded and parsed from sheet "${sheetName}".`);
        return;
      } catch (err) {
        console.error(err);
        alert("Error parsing Excel file.");
        return;
      }
    }

    // Plain text / CSV / TSV
    try {
      const text = await file.text();
      await handleImportText(text, { source: `upload:${file.name}` });
      showToast("File uploaded and parsed.");
    } catch (e) {
      console.error(e);
      alert("Unable to read the uploaded file.");
    }
  }

  function initImportPipelineBindings() {
    // Optional upload input
    const fileInput = $("#dataUpload") || $("#uploadData") || $("#importFile") || $("#trialDataUpload");
    if (fileInput && fileInput.tagName === "INPUT" && fileInput.type === "file") {
      fileInput.addEventListener("change", async e => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        await handleUploadFile(f);
        e.target.value = "";
      });
    }

    // Optional upload button that triggers hidden input
    const uploadBtn = $("#uploadBtn") || $("#importUploadBtn") || $("#chooseFileBtn");
    if (uploadBtn && fileInput && fileInput.tagName === "INPUT" && fileInput.type === "file") {
      uploadBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        fileInput.click();
      });
    }

    // Optional paste parse + commit
    const pasteBox = $("#dataPaste") || $("#pasteData") || $("#trialPaste") || $("#pasteBox");
    const parsePasteBtn = $("#parsePaste") || $("#parsePastedData");
    const commitPasteBtn = $("#commitPaste") || $("#importPaste") || $("#commitPastedData");

    let cachedPasteText = null;

    if (parsePasteBtn && pasteBox) {
      parsePasteBtn.addEventListener("click", async e => {
        e.preventDefault();
        e.stopPropagation();
        cachedPasteText = String(pasteBox.value || "");
        if (!cachedPasteText.trim()) {
          alert("Paste data first.");
          return;
        }
        await handleImportText(cachedPasteText, { source: "paste" });
        showToast("Pasted data parsed and validated.");
      });
    }

    if (commitPasteBtn && pasteBox) {
      commitPasteBtn.addEventListener("click", async e => {
        e.preventDefault();
        e.stopPropagation();
        const text = cachedPasteText != null ? cachedPasteText : String(pasteBox.value || "");
        if (!text.trim()) {
          alert("Paste data first.");
          return;
        }
        await handleImportText(text, { source: "paste" });
        showToast("Pasted data committed.");
      });
    }

    // Optional dictionary upload + parse
    const dictInput = $("#dictUpload") || $("#dataDictionaryUpload") || $("#dictionaryFile");
    if (dictInput && dictInput.tagName === "INPUT" && dictInput.type === "file") {
      dictInput.addEventListener("change", async e => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
          const text = await f.text();
          const dict = parseDictionaryCsv(text);
          trialState.dictionary = dict;
          showToast("Data dictionary parsed and stored.");
          e.target.value = "";
        } catch (err) {
          console.error(err);
          alert("Unable to parse the dictionary file. Please provide a CSV with column metadata.");
          e.target.value = "";
        }
      });
    }

    // Optional hint/column mapping inputs
    const hintInput = $("#controlHint") || $("#controlNameHint");
    if (hintInput) {
      hintInput.addEventListener("input", e => {
        model.calibration.controlNameHint = String(e.target.value || "control");
      });
    }

    const mapTreatment = $("#mapTreatmentCol");
    const mapReplicate = $("#mapReplicateCol");
    const mapPlot = $("#mapPlotCol");
    const mapYield = $("#mapYieldCol");
    const applyMapBtn = $("#applyColumnMap");

    if (applyMapBtn) {
      applyMapBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        model.calibration.columnMap = {
          treatment: mapTreatment ? mapTreatment.value : model.calibration.columnMap.treatment,
          replicate: mapReplicate ? mapReplicate.value : model.calibration.columnMap.replicate,
          plot: mapPlot ? mapPlot.value : model.calibration.columnMap.plot,
          yield: mapYield ? mapYield.value : model.calibration.columnMap.yield
        };
        if (trialState.dataRows && trialState.dataRows.length) {
          // Re-run detection + metrics on the same imported data
          const detected = detectTrialColumns(trialState.dataRows, trialState.dictionary);
          trialState.detected = detected;
          const cleaned = cleanTrialRows(trialState.dataRows, detected);
          trialState.cleaned = cleaned;
          const metrics = computeTrialMetrics(cleaned, detected);
          trialState.metrics = metrics;
          const combinedIssues = [...(detected.issues || []), ...(metrics.issues || [])];
          trialState.validation.issues = combinedIssues;
          renderDataChecksPanel();
          applyTrialCalibrationToModel();
          renderAll();
          calcAndRender();
          showToast("Column mapping applied and trial calibration refreshed.");
        } else {
          showToast("Column mapping saved.");
        }
      });
    }
  }

  // =========================
  // 5) APPLY TRIAL CALIBRATION INTO MODEL (TREATMENTS, COSTS, DELTAS)
  // =========================
  function getYieldOutput() {
    return model.outputs.find(o => o.name.toLowerCase().includes("yield")) || model.outputs[0];
  }

  function applyTrialCalibrationToModel() {
    if (model.calibration.mode !== "trial") return;
    const metrics = trialState.metrics;
    if (!metrics || !metrics.treatments || !metrics.treatments.length) return;

    const yOut = getYieldOutput();
    const yId = yOut ? yOut.id : null;

    // Choose a single control treatment name
    const ctrlName = metrics.summary.globalControlName;
    if (!ctrlName) return;

    // Build treatments from trial summaries
    const nextTreatments = metrics.treatments.map(tt => {
      const isCtrl = tt.treatment === ctrlName;
      const mats = Number.isFinite(tt.meanOperatingCostPerHa) ? tt.meanOperatingCostPerHa : 0;
      const lab = Number.isFinite(tt.meanLabourCostPerHa) ? tt.meanLabourCostPerHa : 0;
      const svc = Number.isFinite(tt.meanServicesCostPerHa) ? tt.meanServicesCostPerHa : 0;

      // Delta yield vs control (replicate-mean based where possible)
      const dy = Number.isFinite(tt.meanDeltaYieldVsControlTHa) ? tt.meanDeltaYieldVsControlTHa : (isCtrl ? 0 : 0);

      const t = {
        id: uid(),
        name: tt.treatment,
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: lab,
        materialsCost: mats,
        servicesCost: svc,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: isCtrl,
        notes: isCtrl ? "Control baseline derived from trial data." : "Calibrated from imported trial plot data.",
        recurrence: {
          cost: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" },
          benefit: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" }
        }
      };

      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      if (yId) t.deltas[yId] = isCtrl ? 0 : dy;

      return t;
    });

    // Ensure only one control in the model
    let foundControl = false;
    nextTreatments.forEach(t => {
      if (t.isControl) {
        if (!foundControl) foundControl = true;
        else t.isControl = false;
      }
    });

    model.treatments = nextTreatments;
    initTreatmentDeltasAndRecurrence();
  }

  // =========================
  // 6) CASHFLOWS + DISCOUNTED CBA ENGINE (FULL)
  // =========================
  function recurrenceYearsFromRule(rule, N) {
    const mode = (rule && rule.mode ? rule.mode : "annual").toLowerCase();
    const everyN = Math.max(1, parseInt(rule && rule.everyN != null ? rule.everyN : 1, 10) || 1);
    const start = Math.max(1, parseInt(rule && rule.startYearOffset != null ? rule.startYearOffset : 1, 10) || 1);
    const endOffset = rule && rule.endYearOffset != null && rule.endYearOffset !== "" ? parseInt(rule.endYearOffset, 10) : null;
    const end = endOffset == null || !Number.isFinite(endOffset) ? N : clamp(endOffset, start, N);
    const yearsCsv = (rule && rule.yearsCsv ? String(rule.yearsCsv) : "").trim();

    const set = new Set();
    if (mode === "once") {
      set.add(start);
      return Array.from(set).filter(y => y >= 0 && y <= N).sort((a, b) => a - b);
    }

    if (mode === "custom") {
      yearsCsv
        .split(/[,\s]+/g)
        .map(s => parseInt(s, 10))
        .filter(n => Number.isFinite(n))
        .forEach(n => {
          if (n >= 0 && n <= N) set.add(n);
        });
      return Array.from(set).sort((a, b) => a - b);
    }

    // Annual or every_n
    for (let y = start; y <= end; y++) {
      if (mode === "annual") set.add(y);
      else if (mode === "every_n_years") {
        if ((y - start) % everyN === 0) set.add(y);
      } else {
        // default annual
        set.add(y);
      }
    }
    return Array.from(set).sort((a, b) => a - b);
  }

  function additionalBenefitsSeries(N, baseYear, adoptMul, risk) {
    const series = new Array(N + 1).fill(0);
    model.benefits.forEach(b => {
      const cat = String(b.category || "").toUpperCase();
      const linkA = !!b.linkAdoption;
      const linkR = !!b.linkRisk;
      const A = linkA ? clamp(adoptMul, 0, 1) : 1;
      const R = linkR ? 1 - clamp(risk, 0, 1) : 1;
      const g = Number(b.growthPct) || 0;

      const addAnnual = (yearIndex, baseAmount, tFromStart) => {
        const grown = baseAmount * Math.pow(1 + g / 100, tFromStart);
        if (yearIndex >= 1 && yearIndex <= N) series[yearIndex] += grown * A * R;
      };
      const addOnce = (absYear, amount) => {
        const idx = absYear - baseYear + 1;
        if (idx >= 0 && idx <= N) series[idx] += amount * A * R;
      };

      const sy = Number(b.startYear) || baseYear;
      const ey = Number(b.endYear) || sy;
      const yr = Number(b.year) || sy;

      if (b.frequency === "Once" || cat === "C6") {
        const amount = Number(b.annualAmount) || 0;
        addOnce(yr, amount);
        return;
      }

      for (let y = sy; y <= ey; y++) {
        const idx = y - baseYear + 1;
        const tFromStart = y - sy;
        let amt = 0;
        switch (cat) {
          case "C1":
          case "C2":
          case "C3": {
            const v = Number(b.unitValue) || 0;
            const q = Number(cat === "C3" ? b.abatement : b.quantity) || 0;
            amt = v * q;
            break;
          }
          case "C4":
          case "C5":
          case "C8":
            amt = Number(b.annualAmount) || 0;
            break;
          case "C7": {
            const p0 = Number(b.p0) || 0;
            const p1 = Number(b.p1) || 0;
            const c = Number(b.consequence) || 0;
            amt = Math.max(p0 - p1, 0) * c;
            break;
          }
          default:
            amt = 0;
        }
        addAnnual(idx, amt, tFromStart);
      }
    });
    return series;
  }

  // Cost scaling rule (implemented explicitly):
  // - Capital cost is year 0, not discounted within the year 0 period; it is included in PV costs as-is.
  // - Operating costs are applied in the years specified by the treatment's cost recurrence rule.
  // - Benefits are applied in the years specified by the treatment's benefit recurrence rule, capped by persistenceYears.
  // - Benefits and costs are scaled by: areaHa * adoptionMultiplier * treatmentSpecificAdoption.
  // - Risk reduces benefits only: benefit * (1 - risk).
  // - Missing numeric values are treated as missing and excluded; they never become zero by accident.
  function buildTreatmentCashflows(t, opts) {
    const rate = opts.ratePct;
    const N = opts.years;
    const adoptMul = clamp(opts.adoptMul, 0, 1);
    const risk = clamp(opts.risk, 0, 1);
    const priceMultiplier = Number.isFinite(opts.priceMultiplier) ? opts.priceMultiplier : 1.0;
    const persistenceYears = Number.isFinite(opts.persistenceYears) ? Math.max(0, Math.floor(opts.persistenceYears)) : N;
    const recurrenceMultiplier = Number.isFinite(opts.recurrenceMultiplier) ? opts.recurrenceMultiplier : 1.0;

    const effectiveAdoption = adoptMul * clamp(Number.isFinite(t.adoption) ? t.adoption : 1, 0, 1);
    const area = Number.isFinite(+t.area) ? +t.area : 0;
    const scale = area * effectiveAdoption;

    // Benefit per ha: sum of output deltas * value
    let valuePerHa = 0;
    model.outputs.forEach(o => {
      const delta = Number.isFinite(parseFloat(t.deltas[o.id])) ? +t.deltas[o.id] : 0;
      let v = Number.isFinite(+o.value) ? +o.value : 0;
      // price multiplier applies to yield value if output is yield-like, else keep as given
      if (o.name && o.name.toLowerCase().includes("yield")) v = v * priceMultiplier;
      valuePerHa += delta * v;
    });

    const annualBenefit = valuePerHa * scale * (1 - risk);

    // Costs per ha
    const mats = Number.isFinite(+t.materialsCost) ? +t.materialsCost : NaN;
    const serv = Number.isFinite(+t.servicesCost) ? +t.servicesCost : NaN;
    const lab = Number.isFinite(+t.labourCost) ? +t.labourCost : NaN;

    const perHaOp = safeSum([mats, serv, lab]); // missing-safe sum (NaN excluded)
    const annualOpCost = perHaOp * scale * recurrenceMultiplier;

    const cap = Number.isFinite(+t.capitalCost) ? +t.capitalCost : 0;
    const capScaled = cap * effectiveAdoption; // capital scales with adoption, not with area by default (cap is a lump sum). If capital is per-ha, users should encode it as op cost.

    const cf = new Array(N + 1).fill(0);
    const ben = new Array(N + 1).fill(0);
    const cost = new Array(N + 1).fill(0);

    // Year 0 capital
    if (capScaled !== 0) {
      cf[0] -= capScaled;
      cost[0] += capScaled;
    }

    // Recurrence years
    const costYears = recurrenceYearsFromRule(t.recurrence && t.recurrence.cost, N);
    const benefitYears = recurrenceYearsFromRule(t.recurrence && t.recurrence.benefit, N);

    // Benefits only accrue within persistence window (year indices 1..persistenceYears)
    const benefitYearsFiltered = benefitYears.filter(y => y >= 1 && y <= Math.min(N, persistenceYears));

    benefitYearsFiltered.forEach(y => {
      ben[y] += annualBenefit;
      cf[y] += annualBenefit;
    });

    costYears.forEach(y => {
      if (y >= 1 && y <= N) {
        cost[y] += annualOpCost;
        cf[y] -= annualOpCost;
      }
    });

    return { cf, benefitByYear: ben, costByYear: cost, annualBenefit, annualOpCost, capScaled, scale };
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) {
      pv += series[t] / Math.pow(1 + ratePct / 100, t);
    }
    return pv;
  }

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

  function additionalProjectCostsSeries(N, baseYear) {
    const costByYear = new Array(N + 1).fill(0);
    const constrainedByYear = new Array(N + 1).fill(0);

    let capY0 = 0;
    let capY0Con = 0;

    model.otherCosts.forEach(c => {
      if (c.type === "annual") {
        const a = Number.isFinite(+c.annual) ? +c.annual : 0;
        const sy = Number.isFinite(+c.startYear) ? +c.startYear : baseYear;
        const ey = Number.isFinite(+c.endYear) ? +c.endYear : sy;
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) {
            costByYear[idx] += a;
            if (c.constrained) constrainedByYear[idx] += a;
          }
        }
      } else if (c.type === "capital") {
        const cap = Number.isFinite(+c.capital) ? +c.capital : 0;
        const cy = Number.isFinite(+c.year) ? +c.year : baseYear;
        const idx = cy - baseYear;
        if (idx === 0) {
          capY0 += cap;
          if (c.constrained) capY0Con += cap;
        } else if (idx > 0 && idx <= N) {
          costByYear[idx] += cap;
          if (c.constrained) constrainedByYear[idx] += cap;
        }
      }
    });

    costByYear[0] += capY0;
    constrainedByYear[0] += capY0Con;

    return { costByYear, constrainedByYear };
  }

  function computeTreatmentCBA(t, opts) {
    const years = opts.years;
    const rate = opts.ratePct;

    const flows = buildTreatmentCashflows(t, opts);
    const pvBen = presentValue(flows.benefitByYear, rate);
    const pvCost = presentValue(flows.costByYear, rate);
    const npv = pvBen - pvCost;
    const bcr = pvCost > 0 ? pvBen / pvCost : NaN;
    const roi = pvCost > 0 ? (npv / pvCost) * 100 : NaN;
    const irrVal = irr(flows.cf);
    const mirrVal = mirr(flows.cf, model.time.mirrFinance, model.time.mirrReinvest);
    const pb = payback(flows.cf, rate);

    return {
      pvBenefits: pvBen,
      pvCosts: pvCost,
      npv,
      bcr,
      roiPct: roi,
      irrPct: irrVal,
      mirrPct: mirrVal,
      paybackYears: pb,
      annualBenefit: flows.annualBenefit,
      annualOpCost: flows.annualOpCost,
      capitalCostY0: flows.capScaled,
      scale: flows.scale,
      cf: flows.cf,
      benefitByYear: flows.benefitByYear,
      costByYear: flows.costByYear
    };
  }

  function computeProjectCBA(opts) {
    const N = opts.years;
    const baseYear = model.time.startYear;

    const projectBenefit = new Array(N + 1).fill(0);
    const projectCost = new Array(N + 1).fill(0);
    const projectCostCon = new Array(N + 1).fill(0);

    // Treatments (all included in project totals)
    model.treatments.forEach(t => {
      const f = buildTreatmentCashflows(t, opts);
      for (let i = 0; i <= N; i++) {
        projectBenefit[i] += f.benefitByYear[i];
        projectCost[i] += f.costByYear[i];
        if (t.constrained) projectCostCon[i] += f.costByYear[i];
      }
    });

    // Other project costs
    const other = additionalProjectCostsSeries(N, baseYear);
    for (let i = 0; i <= N; i++) {
      projectCost[i] += other.costByYear[i];
      projectCostCon[i] += other.constrainedByYear[i];
    }

    // Additional benefits
    const extraBen = additionalBenefitsSeries(N, baseYear, opts.adoptMul, opts.risk);
    for (let i = 0; i <= N; i++) projectBenefit[i] += extraBen[i];

    const cf = projectBenefit.map((b, i) => b - projectCost[i]);
    const pvBenefits = presentValue(projectBenefit, opts.ratePct);
    const pvCosts = presentValue(projectCost, opts.ratePct);
    const pvCostsConstrained = presentValue(projectCostCon, opts.ratePct);
    const npv = pvBenefits - pvCosts;
    const denom = opts.bcrMode === "constrained" ? pvCostsConstrained : pvCosts;
    const bcr = denom > 0 ? pvBenefits / denom : NaN;

    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;

    // annual gross margin (year 1)
    const annualGM = projectBenefit[1] - (projectCost[1] - other.costByYear[1]); // exclude other project costs from GM
    const profitMargin = projectBenefit[1] > 0 ? (annualGM / projectBenefit[1]) * 100 : NaN;
    const pb = payback(cf, opts.ratePct);

    return {
      pvBenefits,
      pvCosts,
      pvCostsConstrained,
      npv,
      bcr,
      irrPct: irrVal,
      mirrPct: mirrVal,
      roiPct: roi,
      annualGM,
      profitMarginPct: profitMargin,
      paybackYears: pb,
      cf,
      benefitByYear: projectBenefit,
      costByYear: projectCost
    };
  }

  // =========================
  // 7) RESULTS: COMPARISON TO CONTROL GRID + FILTERS + NARRATIVE
  // =========================
  const resultsState = {
    filter: "all" // all | top_npv | top_bcr | improve_only
  };

  function getControlTreatment() {
    const ctrl = model.treatments.find(t => t.isControl);
    if (ctrl) return ctrl;
    // fallback: choose first control-like name
    const hint = model.calibration.controlNameHint || "control";
    const guess = model.treatments.find(t => identifyControlTreatments(t.name, hint));
    return guess || null;
  }

  function buildPerTreatmentResults(opts) {
    const ctrl = getControlTreatment();
    const controlMetrics = ctrl ? computeTreatmentCBA(ctrl, opts) : null;

    const rows = model.treatments.map(t => {
      const m = computeTreatmentCBA(t, opts);
      const d = controlMetrics
        ? {
            pvBenefits: m.pvBenefits - controlMetrics.pvBenefits,
            pvCosts: m.pvCosts - controlMetrics.pvCosts,
            npv: m.npv - controlMetrics.npv,
            bcr: (Number.isFinite(m.bcr) && Number.isFinite(controlMetrics.bcr)) ? (m.bcr - controlMetrics.bcr) : NaN,
            roiPct: (Number.isFinite(m.roiPct) && Number.isFinite(controlMetrics.roiPct)) ? (m.roiPct - controlMetrics.roiPct) : NaN
          }
        : null;

      const deltaPct = controlMetrics && Number.isFinite(controlMetrics.npv) && controlMetrics.npv !== 0
        ? (d ? (d.npv / controlMetrics.npv) * 100 : NaN)
        : NaN;

      return { t, m, controlMetrics, delta: d, deltaNpvPct: deltaPct, isControl: !!t.isControl };
    });

    // Ranking: by NPV desc (excluding control)
    const ranked = rows
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (Number.isFinite(b.m.npv) ? b.m.npv : -Infinity) - (Number.isFinite(a.m.npv) ? a.m.npv : -Infinity))
      .map((r, idx) => ({ ...r, rank: idx + 1 }));

    // Control gets rank 0
    rows.forEach(r => {
      const hit = ranked.find(x => x.t.id === r.t.id);
      r.rank = r.isControl ? 0 : (hit ? hit.rank : null);
    });

    return { control: ctrl, controlMetrics, rows };
  }

  function formatDeltaCell(value, isMoney, isPct) {
    if (!Number.isFinite(value)) return `<span class="muted">n/a</span>`;
    let txt = "";
    if (isMoney) txt = money(value);
    else if (isPct) txt = percent(value);
    else txt = fmt(value);

    const cls = value > 0 ? "pos" : value < 0 ? "neg" : "neu";
    return `<span class="delta ${cls}">${esc(txt)}</span>`;
  }

  function renderLeaderboard(perTreatment) {
    const root = $("#leaderboard") || $("#resultsLeaderboard") || $("#leaderboardList");
    if (!root) return;

    const rows = perTreatment.rows.filter(r => !r.isControl);
    let filtered = rows;

    if (resultsState.filter === "top_npv") {
      filtered = rows.slice().sort((a, b) => (b.m.npv || -Infinity) - (a.m.npv || -Infinity)).slice(0, 5);
    } else if (resultsState.filter === "top_bcr") {
      filtered = rows.slice().sort((a, b) => (Number.isFinite(b.m.bcr) ? b.m.bcr : -Infinity) - (Number.isFinite(a.m.bcr) ? a.m.bcr : -Infinity)).slice(0, 5);
    } else if (resultsState.filter === "improve_only") {
      filtered = rows.filter(r => Number.isFinite(r.delta?.npv) && r.delta.npv > 0);
    }

    filtered = filtered.slice().sort((a, b) => (a.rank || Infinity) - (b.rank || Infinity));

    root.innerHTML = "";
    if (!filtered.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "No treatments to display for the current filter.";
      root.appendChild(p);
      return;
    }

    const wrap = document.createElement("div");
    wrap.className = "leaderboard";
    filtered.forEach(r => {
      const card = document.createElement("div");
      card.className = "lb-row";
      card.innerHTML = `
        <div class="lb-rank">${r.rank != null ? r.rank : ""}</div>
        <div class="lb-name">${esc(r.t.name)}</div>
        <div class="lb-metric">${money(r.m.npv)}</div>
        <div class="lb-metric">${Number.isFinite(r.m.bcr) ? fmt(r.m.bcr) : "n/a"}</div>
      `;
      wrap.appendChild(card);
    });
    root.appendChild(wrap);
  }

  function renderComparisonToControlGrid(perTreatment) {
    const root =
      $("#comparisonToControl") ||
      $("#comparisonToControlTable") ||
      $("#comparisonTable") ||
      $("#comparisonGrid") ||
      $("#resultsComparison");
    if (!root) return;

    const ctrl = perTreatment.control;
    const controlMetrics = perTreatment.controlMetrics;

    root.innerHTML = "";

    if (!ctrl || !controlMetrics) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "Select a control treatment to view comparison-to-control results.";
      root.appendChild(p);
      return;
    }

    // Filter columns
    const allRows = perTreatment.rows.slice();
    let treatments = allRows.slice();

    const nonControl = treatments.filter(r => !r.isControl);

    if (resultsState.filter === "top_npv") {
      treatments = [allRows.find(r => r.isControl)]
        .concat(nonControl.slice().sort((a, b) => (b.m.npv || -Infinity) - (a.m.npv || -Infinity)).slice(0, 5));
    } else if (resultsState.filter === "top_bcr") {
      treatments = [allRows.find(r => r.isControl)]
        .concat(nonControl.slice().sort((a, b) => (Number.isFinite(b.m.bcr) ? b.m.bcr : -Infinity) - (Number.isFinite(a.m.bcr) ? a.m.bcr : -Infinity)).slice(0, 5));
    } else if (resultsState.filter === "improve_only") {
      treatments = [allRows.find(r => r.isControl)]
        .concat(nonControl.filter(r => Number.isFinite(r.delta?.npv) && r.delta.npv > 0));
    } else {
      // all
      treatments = treatments.slice().sort((a, b) => (a.isControl ? -1 : 1) - (b.isControl ? -1 : 1) || (a.rank || 9999) - (b.rank || 9999));
    }

    const indicators = [
      { key: "pvBenefits", label: "PV Benefits", fmt: (v) => money(v), isMoney: true },
      { key: "pvCosts", label: "PV Costs", fmt: (v) => money(v), isMoney: true },
      { key: "npv", label: "NPV", fmt: (v) => money(v), isMoney: true },
      { key: "bcr", label: "BCR", fmt: (v) => (Number.isFinite(v) ? fmt(v) : "n/a"), isMoney: false },
      { key: "roiPct", label: "ROI", fmt: (v) => (Number.isFinite(v) ? percent(v) : "n/a"), isMoney: false },
      { key: "rank", label: "Rank", fmt: (v) => (v != null ? String(v) : ""), isMoney: false },
      { key: "delta_npv", label: "Δ NPV vs Control", fmt: (v) => money(v), isMoney: true, isDelta: true },
      { key: "delta_npv_pct", label: "Δ NPV vs Control (%)", fmt: (v) => (Number.isFinite(v) ? percent(v) : "n/a"), isMoney: false, isDelta: true, isPct: true },
      { key: "delta_pvcost", label: "Δ PV Cost vs Control", fmt: (v) => money(v), isMoney: true, isDelta: true }
    ];

    const table = document.createElement("div");
    table.className = "comparison-grid";

    // Build header row
    const header = document.createElement("div");
    header.className = "cg-row cg-header";
    header.innerHTML = `<div class="cg-cell cg-sticky cg-indicator">Indicator</div>` +
      treatments
        .map(r => {
          const name = r.isControl ? "Control (baseline)" : r.t.name;
          const cls = r.isControl ? "cg-treatment cg-control" : "cg-treatment";
          return `<div class="cg-cell cg-top ${cls}">${esc(name)}</div>`;
        })
        .join("") +
      // Add delta columns per treatment (absolute and percent handled in indicator rows, not extra columns)
      "";
    table.appendChild(header);

    function cellClassForDelta(val) {
      if (!Number.isFinite(val)) return "cg-cell";
      if (val > 0) return "cg-cell cg-pos";
      if (val < 0) return "cg-cell cg-neg";
      return "cg-cell";
    }

    // Body rows
    indicators.forEach(ind => {
      const row = document.createElement("div");
      row.className = "cg-row";
      row.innerHTML = `<div class="cg-cell cg-sticky cg-indicator">${esc(ind.label)}</div>` +
        treatments
          .map(r => {
            if (r.isControl) {
              if (ind.key === "rank") return `<div class="cg-cell cg-control">${esc("Baseline")}</div>`;
              if (ind.key.startsWith("delta_")) return `<div class="cg-cell cg-control muted">n/a</div>`;
              const v = ind.key === "roiPct" ? r.m.roiPct : ind.key === "bcr" ? r.m.bcr : ind.key === "pvBenefits" ? r.m.pvBenefits : ind.key === "pvCosts" ? r.m.pvCosts : ind.key === "npv" ? r.m.npv : (ind.key === "rank" ? 0 : NaN);
              return `<div class="cg-cell cg-control">${esc(ind.fmt(v))}</div>`;
            }

            let v = NaN;
            if (ind.key === "pvBenefits") v = r.m.pvBenefits;
            else if (ind.key === "pvCosts") v = r.m.pvCosts;
            else if (ind.key === "npv") v = r.m.npv;
            else if (ind.key === "bcr") v = r.m.bcr;
            else if (ind.key === "roiPct") v = r.m.roiPct;
            else if (ind.key === "rank") v = r.rank;
            else if (ind.key === "delta_npv") v = r.delta ? r.delta.npv : NaN;
            else if (ind.key === "delta_npv_pct") v = r.deltaNpvPct;
            else if (ind.key === "delta_pvcost") v = r.delta ? r.delta.pvCosts : NaN;

            const cls = ind.key.startsWith("delta_") ? cellClassForDelta(v) : "cg-cell";
            const txt = ind.key === "rank" ? (v != null ? String(v) : "") : ind.fmt(v);
            return `<div class="${cls}">${esc(txt)}</div>`;
          })
          .join("");

      table.appendChild(row);
    });

    root.appendChild(table);
  }

  function renderWhatThisMeans(perTreatment, opts) {
    const root = $("#whatThisMeans") || $("#resultsNarrative") || $("#whatThisMeansText");
    if (!root) return;

    const ctrl = perTreatment.control;
    const controlMetrics = perTreatment.controlMetrics;
    if (!ctrl || !controlMetrics) {
      root.textContent = "Select a control treatment to generate an interpretation.";
      return;
    }

    const nonControl = perTreatment.rows.filter(r => !r.isControl);
    const improving = nonControl.filter(r => Number.isFinite(r.delta?.npv) && r.delta.npv > 0);
    const worsening = nonControl.filter(r => Number.isFinite(r.delta?.npv) && r.delta.npv < 0);

    const bestNpv = nonControl.slice().sort((a, b) => (b.m.npv || -Infinity) - (a.m.npv || -Infinity))[0] || null;
    const bestBcr = nonControl.slice().sort((a, b) => (Number.isFinite(b.m.bcr) ? b.m.bcr : -Infinity) - (Number.isFinite(a.m.bcr) ? a.m.bcr : -Infinity))[0] || null;

    const rate = opts.ratePct;
    const years = opts.years;
    const price = opts.pricePerTonne;

    const ctrlAnnual = controlMetrics.annualBenefit - controlMetrics.annualOpCost;

    // No bullets, no em dash, no abbreviations
    const parts = [];
    parts.push(
      `This view compares each treatment to the control baseline, using a ${years} year horizon and a discount rate of ${fmt(rate)} percent per year. The grain price used for yield benefits is ${money(price)} per tonne.`
    );
    parts.push(
      `The control baseline has present value benefits of ${money(controlMetrics.pvBenefits)} and present value costs of ${money(controlMetrics.pvCosts)}, giving a net present value of ${money(controlMetrics.npv)}.`
    );

    if (bestNpv) {
      const d = bestNpv.delta;
      parts.push(
        `The strongest net present value relative to the control in the base case is ${esc(bestNpv.t.name)}, with a net present value of ${money(bestNpv.m.npv)}. Relative to the control, its change in net present value is ${money(d ? d.npv : NaN)}.`
      );
    }
    if (bestBcr) {
      parts.push(
        `The strongest benefit cost ratio in the base case is ${esc(bestBcr.t.name)}, with a benefit cost ratio of ${Number.isFinite(bestBcr.m.bcr) ? fmt(bestBcr.m.bcr) : "n/a"}.`
      );
    }

    parts.push(
      `Treatments can outperform the control by increasing benefits, by reducing costs, or by both. Use the change in present value of benefits and the change in present value of costs to see which mechanism drives each result.`
    );

    if (improving.length) {
      const share = (improving.length / nonControl.length) * 100;
      parts.push(
        `${improving.length} out of ${nonControl.length} treatments have a higher net present value than the control, which is ${fmt(share)} percent of treatments.`
      );
    } else {
      parts.push("No treatments have a higher net present value than the control in the base case settings.");
    }

    if (worsening.length) {
      parts.push(
        `Some treatments underperform because the additional costs are not offset by enough yield benefit at the current price and persistence settings. In those cases, results are most sensitive to the grain price, the duration that benefits persist, and how often costs recur.`
      );
    }

    parts.push(
      `The control annual net position, defined as annual benefits minus annual operating costs, is ${money(ctrlAnnual)} in the base case. This provides a reference point for how much extra annual benefit a treatment needs to justify extra costs over time.`
    );

    root.textContent = parts.join("\n\n");
  }

  function initResultsFilters() {
    const setFilter = (f) => {
      resultsState.filter = f;
      calcAndRender();
      showToast("Results filter applied.");
    };

    const btnAll = $("#filterShowAll");
    if (btnAll) btnAll.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); setFilter("all"); });

    const btnTopNpv = $("#filterTopNpv");
    if (btnTopNpv) btnTopNpv.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); setFilter("top_npv"); });

    const btnTopBcr = $("#filterTopBcr");
    if (btnTopBcr) btnTopBcr.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); setFilter("top_bcr"); });

    const btnImprove = $("#filterImproveOnly");
    if (btnImprove) btnImprove.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); setFilter("improve_only"); });
  }

  // =========================
  // 8) SENSITIVITY GRID ENGINE
  // =========================
  function computeSensitivityGrid() {
    const rateList = model.sensitivity.discountRatesPct && Array.isArray(model.sensitivity.discountRatesPct) && model.sensitivity.discountRatesPct.length
      ? model.sensitivity.discountRatesPct
      : [model.time.discLow, model.time.discBase, model.time.discHigh].filter(v => Number.isFinite(+v)).map(v => +v);

    const priceMultipliers = (model.sensitivity.priceMultipliers || DEFAULT_SENSITIVITY.priceMultipliers).filter(v => Number.isFinite(+v)).map(v => +v);
    const persistenceYears = (model.sensitivity.persistenceYears || DEFAULT_SENSITIVITY.persistenceYears).filter(v => Number.isFinite(+v)).map(v => Math.max(0, Math.floor(+v)));
    const recurrenceMultipliers = (model.sensitivity.recurrenceMultipliers || DEFAULT_SENSITIVITY.recurrenceMultipliers).filter(v => Number.isFinite(+v)).map(v => +v);

    const ctrl = getControlTreatment();
    const grid = [];
    const years = model.time.years;
    const adoptMul = model.adoption.base;
    const risk = model.risk.base;

    const basePrice = getYieldOutput()?.value || 0;

    model.treatments.forEach(t => {
      if (t.isControl) return;
      rateList.forEach(rate => {
        priceMultipliers.forEach(pm => {
          persistenceYears.forEach(py => {
            recurrenceMultipliers.forEach(rm => {
              const opts = {
                ratePct: rate,
                years,
                adoptMul,
                risk,
                priceMultiplier: pm,
                persistenceYears: py,
                recurrenceMultiplier: rm
              };
              const mT = computeTreatmentCBA(t, opts);
              const mC = ctrl ? computeTreatmentCBA(ctrl, opts) : null;

              const deltaNpv = mC ? (mT.npv - mC.npv) : NaN;
              const deltaPvCost = mC ? (mT.pvCosts - mC.pvCosts) : NaN;
              const deltaPvBen = mC ? (mT.pvBenefits - mC.pvBenefits) : NaN;

              grid.push({
                treatment: t.name,
                control: ctrl ? ctrl.name : "",
                discountRatePct: rate,
                pricePerTonne: Number.isFinite(+basePrice) ? +basePrice * pm : NaN,
                priceMultiplier: pm,
                persistenceYears: py,
                recurrenceMultiplier: rm,
                pvBenefits: mT.pvBenefits,
                pvCosts: mT.pvCosts,
                npv: mT.npv,
                bcr: mT.bcr,
                roiPct: mT.roiPct,
                deltaNpvVsControl: deltaNpv,
                deltaPvCostsVsControl: deltaPvCost,
                deltaPvBenefitsVsControl: deltaPvBen
              });
            });
          });
        });
      });
    });

    // Rank within each scenario cell by NPV
    const keyOf = r => [r.discountRatePct, r.priceMultiplier, r.persistenceYears, r.recurrenceMultiplier].join("|");
    const grouped = new Map();
    grid.forEach(r => {
      const k = keyOf(r);
      if (!grouped.has(k)) grouped.set(k, []);
      grouped.get(k).push(r);
    });
    grouped.forEach(list => {
      list.sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));
      list.forEach((r, i) => (r.rankByNpv = i + 1));
    });

    return grid;
  }

  function renderSensitivityGrid(grid) {
    const root = $("#sensitivityGrid") || $("#sensitivityTable") || $("#sensitivityResults");
    if (!root) return;

    root.innerHTML = "";

    if (!grid || !grid.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "No sensitivity results to display.";
      root.appendChild(p);
      return;
    }

    const maxRows = Math.min(grid.length, 2500);
    const view = grid.slice(0, maxRows);

    const table = document.createElement("table");
    table.className = "summary-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Treatment</th>
          <th>Discount rate (%)</th>
          <th>Price multiplier</th>
          <th>Price ($/t)</th>
          <th>Persistence (years)</th>
          <th>Recurrence multiplier</th>
          <th>NPV</th>
          <th>PV Benefits</th>
          <th>PV Costs</th>
          <th>BCR</th>
          <th>ROI (%)</th>
          <th>Δ NPV vs control</th>
          <th>Rank by NPV</th>
        </tr>
      </thead>
      <tbody>
        ${view
          .map(r => {
            const dCls = Number.isFinite(r.deltaNpvVsControl) ? (r.deltaNpvVsControl > 0 ? "pos" : r.deltaNpvVsControl < 0 ? "neg" : "neu") : "";
            return `
              <tr>
                <td>${esc(r.treatment)}</td>
                <td>${fmt(r.discountRatePct)}</td>
                <td>${fmt(r.priceMultiplier)}</td>
                <td>${money(r.pricePerTonne)}</td>
                <td>${esc(String(r.persistenceYears))}</td>
                <td>${fmt(r.recurrenceMultiplier)}</td>
                <td>${money(r.npv)}</td>
                <td>${money(r.pvBenefits)}</td>
                <td>${money(r.pvCosts)}</td>
                <td>${Number.isFinite(r.bcr) ? fmt(r.bcr) : "n/a"}</td>
                <td>${Number.isFinite(r.roiPct) ? fmt(r.roiPct) : "n/a"}</td>
                <td class="${dCls}">${money(r.deltaNpvVsControl)}</td>
                <td>${r.rankByNpv || ""}</td>
              </tr>
            `;
          })
          .join("")}
      </tbody>
    `;
    root.appendChild(table);

    const note = document.createElement("p");
    note.className = "small muted";
    note.textContent =
      grid.length > maxRows
        ? `Showing the first ${maxRows.toLocaleString()} rows out of ${grid.length.toLocaleString()} results. Use CSV or workbook export to obtain all rows.`
        : `Showing ${grid.length.toLocaleString()} sensitivity results.`;
    root.appendChild(note);
  }

  // =========================
  // 9) EXPORTS: CLEANED TSV, TREATMENT CSV, SENSITIVITY CSV, WORKBOOK
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

  function toCsvRow(arr) {
    return arr.map(x => {
      const s = x == null ? "" : String(x);
      if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
      return s;
    }).join(",");
  }

  function exportCleanedDatasetTsv() {
    const cleaned = trialState.cleaned;
    if (!cleaned || !cleaned.length) {
      alert("No cleaned dataset available. Import trial data first.");
      return;
    }
    const cols = ["__row", "replicate", "plot", "treatment", "yield_t_ha", "delta_yield_vs_control_t_ha", "labour_cost_per_ha", "operating_cost_per_ha", "services_cost_per_ha"];
    const lines = [];
    lines.push(cols.join("\t"));
    cleaned.forEach(r => {
      const rep = r.replicate || "";
      const plot = r.plot || "";
      const treat = r.treatment || "";
      const y = Number.isFinite(r.yield_t_ha) ? r.yield_t_ha : "";
      const dy = Number.isFinite(r.delta_yield_vs_control_t_ha) ? r.delta_yield_vs_control_t_ha : "";
      const lab = Number.isFinite(r.labour_cost_per_ha) ? r.labour_cost_per_ha : "";
      const op = Number.isFinite(r.operating_cost_per_ha) ? r.operating_cost_per_ha : "";
      const svc = Number.isFinite(r.services_cost_per_ha) ? r.services_cost_per_ha : "";
      lines.push([r.__row, rep, plot, treat, y, dy, lab, op, svc].join("\t"));
    });
    downloadFile(`${slug(model.project.name)}_cleaned_dataset.tsv`, lines.join("\n"), "text/tab-separated-values");
    showToast("Cleaned dataset TSV downloaded.");
  }

  function exportTreatmentSummaryCsv(baseResults) {
    const rows = [];
    rows.push([
      "Treatment",
      "Is control",
      "Rank by NPV",
      "PV Benefits",
      "PV Costs",
      "NPV",
      "BCR",
      "ROI (%)",
      "IRR (%)",
      "MIRR (%)",
      "Payback (years)",
      "Annual benefit",
      "Annual operating cost",
      "Capital cost (year 0)",
      "Δ PV Benefits vs control",
      "Δ PV Costs vs control",
      "Δ NPV vs control",
      "Δ NPV vs control (%)"
    ]);

    const ctrl = baseResults.control;
    const cM = baseResults.controlMetrics;

    baseResults.rows.slice().sort((a, b) => (a.isControl ? -1 : 1) - (b.isControl ? -1 : 1) || (a.rank || 9999) - (b.rank || 9999)).forEach(r => {
      const dBen = cM && !r.isControl ? (r.m.pvBenefits - cM.pvBenefits) : "";
      const dCost = cM && !r.isControl ? (r.m.pvCosts - cM.pvCosts) : "";
      const dNpv = cM && !r.isControl ? (r.m.npv - cM.npv) : "";
      const dPct = cM && !r.isControl && Number.isFinite(cM.npv) && cM.npv !== 0 ? (dNpv / cM.npv) * 100 : "";

      rows.push([
        r.t.name,
        r.isControl ? "Yes" : "No",
        r.isControl ? "Baseline" : (r.rank || ""),
        r.m.pvBenefits,
        r.m.pvCosts,
        r.m.npv,
        r.m.bcr,
        r.m.roiPct,
        r.m.irrPct,
        r.m.mirrPct,
        r.m.paybackYears != null ? r.m.paybackYears : "",
        r.m.annualBenefit,
        r.m.annualOpCost,
        r.m.capitalCostY0,
        dBen,
        dCost,
        dNpv,
        dPct
      ]);
    });

    const csv = rows.map(toCsvRow).join("\r\n");
    downloadFile(`${slug(model.project.name)}_treatment_summary.csv`, csv, "text/csv");
    showToast("Treatment summary CSV downloaded.");
  }

  function exportSensitivityCsv(grid) {
    if (!grid || !grid.length) {
      alert("No sensitivity grid available. Recalculate results first.");
      return;
    }
    const rows = [];
    rows.push([
      "Treatment",
      "Control",
      "Discount rate (%)",
      "Price multiplier",
      "Price ($/t)",
      "Persistence (years)",
      "Recurrence multiplier",
      "PV Benefits",
      "PV Costs",
      "NPV",
      "BCR",
      "ROI (%)",
      "Δ PV Benefits vs control",
      "Δ PV Costs vs control",
      "Δ NPV vs control",
      "Rank by NPV"
    ]);

    grid.forEach(r => {
      rows.push([
        r.treatment,
        r.control,
        r.discountRatePct,
        r.priceMultiplier,
        r.pricePerTonne,
        r.persistenceYears,
        r.recurrenceMultiplier,
        r.pvBenefits,
        r.pvCosts,
        r.npv,
        r.bcr,
        r.roiPct,
        r.deltaPvBenefitsVsControl,
        r.deltaPvCostsVsControl,
        r.deltaNpvVsControl,
        r.rankByNpv
      ]);
    });

    const csv = rows.map(toCsvRow).join("\r\n");
    downloadFile(`${slug(model.project.name)}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Sensitivity grid CSV downloaded.");
  }

  function exportWorkbook(baseResults, grid) {
    if (typeof XLSX === "undefined") {
      alert("The SheetJS XLSX library is required for workbook export.");
      return;
    }

    const wb = XLSX.utils.book_new();

    // Base case sheet
    const baseAoA = [];
    baseAoA.push(["Project", model.project.name]);
    baseAoA.push(["Organisation", model.project.organisation]);
    baseAoA.push(["Start year", model.time.startYear]);
    baseAoA.push(["Years", model.time.years]);
    baseAoA.push(["Discount rate (%)", model.time.discBase]);
    baseAoA.push(["Adoption multiplier", model.adoption.base]);
    baseAoA.push(["Risk", model.risk.base]);
    baseAoA.push([]);
    baseAoA.push(["Treatment", "Is control", "Rank by NPV", "PV Benefits", "PV Costs", "NPV", "BCR", "ROI (%)", "IRR (%)", "Payback (years)", "Δ NPV vs control", "Δ PV Costs vs control"]);
    const cM = baseResults.controlMetrics;
    baseResults.rows.slice().sort((a, b) => (a.isControl ? -1 : 1) - (b.isControl ? -1 : 1) || (a.rank || 9999) - (b.rank || 9999)).forEach(r => {
      const dNpv = cM && !r.isControl ? (r.m.npv - cM.npv) : "";
      const dCost = cM && !r.isControl ? (r.m.pvCosts - cM.pvCosts) : "";
      baseAoA.push([
        r.t.name,
        r.isControl ? "Yes" : "No",
        r.isControl ? "Baseline" : (r.rank || ""),
        r.m.pvBenefits,
        r.m.pvCosts,
        r.m.npv,
        r.m.bcr,
        r.m.roiPct,
        r.m.irrPct,
        r.m.paybackYears != null ? r.m.paybackYears : "",
        dNpv,
        dCost
      ]);
    });

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(baseAoA), "BaseCase");

    // Sensitivity sheet
    if (grid && grid.length) {
      const sensAoA = [];
      sensAoA.push([
        "Treatment",
        "Discount rate (%)",
        "Price multiplier",
        "Price ($/t)",
        "Persistence (years)",
        "Recurrence multiplier",
        "NPV",
        "PV Benefits",
        "PV Costs",
        "BCR",
        "ROI (%)",
        "Δ NPV vs control",
        "Rank by NPV"
      ]);
      grid.forEach(r => {
        sensAoA.push([
          r.treatment,
          r.discountRatePct,
          r.priceMultiplier,
          r.pricePerTonne,
          r.persistenceYears,
          r.recurrenceMultiplier,
          r.npv,
          r.pvBenefits,
          r.pvCosts,
          r.bcr,
          r.roiPct,
          r.deltaNpvVsControl,
          r.rankByNpv
        ]);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sensAoA), "Sensitivity");
    }

    // Data checks sheet
    const checksAoA = [];
    checksAoA.push(["Severity", "Code", "Message"]);
    (trialState.validation.issues || []).forEach(i => {
      checksAoA.push([i.severity || "", i.code || "", i.message || ""]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(checksAoA), "DataChecks");

    // Cleaned data sheet
    if (trialState.cleaned && trialState.cleaned.length) {
      const cols = ["__row", "replicate", "plot", "treatment", "yield_t_ha", "delta_yield_vs_control_t_ha", "labour_cost_per_ha", "operating_cost_per_ha", "services_cost_per_ha"];
      const dataAoA = [cols];
      trialState.cleaned.forEach(r => {
        dataAoA.push([
          r.__row,
          r.replicate || "",
          r.plot || "",
          r.treatment || "",
          Number.isFinite(r.yield_t_ha) ? r.yield_t_ha : "",
          Number.isFinite(r.delta_yield_vs_control_t_ha) ? r.delta_yield_vs_control_t_ha : "",
          Number.isFinite(r.labour_cost_per_ha) ? r.labour_cost_per_ha : "",
          Number.isFinite(r.operating_cost_per_ha) ? r.operating_cost_per_ha : "",
          Number.isFinite(r.services_cost_per_ha) ? r.services_cost_per_ha : ""
        ]);
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dataAoA), "CleanedData");
    }

    // Write
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slug(model.project.name)}_results_workbook.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Workbook export downloaded.");
  }

  // Bind export buttons (optional IDs)
  function initExportBindings() {
    const btnClean = $("#exportCleanedTsv");
    if (btnClean) btnClean.addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); exportCleanedDatasetTsv(); });

    const btnTreat = $("#exportTreatmentSummary");
    if (btnTreat) btnTreat.addEventListener("click", e => {
      e.preventDefault(); e.stopPropagation();
      const opts = getBaseCaseOpts();
      const per = buildPerTreatmentResults(opts);
      exportTreatmentSummaryCsv(per);
    });

    const btnSens = $("#exportSensitivityGrid");
    if (btnSens) btnSens.addEventListener("click", e => {
      e.preventDefault(); e.stopPropagation();
      const grid = computeSensitivityGrid();
      exportSensitivityCsv(grid);
    });

    const btnWb = $("#exportWorkbook");
    if (btnWb) btnWb.addEventListener("click", e => {
      e.preventDefault(); e.stopPropagation();
      const opts = getBaseCaseOpts();
      const per = buildPerTreatmentResults(opts);
      const grid = computeSensitivityGrid();
      exportWorkbook(per, grid);
    });
  }

  // =========================
  // 10) AI BRIEFING TAB: NARRATIVE PROMPT + COPY RESULTS JSON
  // =========================
  function buildAiBriefingPayload() {
    const opts = getBaseCaseOpts();
    const project = computeProjectCBA(opts);
    const per = buildPerTreatmentResults(opts);
    const grid = computeSensitivityGrid();

    // Summaries for trial calibration
    const trialSummary = trialState.metrics ? trialState.metrics.summary : null;

    const ranked = per.rows
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (Number.isFinite(b.m.npv) ? b.m.npv : -Infinity) - (Number.isFinite(a.m.npv) ? a.m.npv : -Infinity))
      .map(r => ({
        name: r.t.name,
        rankByNpv: r.rank,
        pvBenefits: r.m.pvBenefits,
        pvCosts: r.m.pvCosts,
        npv: r.m.npv,
        bcr: r.m.bcr,
        roiPct: r.m.roiPct,
        deltaNpvVsControl: r.delta ? r.delta.npv : NaN,
        deltaPvCostsVsControl: r.delta ? r.delta.pvCosts : NaN,
        deltaPvBenefitsVsControl: r.delta ? r.delta.pvBenefits : NaN
      }));

    return {
      toolName: "Farming CBA Decision Tool",
      timestamp: new Date().toISOString(),
      project: {
        name: model.project.name,
        organisation: model.project.organisation,
        summary: model.project.summary,
        objectives: model.project.objectives,
        stakeholders: model.project.stakeholders,
        lastUpdated: model.project.lastUpdated
      },
      calibration: {
        mode: model.calibration.mode,
        controlHint: model.calibration.controlNameHint,
        trialSummary
      },
      settings: {
        startYear: model.time.startYear,
        years: model.time.years,
        discountRateBasePct: opts.ratePct,
        adoptionMultiplier: opts.adoptMul,
        risk: opts.risk,
        pricePerTonne: opts.pricePerTonne,
        persistenceYears: opts.persistenceYears
      },
      baseCaseProjectResults: project,
      control: per.control ? { name: per.control.name } : null,
      baseCaseTreatmentResults: ranked,
      dataChecks: trialState.validation.issues || [],
      sensitivityGrid: grid
    };
  }

  function buildAiBriefingPrompt(payload) {
    const lines = [];
    // No bullets, no em dash, no abbreviations, no decision rules, no hard thresholds.
    lines.push(`Write a clear decision support brief for a farm manager using the results provided in the JSON that follows.`);
    lines.push(`Use full sentences and paragraphs only. Do not use bullet points. Do not use em dash punctuation. Do not use abbreviations.`);
    lines.push(`Do not tell the reader which treatment to choose. Do not impose decision rules or thresholds.`);
    lines.push(`Explain what drives differences relative to the control baseline. Separate whether gains come from higher benefits, lower costs, or both.`);
    lines.push(`Make sure you describe uncertainty using the sensitivity grid and the data checks.`);
    lines.push(`Where you refer to outcomes, write the numbers in dollars and percentages, and state the horizon and discount rate used.`);
    lines.push(`If a treatment appears strong, explain what assumptions are required for that to hold. If it appears weak, explain what could change to improve it.`);
    lines.push(`If trial calibration was used, explain that yield uplifts and costs were derived from plot data and that replicate specific control baselines were used when replicate information was available.`);
    lines.push(`Include a short section that lists what additional information would most improve confidence in these estimates.`);
    lines.push(`Below is the JSON input. Use only the information in that JSON. Do not invent any data.`);
    lines.push("");
    lines.push(JSON.stringify(payload, null, 2));
    return lines.join("\n");
  }

  function renderAiBriefingTab() {
    const preview = $("#aiBriefingPreview") || $("#copilotPreview");
    if (!preview) return;

    const payload = buildAiBriefingPayload();
    const prompt = buildAiBriefingPrompt(payload);
    preview.value = prompt;

    // Optional: separate JSON preview
    const jsonBox = $("#resultsJsonPreview");
    if (jsonBox) jsonBox.value = JSON.stringify(payload, null, 2);
  }

  function initAiBriefingBindings() {
    const btnCopyPrompt = $("#copyAiPrompt") || $("#copyBriefingPrompt");
    if (btnCopyPrompt) {
      btnCopyPrompt.addEventListener("click", async e => {
        e.preventDefault();
        e.stopPropagation();
        const preview = $("#aiBriefingPreview") || $("#copilotPreview");
        const text = preview ? String(preview.value || "") : "";
        if (!text.trim()) {
          renderAiBriefingTab();
        }
        const ok = await copyToClipboard(text || (preview ? preview.value : ""));
        showToast(ok ? "AI briefing prompt copied." : "Unable to copy automatically. Copy from the text box.");
      });
    }

    const btnCopyJson = $("#copyResultsJson") || $("#copyResultsJSON");
    if (btnCopyJson) {
      btnCopyJson.addEventListener("click", async e => {
        e.preventDefault();
        e.stopPropagation();
        const payload = buildAiBriefingPayload();
        const json = JSON.stringify(payload, null, 2);
        const ok = await copyToClipboard(json);
        showToast(ok ? "Results JSON copied." : "Unable to copy automatically. Copy from the JSON box if available.");
      });
    }

    const btnRefresh = $("#refreshAiBriefing") || $("#buildAiBriefing");
    if (btnRefresh) {
      btnRefresh.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        renderAiBriefingTab();
        showToast("AI briefing refreshed.");
      });
    }
  }

  // =========================
  // 11) SCENARIO SAVE / LOAD (LOCALSTORAGE JSON)
  // =========================
  function getScenarioIndex() {
    try {
      const raw = localStorage.getItem(STORAGE_KEYS.scenarioIndex);
      const obj = raw ? JSON.parse(raw) : null;
      return Array.isArray(obj) ? obj : [];
    } catch (e) {
      return [];
    }
  }

  function setScenarioIndex(list) {
    try {
      localStorage.setItem(STORAGE_KEYS.scenarioIndex, JSON.stringify(list));
    } catch (e) {
      // ignore
    }
  }

  function scenarioKey(name) {
    return `${APP_STORAGE_PREFIX}.scenario.${slug(name || "scenario")}`;
  }

  function saveScenario(name) {
    const safeName = (name || "").trim() || `scenario_${new Date().toISOString().slice(0, 10)}`;
    const key = scenarioKey(safeName);
    const payload = {
      savedAt: new Date().toISOString(),
      name: safeName,
      model,
      resultsFilter: resultsState.filter
    };
    try {
      localStorage.setItem(key, JSON.stringify(payload));
      const idx = getScenarioIndex();
      if (!idx.includes(safeName)) idx.unshift(safeName);
      setScenarioIndex(idx.slice(0, 50));
      localStorage.setItem(STORAGE_KEYS.lastScenario, safeName);
      showToast("Scenario saved.");
      renderScenarioList();
    } catch (e) {
      alert("Unable to save scenario. Storage may be full.");
    }
  }

  function loadScenario(name) {
    const safeName = (name || "").trim();
    if (!safeName) return;
    const key = scenarioKey(safeName);
    const raw = localStorage.getItem(key);
    if (!raw) {
      alert("Scenario not found.");
      return;
    }
    try {
      const obj = JSON.parse(raw);
      if (obj && obj.model) {
        // Replace model fields in place to keep references stable
        Object.keys(model).forEach(k => delete model[k]);
        Object.assign(model, obj.model);

        // Re-init guards
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        if (!model.sensitivity) model.sensitivity = JSON.parse(JSON.stringify(DEFAULT_SENSITIVITY));
        initTreatmentDeltasAndRecurrence();

        resultsState.filter = obj.resultsFilter || "all";

        // Reset UI
        renderAll();
        setBasicsFieldsFromModel();
        calcAndRender();
        showToast("Scenario loaded.");
      }
    } catch (e) {
      console.error(e);
      alert("Unable to load scenario.");
    }
  }

  function deleteScenario(name) {
    const safeName = (name || "").trim();
    if (!safeName) return;
    const key = scenarioKey(safeName);
    try {
      localStorage.removeItem(key);
      const idx = getScenarioIndex().filter(n => n !== safeName);
      setScenarioIndex(idx);
      showToast("Scenario deleted.");
      renderScenarioList();
    } catch (e) {
      // ignore
    }
  }

  function renderScenarioList() {
    const sel = $("#scenarioSelect");
    if (!sel) return;
    const idx = getScenarioIndex();
    sel.innerHTML = "";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = "Select a saved scenario";
    sel.appendChild(opt0);
    idx.forEach(name => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sel.appendChild(opt);
    });
    const last = localStorage.getItem(STORAGE_KEYS.lastScenario) || "";
    if (last && idx.includes(last)) sel.value = last;
  }

  function initScenarioBindings() {
    const btnSave = $("#saveScenario");
    const btnLoad = $("#loadScenario");
    const btnDelete = $("#deleteScenario");
    const nameInput = $("#scenarioName");
    const select = $("#scenarioSelect");

    if (btnSave) {
      btnSave.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const name = nameInput ? nameInput.value : (select ? select.value : "");
        saveScenario(name);
      });
    }

    if (btnLoad) {
      btnLoad.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const name = select ? select.value : (nameInput ? nameInput.value : "");
        if (!name) {
          alert("Select a scenario to load.");
          return;
        }
        loadScenario(name);
      });
    }

    if (btnDelete) {
      btnDelete.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const name = select ? select.value : (nameInput ? nameInput.value : "");
        if (!name) {
          alert("Select a scenario to delete.");
          return;
        }
        if (!confirm("Delete this scenario?")) return;
        deleteScenario(name);
      });
    }

    if (select) {
      select.addEventListener("change", e => {
        if (nameInput) nameInput.value = e.target.value;
      });
    }

    renderScenarioList();
  }

  // =========================
  // 12) UI: TABS + ACTIONS + FORMS (EXISTING + EXTENDED)
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

  // Base case options
  function getBaseCaseOpts() {
    const yOut = getYieldOutput();
    const price = model.calibration.pricePerTonne != null && Number.isFinite(+model.calibration.pricePerTonne)
      ? +model.calibration.pricePerTonne
      : (Number.isFinite(+yOut.value) ? +yOut.value : 0);

    const persistence = model.calibration.persistenceYears != null && Number.isFinite(+model.calibration.persistenceYears)
      ? Math.max(0, Math.floor(+model.calibration.persistenceYears))
      : model.time.years;

    return {
      ratePct: model.time.discBase,
      years: model.time.years,
      adoptMul: model.adoption.base,
      risk: model.risk.base,
      bcrMode: model.sim.bcrMode,
      pricePerTonne: price,
      priceMultiplier: 1.0,
      persistenceYears: persistence,
      recurrenceMultiplier: 1.0
    };
  }

  function initActions() {
    document.addEventListener("click", e => {
      const el = e.target.closest("#recalc, #getResults, [data-action='recalc']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      calcAndRender();
      showToast("Base case economic indicators recalculated.");
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#runSim, [data-action='run-sim']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      runSimulation();
    });

    // Optional: run sensitivity button
    const btnSens = $("#runSensitivity");
    if (btnSens) {
      btnSens.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        const grid = computeSensitivityGrid();
        renderSensitivityGrid(grid);
        showToast("Sensitivity grid calculated.");
      });
    }

    // Optional: import file button (if present)
    const importBtn = $("#importData") || $("#commitImport");
    if (importBtn) {
      importBtn.addEventListener("click", e => {
        e.preventDefault();
        e.stopPropagation();
        // If paste box exists, commit from paste; otherwise rely on file input handler
        const pasteBox = $("#dataPaste") || $("#pasteData") || $("#trialPaste") || $("#pasteBox");
        if (pasteBox && String(pasteBox.value || "").trim()) {
          handleImportText(String(pasteBox.value || ""), { source: "paste" });
        } else {
          showToast("Use the upload control to import a file, or paste data and commit.");
        }
      });
    }
  }

  // =========================
  // 13) FORMS: SET + BIND BASIC FIELDS
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

    if ($("#simN")) $("#simN").value = model.sim.n;
    if ($("#targetBCR")) $("#targetBCR").value = model.sim.targetBCR;
    if ($("#bcrMode")) $("#bcrMode").value = model.sim.bcrMode;
    if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = model.sim.targetBCR;

    if ($("#simVarPct")) $("#simVarPct").value = String(model.sim.variationPct || 20);
    if ($("#simVaryOutputs")) $("#simVaryOutputs").value = model.sim.varyOutputs ? "true" : "false";
    if ($("#simVaryTreatCosts")) $("#simVaryTreatCosts").value = model.sim.varyTreatCosts ? "true" : "false";
    if ($("#simVaryInputCosts")) $("#simVaryInputCosts").value = model.sim.varyInputCosts ? "true" : "false";

    if ($("#systemType")) $("#systemType").value = model.outputsMeta.systemType || "single";
    if ($("#outputAssumptions")) $("#outputAssumptions").value = model.outputsMeta.assumptions || "";

    // Calibration controls (optional)
    if ($("#calibrationMode")) $("#calibrationMode").value = model.calibration.mode || "model";
    if ($("#pricePerTonne")) $("#pricePerTonne").value = model.calibration.pricePerTonne != null ? model.calibration.pricePerTonne : "";
    if ($("#persistenceYears")) $("#persistenceYears").value = model.calibration.persistenceYears != null ? model.calibration.persistenceYears : "";
    if ($("#controlHint")) $("#controlHint").value = model.calibration.controlNameHint || "control";

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
  }

  let debTimer = null;
  function calcAndRenderDebounced() {
    clearTimeout(debTimer);
    debTimer = setTimeout(calcAndRender, 120);
  }

  function bindBasics() {
    setBasicsFieldsFromModel();
    initActions();

    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", e => {
        e.stopPropagation();
        const r =
          1 -
          (1 - num("#rTech")) *
            (1 - num("#rNonCoop")) *
            (1 - num("#rSocio")) *
            (1 - num("#rFin")) *
            (1 - num("#rMan"));
        if ($("#combinedRiskOut")) $("#combinedRiskOut").textContent = "Combined: " + (r * 100).toFixed(2) + "%";
        if ($("#riskBase")) $("#riskBase").value = r.toFixed(3);
        model.risk.base = r;
        calcAndRender();
        showToast("Combined risk updated from component risks.");
      });
    }

    const addCostBtn = $("#addCost");
    if (addCostBtn) {
      addCostBtn.addEventListener("click", e => {
        e.stopPropagation();
        const c = {
          id: uid(),
          label: "New cost",
          type: "annual",
          category: "Services",
          annual: 0,
          startYear: model.time.startYear,
          endYear: model.time.startYear,
          capital: 0,
          year: model.time.startYear,
          constrained: true,
          depMethod: "none",
          depLife: 5,
          depRate: 30
        };
        model.otherCosts.push(c);
        renderCosts();
        calcAndRender();
        showToast("New cost item added.");
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

        case "simN": model.sim.n = +t.value; break;
        case "targetBCR":
          model.sim.targetBCR = +t.value;
          if ($("#simBcrTargetLabel")) $("#simBcrTargetLabel").textContent = t.value;
          break;
        case "bcrMode": model.sim.bcrMode = t.value; break;
        case "randSeed": model.sim.seed = t.value ? +t.value : null; break;

        case "simVarPct": model.sim.variationPct = +t.value || 20; break;
        case "simVaryOutputs": model.sim.varyOutputs = t.value === "true"; break;
        case "simVaryTreatCosts": model.sim.varyTreatCosts = t.value === "true"; break;
        case "simVaryInputCosts": model.sim.varyInputCosts = t.value === "true"; break;

        case "systemType": model.outputsMeta.systemType = t.value; break;
        case "outputAssumptions": model.outputsMeta.assumptions = t.value; break;

        // Calibration
        case "calibrationMode":
          model.calibration.mode = t.value;
          if (model.calibration.mode === "trial") applyTrialCalibrationToModel();
          renderAll();
          calcAndRender();
          showToast("Calibration mode updated.");
          return;
        case "pricePerTonne":
          model.calibration.pricePerTonne = t.value === "" ? null : +t.value;
          break;
        case "persistenceYears":
          model.calibration.persistenceYears = t.value === "" ? null : +t.value;
          break;
        case "controlHint":
          model.calibration.controlNameHint = t.value;
          break;

        // Sensitivity settings (optional)
        case "sensPriceMultipliers":
          model.sensitivity.priceMultipliers = String(t.value || "")
            .split(/[,\s]+/g)
            .map(x => +x)
            .filter(x => Number.isFinite(x));
          break;
        case "sensDiscountRates":
          model.sensitivity.discountRatesPct = String(t.value || "")
            .split(/[,\s]+/g)
            .map(x => +x)
            .filter(x => Number.isFinite(x));
          break;
        case "sensPersistenceYears":
          model.sensitivity.persistenceYears = String(t.value || "")
            .split(/[,\s]+/g)
            .map(x => +x)
            .filter(x => Number.isFinite(x))
            .map(x => Math.max(0, Math.floor(x)));
          break;
        case "sensRecurrenceMultipliers":
          model.sensitivity.recurrenceMultipliers = String(t.value || "")
            .split(/[,\s]+/g)
            .map(x => +x)
            .filter(x => Number.isFinite(x));
          break;
      }

      calcAndRenderDebounced();
    });

    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        const data = JSON.stringify(model, null, 2);
        downloadFile(
          "cba_" + (model.project.name || "project").replace(/\s+/g, "_") + ".json",
          data,
          "application/json"
        );
        showToast("Project JSON downloaded.");
      });
    }

    const loadProjectBtn = $("#loadProject");
    const loadFileInput = $("#loadFile");
    if (loadProjectBtn && loadFileInput) {
      loadProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        loadFileInput.click();
      });
      loadFileInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          Object.keys(model).forEach(k => delete model[k]);
          Object.assign(model, obj);

          if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          if (!model.sensitivity) model.sensitivity = JSON.parse(JSON.stringify(DEFAULT_SENSITIVITY));
          initTreatmentDeltasAndRecurrence();

          renderAll();
          setBasicsFieldsFromModel();
          calcAndRender();
          showToast("Project JSON loaded and applied.");
        } catch (err) {
          alert("Invalid JSON file.");
          console.error(err);
        } finally {
          e.target.value = "";
        }
      });
    }

    // Existing CSV/PDF exports (kept)
    const exportCsvBtn = $("#exportCsv");
    const exportCsvFootBtn = $("#exportCsvFoot");
    if (exportCsvBtn) exportCsvBtn.addEventListener("click", e => { e.stopPropagation(); exportAllCsvLegacy(); });
    if (exportCsvFootBtn) exportCsvFootBtn.addEventListener("click", e => { e.stopPropagation(); exportAllCsvLegacy(); });

    const exportPdfBtn = $("#exportPdf");
    const exportPdfFootBtn = $("#exportPdfFoot");
    if (exportPdfBtn) exportPdfBtn.addEventListener("click", e => { e.stopPropagation(); exportPdf(); showToast("Print dialog opened for PDF export."); });
    if (exportPdfFootBtn) exportPdfFootBtn.addEventListener("click", e => { e.stopPropagation(); exportPdf(); showToast("Print dialog opened for PDF export."); });

    // Excel parse/import retained but expanded for csv/tsv/txt as well
    const parseExcelBtn = $("#parseExcel");
    const importExcelBtn = $("#importExcel");
    if (parseExcelBtn) parseExcelBtn.addEventListener("click", e => { e.stopPropagation(); handleParseExcelOrText(); });
    if (importExcelBtn) importExcelBtn.addEventListener("click", e => { e.stopPropagation(); commitParsedToImport(); });

    const downloadTemplateBtn = $("#downloadTemplate");
    const downloadSampleBtn = $("#downloadSample");
    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", e => { e.stopPropagation(); downloadExcelTemplate(); });
    if (downloadSampleBtn) downloadSampleBtn.addEventListener("click", e => { e.stopPropagation(); downloadSampleDataset(); });

    const startBtn = $("#startBtn");
    if (startBtn) {
      startBtn.addEventListener("click", e => {
        e.stopPropagation();
        switchTab("project");
        showToast("Welcome. Start with the Project tab.");
      });
    }

    const openCopilotBtns = $$("#openCopilot");
    if (openCop.value; break;

        // Calibration + sensitivity (optional IDs)
        case "calibrationMode":
          model.calibration.mode = t.value === "trial" ? "trial" : "model";
          if (model.calibration.mode === "trial") applyTrialCalibrationToModel();
          break;
        case "pricePerTonne":
          model.calibration.pricePerTonne = t.value && t.value.trim() !== "" ? parseNumber(t.value) : null;
          break;
        case "persistenceYears":
          model.calibration.persistenceYears = t.value && t.value.trim() !== "" ? parseInt(t.value, 10) : null;
          break;
        case "controlHint":
          model.calibration.controlNameHint = t.value || "control";
          break;

        case "sensDiscountRates":
          // comma-separated list, optional
          {
            const list = String(t.value || "")
              .split(/[,\s]+/g)
              .map(x => parseNumber(x))
              .filter(v => Number.isFinite(v));
            model.sensitivity.discountRatesPct = list.length ? list : null;
          }
          break;
        case "sensPriceMultipliers":
          {
            const list = String(t.value || "")
              .split(/[,\s]+/g)
              .map(x => parseNumber(x))
              .filter(v => Number.isFinite(v));
            if (list.length) model.sensitivity.priceMultipliers = list;
          }
          break;
        case "sensPersistenceYears":
          {
            const list = String(t.value || "")
              .split(/[,\s]+/g)
              .map(x => parseInt(x, 10))
              .filter(v => Number.isFinite(v))
              .map(v => Math.max(0, Math.floor(v)));
            if (list.length) model.sensitivity.persistenceYears = list;
          }
          break;
        case "sensRecurrenceMultipliers":
          {
            const list = String(t.value || "")
              .split(/[,\s]+/g)
              .map(x => parseNumber(x))
              .filter(v => Number.isFinite(v));
            if (list.length) model.sensitivity.recurrenceMultipliers = list;
          }
          break;

        default:
          break;
      }

      // If years changed, keep recurrence bounds sensible
      if (id === "years") {
        const N = Math.max(1, Math.floor(model.time.years || 1));
        model.treatments.forEach(tr => {
          ["cost", "benefit"].forEach(kind => {
            const rule = tr.recurrence && tr.recurrence[kind];
            if (!rule) return;
            const start = Math.max(0, Math.floor(parseInt(rule.startYearOffset ?? 1, 10) || 1));
            if (rule.endYearOffset != null && rule.endYearOffset !== "") {
              const end = Math.floor(parseInt(rule.endYearOffset, 10));
              rule.endYearOffset = clamp(end, start, N);
            }
          });
        });
      }

      calcAndRenderDebounced();
    });
  }

  // =========================
  // 14) RENDERERS (MISSING-SAFE, OPTIONAL DOM)
  // =========================
  function renderProjectSummaryCards(projectRes) {
    const root = $("#projectSummaryCards") || $("#projectSummary") || $("#projectSummaryPanel");
    if (!root) return;

    root.innerHTML = "";
    const cards = [
      { label: "PV benefits", value: money(projectRes.pvBenefits) },
      { label: "PV costs", value: money(projectRes.pvCosts) },
      { label: "NPV", value: money(projectRes.npv) },
      { label: "Benefit cost ratio", value: Number.isFinite(projectRes.bcr) ? fmt(projectRes.bcr) : "n/a" },
      { label: "Return on investment", value: Number.isFinite(projectRes.roiPct) ? percent(projectRes.roiPct) : "n/a" },
      { label: "Payback year", value: projectRes.paybackYears != null ? String(projectRes.paybackYears) : "n/a" }
    ];

    const grid = document.createElement("div");
    grid.className = "kpi-grid";
    cards.forEach(c => {
      const el = document.createElement("div");
      el.className = "kpi";
      el.innerHTML = `<div class="kpi-label">${esc(c.label)}</div><div class="kpi-value">${esc(c.value)}</div>`;
      grid.appendChild(el);
    });
    root.appendChild(grid);
  }

  function renderOutputs() {
    const root = $("#outputsList") || $("#outputs") || $("#outputsTable");
    if (!root) return;

    root.innerHTML = "";

    const table = document.createElement("table");
    table.className = "summary-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Output</th>
          <th>Unit</th>
          <th>Value (used as price or unit value)</th>
          <th>Source</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;

    const tb = table.querySelector("tbody");
    model.outputs.forEach(o => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${esc(o.name)}</td>
        <td>${esc(o.unit || "")}</td>
        <td>
          <input class="cell-input" data-output-id="${esc(o.id)}" data-field="value" value="${esc(o.value ?? "")}" />
        </td>
        <td>
          <input class="cell-input" data-output-id="${esc(o.id)}" data-field="source" value="${esc(o.source ?? "")}" />
        </td>
      `;
      tb.appendChild(tr);
    });

    root.appendChild(table);

    // Bind inline edits
    $$(".cell-input[data-output-id]").forEach(inp => {
      inp.addEventListener("input", e => {
        const id = e.target.dataset.outputId;
        const field = e.target.dataset.field;
        const out = model.outputs.find(x => x.id === id);
        if (!out) return;
        if (field === "value") out.value = parseNumber(e.target.value);
        else out[field] = e.target.value;
        calcAndRenderDebounced();
      });
    });

    const btnAdd = $("#addOutput");
    if (btnAdd && !btnAdd.__bound) {
      btnAdd.__bound = true;
      btnAdd.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        model.outputs.push({ id: uid(), name: "New output", unit: "", value: 0, source: "" });
        initTreatmentDeltasAndRecurrence();
        renderOutputs();
        renderTreatments();
        calcAndRender();
        showToast("Output added.");
      });
    }
  }

  function renderTreatments() {
    const root = $("#treatmentsList") || $("#treatments") || $("#treatmentsPanel");
    if (!root) return;

    root.innerHTML = "";

    const wrap = document.createElement("div");
    wrap.className = "treatments-wrap";

    const yOut = getYieldOutput();
    const yId = yOut ? yOut.id : null;

    model.treatments.forEach(trt => {
      const card = document.createElement("div");
      card.className = "treatment-card";

      const title = document.createElement("div");
      title.className = "treatment-head";
      title.innerHTML = `
        <div class="treatment-title">
          <input class="cell-input" data-trt-id="${esc(trt.id)}" data-field="name" value="${esc(trt.name)}" />
          <span class="badge ${trt.isControl ? "badge-control" : "badge-treatment"}">${trt.isControl ? "Control" : "Treatment"}</span>
        </div>
        <div class="treatment-meta">
          <label class="small muted">Area (ha)
            <input class="cell-input small-input" data-trt-id="${esc(trt.id)}" data-field="area" value="${esc(trt.area ?? 0)}" />
          </label>
          <label class="small muted">Adoption (0 to 1)
            <input class="cell-input small-input" data-trt-id="${esc(trt.id)}" data-field="adoption" value="${esc(trt.adoption ?? 1)}" />
          </label>
          <label class="small muted">Control
            <input type="radio" name="controlPick" data-action="set-control" data-trt-id="${esc(trt.id)}" ${trt.isControl ? "checked" : ""} />
          </label>
        </div>
      `;
      card.appendChild(title);

      // Deltas table
      const dTable = document.createElement("table");
      dTable.className = "summary-table tight";
      dTable.innerHTML = `
        <thead>
          <tr>
            <th>Delta (per ha)</th>
            <th>Value</th>
          </tr>
        </thead>
        <tbody>
          ${
            model.outputs
              .map(o => {
                const d = Number.isFinite(parseNumber(trt.deltas[o.id])) ? +trt.deltas[o.id] : 0;
                const label = o.id === yId ? `${o.name} (drives revenue)` : o.name;
                return `
                  <tr>
                    <td>${esc(label)}</td>
                    <td><input class="cell-input" data-trt-id="${esc(trt.id)}" data-delta-id="${esc(o.id)}" value="${esc(d)}" /></td>
                  </tr>
                `;
              })
              .join("")
          }
        </tbody>
      `;
      card.appendChild(dTable);

      // Costs + recurrence
      const costBlock = document.createElement("div");
      costBlock.className = "treatment-costs";
      costBlock.innerHTML = `
        <div class="grid-3">
          <label class="small muted">Labour cost per ha
            <input class="cell-input" data-trt-id="${esc(trt.id)}" data-field="labourCost" value="${esc(trt.labourCost ?? 0)}" />
          </label>
          <label class="small muted">Materials cost per ha
            <input class="cell-input" data-trt-id="${esc(trt.id)}" data-field="materialsCost" value="${esc(trt.materialsCost ?? 0)}" />
          </label>
          <label class="small muted">Services cost per ha
            <input class="cell-input" data-trt-id="${esc(trt.id)}" data-field="servicesCost" value="${esc(trt.servicesCost ?? 0)}" />
          </label>
        </div>

        <div class="grid-2">
          <label class="small muted">Capital cost (year 0, lump sum)
            <input class="cell-input" data-trt-id="${esc(trt.id)}" data-field="capitalCost" value="${esc(trt.capitalCost ?? 0)}" />
          </label>
          <label class="small muted">Constrained cost
            <select class="cell-input" data-trt-id="${esc(trt.id)}" data-field="constrained">
              <option value="true" ${trt.constrained ? "selected" : ""}>Yes</option>
              <option value="false" ${!trt.constrained ? "selected" : ""}>No</option>
            </select>
          </label>
        </div>

        <div class="recurrence">
          <div class="recurrence-col">
            <div class="small strong">Cost recurrence</div>
            <div class="grid-4">
              <label class="small muted">Mode
                <select class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="cost" data-rec-field="mode">
                  <option value="annual" ${(trt.recurrence?.cost?.mode || "annual") === "annual" ? "selected" : ""}>Annual</option>
                  <option value="every_n_years" ${(trt.recurrence?.cost?.mode || "") === "every_n_years" ? "selected" : ""}>Every N years</option>
                  <option value="once" ${(trt.recurrence?.cost?.mode || "") === "once" ? "selected" : ""}>Once</option>
                  <option value="custom" ${(trt.recurrence?.cost?.mode || "") === "custom" ? "selected" : ""}>Custom years</option>
                </select>
              </label>
              <label class="small muted">Every N
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="cost" data-rec-field="everyN" value="${esc(trt.recurrence?.cost?.everyN ?? 1)}" />
              </label>
              <label class="small muted">Start year
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="cost" data-rec-field="startYearOffset" value="${esc(trt.recurrence?.cost?.startYearOffset ?? 1)}" />
              </label>
              <label class="small muted">End year
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="cost" data-rec-field="endYearOffset" value="${esc(trt.recurrence?.cost?.endYearOffset ?? "")}" />
              </label>
            </div>
            <label class="small muted">Custom years (comma-separated)
              <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="cost" data-rec-field="yearsCsv" value="${esc(trt.recurrence?.cost?.yearsCsv ?? "")}" />
            </label>
          </div>

          <div class="recurrence-col">
            <div class="small strong">Benefit recurrence</div>
            <div class="grid-4">
              <label class="small muted">Mode
                <select class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="benefit" data-rec-field="mode">
                  <option value="annual" ${(trt.recurrence?.benefit?.mode || "annual") === "annual" ? "selected" : ""}>Annual</option>
                  <option value="every_n_years" ${(trt.recurrence?.benefit?.mode || "") === "every_n_years" ? "selected" : ""}>Every N years</option>
                  <option value="once" ${(trt.recurrence?.benefit?.mode || "") === "once" ? "selected" : ""}>Once</option>
                  <option value="custom" ${(trt.recurrence?.benefit?.mode || "") === "custom" ? "selected" : ""}>Custom years</option>
                </select>
              </label>
              <label class="small muted">Every N
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="benefit" data-rec-field="everyN" value="${esc(trt.recurrence?.benefit?.everyN ?? 1)}" />
              </label>
              <label class="small muted">Start year
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="benefit" data-rec-field="startYearOffset" value="${esc(trt.recurrence?.benefit?.startYearOffset ?? 1)}" />
              </label>
              <label class="small muted">End year
                <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="benefit" data-rec-field="endYearOffset" value="${esc(trt.recurrence?.benefit?.endYearOffset ?? "")}" />
              </label>
            </div>
            <label class="small muted">Custom years (comma-separated)
              <input class="cell-input" data-trt-id="${esc(trt.id)}" data-rec-kind="benefit" data-rec-field="yearsCsv" value="${esc(trt.recurrence?.benefit?.yearsCsv ?? "")}" />
            </label>
          </div>
        </div>

        <label class="small muted">Notes
          <textarea class="cell-input" data-trt-id="${esc(trt.id)}" data-field="notes" rows="2">${esc(trt.notes || "")}</textarea>
        </label>

        <div class="treatment-actions">
          <button class="btn small" data-action="duplicate-treatment" data-trt-id="${esc(trt.id)}">Duplicate</button>
          <button class="btn small danger" data-action="delete-treatment" data-trt-id="${esc(trt.id)}" ${trt.isControl ? "disabled" : ""}>Delete</button>
        </div>
      `;
      card.appendChild(costBlock);

      wrap.appendChild(card);
    });

    root.appendChild(wrap);

    // Bind events
    $$(".cell-input[data-trt-id][data-field]").forEach(inp => {
      inp.addEventListener("input", e => {
        const id = e.target.dataset.trtId;
        const field = e.target.dataset.field;
        const trt = model.treatments.find(x => x.id === id);
        if (!trt) return;

        if (field === "area" || field === "adoption" || field === "labourCost" || field === "materialsCost" || field === "servicesCost" || field === "capitalCost") {
          trt[field] = parseNumber(e.target.value);
          if (!Number.isFinite(trt[field])) trt[field] = 0;
        } else if (field === "constrained") {
          trt.constrained = String(e.target.value) === "true";
        } else {
          trt[field] = e.target.value;
        }
        calcAndRenderDebounced();
      });
    });

    $$(".cell-input[data-trt-id][data-delta-id]").forEach(inp => {
      inp.addEventListener("input", e => {
        const id = e.target.dataset.trtId;
        const dId = e.target.dataset.deltaId;
        const trt = model.treatments.find(x => x.id === id);
        if (!trt) return;
        trt.deltas[dId] = parseNumber(e.target.value);
        if (!Number.isFinite(trt.deltas[dId])) trt.deltas[dId] = 0;
        calcAndRenderDebounced();
      });
    });

    $$(".cell-input[data-trt-id][data-rec-kind][data-rec-field]").forEach(inp => {
      inp.addEventListener("input", e => {
        const id = e.target.dataset.trtId;
        const kind = e.target.dataset.recKind;
        const field = e.target.dataset.recField;
        const trt = model.treatments.find(x => x.id === id);
        if (!trt) return;
        if (!trt.recurrence) trt.recurrence = { cost: {}, benefit: {} };
        if (!trt.recurrence[kind]) trt.recurrence[kind] = {};

        let v = e.target.value;
        if (field === "everyN" || field === "startYearOffset" || field === "endYearOffset") {
          if (v === "" && field === "endYearOffset") trt.recurrence[kind][field] = null;
          else {
            const n = parseInt(v, 10);
            trt.recurrence[kind][field] = Number.isFinite(n) ? n : (field === "endYearOffset" ? null : 1);
          }
        } else {
          trt.recurrence[kind][field] = v;
        }
        calcAndRenderDebounced();
      });
    });

    $$("input[type='radio'][data-action='set-control']").forEach(r => {
      r.addEventListener("change", e => {
        const id = e.target.dataset.trtId;
        model.treatments.forEach(t => (t.isControl = t.id === id));
        renderTreatments();
        calcAndRender();
        showToast("Control baseline updated.");
      });
    });

    $$("[data-action='duplicate-treatment']").forEach(btn => {
      if (btn.__bound) return;
      btn.__bound = true;
      btn.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const id = e.target.closest("[data-action='duplicate-treatment']").dataset.trtId;
        const trt = model.treatments.find(x => x.id === id);
        if (!trt) return;
        const copy = JSON.parse(JSON.stringify(trt));
        copy.id = uid();
        copy.name = `${copy.name} (copy)`;
        copy.isControl = false;
        model.treatments.push(copy);
        initTreatmentDeltasAndRecurrence();
        renderTreatments();
        calcAndRender();
        showToast("Treatment duplicated.");
      });
    });

    $$("[data-action='delete-treatment']").forEach(btn => {
      if (btn.__bound) return;
      btn.__bound = true;
      btn.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const id = e.target.closest("[data-action='delete-treatment']").dataset.trtId;
        const trt = model.treatments.find(x => x.id === id);
        if (!trt || trt.isControl) return;
        model.treatments = model.treatments.filter(x => x.id !== id);
        renderTreatments();
        calcAndRender();
        showToast("Treatment deleted.");
      });
    });

    const btnAdd = $("#addTreatment");
    if (btnAdd && !btnAdd.__bound) {
      btnAdd.__bound = true;
      btnAdd.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const t = {
          id: uid(),
          name: "New treatment",
          area: 100,
          adoption: 1,
          deltas: {},
          labourCost: 0,
          materialsCost: 0,
          servicesCost: 0,
          capitalCost: 0,
          constrained: true,
          source: "User",
          isControl: false,
          notes: "",
          recurrence: {
            cost: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" },
            benefit: { mode: "annual", everyN: 1, startYearOffset: 1, endYearOffset: null, yearsCsv: "" }
          }
        };
        model.outputs.forEach(o => (t.deltas[o.id] = 0));
        model.treatments.push(t);
        renderTreatments();
        calcAndRender();
        showToast("Treatment added.");
      });
    }
  }

  function renderCosts() {
    const root = $("#costItems") || $("#costsList") || $("#costsPanel");
    if (!root) return;

    root.innerHTML = "";

    const table = document.createElement("table");
    table.className = "summary-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Cost item</th>
          <th>Type</th>
          <th>Annual</th>
          <th>Start year</th>
          <th>End year</th>
          <th>Capital</th>
          <th>Capital year</th>
          <th>Constrained</th>
          <th></th>
        </tr>
      </thead>
      <tbody></tbody>
    `;
    const tb = table.querySelector("tbody");

    model.otherCosts.forEach(c => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="label" value="${esc(c.label || "")}" /></td>
        <td>
          <select class="cell-input" data-cost-id="${esc(c.id)}" data-field="type">
            <option value="annual" ${c.type === "annual" ? "selected" : ""}>Annual</option>
            <option value="capital" ${c.type === "capital" ? "selected" : ""}>Capital</option>
          </select>
        </td>
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="annual" value="${esc(c.annual ?? 0)}" /></td>
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="startYear" value="${esc(c.startYear ?? model.time.startYear)}" /></td>
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="endYear" value="${esc(c.endYear ?? model.time.startYear)}" /></td>
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="capital" value="${esc(c.capital ?? 0)}" /></td>
        <td><input class="cell-input" data-cost-id="${esc(c.id)}" data-field="year" value="${esc(c.year ?? model.time.startYear)}" /></td>
        <td>
          <select class="cell-input" data-cost-id="${esc(c.id)}" data-field="constrained">
            <option value="true" ${c.constrained ? "selected" : ""}>Yes</option>
            <option value="false" ${!c.constrained ? "selected" : ""}>No</option>
          </select>
        </td>
        <td><button class="btn small danger" data-action="delete-cost" data-cost-id="${esc(c.id)}">Delete</button></td>
      `;
      tb.appendChild(tr);
    });

    root.appendChild(table);

    $$(".cell-input[data-cost-id]").forEach(inp => {
      inp.addEventListener("input", e => {
        const id = e.target.dataset.costId;
        const field = e.target.dataset.field;
        const c = model.otherCosts.find(x => x.id === id);
        if (!c) return;

        if (["annual", "capital"].includes(field)) c[field] = parseNumber(e.target.value);
        else if (["startYear", "endYear", "year"].includes(field)) c[field] = parseInt(e.target.value, 10) || model.time.startYear;
        else if (field === "constrained") c.constrained = String(e.target.value) === "true";
        else c[field] = e.target.value;

        calcAndRenderDebounced();
      });
    });

    $$("[data-action='delete-cost']").forEach(btn => {
      if (btn.__bound) return;
      btn.__bound = true;
      btn.addEventListener("click", e => {
        e.preventDefault(); e.stopPropagation();
        const id = e.target.closest("[data-action='delete-cost']").dataset.costId;
        model.otherCosts = model.otherCosts.filter(x => x.id !== id);
        renderCosts();
        calcAndRender();
        showToast("Cost item deleted.");
      });
    });
  }

  function renderAll() {
    // These are all missing-safe and only render if the target exists
    renderOutputs();
    renderTreatments();
    renderCosts();
    renderDataChecksPanel();
    renderScenarioList();
    renderAiBriefingTab();
  }

  // =========================
  // 15) CALCULATE + RENDER RESULTS (BASE + OPTIONAL SENSITIVITY)
  // =========================
  function calcAndRender() {
    try {
      // Ensure recurrence and deltas integrity
      initTreatmentDeltasAndRecurrence();

      // Apply trial calibration automatically when in trial mode
      if (model.calibration.mode === "trial") applyTrialCalibrationToModel();

      const opts = getBaseCaseOpts();

      // Base results
      const projectRes = computeProjectCBA(opts);
      const per = buildPerTreatmentResults(opts);

      renderProjectSummaryCards(projectRes);
      renderLeaderboard(per);
      renderComparisonToControlGrid(per);
      renderWhatThisMeans(per, opts);

      // Sensitivity, only if a container exists
      const sensRoot = $("#sensitivityGrid") || $("#sensitivityTable") || $("#sensitivityResults");
      if (sensRoot) {
        const grid = computeSensitivityGrid();
        renderSensitivityGrid(grid);
      }

      // AI briefing preview (cheap to refresh; grid generation happens inside payload, so only do if tab exists)
      if ($("#aiBriefingPreview") || $("#copilotPreview") || $("#resultsJsonPreview")) {
        renderAiBriefingTab();
      }

      // Optional: expose control name in UI
      const ctrl = per.control;
      if ($("#controlNameOut")) $("#controlNameOut").textContent = ctrl ? ctrl.name : "n/a";
    } catch (e) {
      console.error(e);
      // Avoid hard failure in production UI; show minimal message if a status element exists
      const status = $("#calcStatus") || $("#resultsStatus");
      if (status) status.textContent = "Unable to calculate results due to an internal error. Check console for details.";
    }
  }

  // =========================
  // 16) SIMULATION (MONTE CARLO)
  // =========================
  function quantile(arr, q) {
    const clean = arr.filter(v => Number.isFinite(v)).slice().sort((a, b) => a - b);
    if (!clean.length) return NaN;
    const pos = (clean.length - 1) * q;
    const base = Math.floor(pos);
    const rest = pos - base;
    if (clean[base + 1] === undefined) return clean[base];
    return clean[base] + rest * (clean[base + 1] - clean[base]);
  }

  function runSimulation() {
    const root = $("#simResults") || $("#simulationResults") || $("#simSummary");
    const n = Math.max(200, Math.floor(model.sim.n || 1000));
    const seed = model.sim.seed != null && model.sim.seed !== "" ? +model.sim.seed : null;
    const rand = rng(seed || undefined);

    const baseOpts = getBaseCaseOpts();
    const varPct = clamp(Number(model.sim.variationPct || 20), 0, 200) / 100;

    const npvArr = [];
    const bcrArr = [];
    const details = [];

    // Snapshot model outputs for variation
    const baseOutputs = model.outputs.map(o => ({ id: o.id, value: Number.isFinite(+o.value) ? +o.value : 0, name: o.name }));
    const baseTreatCosts = model.treatments.map(t => ({
      id: t.id,
      labourCost: Number.isFinite(+t.labourCost) ? +t.labourCost : 0,
      materialsCost: Number.isFinite(+t.materialsCost) ? +t.materialsCost : 0,
      servicesCost: Number.isFinite(+t.servicesCost) ? +t.servicesCost : 0,
      capitalCost: Number.isFinite(+t.capitalCost) ? +t.capitalCost : 0
    }));

    for (let i = 0; i < n; i++) {
      // Adoption and risk draws: triangular around base, within [low, high]
      const adopt = triangular(rand(), model.adoption.low, model.adoption.base, model.adoption.high);
      const risk = triangular(rand(), model.risk.low, model.risk.base, model.risk.high);

      // Temporarily vary outputs and costs if configured
      if (model.sim.varyOutputs) {
        baseOutputs.forEach(bo => {
          const mult = 1 + (rand() * 2 - 1) * varPct;
          const out = model.outputs.find(o => o.id === bo.id);
          if (out) out.value = bo.value * mult;
        });
      }

      if (model.sim.varyTreatCosts) {
        baseTreatCosts.forEach(bt => {
          const mult = 1 + (rand() * 2 - 1) * varPct;
          const t = model.treatments.find(x => x.id === bt.id);
          if (!t) return;
          t.labourCost = bt.labourCost * mult;
          t.materialsCost = bt.materialsCost * mult;
          t.servicesCost = bt.servicesCost * mult;
          t.capitalCost = bt.capitalCost * mult;
        });
      }

      // Optional: vary other costs
      const otherBackup = model.otherCosts.map(c => ({ id: c.id, annual: c.annual, capital: c.capital }));
      if (model.sim.varyInputCosts) {
        model.otherCosts.forEach(c => {
          const mult = 1 + (rand() * 2 - 1) * varPct;
          if (Number.isFinite(+c.annual)) c.annual = +c.annual * mult;
          if (Number.isFinite(+c.capital)) c.capital = +c.capital * mult;
        });
      }

      const opts = { ...baseOpts, adoptMul: clamp(adopt, 0, 1), risk: clamp(risk, 0, 1) };
      const res = computeProjectCBA(opts);

      npvArr.push(res.npv);
      bcrArr.push(res.bcr);
      if (i < 25) details.push({ i: i + 1, adoptMul: opts.adoptMul, risk: opts.risk, npv: res.npv, bcr: res.bcr });

      // Restore other costs if varied
      if (model.sim.varyInputCosts) {
        otherBackup.forEach(bk => {
          const c = model.otherCosts.find(x => x.id === bk.id);
          if (!c) return;
          c.annual = bk.annual;
          c.capital = bk.capital;
        });
      }
    }

    // Restore baseline after sim
    baseOutputs.forEach(bo => {
      const out = model.outputs.find(o => o.id === bo.id);
      if (out) out.value = bo.value;
    });
    baseTreatCosts.forEach(bt => {
      const t = model.treatments.find(x => x.id === bt.id);
      if (!t) return;
      t.labourCost = bt.labourCost;
      t.materialsCost = bt.materialsCost;
      t.servicesCost = bt.servicesCost;
      t.capitalCost = bt.capitalCost;
    });

    model.sim.results = { npv: npvArr, bcr: bcrArr };
    model.sim.details = details;

    const summary = {
      n,
      npv_mean: safeMean(npvArr),
      npv_median: safeMedian(npvArr),
      npv_p10: quantile(npvArr, 0.1),
      npv_p90: quantile(npvArr, 0.9),
      bcr_mean: safeMean(bcrArr),
      bcr_median: safeMedian(bcrArr),
      bcr_p10: quantile(bcrArr, 0.1),
      bcr_p90: quantile(bcrArr, 0.9)
    };

    if (root) {
      root.innerHTML = `
        <div class="sim-box">
          <div class="small muted">Simulation runs</div>
          <div class="big">${esc(n.toLocaleString())}</div>
          <div class="small muted">Assumptions</div>
          <div class="small">
            Adoption and risk vary between their low and high values. Outputs and costs vary by ±${esc(String(model.sim.variationPct || 20))} percent if enabled.
          </div>
        </div>

        <div class="sim-box">
          <div class="small muted">NPV distribution</div>
          <div class="small">Mean ${esc(money(summary.npv_mean))}, median ${esc(money(summary.npv_median))}</div>
          <div class="small">10th percentile ${esc(money(summary.npv_p10))}, 90th percentile ${esc(money(summary.npv_p90))}</div>
        </div>

        <div class="sim-box">
          <div class="small muted">Benefit cost ratio distribution</div>
          <div class="small">Mean ${esc(Number.isFinite(summary.bcr_mean) ? fmt(summary.bcr_mean) : "n/a")}, median ${esc(Number.isFinite(summary.bcr_median) ? fmt(summary.bcr_median) : "n/a")}</div>
          <div class="small">10th percentile ${esc(Number.isFinite(summary.bcr_p10) ? fmt(summary.bcr_p10) : "n/a")}, 90th percentile ${esc(Number.isFinite(summary.bcr_p90) ? fmt(summary.bcr_p90) : "n/a")}</div>
        </div>

        <div class="sim-box">
          <div class="small muted">First 25 draws (for checking)</div>
          <table class="summary-table tight">
            <thead><tr><th>Draw</th><th>Adoption</th><th>Risk</th><th>NPV</th><th>BCR</th></tr></thead>
            <tbody>
              ${details
                .map(d => `<tr><td>${d.i}</td><td>${fmt(d.adoptMul)}</td><td>${fmt(d.risk)}</td><td>${money(d.npv)}</td><td>${Number.isFinite(d.bcr) ? fmt(d.bcr) : "n/a"}</td></tr>`)
                .join("")}
            </tbody>
          </table>
        </div>
      `;
    }

    calcAndRender();
    showToast("Simulation complete.");
  }

  // =========================
  // 17) STARTUP
  // =========================
  function buildEmbeddedTrialText() {
    const cols = ["Amendment", "Yield t/ha", "Pre sowing Labour", "Treatment Input Cost Only /Ha"];
    const lines = [cols.join("\t")];
    RAW_PLOTS.forEach(r => {
      lines.push([r["Amendment"], r["Yield t/ha"], r["Pre sowing Labour"], r["Treatment Input Cost Only /Ha"]].join("\t"));
    });
    return lines.join("\n");
  }

  function initOnce() {
    initTabs();
    initImportPipelineBindings();
    initResultsFilters();
    initExportBindings();
    initAiBriefingBindings();
    initScenarioBindings();
    bindBasics();

    // Ensure toast root exists if the HTML did not include it
    if (!$("#toast-root")) {
      const div = document.createElement("div");
      div.id = "toast-root";
      document.body.appendChild(div);
    }

    renderAll();
    calcAndRender();

    // Load embedded trial defaults through the same pipeline if no import has occurred
    const alreadyImported = !!(trialState.dataRows && trialState.dataRows.length);
    const hasSnapshot = (() => {
      try { return !!localStorage.getItem(STORAGE_KEYS.lastImport); } catch (e) { return false; }
    })();

    if (!alreadyImported && !hasSnapshot && RAW_PLOTS && RAW_PLOTS.length) {
      handleImportText(buildEmbeddedTrialText(), { source: "embedded", delimiter: "\t" }).catch(err => console.error(err));
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initOnce);
  } else {
    initOnce();
  }
})();

