// Farming CBA Decision Tool 2 (Commercial-grade)
// Newcastle Business School | The University of Newcastle
// App.js fully upgraded: robust tabs, full default 2022 faba bean dataset ingestion,
// treatment-vs-control results table (indicators as rows, treatments as columns),
// Excel-first workflow (XLSX if available, CSV fallback), clean exports (Excel/CSV/PDF),
// AI prompt generator (non-prescriptive; includes improvement suggestions), tooltips near indicators,
// and fully dynamic cost component handling with Capital cost (year 0) before Total cost ($/ha).

(() => {
  "use strict";

  // =========================
  // GLOBAL CONFIG
  // =========================
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const TOOL_VERSION = "2.0.0";
  const ORG = "Newcastle Business School, The University of Newcastle";

  // Attempt to set document title + any visible header placeholders (non-breaking if missing)
  try {
    document.title = TOOL_NAME;
    const titleEls = [
      document.getElementById("toolTitle"),
      document.getElementById("appTitle"),
      document.getElementById("brandTitle"),
      document.querySelector("[data-role='tool-title']"),
      document.querySelector(".tool-title")
    ].filter(Boolean);
    titleEls.forEach(el => (el.textContent = TOOL_NAME));
  } catch (_) {}

  // =========================
  // CONSTANTS
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  // =========================
  // ID HELPER
  // =========================
  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  // =========================
  // UTILS
  // =========================
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const nowISO = () => new Date().toISOString().slice(0, 10);

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
    if (value === null || value === undefined || value === "") return NaN;
    if (typeof value === "number") return value;
    const s = String(value).trim();
    if (!s) return NaN;
    if (s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const n = parseFloat(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function downloadFile(filename, content, mime) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: mime || "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
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

  // lightweight debounce
  function debounce(fn, wait = 250) {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), wait);
    };
  }

  // RNG + triangular (for simulation)
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

  // =========================
  // MODEL (BASE)
  // =========================
  const model = {
    meta: {
      toolName: TOOL_NAME,
      version: TOOL_VERSION
    },
    project: {
      name: TOOL_NAME,
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: ORG,
      contactEmail: "",
      contactPhone: "",
      summary:
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum and fertiliser treatments against a control.",
      objectives: "Quantify yield and gross margin impacts of alternative soil amendment strategies.",
      activities:
        "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders: "Producers, agronomists, government agencies, research partners.",
      lastUpdated: nowISO(),
      goal:
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers adopt high-performing amendment packages on trial farms and similar soils in the region.",
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
    outputsMeta: { systemType: "single", assumptions: "" },

    // Output values should be interpretable per unit (e.g. grain price $/t)
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Default (editable)" },
      { id: uid(), name: "Screenings", unit: "percentage point", value: -20, source: "Default (editable)" },
      { id: uid(), name: "Protein", unit: "percentage point", value: 10, source: "Default (editable)" }
    ],

    // Treatments will be populated from default 2022 dataset on load
    treatments: [],

    // Optional project-wide benefits/costs remain supported (Excel workflow included)
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
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false,
      results: [],
      details: []
    }
  };

  // Ensure treatment deltas exist for each output
  function initTreatmentDeltas() {
    model.treatments.forEach(t => {
      t.deltas = t.deltas || {};
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (typeof t.adoption !== "number" || isNaN(t.adoption)) t.adoption = 1;
      if (typeof t.labourCost === "undefined") t.labourCost = 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0; // $ year 0 (total)
      if (typeof t.area === "undefined") t.area = 100;
      if (typeof t.constrained === "undefined") t.constrained = true;
      if (typeof t.isControl === "undefined") t.isControl = false;
    });
  }

  // =========================
  // DEFAULT DATASET (FULL) – 2022 Faba Beans
  // Stored as tab-separated rows with a canonical header line.
  // (The tool parses + keeps all columns; calculations use yield + cost columns heuristically,
  // while still retaining the full dataset for export and Excel-first workflow.)
  // =========================

  // Canonical header line (from the dataset’s final header row; duplicates are made unique at parse time).
  const FABA_2022_HEADERS_RAW = [
    "Plot",
    "Trt",
    "Rep",
    "Amendment",
    "Practice Change",
    "Plot Length (m)",
    "Plot Width (m)",
    "Plot Area (m^2)",
    "Plants/1m^2",
    "Yield t/ha",
    "Moisture",
    "Protein",
    "Anthesis Biomass t/ha",
    "Harvest Biomass t/ha",
    "Practice Change (Label)",
    "Application rate",
    "Treatment Input Cost Only /Ha",
    "Labour per Ha application could be included in next column",
    "Prototype Machinery for Adding amendments",
    "500hp tractor + Speed tiller task",
    "Tractor and 12 m air-seeder wet hire",
    "Sowing Labour included in wet hire",
    "Amberly Faba Bean",
    "Amberly Faba Bean",
    "DAP Fertiliser treated",
    "Inoculant F Pea/Faba",
    "4Farmers Ammonium Sulphate Herbicide Adjuvant",
    "4Farmers Ammonium Sulphate Herbicide Adjuvant",
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
    "Barrack fungicide discontinued",
    "Talstar",
    "Talstar",
    "Pre sowing Labour",
    "Amendment Labour",
    "Sowing Labour",
    "Herbicide Labour",
    "Herbicide Labour",
    "Herbicide Labour",
    "Harvesting Labour",
    "Harvesting Labour",
    "Pre sowing amendment 5 tyne ripper",
    "Speed tiller 10 m",
    "Air seeder 12 m",
    "36 m Boomspray",
    "Smaller tractor 150 hp",
    "Large Tractor 500hp",
    "Header 12 m front",
    "Ute",
    "Truck",
    "Utes $ per kilometer",
    "Trucks $ per kilometer",
    "Tractor",
    "Speed tiller",
    "Air seeder",
    "Boom spray",
    "Header",
    "Truck (Asset)",
    "Note"
  ];

  // Full 48 rows as provided (tab-separated). Keep as strings; parsing handles $ and commas and ?.
  const FABA_2022_ROWS_TSV = `
1\t12\t1\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 1\t20\t2.5\t50\t34\t7.03\t11.8\t23.2\t8.40\t15.51\tCrop 1\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t$210.00\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$4.24\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,945
2\t3\t1\tDeep OM (CP1)\tCrop 2\t20\t2.5\t50\t27\t5.18\t10.6\t23.6\t14.83\t16.46\tCrop 2\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385
3\t11\t1\tDeep Ripping\tCrop 3\t20\t2.5\t50\t33\t7.26\t10.7\t23.4\t17.89\t16.41\tCrop 3\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
4\t1\t1\tControl\tCrop 4\t20\t2.5\t50\t29\t6.20\t10\t22.7\t12.28\t15.19\tCrop 4\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695
5\t5\t1\tDeep Carbon-coated mineral (CCM)\tCrop 5\t20\t2.5\t50\t28\t6.13\t10.2\t22.8\t12.69\t13.28\tCrop 5\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110
6\t10\t1\tDeep OM (CP1) + PAM\tCrop 6\t20\t2.5\t50\t28\t7.27\t11.6\t23.4\t16.13\t15.20\tCrop 6\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
7\t9\t1\tSurface Silicon\tCrop 7\t20\t2.5\t50\t29\t6.78\t10.5\t23.4\t12.23\t15.29\tCrop 7\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835
8\t4\t1\tDeep OM + Gypsum (CP2)\tCrop 8\t20\t2.5\t50\t31\t7.60\t10.3\t25.2\t13.87\t14.46\tCrop 8\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885
9\t6\t1\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 9\t20\t2.5\t50\t31\t5.88\t10.3\t24.4\t14.19\t17.95\tCrop 9\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110
10\t7\t1\tDeep liq. NPKS\tCrop 10\t20\t2.5\t50\t25\t7.23\t11.5\t23.2\t12.12\t15.57\tCrop 10\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
11\t2\t1\tDeep Gypsum\tCrop 11\t20\t2.5\t50\t22\t6.29\t9.9\t22.8\t9.85\t15.45\tCrop 11\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385
12\t8\t1\tDeep liq. Gypsum (CHT)\tCrop 12\t20\t2.5\t50\t26\t5.88\t9.9\t23.5\t10.48\t11.69\tCrop 12\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235
13\t6\t2\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 13\t20\t2.5\t50\t33\t4.79\t9.8\t24.4\t14.49\t13.62\tCrop 13\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110
14\t7\t2\tDeep liq. NPKS\tCrop 14\t20\t2.5\t50\t29\t4.88\t10.4\t23.7\t12.81\t13.49\tCrop 14\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
15\t5\t2\tDeep Carbon-coated mineral (CCM)\tCrop 15\t20\t2.5\t50\t26\t5.39\t10.5\t23.7\t11.97\t12.77\tCrop 15\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110
16\t3\t2\tDeep OM (CP1)\tCrop 16\t20\t2.5\t50\t24\t4.96\t10.2\t23.2\t13.85\t14.44\tCrop 16\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385
17\t1\t2\tControl\tCrop 17\t20\t2.5\t50\t24\t4.99\t10.3\t23.3\t15.61\t10.63\tCrop 17\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695
18\t9\t2\tSurface Silicon\tCrop 18\t20\t2.5\t50\t27\t5.79\t10.6\t21.1\t8.59\t10.63\tCrop 18\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835
19\t4\t2\tDeep OM + Gypsum (CP2)\tCrop 19\t20\t2.5\t50\t22\t5.45\t11.2\t23\t12.34\t15.59\tCrop 19\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885
20\t11\t2\tDeep Ripping\tCrop 20\t20\t2.5\t50\t27\t6.30\t10.4\t22.9\t12.34\t15.28\tCrop 20\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
21\t8\t2\tDeep liq. Gypsum (CHT)\tCrop 21\t20\t2.5\t50\t24\t6.57\t9.8\t23.2\t16.16\t11.35\tCrop 21\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235
22\t12\t2\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 22\t20\t2.5\t50\t26\t6.10\t10.3\t23.6\t14.16\t12.21\tCrop 22\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735
23\t10\t2\tDeep OM (CP1) + PAM\tCrop 23\t20\t2.5\t50\t24\t6.34\t10\t22.8\t15.68\t12.70\tCrop 23\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
24\t2\t2\tDeep Gypsum\tCrop 24\t20\t2.5\t50\t25\t5.44\t9.8\t23.1\t12.70\t13.24\tCrop 24\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385
25\t6\t3\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 25\t20\t2.5\t50\t19\t5.04\t11.2\t24.5\t15.45\t10.97\tCrop 25\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110
26\t11\t3\tDeep Ripping\tCrop 26\t20\t2.5\t50\t21\t6.35\t11.2\t27.3\t19.73\t20.65\tCrop 26\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
27\t2\t3\tDeep Gypsum\tCrop 27\t20\t2.5\t50\t21\t6.94\t10.2\t24.6\t16.39\t14.52\tCrop 27\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385
28\t5\t3\tDeep Carbon-coated mineral (CCM)\tCrop 28\t20\t2.5\t50\t19\t6.31\t10.2\t23\t11.23\t15.58\tCrop 28\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110
29\t8\t3\tDeep liq. Gypsum (CHT)\tCrop 29\t20\t2.5\t50\t26\t6.64\t11.2\t23.5\t13.36\t14.23\tCrop 29\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235
30\t12\t3\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 30\t20\t2.5\t50\t22\t5.96\t10.4\t23.8\t12.01\t13.71\tCrop 30\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735
31\t10\t3\tDeep OM (CP1) + PAM\tCrop 31\t20\t2.5\t50\t22\t7.58\t10.2\t24.2\t12.73\t11.98\tCrop 31\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
32\t4\t3\tDeep OM + Gypsum (CP2)\tCrop 32\t20\t2.5\t50\t25\t6.68\t10.3\t24.6\t13.34\t13.12\tCrop 32\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885
33\t7\t3\tDeep liq. NPKS\tCrop 33\t20\t2.5\t50\t23\t7.33\t10.1\t23.3\t13.06\t12.18\tCrop 33\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
34\t1\t3\tControl\tCrop 34\t20\t2.5\t50\t25\t7.37\t10.3\t23.3\t15.30\t9.52\tCrop 34\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695
35\t3\t3\tDeep OM (CP1)\tCrop 35\t20\t2.5\t50\t23\t5.29\t10.5\t23.7\t12.61\t11.73\tCrop 35\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385
36\t9\t3\tSurface Silicon\tCrop 36\t20\t2.5\t50\t18\t6.81\t10\t23.8\t14.04\t17.68\tCrop 36\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835
37\t5\t4\tDeep Carbon-coated mineral (CCM)\tCrop 37\t20\t2.5\t50\t20\t6.42\t11.1\t23.4\t13.51\t13.34\tCrop 37\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110
38\t6\t4\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 38\t20\t2.5\t50\t20\t6.18\t10.6\t24.9\t14.50\t13.16\tCrop 38\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110
39\t9\t4\tSurface Silicon\tCrop 39\t20\t2.5\t50\t21\t6.69\t10.8\t24.6\t13.72\t15.00\tCrop 39\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835
40\t10\t4\tDeep OM (CP1) + PAM\tCrop 40\t20\t2.5\t50\t21\t7.72\t10.2\t23.3\t16.55\t18.02\tCrop 40\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
41\t11\t4\tDeep Ripping\tCrop 41\t20\t2.5\t50\t23\t6.28\t10.6\t23.4\t10.25\t14.71\tCrop 41\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
42\t2\t4\tDeep Gypsum\tCrop 42\t20\t2.5\t50\t19\t5.85\t9.8\t23.1\t10.66\t11.19\tCrop 42\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385
43\t7\t4\tDeep liq. NPKS\tCrop 43\t20\t2.5\t50\t23\t6.40\t10.1\t23.6\t13.28\t10.18\tCrop 43\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885
44\t4\t4\tDeep OM + Gypsum (CP2)\tCrop 44\t20\t2.5\t50\t33\t5.30\t9.7\t25.5\t16.80\t13.87\tCrop 44\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885
45\t1\t4\tControl\tCrop 45\t20\t2.5\t50\t24\t6.21\t9.9\t22.1\t10.02\t14.31\tCrop 45\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695
46\t3\t4\tDeep OM (CP1)\tCrop 46\t20\t2.5\t50\t28\t5.85\t10.9\t23.9\t13.05\t13.28\tCrop 46\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385
47\t8\t4\tDeep liq. Gypsum (CHT)\tCrop 47\t20\t2.5\t50\t27\t5.85\t9.6\t24.2\t20.66\t12.83\tCrop 47\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235
48\t12\t4\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 48\t20\t2.5\t50\t23\t6.06\t10\t25.1\t15.65\t11.32\tCrop 48\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735
`.trim();

  function makeUniqueHeaders(headers) {
    const seen = new Map();
    return headers.map(h => {
      const key = String(h || "").trim() || "col";
      const n = seen.get(key) || 0;
      seen.set(key, n + 1);
      return n === 0 ? key : `${key} (${n + 1})`;
    });
  }

  const FABA_2022_HEADERS = makeUniqueHeaders(FABA_2022_HEADERS_RAW);

  function parseTSVRows(headers, tsv) {
    const lines = String(tsv || "")
      .split(/\r?\n/)
      .map(s => s.trim())
      .filter(Boolean);

    const rows = [];
    for (const line of lines) {
      const parts = line.split("\t");
      const obj = {};
      for (let i = 0; i < headers.length; i++) {
        obj[headers[i]] = parts[i] !== undefined ? parts[i] : "";
      }
      rows.push(obj);
    }
    return rows;
  }

  // Keep the full raw dataset accessible for export + Excel workflow
  const defaultRaw = parseTSVRows(FABA_2022_HEADERS, FABA_2022_ROWS_TSV);

  // =========================
  // TREATMENT CALIBRATION FROM FULL DATASET
  // =========================

  function isLikelyLabourKey(k) {
    const s = String(k || "").toLowerCase();
    return s.includes("labour");
  }
  function isLikelySmallCost(v) {
    // Per-ha operational lines in the dataset are small ($0-$300-ish); capital assets are huge.
    return Number.isFinite(v) && Math.abs(v) <= 5000;
  }
  function isLikelyOperatingKey(k) {
    const s = String(k || "").toLowerCase();
    // chemicals + inputs + wet hire + machinery operations often appear as named products
    return (
      s.includes("fertil") ||
      s.includes("inocul") ||
      s.includes("herb") ||
      s.includes("fung") ||
      s.includes("insect") ||
      s.includes("roundup") ||
      s.includes("cavalier") ||
      s.includes("simazine") ||
      s.includes("veritas") ||
      s.includes("talstar") ||
      s.includes("mentor") ||
      s.includes("factor") ||
      s.includes("adjuvant") ||
      s.includes("wet hire") ||
      s.includes("boomspray") ||
      s.includes("seeder") ||
      s.includes("tiller") ||
      s.includes("tractor") ||
      s.includes("header") ||
      s.includes("ute") ||
      s.includes("truck") ||
      s.includes("treatment input cost")
    );
  }

  function computeTreatmentGroups(rawRows) {
    const groups = new Map();
    for (const row of rawRows) {
      const name = String(row["Amendment"] || "").trim();
      if (!name) continue;

      let g = groups.get(name);
      if (!g) {
        g = {
          name,
          n: 0,
          sumYield: 0,
          nYield: 0,
          sumArea: 0,
          sumLabour: 0,
          sumMaterials: 0,
          sumServices: 0,
          missingInputCostRows: 0,
          rows: 0
        };
        groups.set(name, g);
      }
      g.rows += 1;

      const y = parseNumber(row["Yield t/ha"]);
      if (!Number.isNaN(y)) {
        g.sumYield += y;
        g.nYield += 1;
      }

      const area_m2 = parseNumber(row["Plot Area (m^2)"]);
      if (!Number.isNaN(area_m2)) g.sumArea += area_m2;

      // Heuristic component aggregation from the full dataset (retains robustness across header variations)
      let labour = 0;
      let materials = 0;
      let services = 0;

      for (const k of Object.keys(row)) {
        const v = parseNumber(row[k]);
        if (!isLikelySmallCost(v)) continue;

        if (isLikelyLabourKey(k)) {
          labour += v;
          continue;
        }
        // Separate the explicit treatment input cost for transparency (often large but still per ha)
        if (String(k).toLowerCase().includes("treatment input cost")) {
          if (Number.isNaN(v)) g.missingInputCostRows += 1;
          else materials += v;
          continue;
        }
        if (isLikelyOperatingKey(k)) {
          // operational lines not labelled labour get treated as materials/services based on keyword
          const s = String(k).toLowerCase();
          if (
            s.includes("tractor") ||
            s.includes("tiller") ||
            s.includes("seeder") ||
            s.includes("boomspray") ||
            s.includes("boom spray") ||
            s.includes("header") ||
            s.includes("ute") ||
            (s.includes("truck") && !s.includes("asset"))
          ) {
            services += v;
          } else {
            materials += v;
          }
        }
      }

      g.sumLabour += labour;
      g.sumMaterials += materials;
      g.sumServices += services;
      g.n += 1;
    }

    const out = [];
    for (const [, g] of groups.entries()) {
      const meanYield = g.nYield ? g.sumYield / g.nYield : 0;
      const meanLabour = g.n ? g.sumLabour / g.n : 0;
      const meanMaterials = g.n ? g.sumMaterials / g.n : 0;
      const meanServices = g.n ? g.sumServices / g.n : 0;

      out.push({
        name: g.name,
        meanYieldTHa: meanYield,
        labourPerHa: meanLabour,
        materialsPerHa: meanMaterials,
        servicesPerHa: meanServices,
        rows: g.rows,
        missingInputCostRows: g.missingInputCostRows
      });
    }
    return out;
  }

  function applyDefaultDatasetCalibration() {
    const stats = computeTreatmentGroups(defaultRaw);
    if (!stats.length) {
      showToast("Default dataset could not be parsed.");
      return;
    }

    const control = stats.find(s => s.name.toLowerCase().includes("control"));
    const controlYield = control ? control.meanYieldTHa : null;

    const yieldOutput = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    const yieldId = yieldOutput ? yieldOutput.id : null;

    // Default area basis: 100 ha (as per sheet notes)
    const defaultAreaHa = 100;

    model.treatments = stats
      .slice()
      .sort((a, b) => (a.name.toLowerCase().includes("control") ? -1 : 1))
      .map(s => {
        const isControl = s.name.toLowerCase().includes("control");
        const t = {
          id: uid(),
          name: s.name,
          area: defaultAreaHa,
          adoption: 1,
          deltas: {},
          labourCost: s.labourPerHa || 0, // $/ha
          materialsCost: s.materialsPerHa || 0, // $/ha
          servicesCost: s.servicesPerHa || 0, // $/ha
          capitalCost: 0, // $ year 0 total (editable)
          constrained: true,
          source: "2022 Faba beans (default dataset)",
          isControl,
          notes: `Calibrated from ${s.rows} plot rows.${s.missingInputCostRows ? " Some input costs were missing ('?') and treated as 0." : ""}`
        };
        model.outputs.forEach(o => (t.deltas[o.id] = 0));

        if (yieldId && controlYield !== null && Number.isFinite(s.meanYieldTHa)) {
          // delta yield vs control (t/ha)
          t.deltas[yieldId] = isControl ? 0 : s.meanYieldTHa - controlYield;
        }
        return t;
      });

    initTreatmentDeltas();
    showToast("Default 2022 dataset loaded. Treatments calibrated vs control.");
  }

  // =========================
  // CORE ECONOMIC CALCULATIONS (TREATMENT vs CONTROL)
  // =========================

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

  function treatmentTotalCostPerHa(t) {
    return (Number(t.labourCost) || 0) + (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0);
  }

  // Optional: project-wide benefits/costs (kept, but applied only for non-control options by default)
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

  function otherCostsSeries(N) {
    const baseYear = model.time.startYear;
    const costByYear = new Array(N + 1).fill(0);
    model.otherCosts.forEach(c => {
      if (c.type === "annual") {
        const a = Number(c.annual) || 0;
        const sy = Number(c.startYear) || baseYear;
        const ey = Number(c.endYear) || sy;
        for (let y = sy; y <= ey; y++) {
          const idx = y - baseYear + 1;
          if (idx >= 1 && idx <= N) costByYear[idx] += a;
        }
      } else if (c.type === "capital") {
        const cap = Number(c.capital) || 0;
        const cy = Number(c.year) || baseYear;
        const idx = cy - baseYear;
        if (idx >= 0 && idx <= N) costByYear[idx] += cap;
      }
    });
    return costByYear;
  }

  // Core: compute option cashflows vs control, with control column shown (zeros)
  function computeOption(optionT, { ratePct, adoptMul, risk }) {
    const N = Number(model.time.years) || 10;
    const baseYear = Number(model.time.startYear) || new Date().getFullYear();
    const control = model.treatments.find(t => t.isControl) || null;

    const benefitByYear = new Array(N + 1).fill(0);
    const costByYear = new Array(N + 1).fill(0);

    // Control option: baseline in this comparison is 0 (incremental)
    if (!optionT || !control) {
      return {
        pvBenefits: 0,
        pvCosts: 0,
        npv: 0,
        bcr: NaN,
        roi: NaN,
        irrVal: NaN,
        mirrVal: NaN,
        paybackYears: null,
        cf: new Array(N + 1).fill(0),
        benefitByYear,
        costByYear
      };
    }

    const adopt = clamp(adoptMul, 0, 1);
    const r = clamp(risk, 0, 1);
    const area = Number(optionT.area) || 0;

    // Incremental benefit per ha from output deltas (treatment deltas are calibrated vs control for yield)
    let incValuePerHa = 0;
    model.outputs.forEach(o => {
      const delta = Number(optionT.deltas[o.id]) || 0;
      const v = Number(o.value) || 0;
      incValuePerHa += delta * v;
    });

    // Incremental annual benefits (year 1..N)
    const annualIncBenefit = incValuePerHa * area * adopt * (1 - r);

    for (let t = 1; t <= N; t++) benefitByYear[t] += annualIncBenefit;

    // Incremental costs per ha vs control (operating) + capital (year 0)
    const incCostPerHa = treatmentTotalCostPerHa(optionT) - treatmentTotalCostPerHa(control);
    const annualIncCost = incCostPerHa * area * adopt; // adoption scales implementation

    // Capital cost is year 0 cash outlay (total $). Compare vs control capital (usually 0).
    const incCapY0 = (Number(optionT.capitalCost) || 0) - (Number(control.capitalCost) || 0);

    costByYear[0] += incCapY0;
    for (let t = 1; t <= N; t++) costByYear[t] += annualIncCost;

    // Apply project-wide benefits/costs only if non-control option (implementation scenario)
    if (!optionT.isControl) {
      const extraB = additionalBenefitsSeries(N, baseYear, adoptMul, risk);
      for (let i = 0; i < extraB.length; i++) benefitByYear[i] += extraB[i];

      const extraC = otherCostsSeries(N);
      for (let i = 0; i < extraC.length; i++) costByYear[i] += extraC[i];
    }

    const cf = benefitByYear.map((b, i) => b - costByYear[i]);

    const pvBenefits = presentValue(benefitByYear, ratePct);
    const pvCosts = presentValue(costByYear, ratePct);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;
    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const pb = payback(cf, ratePct);

    return { pvBenefits, pvCosts, npv, bcr, roi, irrVal, mirrVal, paybackYears: pb, cf, benefitByYear, costByYear };
  }

  function computeComparison() {
    const ratePct = Number(model.time.discBase) || 7;
    const adoptMul = Number(model.adoption.base) || 1;
    const risk = Number(model.risk.base) || 0;

    const control = model.treatments.find(t => t.isControl) || null;
    const options = model.treatments.slice();

    const results = options.map(t => ({
      id: t.id,
      name: t.name,
      isControl: !!t.isControl,
      res: computeOption(t, { ratePct, adoptMul, risk })
    }));

    // Ranking by NPV (descending), excluding control
    const nonControl = results.filter(r => !r.isControl);
    nonControl.sort((a, b) => (b.res.npv || -Infinity) - (a.res.npv || -Infinity));
    const rankMap = new Map();
    nonControl.forEach((r, i) => rankMap.set(r.id, i + 1));

    results.forEach(r => {
      r.rank = r.isControl ? null : rankMap.get(r.id) || null;
      r.deltaYield = (() => {
        const yieldOut = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
        if (!yieldOut) return NaN;
        const t = model.treatments.find(x => x.id === r.id);
        return t ? Number(t.deltas[yieldOut.id]) || 0 : NaN;
      })();
      r.totalCostPerHa = (() => {
        const t = model.treatments.find(x => x.id === r.id);
        return t ? treatmentTotalCostPerHa(t) : NaN;
      })();
    });

    // Ensure control appears first in display
    results.sort((a, b) => (a.isControl ? -1 : b.isControl ? 1 : 0));

    return { results, controlName: control ? control.name : "Control" };
  }

  // =========================
  // DOM HELPERS
  // =========================
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  function ensureMount(containerCandidates, { title, className } = {}) {
    for (const sel of containerCandidates) {
      const el = typeof sel === "string" ? document.querySelector(sel) : sel;
      if (el) return el;
    }
    // fallback: create in results panel if possible
    const panel = document.querySelector("#tab-results, [data-tab-panel='results'], .tab-panel#results") || document.body;
    const box = document.createElement("div");
    if (className) box.className = className;
    if (title) {
      const h = document.createElement("h3");
      h.textContent = title;
      box.appendChild(h);
    }
    panel.appendChild(box);
    return box;
  }

  // =========================
  // UI: TOOLTIP + STYLING (injected, minimal but commercial clean)
  // =========================
  function injectStyle() {
    if (document.getElementById("fcdt2-style")) return;
    const css = `
      .fcdt2-card{background:#fff;border:1px solid rgba(0,0,0,.08);border-radius:14px;box-shadow:0 6px 18px rgba(0,0,0,.06);padding:14px;margin:12px 0}
      .fcdt2-row{display:flex;gap:10px;flex-wrap:wrap;align-items:stretch}
      .fcdt2-kpi{flex:1;min-width:210px;padding:12px;border-radius:14px;border:1px solid rgba(0,0,0,.08);background:rgba(0,0,0,.02)}
      .fcdt2-kpi .k{font-size:12px;opacity:.7;margin-bottom:6px}
      .fcdt2-kpi .v{font-size:20px;font-weight:700}
      .fcdt2-muted{opacity:.75}
      .fcdt2-tablewrap{overflow:auto;border:1px solid rgba(0,0,0,.08);border-radius:14px}
      table.fcdt2-table{border-collapse:separate;border-spacing:0;width:100%;min-width:900px;background:#fff}
      table.fcdt2-table th, table.fcdt2-table td{padding:10px 10px;border-bottom:1px solid rgba(0,0,0,.06);vertical-align:top;white-space:nowrap}
      table.fcdt2-table thead th{position:sticky;top:0;background:#fafafa;z-index:2;font-weight:700}
      table.fcdt2-table tbody tr:nth-child(even) td{background:rgba(0,0,0,.015)}
      table.fcdt2-table th:first-child, table.fcdt2-table td:first-child{position:sticky;left:0;background:#fff;z-index:1;border-right:1px solid rgba(0,0,0,.06);white-space:normal;min-width:240px}
      table.fcdt2-table thead th:first-child{z-index:3;background:#fafafa}
      .fcdt2-badge{display:inline-flex;align-items:center;gap:6px;padding:3px 8px;border-radius:999px;border:1px solid rgba(0,0,0,.14);font-size:12px}
      .fcdt2-badge.good{background:rgba(0,0,0,.03)}
      .fcdt2-badge.warn{background:rgba(0,0,0,.03)}
      .fcdt2-actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:10px 0}
      .fcdt2-btn{display:inline-flex;align-items:center;gap:8px;padding:8px 12px;border-radius:12px;border:1px solid rgba(0,0,0,.15);background:#fff;cursor:pointer}
      .fcdt2-btn:hover{background:rgba(0,0,0,.03)}
      .fcdt2-tip{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:999px;border:1px solid rgba(0,0,0,.18);font-size:12px;margin-left:8px;cursor:help;position:relative}
      .fcdt2-tip[data-tip]:hover:after{content:attr(data-tip);position:absolute;left:50%;top:110%;transform:translateX(-50%);background:#111;color:#fff;padding:8px 10px;border-radius:10px;max-width:320px;white-space:normal;z-index:9999;box-shadow:0 12px 30px rgba(0,0,0,.25)}
      .fcdt2-tip[data-tip]:hover:before{content:"";position:absolute;left:50%;top:102%;transform:translateX(-50%);border:7px solid transparent;border-bottom-color:#111}
      .fcdt2-note{font-size:13px;line-height:1.35;opacity:.82}
      .fcdt2-hr{height:1px;background:rgba(0,0,0,.08);margin:12px 0}
      .fcdt2-textarea{width:100%;min-height:220px;border-radius:12px;border:1px solid rgba(0,0,0,.15);padding:10px;font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;font-size:12px}
    `;
    const style = document.createElement("style");
    style.id = "fcdt2-style";
    style.textContent = css;
    document.head.appendChild(style);
  }

  // =========================
  // RESULTS RENDERING (Snapshot-first)
  // =========================
  const INDICATOR_TIPS = {
    ranking:
      "Ranked by Net Present Value (NPV), highest to lowest, excluding the control. Ranking is a summary, not a rule.",
    npv:
      "NPV is the present value of benefits minus costs over the chosen horizon, discounted at the selected rate. Positive NPV suggests the treatment outperforms control economically under current assumptions.",
    pvBenefits:
      "Present value of incremental benefits versus control over the horizon. For this tool, yield changes are monetised using the output values you set (e.g. grain price).",
    pvCosts:
      "Present value of incremental costs versus control over the horizon, including annual per-ha costs and any capital cost entered as a year 0 outlay.",
    bcr:
      "Benefit–cost ratio equals PV(Benefits) divided by PV(Costs). A higher BCR indicates more benefit per dollar of cost, but it must be interpreted alongside scale and NPV.",
    roi:
      "ROI equals (NPV / PV(Costs)) × 100. It expresses the net gain as a percent of the present value of costs.",
    irr:
      "IRR is the discount rate at which the net present value of incremental cashflows equals zero. It can be unstable if cashflows are irregular or near-zero.",
    mirr:
      "MIRR is a more conservative IRR variant that assumes financing and reinvestment rates (editable in Settings).",
    payback:
      "Discounted payback is the first year in which cumulative discounted net benefits become non-negative (if it occurs).",
    deltaYield:
      "Incremental yield (t/ha) relative to the control, estimated from the trial dataset and averaged across replicates."
  };

  function renderResults() {
    injectStyle();

    const mount = ensureMount(
      ["#resultsComparison", "#resultsTable", "#results", "#tab-results .results", "#tab-results"],
      { className: "fcdt2-card", title: "Results: Treatment comparison vs control" }
    );

    const { results, controlName } = computeComparison();

    const nonControl = results.filter(r => !r.isControl);
    const top3 = nonControl
      .slice()
      .sort((a, b) => (b.res.npv || -Infinity) - (a.res.npv || -Infinity))
      .slice(0, 3);

    const best = top3[0] || null;

    const rate = Number(model.time.discBase) || 7;
    const horizon = Number(model.time.years) || 10;

    const kpiHtml = `
      <div class="fcdt2-row">
        <div class="fcdt2-kpi">
          <div class="k">Discount rate (base)</div>
          <div class="v">${fmt(rate)}%</div>
          <div class="fcdt2-note">Horizon: ${horizon} years</div>
        </div>
        <div class="fcdt2-kpi">
          <div class="k">Top-ranked treatment (NPV)</div>
          <div class="v">${best ? esc(best.name) : "n/a"}</div>
          <div class="fcdt2-note">${best ? `NPV: ${money(best.res.npv)} | BCR: ${fmt(best.res.bcr)}` : ""}</div>
        </div>
        <div class="fcdt2-kpi">
          <div class="k">Control</div>
          <div class="v">${esc(controlName)}</div>
          <div class="fcdt2-note">All results are incremental vs control. Control column shows zeros.</div>
        </div>
      </div>
    `;

    const actionsHtml = `
      <div class="fcdt2-actions">
        <button class="fcdt2-btn" type="button" data-action="copy-results-tsv">Copy table (Word/Excel)</button>
        <button class="fcdt2-btn" type="button" data-action="export-results-xlsx">Export results (Excel)</button>
        <button class="fcdt2-btn" type="button" data-action="export-results-csv">Export results (CSV)</button>
        <button class="fcdt2-btn" type="button" data-action="jump-ai">AI interpretation prompt</button>
      </div>
    `;

    // Build comparison table: indicators as rows; treatments as columns (control + all)
    const cols = results.map(r => r);
    const colHeaders = cols
      .map(r => {
        const tag = r.isControl ? `<span class="fcdt2-badge">Control</span>` : `<span class="fcdt2-badge">Treatment</span>`;
        const rank = r.isControl
          ? `<span class="fcdt2-muted">Baseline</span>`
          : `<span class="fcdt2-badge good">Rank ${r.rank ?? "–"}</span>`;
        return `<th>${esc(r.name)}<div style="margin-top:6px;display:flex;gap:6px;flex-wrap:wrap">${tag}${rank}</div></th>`;
      })
      .join("");

    const row = (key, label, tipKey, valueFn) => {
      const tip = INDICATOR_TIPS[tipKey] || "";
      return `
        <tr>
          <td><strong>${esc(label)}</strong><span class="fcdt2-tip" data-tip="${esc(tip)}">i</span></td>
          ${cols
            .map(c => {
              const v = valueFn(c);
              return `<td>${v}</td>`;
            })
            .join("")}
        </tr>
      `;
    };

    const tableHtml = `
      <div class="fcdt2-tablewrap">
        <table class="fcdt2-table" aria-label="Treatment comparison table">
          <thead>
            <tr>
              <th>Economic indicator</th>
              ${colHeaders}
            </tr>
          </thead>
          <tbody>
            ${row(
              "deltaYield",
              "Incremental yield vs control (t/ha)",
              "deltaYield",
              c => (c.isControl ? "0.00" : fmt(c.deltaYield))
            )}
            ${row("pvBenefits", "Present value of benefits ($)", "pvBenefits", c => money(c.res.pvBenefits))}
            ${row("pvCosts", "Present value of costs ($)", "pvCosts", c => money(c.res.pvCosts))}
            ${row("npv", "Net present value (NPV) ($)", "npv", c => money(c.res.npv))}
            ${row("bcr", "Benefit–cost ratio (BCR)", "bcr", c => (isFinite(c.res.bcr) ? fmt(c.res.bcr) : "n/a"))}
            ${row("roi", "Return on investment (ROI)", "roi", c => (isFinite(c.res.roi) ? percent(c.res.roi) : "n/a"))}
            ${row("irr", "Internal rate of return (IRR)", "irr", c => (isFinite(c.res.irrVal) ? percent(c.res.irrVal) : "n/a"))}
            ${row("mirr", "Modified IRR (MIRR)", "mirr", c => (isFinite(c.res.mirrVal) ? percent(c.res.mirrVal) : "n/a"))}
            ${row(
              "payback",
              "Discounted payback (years)",
              "payback",
              c => (c.res.paybackYears === null ? "n/a" : String(c.res.paybackYears))
            )}
            ${row(
              "ranking",
              "Ranking (by NPV, excluding control)",
              "ranking",
              c => (c.isControl ? "Baseline" : (c.rank ?? "–"))
            )}
          </tbody>
        </table>
      </div>
      <div class="fcdt2-note" style="margin-top:10px">
        Notes: Costs are built from Labour + Materials + Services ($/ha) and scaled by treatment area and adoption. Capital cost is a year 0 outlay (total $) and is included in PV costs.
      </div>
    `;

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h2 style="margin:0 0 10px 0">${esc(TOOL_NAME)}</h2>
        ${kpiHtml}
        <div class="fcdt2-hr"></div>
        ${actionsHtml}
        ${tableHtml}
      </div>
    `;

    // wire buttons
    mount.querySelector("[data-action='copy-results-tsv']")?.addEventListener("click", () => {
      const tsv = buildResultsTSV(results);
      copyToClipboard(tsv);
      showToast("Results table copied (tab-separated). Paste into Word or Excel.");
    });

    mount.querySelector("[data-action='export-results-csv']")?.addEventListener("click", () => {
      const csv = buildResultsCSV(results);
      downloadFile(`${slug(TOOL_NAME)}_results_comparison.csv`, csv, "text/csv");
      showToast("Results CSV downloaded.");
    });

    mount.querySelector("[data-action='export-results-xlsx']")?.addEventListener("click", () => {
      exportResultsExcel(results);
    });

    mount.querySelector("[data-action='jump-ai']")?.addEventListener("click", () => {
      switchTab("ai");
      renderAIPrompt();
      showToast("AI interpretation prompt ready.");
    });
  }

  function buildResultsMatrix(results) {
    const indicators = [
      { key: "deltaYield", label: "Incremental yield vs control (t/ha)" },
      { key: "pvBenefits", label: "Present value of benefits ($)" },
      { key: "pvCosts", label: "Present value of costs ($)" },
      { key: "npv", label: "Net present value (NPV) ($)" },
      { key: "bcr", label: "Benefit–cost ratio (BCR)" },
      { key: "roi", label: "Return on investment (ROI)" },
      { key: "irr", label: "Internal rate of return (IRR)" },
      { key: "mirr", label: "Modified IRR (MIRR)" },
      { key: "payback", label: "Discounted payback (years)" },
      { key: "rank", label: "Ranking (by NPV, excluding control)" }
    ];

    const cols = results.map(r => r.name);
    const rows = indicators.map(ind => {
      const cells = results.map(r => {
        switch (ind.key) {
          case "deltaYield":
            return r.isControl ? 0 : (Number(r.deltaYield) || 0);
          case "pvBenefits":
            return r.res.pvBenefits;
          case "pvCosts":
            return r.res.pvCosts;
          case "npv":
            return r.res.npv;
          case "bcr":
            return r.res.bcr;
          case "roi":
            return r.res.roi;
          case "irr":
            return r.res.irrVal;
          case "mirr":
            return r.res.mirrVal;
          case "payback":
            return r.res.paybackYears === null ? "" : r.res.paybackYears;
          case "rank":
            return r.isControl ? "Baseline" : (r.rank ?? "");
          default:
            return "";
        }
      });
      return { label: ind.label, cells };
    });

    return { cols, rows };
  }

  function buildResultsCSV(results) {
    const { cols, rows } = buildResultsMatrix(results);
    const lines = [];
    lines.push(["Economic indicator", ...cols].map(csvCell).join(","));
    rows.forEach(r => {
      lines.push([r.label, ...r.cells].map(csvCell).join(","));
    });
    return lines.join("\n");
  }

  function buildResultsTSV(results) {
    const { cols, rows } = buildResultsMatrix(results);
    const lines = [];
    lines.push(["Economic indicator", ...cols].join("\t"));
    rows.forEach(r => {
      lines.push([r.label, ...r.cells].join("\t"));
    });
    return lines.join("\n");
  }

  function csvCell(x) {
    const s = x === null || x === undefined ? "" : String(x);
    if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (_) {
      // fallback
      const ta = document.createElement("textarea");
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
      return true;
    }
  }

  // =========================
  // EXPORT: Excel (XLSX if available)
  // =========================
  function exportResultsExcel(results) {
    if (!window.XLSX) {
      // fallback to CSV (Excel-readable)
      const csv = buildResultsCSV(results);
      downloadFile(`${slug(TOOL_NAME)}_results_comparison.csv`, csv, "text/csv");
      showToast("XLSX library not found. Exported CSV instead (Excel-readable).");
      return;
    }

    const { cols, rows } = buildResultsMatrix(results);

    const aoa = [];
    aoa.push([TOOL_NAME, "", "", "", `Version ${TOOL_VERSION}`]);
    aoa.push([`Export date: ${nowISO()}`]);
    aoa.push([]);
    aoa.push(["Economic indicator", ...cols]);

    rows.forEach(r => aoa.push([r.label, ...r.cells]));

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 38 }, ...cols.map(() => ({ wch: 20 }))];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results_Comparison");

    // Include treatments + raw data for transparency
    const trAoa = buildTreatmentsAOA();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(trAoa), "Treatments");

    const outAoa = buildOutputsAOA();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(outAoa), "Outputs");

    const rawAoa = buildRawAOA(defaultRaw, FABA_2022_HEADERS);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rawAoa), "Default_Raw_Data");

    const readme = buildReadmeAOA();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(readme), "README");

    const filename = `${slug(TOOL_NAME)}_results_${nowISO()}.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast("Results exported to Excel (XLSX).");
  }

  function buildTreatmentsAOA() {
    const aoa = [
      ["Name", "Control?", "Area (ha)", "Adoption (0-1)", "Capital cost ($, year 0)", "Labour ($/ha)", "Materials ($/ha)", "Services ($/ha)", "Total cost ($/ha)", "Yield delta vs control (t/ha)", "Notes"]
    ];
    const yOut = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    const yId = yOut ? yOut.id : null;

    model.treatments.forEach(t => {
      aoa.push([
        t.name,
        t.isControl ? "Yes" : "No",
        Number(t.area) || 0,
        Number(t.adoption) || 0,
        Number(t.capitalCost) || 0,
        Number(t.labourCost) || 0,
        Number(t.materialsCost) || 0,
        Number(t.servicesCost) || 0,
        treatmentTotalCostPerHa(t),
        yId ? (Number(t.deltas[yId]) || 0) : 0,
        t.notes || ""
      ]);
    });
    return aoa;
  }

  function buildOutputsAOA() {
    const aoa = [["Output", "Unit", "Value", "Source"]];
    model.outputs.forEach(o => aoa.push([o.name, o.unit, Number(o.value) || 0, o.source || ""]));
    return aoa;
  }

  function buildRawAOA(rows, headers) {
    const aoa = [headers.slice()];
    rows.forEach(r => {
      aoa.push(headers.map(h => r[h] ?? ""));
    });
    return aoa;
  }

  function buildReadmeAOA() {
    return [
      [TOOL_NAME],
      [`Version: ${TOOL_VERSION}`],
      [`Organisation: ${ORG}`],
      [],
      ["Excel-first workflow"],
      [
        "1) Download a template (or sample) from the tool. 2) Edit rows in Excel. 3) Save. 4) Upload back into the tool. The tool validates, parses, calibrates, and updates results."
      ],
      [],
      ["Interpretation"],
      [
        "The AI prompt tab generates a structured prompt that can be pasted into Copilot or ChatGPT to produce a plain-English interpretation of results. It is non-prescriptive and supports learning and improvement."
      ]
    ];
  }

  // =========================
  // EXCEL-FIRST WORKFLOW (Template + Import)
  // =========================
  let parsedExcel = null;

  function downloadExcelTemplate({ scenarioSpecific = true } = {}) {
    if (!window.XLSX) {
      // fallback template as CSV pack
      const csv = buildTreatmentsCSVTemplate(scenarioSpecific);
      downloadFile(`${slug(TOOL_NAME)}_template_treatments.csv`, csv, "text/csv");
      showToast("XLSX library not found. Downloaded CSV template instead.");
      return;
    }

    const wb = XLSX.utils.book_new();

    const trAoa = scenarioSpecific ? buildTreatmentsAOA() : buildTreatmentsBlankAOA();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(trAoa), "Treatments");

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildOutputsAOA()), "Outputs");

    const settings = [
      ["Setting", "Value"],
      ["Tool", TOOL_NAME],
      ["Version", TOOL_VERSION],
      ["Start year", model.time.startYear],
      ["Horizon (years)", model.time.years],
      ["Discount rate (base, %)", model.time.discBase],
      ["Adoption (base, 0-1)", model.adoption.base],
      ["Risk (base, 0-1)", model.risk.base],
      ["Notes", "Edit in Excel, then upload back into the tool."]
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(settings), "Settings");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildReadmeAOA()), "README");

    const filename = `${slug(TOOL_NAME)}_${scenarioSpecific ? "scenario" : "blank"}_template.xlsx`;
    XLSX.writeFile(wb, filename);
    showToast("Excel template downloaded.");
  }

  function buildTreatmentsBlankAOA() {
    return [
      ["Name", "Control?", "Area (ha)", "Adoption (0-1)", "Capital cost ($, year 0)", "Labour ($/ha)", "Materials ($/ha)", "Services ($/ha)", "Yield delta vs control (t/ha)", "Notes"],
      ["Control", "Yes", 100, 1, 0, 0, 0, 0, 0, "Baseline/control row (must be present)"],
      ["Treatment A", "No", 100, 1, 0, 0, 0, 0, 0.0, ""]
    ];
  }

  function buildTreatmentsCSVTemplate(scenarioSpecific) {
    const aoa = scenarioSpecific ? buildTreatmentsAOA() : buildTreatmentsBlankAOA();
    return aoa.map(row => row.map(csvCell).join(",")).join("\n");
  }

  async function handleParseExcel() {
    const fileInput = document.getElementById("excelFile") || document.getElementById("uploadExcel") || document.getElementById("excelUpload");
    if (!fileInput || !fileInput.files || !fileInput.files[0]) {
      alert("Please choose an Excel file first.");
      return;
    }
    const file = fileInput.files[0];

    // XLSX path
    if (window.XLSX) {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: "array" });
      parsedExcel = wb;
      const sheets = wb.SheetNames || [];
      showToast(`Excel parsed: ${sheets.join(", ")}`);
      renderExcelPreview(wb);
      return;
    }

    // CSV fallback
    const text = await file.text();
    parsedExcel = { csvText: text };
    showToast("CSV parsed (XLSX library not found).");
    renderExcelPreview(parsedExcel);
  }

  function renderExcelPreview(wb) {
    injectStyle();
    const mount = ensureMount(["#excelPreview", "#tab-excel", "[data-tab-panel='excel']"], {
      className: "fcdt2-card",
      title: "Excel import preview"
    });

    const isXlsx = !!(wb && wb.SheetNames);
    let html = `<div class="fcdt2-note">Upload summary</div>`;
    if (isXlsx) {
      html += `<div class="fcdt2-note">Sheets found: <strong>${esc(wb.SheetNames.join(", "))}</strong></div>`;
      html += `<div class="fcdt2-note">Next: click “Import/Apply” to calibrate the tool from the workbook.</div>`;
    } else {
      html += `<div class="fcdt2-note">CSV detected. Expected to match the Treatments template columns.</div>`;
    }

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h3 style="margin:0 0 10px 0">Excel-first workflow</h3>
        ${html}
        <div class="fcdt2-actions" style="margin-top:10px">
          <button class="fcdt2-btn" type="button" data-action="apply-excel">Import/Apply</button>
          <button class="fcdt2-btn" type="button" data-action="download-template">Download scenario template (XLSX/CSV)</button>
          <button class="fcdt2-btn" type="button" data-action="download-blank">Download blank template (XLSX/CSV)</button>
        </div>
      </div>
    `;

    mount.querySelector("[data-action='apply-excel']")?.addEventListener("click", () => commitExcelToModel());
    mount.querySelector("[data-action='download-template']")?.addEventListener("click", () => downloadExcelTemplate({ scenarioSpecific: true }));
    mount.querySelector("[data-action='download-blank']")?.addEventListener("click", () => downloadExcelTemplate({ scenarioSpecific: false }));
  }

  function commitExcelToModel() {
    if (!parsedExcel) {
      alert("No parsed Excel data found. Please parse an uploaded file first.");
      return;
    }

    try {
      if (parsedExcel.SheetNames && window.XLSX) {
        const ws = parsedExcel.Sheets["Treatments"] || parsedExcel.Sheets[parsedExcel.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

        const outWs = parsedExcel.Sheets["Outputs"];
        let outRows = [];
        if (outWs) outRows = XLSX.utils.sheet_to_json(outWs, { defval: "" });

        const validated = applyTreatmentsFromRows(rows);
        if (!validated.ok) {
          alert("Import validation failed:\n\n" + validated.issues.join("\n"));
          return;
        }

        if (outRows.length) applyOutputsFromRows(outRows);

        initTreatmentDeltas();
        renderAll();
        calcAndRender();
        showToast("Excel import applied. Results updated.");
        return;
      }

      // CSV fallback
      if (parsedExcel.csvText) {
        const lines = parsedExcel.csvText.split(/\r?\n/).filter(Boolean);
        const header = lines[0].split(",").map(s => s.replace(/^"|"$/g, "").trim());
        const rows = lines.slice(1).map(line => {
          const parts = splitCsvLine(line);
          const obj = {};
          header.forEach((h, i) => (obj[h] = parts[i] ?? ""));
          return obj;
        });
        const validated = applyTreatmentsFromRows(rows);
        if (!validated.ok) {
          alert("Import validation failed:\n\n" + validated.issues.join("\n"));
          return;
        }
        initTreatmentDeltas();
        renderAll();
        calcAndRender();
        showToast("CSV import applied. Results updated.");
        return;
      }
    } catch (err) {
      console.error(err);
      alert("Import failed. Please check the template format.");
    }
  }

  function splitCsvLine(line) {
    // minimal CSV parser
    const out = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
        continue;
      }
      if (ch === '"') {
        inQ = !inQ;
        continue;
      }
      if (ch === "," && !inQ) {
        out.push(cur.trim());
        cur = "";
        continue;
      }
      cur += ch;
    }
    out.push(cur.trim());
    return out;
  }

  function applyTreatmentsFromRows(rows) {
    const issues = [];
    if (!Array.isArray(rows) || rows.length < 2) {
      return { ok: false, issues: ["Treatments sheet appears empty or invalid."] };
    }

    // Accept both template variants
    const norm = k => String(k || "").trim().toLowerCase();
    const col = name => rows[0] && Object.keys(rows[0]).find(k => norm(k) === norm(name));

    const cName = col("Name") || col("Treatment") || col("Treatment name");
    const cControl = col("Control?") || col("Control");
    const cArea = col("Area (ha)") || col("Area");
    const cAdopt = col("Adoption (0-1)") || col("Adoption");
    const cCap = col("Capital cost ($, year 0)") || col("Capital cost") || col("Capital");
    const cLab = col("Labour ($/ha)") || col("Labour");
    const cMat = col("Materials ($/ha)") || col("Materials");
    const cServ = col("Services ($/ha)") || col("Services");
    const cYieldDelta = col("Yield delta vs control (t/ha)") || col("Yield delta");
    const cNotes = col("Notes") || col("Note");

    if (!cName || !cControl) {
      return { ok: false, issues: ["Treatments sheet must include at least: Name and Control? columns."] };
    }

    // Ensure there is exactly one control row
    const controlRows = rows.filter(r => String(r[cControl] || "").toLowerCase().startsWith("y") || String(r[cControl] || "").toLowerCase().includes("control") || String(r[cControl] || "").toLowerCase().startsWith("t") || String(r[cControl] || "").toLowerCase().startsWith("1"));
    if (controlRows.length !== 1) {
      issues.push("Treatments must contain exactly ONE control row (Control? = Yes).");
    }

    const yieldOut = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    if (!yieldOut) issues.push("Could not find a yield output to link yield deltas to.");

    const newTreatments = [];
    rows.forEach((r, idx) => {
      const name = String(r[cName] || "").trim();
      if (!name) return;

      const isControl =
        String(r[cControl] || "").toLowerCase().startsWith("y") ||
        String(r[cControl] || "").toLowerCase().includes("control") ||
        String(r[cControl] || "").toLowerCase().startsWith("t") ||
        String(r[cControl] || "").toLowerCase().startsWith("1");

      const area = cArea ? (parseNumber(r[cArea]) || 0) : 100;
      const adopt = cAdopt ? clamp(parseNumber(r[cAdopt]) || 1, 0, 1) : 1;

      const capital = cCap ? (parseNumber(r[cCap]) || 0) : 0;

      const lab = cLab ? (parseNumber(r[cLab]) || 0) : 0;
      const mat = cMat ? (parseNumber(r[cMat]) || 0) : 0;
      const serv = cServ ? (parseNumber(r[cServ]) || 0) : 0;

      const yd = cYieldDelta ? parseNumber(r[cYieldDelta]) : 0;
      const notes = cNotes ? String(r[cNotes] || "") : "";

      const t = {
        id: uid(),
        name,
        area,
        adoption: adopt,
        deltas: {},
        capitalCost: capital,
        labourCost: lab,
        materialsCost: mat,
        servicesCost: serv,
        constrained: true,
        source: "Excel import",
        isControl: !!isControl,
        notes
      };

      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      if (yieldOut && Number.isFinite(yd)) t.deltas[yieldOut.id] = isControl ? 0 : yd;

      if (!Number.isFinite(area) || area <= 0) issues.push(`Row ${idx + 2}: invalid area.`);
      if (!Number.isFinite(adopt) || adopt < 0 || adopt > 1) issues.push(`Row ${idx + 2}: invalid adoption.`);
      newTreatments.push(t);
    });

    if (!newTreatments.some(t => t.isControl)) issues.push("No control treatment detected after import.");

    if (issues.length) return { ok: false, issues };

    model.treatments = newTreatments;
    return { ok: true, issues: [] };
  }

  function applyOutputsFromRows(rows) {
    // Expected columns: Output, Unit, Value, Source
    const norm = k => String(k || "").trim().toLowerCase();
    const pick = name => rows[0] && Object.keys(rows[0]).find(k => norm(k) === norm(name));

    const cOut = pick("Output") || pick("Name");
    const cUnit = pick("Unit");
    const cVal = pick("Value");
    const cSrc = pick("Source");

    if (!cOut || !cVal) return;

    const byName = new Map(model.outputs.map(o => [o.name.toLowerCase(), o]));
    rows.forEach(r => {
      const name = String(r[cOut] || "").trim();
      if (!name) return;
      const val = parseNumber(r[cVal]);
      const existing = byName.get(name.toLowerCase());
      if (existing && Number.isFinite(val)) {
        existing.value = val;
        if (cUnit) existing.unit = String(r[cUnit] || existing.unit);
        if (cSrc) existing.source = String(r[cSrc] || existing.source);
      }
    });
  }

  // =========================
  // TREATMENTS TAB (dynamic totals; capital before total)
  // =========================
  function renderTreatments() {
    injectStyle();
    const mount = ensureMount(["#treatmentsTable", "#tab-treatments", "[data-tab-panel='treatments']"], {
      className: "fcdt2-card",
      title: "Treatments"
    });

    const yOut = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    const yId = yOut ? yOut.id : null;

    const rows = model.treatments
      .slice()
      .sort((a, b) => (a.isControl ? -1 : b.isControl ? 1 : 0))
      .map((t, i) => {
        const total = treatmentTotalCostPerHa(t);
        const yd = yId ? (Number(t.deltas[yId]) || 0) : 0;
        return `
          <tr>
            <td style="white-space:normal">
              <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
                <strong>${esc(t.name)}</strong>
                ${t.isControl ? `<span class="fcdt2-badge">Control</span>` : ``}
              </div>
              <div class="fcdt2-note">${esc(t.source || "")}</div>
            </td>
            <td><input data-tid="${t.id}" data-field="area" type="number" step="0.01" value="${esc(t.area)}" style="width:110px"></td>
            <td><input data-tid="${t.id}" data-field="adoption" type="number" step="0.01" min="0" max="1" value="${esc(t.adoption)}" style="width:110px"></td>
            <td><input data-tid="${t.id}" data-field="capitalCost" type="number" step="1" value="${esc(t.capitalCost)}" style="width:140px"></td>
            <td><input data-tid="${t.id}" data-field="labourCost" type="number" step="0.01" value="${esc(t.labourCost)}" style="width:120px"></td>
            <td><input data-tid="${t.id}" data-field="materialsCost" type="number" step="0.01" value="${esc(t.materialsCost)}" style="width:120px"></td>
            <td><input data-tid="${t.id}" data-field="servicesCost" type="number" step="0.01" value="${esc(t.servicesCost)}" style="width:120px"></td>
            <td><strong>${money(total)}</strong></td>
            <td>${t.isControl ? "0.00" : fmt(yd)}</td>
          </tr>
        `;
      })
      .join("");

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h3 style="margin:0 0 10px 0">Treatments (editable)</h3>
        <div class="fcdt2-note">
          Costs are per hectare per year. Capital cost is a total year 0 outlay (not per hectare) and is included in PV costs.
          The results table compares each treatment to the control.
        </div>
        <div class="fcdt2-hr"></div>
        <div class="fcdt2-tablewrap">
          <table class="fcdt2-table" aria-label="Treatments table">
            <thead>
              <tr>
                <th>Treatment</th>
                <th>Area (ha)</th>
                <th>Adoption</th>
                <th>Capital cost ($, year 0)</th>
                <th>Labour ($/ha)</th>
                <th>Materials ($/ha)</th>
                <th>Services ($/ha)</th>
                <th>Total cost ($/ha)</th>
                <th>Yield delta vs control (t/ha)</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
        <div class="fcdt2-actions" style="margin-top:10px">
          <button class="fcdt2-btn" type="button" data-action="add-treatment">Add treatment</button>
          <button class="fcdt2-btn" type="button" data-action="recalc">Recalculate results</button>
        </div>
      </div>
    `;

    mount.querySelector("[data-action='add-treatment']")?.addEventListener("click", () => {
      const t = {
        id: uid(),
        name: "New treatment",
        area: 100,
        adoption: 1,
        deltas: {},
        capitalCost: 0,
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        constrained: true,
        source: "User added",
        isControl: false,
        notes: ""
      };
      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      model.treatments.push(t);
      initTreatmentDeltas();
      renderTreatments();
      calcAndRender();
      showToast("New treatment added.");
    });

    mount.querySelector("[data-action='recalc']")?.addEventListener("click", () => {
      calcAndRender();
      showToast("Recalculated.");
    });

    mount.querySelectorAll("input[data-tid][data-field]")?.forEach(inp => {
      inp.addEventListener("input", e => {
        const el = e.target;
        const tid = el.getAttribute("data-tid");
        const field = el.getAttribute("data-field");
        const t = model.treatments.find(x => x.id === tid);
        if (!t) return;

        const v = parseNumber(el.value);
        if (field === "adoption") t[field] = clamp(Number.isFinite(v) ? v : 0, 0, 1);
        else t[field] = Number.isFinite(v) ? v : 0;

        calcAndRenderDebounced();
      });
    });
  }

  // =========================
  // OUTPUTS TAB (editable)
  // =========================
  function renderOutputs() {
    injectStyle();
    const mount = ensureMount(["#outputsTable", "#tab-outputs", "[data-tab-panel='outputs']"], {
      className: "fcdt2-card",
      title: "Outputs"
    });

    const rows = model.outputs
      .map(o => {
        return `
          <tr>
            <td><strong>${esc(o.name)}</strong><div class="fcdt2-note">${esc(o.unit)}</div></td>
            <td><input data-oid="${o.id}" data-field="value" type="number" step="0.01" value="${esc(o.value)}" style="width:160px"></td>
            <td>${esc(o.source || "")}</td>
          </tr>
        `;
      })
      .join("");

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h3 style="margin:0 0 10px 0">Outputs (values used to monetise changes)</h3>
        <div class="fcdt2-note">
          For grain yield, set the value to the grain price ($/t). Other outputs are optional and only matter if you use them in treatment deltas.
        </div>
        <div class="fcdt2-hr"></div>
        <div class="fcdt2-tablewrap">
          <table class="fcdt2-table">
            <thead><tr><th>Output</th><th>Value</th><th>Source</th></tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
        <div class="fcdt2-actions" style="margin-top:10px">
          <button class="fcdt2-btn" type="button" data-action="recalc">Recalculate results</button>
        </div>
      </div>
    `;

    mount.querySelector("[data-action='recalc']")?.addEventListener("click", () => {
      calcAndRender();
      showToast("Recalculated.");
    });

    mount.querySelectorAll("input[data-oid][data-field='value']")?.forEach(inp => {
      inp.addEventListener("input", e => {
        const el = e.target;
        const id = el.getAttribute("data-oid");
        const o = model.outputs.find(x => x.id === id);
        if (!o) return;
        const v = parseNumber(el.value);
        o.value = Number.isFinite(v) ? v : 0;
        calcAndRenderDebounced();
      });
    });
  }

  // =========================
  // AI PROMPT (non-prescriptive + improvement suggestions)
  // =========================
  function buildAIPromptText() {
    const { results } = computeComparison();
    const rate = Number(model.time.discBase) || 7;
    const horizon = Number(model.time.years) || 10;
    const adopt = Number(model.adoption.base) || 1;
    const risk = Number(model.risk.base) || 0;

    const nonControl = results.filter(r => !r.isControl).slice().sort((a, b) => (b.res.npv || -Infinity) - (a.res.npv || -Infinity));

    const tableMd = (() => {
      const headers = ["Indicator", ...results.map(r => r.name)];
      const lines = [];
      lines.push(`| ${headers.map(escMd).join(" | ")} |`);
      lines.push(`| ${headers.map(() => "---").join(" | ")} |`);

      const put = (label, fn) => {
        lines.push(`| ${escMd(label)} | ${results.map(fn).map(escMd).join(" | ")} |`);
      };

      put("Incremental yield vs control (t/ha)", r => (r.isControl ? "0.00" : fmt(r.deltaYield)));
      put("PV benefits ($)", r => money(r.res.pvBenefits));
      put("PV costs ($)", r => money(r.res.pvCosts));
      put("NPV ($)", r => money(r.res.npv));
      put("BCR", r => (isFinite(r.res.bcr) ? fmt(r.res.bcr) : "n/a"));
      put("ROI", r => (isFinite(r.res.roi) ? percent(r.res.roi) : "n/a"));
      put("Ranking (NPV)", r => (r.isControl ? "Baseline" : (r.rank ?? "–")));

      return lines.join("\n");
    })();

    const improvementRules = `
When BCR is low or NPV is negative, suggest realistic ways to improve performance without imposing any threshold rules:
1) reduce costs (labour, materials, services, machinery passes, logistics),
2) increase yield uplift (better timing, agronomic practice changes, improved establishment),
3) improve prices/quality (market timing, quality management, protein/screenings),
4) reduce risk exposure (sensitivity to yield, input prices, adoption feasibility),
5) explore scale effects (area, adoption, capital utilisation).
Always frame these as options for learning and scenario testing, not as decisions the tool dictates.
`.trim();

    const prompt = `
You are assisting a farmer and an applied agricultural economist to interpret a cost–benefit analysis (CBA) produced by "${TOOL_NAME}". Do not prescribe a decision. Explain results in plain English, highlight what drives each treatment’s performance relative to the control, and suggest practical improvement options for underperforming treatments.

Context and settings
Tool: ${TOOL_NAME} (version ${TOOL_VERSION})
Horizon: ${horizon} years
Discount rate (base): ${rate}%
Adoption (base): ${fmt(adopt)}
Risk factor (base): ${fmt(risk)}
Interpretation focus: incremental results vs the control (control column is baseline).

Key indicators (explain briefly)
NPV: present value of benefits minus costs over the horizon.
PV benefits: present value of incremental benefits vs control.
PV costs: present value of incremental costs vs control including any year 0 capital outlay.
BCR: PV benefits divided by PV costs (if PV costs > 0).
ROI: NPV divided by PV costs (percent).
Ranking: ordered by NPV (excluding control).

Results table (incremental vs control)
${tableMd}

Tasks
1) Summarise which treatments perform better or worse economically and why, using NPV, PV benefits, PV costs, BCR, and ROI together.
2) Identify the main drivers of performance for the top two treatments and the bottom two treatments.
3) Discuss trade-offs: a treatment may have a high BCR but low scale, or high NPV with higher costs. Explain those patterns.
4) For any treatment with low BCR or negative NPV, propose realistic, practical ways a farmer could improve the economics under plausible constraints.

${improvementRules}

Output style
Write clearly for a non-technical farmer. Use short paragraphs and simple explanations. Avoid telling the user what they must do. End with a short checklist of “things to test next” in the tool (e.g., grain price, cost components, yield uplift, risk, adoption, horizon).
`.trim();

    return prompt;
  }

  function escMd(s) {
    return String(s ?? "").replace(/\|/g, "\\|");
  }

  function renderAIPrompt() {
    injectStyle();
    const mount = ensureMount(["#aiPrompt", "#tab-ai", "[data-tab-panel='ai']"], {
      className: "fcdt2-card",
      title: "AI interpretation prompt"
    });

    const prompt = buildAIPromptText();

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h3 style="margin:0 0 10px 0">AI-assisted interpretation (copy/paste to Copilot or ChatGPT)</h3>
        <div class="fcdt2-note">
          This prompt is designed to produce a plain-English interpretation of the CBA results, explain what the indicators mean, highlight trade-offs, and suggest practical improvement options for underperforming treatments. It does not dictate decisions.
        </div>
        <div class="fcdt2-hr"></div>
        <div class="fcdt2-actions">
          <button class="fcdt2-btn" type="button" data-action="copy-ai">Copy prompt</button>
          <button class="fcdt2-btn" type="button" data-action="download-ai">Download prompt (.txt)</button>
        </div>
        <textarea class="fcdt2-textarea" id="aiPromptText">${esc(prompt)}</textarea>
      </div>
    `;

    mount.querySelector("[data-action='copy-ai']")?.addEventListener("click", async () => {
      await copyToClipboard(prompt);
      showToast("AI prompt copied to clipboard.");
    });

    mount.querySelector("[data-action='download-ai']")?.addEventListener("click", () => {
      downloadFile(`${slug(TOOL_NAME)}_ai_prompt_${nowISO()}.txt`, prompt, "text/plain");
      showToast("AI prompt downloaded.");
    });
  }

  // =========================
  // CSV EXPORT (full pack)
  // =========================
  function exportAllCsv() {
    const { results } = computeComparison();
    const files = [];

    files.push({
      name: `${slug(TOOL_NAME)}_results_comparison.csv`,
      content: buildResultsCSV(results),
      mime: "text/csv"
    });

    files.push({
      name: `${slug(TOOL_NAME)}_treatments.csv`,
      content: buildTreatmentsAOA().map(r => r.map(csvCell).join(",")).join("\n"),
      mime: "text/csv"
    });

    files.push({
      name: `${slug(TOOL_NAME)}_outputs.csv`,
      content: buildOutputsAOA().map(r => r.map(csvCell).join(",")).join("\n"),
      mime: "text/csv"
    });

    files.push({
      name: `${slug(TOOL_NAME)}_default_raw_data.csv`,
      content: buildRawAOA(defaultRaw, FABA_2022_HEADERS).map(r => r.map(csvCell).join(",")).join("\n"),
      mime: "text/csv"
    });

    // deliver as multiple downloads (robust; avoids needing zip libs)
    files.forEach(f => downloadFile(f.name, f.content, f.mime));
    showToast("CSV exports downloaded.");
  }

  // =========================
  // PDF EXPORT (print-friendly)
  // =========================
  function exportPdf() {
    // For clean printable output, switch to results tab first
    try {
      switchTab("results");
      renderResults();
    } catch (_) {}
    setTimeout(() => window.print(), 150);
  }

  // =========================
  // SIMULATION (robust, optional; snapshot results remain primary)
  // =========================
  function runSimulation() {
    const N = Math.max(100, Number(model.sim.n) || 1000);
    const seed = model.sim.seed ? Number(model.sim.seed) : null;
    const R = rng(seed || undefined);

    const ratePct = Number(model.time.discBase) || 7;
    const adoptMul = Number(model.adoption.base) || 1;
    const risk = Number(model.risk.base) || 0;
    const varPct = clamp(Number(model.sim.variationPct) || 20, 0, 200) / 100;

    // baseline values (to restore)
    const baseOutputs = model.outputs.map(o => ({ id: o.id, value: o.value }));
    const baseTreats = model.treatments.map(t => ({
      id: t.id,
      labourCost: t.labourCost,
      materialsCost: t.materialsCost,
      servicesCost: t.servicesCost
    }));

    const { results: baseResults } = computeComparison();
    const nonControl = baseResults.filter(r => !r.isControl);
    const primary = nonControl.slice().sort((a, b) => (b.res.npv || -Infinity) - (a.res.npv || -Infinity))[0];

    if (!primary) {
      showToast("Simulation needs at least one non-control treatment.");
      return;
    }

    const targetId = primary.id;

    const draws = [];
    for (let i = 0; i < N; i++) {
      // vary outputs
      if (model.sim.varyOutputs) {
        model.outputs.forEach(o => {
          const base = baseOutputs.find(x => x.id === o.id)?.value ?? o.value;
          const a = base * (1 - varPct);
          const b = base * (1 + varPct);
          const c = base;
          o.value = triangular(R(), a, c, b);
        });
      }

      // vary treatment costs
      if (model.sim.varyTreatCosts) {
        model.treatments.forEach(t => {
          const base = baseTreats.find(x => x.id === t.id);
          if (!base) return;
          const f = triangular(R(), 1 - varPct, 1, 1 + varPct);
          t.labourCost = (Number(base.labourCost) || 0) * f;
          t.materialsCost = (Number(base.materialsCost) || 0) * f;
          t.servicesCost = (Number(base.servicesCost) || 0) * f;
        });
      }

      const { results } = computeComparison();
      const row = results.find(r => r.id === targetId);
      if (row) {
        draws.push({
          npv: row.res.npv,
          bcr: row.res.bcr
        });
      }
    }

    // restore
    baseOutputs.forEach(b => {
      const o = model.outputs.find(x => x.id === b.id);
      if (o) o.value = b.value;
    });
    baseTreats.forEach(b => {
      const t = model.treatments.find(x => x.id === b.id);
      if (t) {
        t.labourCost = b.labourCost;
        t.materialsCost = b.materialsCost;
        t.servicesCost = b.servicesCost;
      }
    });

    model.sim.results = draws;

    renderSimulationSummary(primary.name, draws);
    showToast("Simulation completed.");
  }

  function renderSimulationSummary(treatmentName, draws) {
    injectStyle();
    const mount = ensureMount(["#simOut", "#tab-sim", "[data-tab-panel='sim']", "#simulationResults"], {
      className: "fcdt2-card",
      title: "Simulation"
    });

    const npvs = draws.map(d => d.npv).filter(Number.isFinite).sort((a, b) => a - b);
    const bcrs = draws.map(d => d.bcr).filter(Number.isFinite).sort((a, b) => a - b);

    const q = (arr, p) => {
      if (!arr.length) return NaN;
      const idx = Math.floor((arr.length - 1) * p);
      return arr[idx];
    };

    const pPos = npvs.length ? npvs.filter(x => x > 0).length / npvs.length : NaN;
    const pBcrGt1 = bcrs.length ? bcrs.filter(x => x > 1).length / bcrs.length : NaN;

    mount.innerHTML = `
      <div class="fcdt2-card">
        <h3 style="margin:0 0 10px 0">Uncertainty simulation (optional)</h3>
        <div class="fcdt2-note">
          This simulation varies output values and treatment costs around the base case. It is a learning aid, not a decision rule.
          The summary shown is for the current top-ranked treatment: <strong>${esc(treatmentName)}</strong>.
        </div>
        <div class="fcdt2-hr"></div>
        <div class="fcdt2-row">
          <div class="fcdt2-kpi"><div class="k">P(NPV &gt; 0)</div><div class="v">${isFinite(pPos) ? percent(pPos * 100) : "n/a"}</div></div>
          <div class="fcdt2-kpi"><div class="k">P(BCR &gt; 1)</div><div class="v">${isFinite(pBcrGt1) ? percent(pBcrGt1 * 100) : "n/a"}</div></div>
          <div class="fcdt2-kpi"><div class="k">NPV (median)</div><div class="v">${isFinite(q(npvs, 0.5)) ? money(q(npvs, 0.5)) : "n/a"}</div></div>
        </div>
        <div class="fcdt2-hr"></div>
        <div class="fcdt2-note">
          NPV quantiles: p10=${isFinite(q(npvs, 0.1)) ? money(q(npvs, 0.1)) : "n/a"}, p50=${isFinite(q(npvs, 0.5)) ? money(q(npvs, 0.5)) : "n/a"}, p90=${isFinite(q(npvs, 0.9)) ? money(q(npvs, 0.9)) : "n/a"}.
        </div>
      </div>
    `;
  }

  // =========================
  // BASIC SETTINGS BINDING (keeps existing IDs; non-breaking if missing)
  // =========================
  function setBasicsFieldsFromModel() {
    const set = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.value = v;
    };
    set("projectName", model.project.name || "");
    set("projectLead", model.project.lead || "");
    set("analystNames", model.project.analysts || "");
    set("projectTeam", model.project.team || "");
    set("projectSummary", model.project.summary || "");
    set("projectObjectives", model.project.objectives || "");
    set("projectActivities", model.project.activities || "");
    set("stakeholderGroups", model.project.stakeholders || "");
    set("lastUpdated", model.project.lastUpdated || "");
    set("projectGoal", model.project.goal || "");
    set("withProject", model.project.withProject || "");
    set("withoutProject", model.project.withoutProject || "");
    set("organisation", model.project.organisation || "");
    set("contactEmail", model.project.contactEmail || "");
    set("contactPhone", model.project.contactPhone || "");

    set("startYear", model.time.startYear);
    set("projectStartYear", model.time.projectStartYear || model.time.startYear);
    set("years", model.time.years);
    set("discBase", model.time.discBase);
    set("discLow", model.time.discLow);
    set("discHigh", model.time.discHigh);
    set("mirrFinance", model.time.mirrFinance);
    set("mirrReinvest", model.time.mirrReinvest);

    set("adoptBase", model.adoption.base);
    set("adoptLow", model.adoption.low);
    set("adoptHigh", model.adoption.high);

    set("riskBase", model.risk.base);
    set("riskLow", model.risk.low);
    set("riskHigh", model.risk.high);
    set("rTech", model.risk.tech);
    set("rNonCoop", model.risk.nonCoop);
    set("rSocio", model.risk.socio);
    set("rFin", model.risk.fin);
    set("rMan", model.risk.man);

    set("simN", model.sim.n);
    set("targetBCR", model.sim.targetBCR);
    set("bcrMode", model.sim.bcrMode);
    const lab = document.getElementById("simBcrTargetLabel");
    if (lab) lab.textContent = String(model.sim.targetBCR);

    set("simVarPct", String(model.sim.variationPct || 20));
    const vOut = document.getElementById("simVaryOutputs");
    if (vOut) vOut.value = model.sim.varyOutputs ? "true" : "false";
    const vTr = document.getElementById("simVaryTreatCosts");
    if (vTr) vTr.value = model.sim.varyTreatCosts ? "true" : "false";
    const vIn = document.getElementById("simVaryInputCosts");
    if (vIn) vIn.value = model.sim.varyInputCosts ? "true" : "false";

    const st = document.getElementById("systemType");
    if (st) st.value = model.outputsMeta.systemType || "single";
    const oa = document.getElementById("outputAssumptions");
    if (oa) oa.value = model.outputsMeta.assumptions || "";
  }

  // =========================
  // TABS (robust; non-breaking)
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

    try {
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (_) {}
  }

  function initTabs() {
    document.addEventListener("click", e => {
      const el = e.target.closest("[data-tab],[data-tab-target],[data-tab-jump]");
      if (!el) return;
      const target = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      if (!target) return;
      e.preventDefault();
      switchTab(target);
      // lazy render key tabs
      if (target === "results") renderResults();
      if (target === "treatments") renderTreatments();
      if (target === "outputs") renderOutputs();
      if (target === "ai") renderAIPrompt();
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
  }

  // =========================
  // ACTIONS (keep existing IDs)
  // =========================
  function initActions() {
    document.addEventListener("click", e => {
      const el = e.target.closest("#recalc, #getResults, [data-action='recalc']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      calcAndRender();
      showToast("Economic indicators recalculated.");
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#runSim, [data-action='run-sim']");
      if (!el) return;
      e.preventDefault();
      e.stopPropagation();
      runSimulation();
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#exportCsv, #exportCsvFoot");
      if (!el) return;
      e.preventDefault();
      exportAllCsv();
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#exportPdf, #exportPdfFoot");
      if (!el) return;
      e.preventDefault();
      exportPdf();
      showToast("Print dialog opened for PDF export.");
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#downloadTemplate");
      if (!el) return;
      e.preventDefault();
      downloadExcelTemplate({ scenarioSpecific: true });
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#downloadSample");
      if (!el) return;
      e.preventDefault();
      // sample = scenario template based on default dataset
      downloadExcelTemplate({ scenarioSpecific: true });
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#parseExcel");
      if (!el) return;
      e.preventDefault();
      handleParseExcel();
    });

    document.addEventListener("click", e => {
      const el = e.target.closest("#importExcel");
      if (!el) return;
      e.preventDefault();
      commitExcelToModel();
    });

    // Combined risk button
    const calcRiskBtn = $("#calcCombinedRisk");
    if (calcRiskBtn) {
      calcRiskBtn.addEventListener("click", e => {
        e.stopPropagation();
        const r =
          1 -
          (1 - (Number($("#rTech")?.value) || model.risk.tech)) *
            (1 - (Number($("#rNonCoop")?.value) || model.risk.nonCoop)) *
            (1 - (Number($("#rSocio")?.value) || model.risk.socio)) *
            (1 - (Number($("#rFin")?.value) || model.risk.fin)) *
            (1 - (Number($("#rMan")?.value) || model.risk.man));
        const out = $("#combinedRiskOut");
        if (out) out.textContent = "Combined: " + (r * 100).toFixed(2) + "%";
        const rb = $("#riskBase");
        if (rb) rb.value = r.toFixed(3);
        model.risk.base = r;
        calcAndRender();
        showToast("Combined risk updated.");
      });
    }

    // Save/load JSON project buttons (keep compatibility)
    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.stopPropagation();
        const data = JSON.stringify(model, null, 2);
        downloadFile(`cba_${slug(model.project.name || TOOL_NAME)}.json`, data, "application/json");
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
          // preserve tool identity but accept loaded parameters
          Object.assign(model, obj);
          model.meta = model.meta || {};
          model.meta.toolName = TOOL_NAME;
          model.meta.version = TOOL_VERSION;

          if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          if (!Array.isArray(model.outputs)) model.outputs = [];
          if (!Array.isArray(model.treatments)) model.treatments = [];
          initTreatmentDeltas();
          setBasicsFieldsFromModel();
          renderAll();
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

    // Global input listener for settings fields (non-breaking)
    document.addEventListener("input", e => {
      const t = e.target;
      if (!t || !t.id) return;
      const id = t.id;

      // project
      if (id === "projectName") model.project.name = t.value;
      if (id === "projectLead") model.project.lead = t.value;
      if (id === "analystNames") model.project.analysts = t.value;
      if (id === "projectTeam") model.project.team = t.value;
      if (id === "projectSummary") model.project.summary = t.value;
      if (id === "projectObjectives") model.project.objectives = t.value;
      if (id === "projectActivities") model.project.activities = t.value;
      if (id === "stakeholderGroups") model.project.stakeholders = t.value;
      if (id === "lastUpdated") model.project.lastUpdated = t.value;
      if (id === "projectGoal") model.project.goal = t.value;
      if (id === "withProject") model.project.withProject = t.value;
      if (id === "withoutProject") model.project.withoutProject = t.value;
      if (id === "organisation") model.project.organisation = t.value;
      if (id === "contactEmail") model.project.contactEmail = t.value;
      if (id === "contactPhone") model.project.contactPhone = t.value;

      // time
      if (id === "startYear") model.time.startYear = Number(t.value) || model.time.startYear;
      if (id === "projectStartYear") model.time.projectStartYear = Number(t.value) || model.time.projectStartYear;
      if (id === "years") model.time.years = Number(t.value) || model.time.years;
      if (id === "discBase") model.time.discBase = Number(t.value) || model.time.discBase;
      if (id === "discLow") model.time.discLow = Number(t.value) || model.time.discLow;
      if (id === "discHigh") model.time.discHigh = Number(t.value) || model.time.discHigh;
      if (id === "mirrFinance") model.time.mirrFinance = Number(t.value) || model.time.mirrFinance;
      if (id === "mirrReinvest") model.time.mirrReinvest = Number(t.value) || model.time.mirrReinvest;

      // adoption/risk
      if (id === "adoptBase") model.adoption.base = Number(t.value) || model.adoption.base;
      if (id === "adoptLow") model.adoption.low = Number(t.value) || model.adoption.low;
      if (id === "adoptHigh") model.adoption.high = Number(t.value) || model.adoption.high;

      if (id === "riskBase") model.risk.base = Number(t.value) || model.risk.base;
      if (id === "riskLow") model.risk.low = Number(t.value) || model.risk.low;
      if (id === "riskHigh") model.risk.high = Number(t.value) || model.risk.high;
      if (id === "rTech") model.risk.tech = Number(t.value) || model.risk.tech;
      if (id === "rNonCoop") model.risk.nonCoop = Number(t.value) || model.risk.nonCoop;
      if (id === "rSocio") model.risk.socio = Number(t.value) || model.risk.socio;
      if (id === "rFin") model.risk.fin = Number(t.value) || model.risk.fin;
      if (id === "rMan") model.risk.man = Number(t.value) || model.risk.man;

      // sim
      if (id === "simN") model.sim.n = Number(t.value) || model.sim.n;
      if (id === "targetBCR") model.sim.targetBCR = Number(t.value) || model.sim.targetBCR;
      if (id === "bcrMode") model.sim.bcrMode = t.value;
      if (id === "randSeed") model.sim.seed = t.value ? Number(t.value) : null;
      if (id === "simVarPct") model.sim.variationPct = Number(t.value) || model.sim.variationPct;

      // outputs meta
      if (id === "systemType") model.outputsMeta.systemType = t.value;
      if (id === "outputAssumptions") model.outputsMeta.assumptions = t.value;

      calcAndRenderDebounced();
    });
  }

  const calcAndRenderDebounced = debounce(() => calcAndRender(), 250);

  function calcAndRender() {
    // Snapshot-first outputs: results, AI prompt (if visible), and any existing summary fields
    try {
      renderResults();
    } catch (_) {}

    // Update any legacy summary fields if present
    try {
      const { results } = computeComparison();
      const best = results.filter(r => !r.isControl).slice().sort((a, b) => (b.res.npv || -Infinity) - (a.res.npv || -Infinity))[0];
      if ($("#bestTreatment")) $("#bestTreatment").textContent = best ? best.name : "n/a";
      if ($("#bestNPV")) $("#bestNPV").textContent = best ? money(best.res.npv) : "n/a";
      if ($("#bestBCR")) $("#bestBCR").textContent = best && isFinite(best.res.bcr) ? fmt(best.res.bcr) : "n/a";
    } catch (_) {}

    // If AI tab active, refresh prompt in-place
    const aiPanel = document.querySelector("#tab-ai, [data-tab-panel='ai']");
    if (aiPanel && (aiPanel.classList.contains("active") || aiPanel.classList.contains("show"))) {
      try {
        renderAIPrompt();
      } catch (_) {}
    }
  }

  function renderAll() {
    // Render the main functional tabs if their containers exist
    try { renderTreatments(); } catch (_) {}
    try { renderOutputs(); } catch (_) {}
    try { renderResults(); } catch (_) {}
    try { renderAIPrompt(); } catch (_) {}
  }

  // =========================
  // INIT (DEFAULT DATASET + WIRING)
  // =========================
  function init() {
    injectStyle();
    applyDefaultDatasetCalibration();
    setBasicsFieldsFromModel();
    initTabs();
    initActions();
    renderAll();
    calcAndRender();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
