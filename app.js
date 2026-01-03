/* Farming CBA Decision Tool 2 (commercial-grade)
   Newcastle Business School, The University of Newcastle

   Upgrade focus:
   1) Results redesigned as a vertical comparison table: indicators as rows, treatments as columns, with Control included.
   2) Excel-first workflow: download template (sample or scenario-specific), edit in Excel, upload, auto-validate, auto-calibrate.
   3) Snapshot usability: headline metrics first, optional detail second.
   4) AI-assisted interpretation prompt: flexible, non-prescriptive; includes improvement ideas for low BCR.
   5) Decision support: never hide low-performing treatments; help users explore drivers.
   6) Exports: clean Excel; printable PDF; copy/paste to Word.
   7) Naming consistency: “Farming CBA Decision Tool 2” everywhere.
   8) Treatments tab: Capital cost ($, year 0) appears before Total cost ($/ha); totals update dynamically.

   Notes:
   - This script is robust to missing DOM elements: it will no-op where elements are not present.
   - XLSX (SheetJS) is used if present for Excel import/export. If not present, CSV fallbacks are provided.
*/

(() => {
  "use strict";

  // -----------------------------
  // BRAND / STORAGE
  // -----------------------------
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const LS_KEY = "farming_cba_tool2_model_v1";

  // -----------------------------
  // CONSTANTS
  // -----------------------------
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

  const horizons = [5, 10, 15, 20, 25];

  // -----------------------------
  // DOM HELPERS
  // -----------------------------
  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  function setText(sel, txt) {
    const el = typeof sel === "string" ? $(sel) : sel;
    if (el) el.textContent = txt;
  }

  function setHTML(sel, html) {
    const el = typeof sel === "string" ? $(sel) : sel;
    if (el) el.innerHTML = html;
  }

  function esc(s) {
    return (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function clamp(v, a, b) {
    return Math.max(a, Math.min(b, v));
  }

  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

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

  const fmt = n =>
    isFinite(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 2 })
      : "n/a";

  const money = n => (isFinite(n) ? "$" + fmt(n) : "n/a");
  const pct = n => (isFinite(n) ? fmt(n) + "%" : "n/a");

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
    }, 3000);
  }

  // -----------------------------
  // SMALL CSS INJECTION (table + tooltips + print)
  // -----------------------------
  function injectTool2CSS() {
    if ($("#tool2-css")) return;
    const css = document.createElement("style");
    css.id = "tool2-css";
    css.textContent = `
      .tool2-badge{display:inline-flex;align-items:center;gap:.5rem}
      .tool2-badge .dot{width:.6rem;height:.6rem;border-radius:999px;background:currentColor;opacity:.35}
      .tool2-kpis{display:flex;flex-wrap:wrap;gap:.75rem;margin:.75rem 0 1rem}
      .tool2-kpi{background:rgba(0,0,0,.04);border:1px solid rgba(0,0,0,.08);border-radius:12px;padding:.6rem .75rem;min-width:180px}
      .tool2-kpi .lbl{font-size:.85rem;opacity:.75;margin-bottom:.15rem}
      .tool2-kpi .val{font-size:1.05rem;font-weight:700}
      .tool2-kpi .sub{font-size:.8rem;opacity:.75;margin-top:.1rem}
      .tool2-table-wrap{overflow:auto;border-radius:14px;border:1px solid rgba(0,0,0,.10);background:#fff}
      table.tool2-compare{border-collapse:separate;border-spacing:0;width:max-content;min-width:100%}
      table.tool2-compare th, table.tool2-compare td{padding:.55rem .65rem;border-bottom:1px solid rgba(0,0,0,.08);border-right:1px solid rgba(0,0,0,.06);vertical-align:middle;white-space:nowrap}
      table.tool2-compare th{position:sticky;top:0;background:#fafafa;z-index:3;font-weight:700}
      table.tool2-compare th:first-child{left:0;z-index:4}
      table.tool2-compare td:first-child{position:sticky;left:0;background:#fff;z-index:2;font-weight:600}
      table.tool2-compare tr:last-child td{border-bottom:none}
      table.tool2-compare th:last-child, table.tool2-compare td:last-child{border-right:none}
      .tool2-muted{opacity:.7;font-weight:500}
      .tool2-cell-good{background:rgba(0,128,0,.08)}
      .tool2-cell-bad{background:rgba(178,34,34,.08)}
      .tool2-cell-neutral{background:rgba(0,0,0,.03)}
      .tool2-small{font-size:.82rem}
      .tool2-actions{display:flex;flex-wrap:wrap;gap:.5rem;margin:.75rem 0}
      .tool2-actions button{cursor:pointer}
      .tool2-detail{margin-top:1rem}
      .tool2-detail details{border:1px solid rgba(0,0,0,.10);border-radius:12px;padding:.6rem .75rem;background:#fff}
      .tool2-detail summary{cursor:pointer;font-weight:700}
      .tool2-help{background:rgba(0,0,0,.03);border:1px solid rgba(0,0,0,.08);border-radius:12px;padding:.75rem;margin:.75rem 0}
      .tool2-tipicon{display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;border-radius:999px;border:1px solid rgba(0,0,0,.25);font-size:12px;line-height:18px;margin-left:.35rem;cursor:help;opacity:.75}
      .tool2-tooltip{position:fixed;z-index:9999;max-width:340px;background:#111;color:#fff;padding:.55rem .65rem;border-radius:10px;font-size:.85rem;box-shadow:0 10px 28px rgba(0,0,0,.25);pointer-events:none}
      .tool2-tooltip .t{opacity:.95}
      @media print{
        .no-print{display:none !important}
        .tool2-table-wrap{overflow:visible !important;border:none !important}
        table.tool2-compare{width:100% !important;min-width:100% !important}
        table.tool2-compare th{position:static !important}
        table.tool2-compare td:first-child, table.tool2-compare th:first-child{position:static !important}
      }
    `;
    document.head.appendChild(css);
  }

  // -----------------------------
  // TOOLTIP SYSTEM (close to indicator)
  // -----------------------------
  let tooltipEl = null;

  function ensureTooltip() {
    if (tooltipEl) return tooltipEl;
    tooltipEl = document.createElement("div");
    tooltipEl.className = "tool2-tooltip";
    tooltipEl.style.display = "none";
    document.body.appendChild(tooltipEl);
    return tooltipEl;
  }

  function showTipNear(target, html) {
    const t = ensureTooltip();
    t.innerHTML = `<div class="t">${html}</div>`;
    t.style.display = "block";
    const r = target.getBoundingClientRect();
    const pad = 10;
    const maxW = Math.min(360, window.innerWidth - 2 * pad);
    t.style.maxWidth = maxW + "px";

    // Initial position: below-left of the target
    let x = r.left;
    let y = r.bottom + 8;

    // Measure and keep within viewport
    const tr = t.getBoundingClientRect();
    if (x + tr.width > window.innerWidth - pad) x = window.innerWidth - pad - tr.width;
    if (x < pad) x = pad;
    if (y + tr.height > window.innerHeight - pad) y = r.top - tr.height - 8;
    if (y < pad) y = pad;

    t.style.left = `${x}px`;
    t.style.top = `${y}px`;
  }

  function hideTip() {
    if (!tooltipEl) return;
    tooltipEl.style.display = "none";
  }

  function bindTooltips() {
    document.addEventListener("mouseenter", e => {
      const el = e.target.closest("[data-tip]");
      if (!el) return;
      showTipNear(el, esc(el.getAttribute("data-tip")));
    }, true);

    document.addEventListener("mousemove", e => {
      if (!tooltipEl || tooltipEl.style.display === "none") return;
      // Keep anchored; do nothing on move for stability
    }, true);

    document.addEventListener("mouseleave", e => {
      const el = e.target.closest("[data-tip]");
      if (!el) return;
      hideTip();
    }, true);

    document.addEventListener("focusin", e => {
      const el = e.target.closest("[data-tip]");
      if (!el) return;
      showTipNear(el, esc(el.getAttribute("data-tip")));
    });

    document.addEventListener("focusout", e => {
      const el = e.target.closest("[data-tip]");
      if (!el) return;
      hideTip();
    });
  }

  // -----------------------------
  // MODEL
  // -----------------------------
  const model = {
    toolName: TOOL_NAME,
    project: {
      name: "Faba Beans: Soil amendment trial economics",
      lead: "Project lead",
      analysts: "Farm economics team",
      team: "Trial team",
      organisation: "Newcastle Business School, The University of Newcastle",
      contactEmail: "",
      contactPhone: "",
      summary:
        "Decision support comparing faba bean soil amendment treatments against a control using cost–benefit analysis.",
      objectives: "Compare treatments to control using NPV, PV benefits, PV costs, BCR, ROI, and ranking.",
      activities: "Load treatments (Excel-first), validate inputs, calculate CBA, export results, and generate AI interpretation prompts.",
      stakeholders: "Producers, agronomists, agencies, research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10),
      goal:
        "Support farmers to understand why treatments perform well or poorly economically, and what drivers matter most.",
      withProject:
        "Growers use the tool to compare treatment economics transparently, including underperforming options, and explore improvements.",
      withoutProject:
        "Decisions rely on partial information with limited comparability across treatments."
    },
    time: {
      startYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE))
    },
    adoption: { base: 0.9, low: 0.6, high: 1.0 },
    risk: {
      base: 0.15,
      low: 0.05,
      high: 0.3
    },
    outputsMeta: {
      systemType: "single",
      assumptions:
        "Outputs are valued per hectare using user-specified unit values. Treatment deltas represent differences versus control (e.g. yield uplift t/ha)."
    },
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t/ha", value: 450, source: "Excel or user input" }
    ],
    // Each treatment is an alternative; comparisons are computed versus the control treatment.
    treatments: [
      {
        id: uid(),
        name: "Control",
        isControl: true,
        area: 100,
        adoption: 1,
        deltas: {}, // per-output delta vs control (control should be 0)
        labourCost: 0, // $/ha per year (operating)
        materialsCost: 0, // $/ha per year (operating)
        servicesCost: 0, // $/ha per year (operating)
        capitalCost: 0, // $ in year 0 (lump sum)
        notes: "Baseline practice."
      }
    ],
    sim: {
      n: 1000,
      seed: null,
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true
    }
  };

  // Ensure deltas exist for outputs
  function ensureTreatmentDeltas() {
    model.treatments.forEach(t => {
      if (!t.deltas || typeof t.deltas !== "object") t.deltas = {};
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
    });
  }

  // -----------------------------
  // DEFAULT DATA (Faba beans 2022) — clean, tool-ready seed
  // - This is a fallback. Uploaded Excel always overrides the default for analysis.
  // -----------------------------
  const DEFAULT_FABA_2022_TREATMENTS = [
    { name: "Control", isControl: true, yieldUpliftTHa: 0, treatInputCostPerHa: 0 },
    { name: "Deep OM (CP1)", isControl: false, yieldUpliftTHa: -1.02, treatInputCostPerHa: 16500 },
    { name: "Deep OM (CP1) + liq. Gypsum (CHT)", isControl: false, yieldUpliftTHa: -0.13, treatInputCostPerHa: 16850 },
    { name: "Deep Ripping", isControl: false, yieldUpliftTHa: 0.54, treatInputCostPerHa: 0 },
    { name: "Deep Carbon-coated mineral (CCM)", isControl: false, yieldUpliftTHa: -0.07, treatInputCostPerHa: 3225 },
    { name: "Deep OM + Gypsum (CP2)", isControl: false, yieldUpliftTHa: 0.55, treatInputCostPerHa: 24000 },
    { name: "Surface Silicon", isControl: false, yieldUpliftTHa: -0.02, treatInputCostPerHa: 1000 },
    { name: "Deep liq. NPKS", isControl: false, yieldUpliftTHa: 0.58, treatInputCostPerHa: 2200 },
    { name: "Deep Gypsum", isControl: false, yieldUpliftTHa: 0.09, treatInputCostPerHa: 500 },
    { name: "Deep liq. Gypsum (CHT)", isControl: false, yieldUpliftTHa: -0.20, treatInputCostPerHa: 350 },
    { name: "Deep OM (CP1) + PAM", isControl: false, yieldUpliftTHa: 1.07, treatInputCostPerHa: 18000 },
    { name: "Deep OM (CP1) + Carbon-coated mineral (CCM)", isControl: false, yieldUpliftTHa: -0.32, treatInputCostPerHa: 21225 }
  ];

  function seedDefaultTreatmentsIfNeeded() {
    const hasMany = Array.isArray(model.treatments) && model.treatments.length > 1;
    if (hasMany) return;

    const yieldOutput = model.outputs.find(o => o.name.toLowerCase().includes("yield"));
    const yieldId = yieldOutput ? yieldOutput.id : model.outputs[0]?.id;

    model.treatments = DEFAULT_FABA_2022_TREATMENTS.map(row => {
      const t = {
        id: uid(),
        name: row.name,
        isControl: !!row.isControl,
        area: 100,
        adoption: 1,
        deltas: {},
        labourCost: 0,
        materialsCost: Number(row.treatInputCostPerHa || 0), // operating $/ha (treatment input)
        servicesCost: 0,
        capitalCost: 0,
        notes: row.isControl ? "Baseline practice." : "Seeded from the 2022 faba bean table (fallback)."
      };
      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      if (yieldId) t.deltas[yieldId] = Number(row.yieldUpliftTHa || 0);
      return t;
    });

    ensureTreatmentDeltas();
  }

  // -----------------------------
  // STORAGE (Excel-first default: uploaded data is the new default)
  // -----------------------------
  function saveModel() {
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(model));
    } catch (e) {
      // ignore
    }
  }

  function loadModel() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      if (!raw) return false;
      const obj = JSON.parse(raw);
      if (obj && typeof obj === "object") {
        // shallow merge, keep functions local
        Object.assign(model, obj);
        // guard required nested objects
        if (!model.toolName) model.toolName = TOOL_NAME;
        if (!model.time) model.time = {};
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        if (!model.project) model.project = {};
        ensureTreatmentDeltas();
        return true;
      }
      return false;
    } catch (e) {
      return false;
    }
  }

  // -----------------------------
  // CASHFLOW + CBA (per treatment, compared to control)
  // -----------------------------
  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
    return pv;
  }

  function irr(cf) {
    const hasPos = cf.some(v => v > 0);
    const hasNeg = cf.some(v => v < 0);
    if (!hasPos || !hasNeg) return NaN;

    const npvAt = r => cf.reduce((acc, v, t) => acc + v / Math.pow(1 + r, t), 0);

    let lo = -0.99, hi = 5.0;
    let nLo = npvAt(lo), nHi = npvAt(hi);
    if (nLo * nHi > 0) {
      for (let k = 0; k < 30 && nLo * nHi > 0; k++) {
        hi *= 1.5;
        nHi = npvAt(hi);
      }
      if (nLo * nHi > 0) return NaN;
    }

    for (let i = 0; i < 90; i++) {
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
    const m = Math.pow(-fvPos / pvNeg, 1 / n) - 1;
    return m * 100;
  }

  function payback(cf, ratePct) {
    let cum = 0;
    for (let t = 0; t < cf.length; t++) {
      cum += cf[t] / Math.pow(1 + ratePct / 100, t);
      if (cum >= 0) return t;
    }
    return null;
  }

  function computeTreatmentCashflows(t, opts) {
    const N = Number(opts?.years ?? model.time.years) || 10;
    const adoptMul = clamp(Number(opts?.adoptMul ?? model.adoption.base) || 0, 0, 1);
    const risk = clamp(Number(opts?.risk ?? model.risk.base) || 0, 0, 1);

    const benefit = new Array(N + 1).fill(0);
    const cost = new Array(N + 1).fill(0);

    const area = Number(t.area) || 0;
    const adoption = clamp(Number(t.adoption ?? 1), 0, 1);
    const A = adoptMul * adoption;
    const R = 1 - risk;

    // Benefits are annual for years 1..N.
    let valuePerHa = 0;
    model.outputs.forEach(o => {
      const delta = Number(t.deltas?.[o.id] ?? 0) || 0;
      const v = Number(o.value) || 0;
      valuePerHa += delta * v;
    });

    const annualBenefit = valuePerHa * area * A * R;

    // Costs: operating ($/ha) annually (years 1..N), capital year 0 (lump sum).
    const opPerHa =
      (Number(t.labourCost) || 0) +
      (Number(t.materialsCost) || 0) +
      (Number(t.servicesCost) || 0);

    const annualCost = opPerHa * area * A; // if adoption is partial, operating scales with adoption
    const capY0 = (Number(t.capitalCost) || 0); // capital is year-0; assumed incurred if adopting at all
    const capScaled = capY0 * clamp(A, 0, 1);

    cost[0] += capScaled;
    for (let y = 1; y <= N; y++) {
      benefit[y] += annualBenefit;
      cost[y] += annualCost;
    }

    const cf = benefit.map((b, i) => b - cost[i]);
    return { benefit, cost, cf };
  }

  function computeMetricsFromSeries(benefit, cost, cf, ratePct) {
    const pvBenefits = presentValue(benefit, ratePct);
    const pvCosts = presentValue(cost, ratePct);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? ((pvBenefits - pvCosts) / pvCosts) * 100 : NaN;
    const irrVal = irr(cf);
    const mirrVal = mirr(cf, model.time.mirrFinance, model.time.mirrReinvest);
    const pb = payback(cf, ratePct);
    return { pvBenefits, pvCosts, npv, bcr, roi, irrVal, mirrVal, paybackYears: pb };
  }

  function getControlTreatment() {
    const t = model.treatments.find(x => !!x.isControl);
    return t || model.treatments[0] || null;
  }

  function computeComparison(opts) {
    const rate = Number(opts?.rate ?? model.time.discBase) || 7;
    const years = Number(opts?.years ?? model.time.years) || 10;
    const adoptMul = Number(opts?.adoptMul ?? model.adoption.base) ?? 0.9;
    const risk = Number(opts?.risk ?? model.risk.base) ?? 0.15;

    const control = getControlTreatment();
    if (!control) return { control: null, rows: [], meta: { rate, years, adoptMul, risk } };

    const cSeries = computeTreatmentCashflows(control, { years, adoptMul, risk });
    const cAbs = computeMetricsFromSeries(cSeries.benefit, cSeries.cost, cSeries.cf, rate);

    const rows = model.treatments.map(t => {
      const s = computeTreatmentCashflows(t, { years, adoptMul, risk });
      const abs = computeMetricsFromSeries(s.benefit, s.cost, s.cf, rate);

      // Incremental vs control (difference)
      const dBenefit = s.benefit.map((v, i) => v - (cSeries.benefit[i] || 0));
      const dCost = s.cost.map((v, i) => v - (cSeries.cost[i] || 0));
      const dCf = dBenefit.map((v, i) => v - (dCost[i] || 0));

      const inc = computeMetricsFromSeries(dBenefit, dCost, dCf, rate);

      // For deltas: control should be 0
      return {
        id: t.id,
        name: t.name,
        isControl: !!t.isControl,
        area: Number(t.area) || 0,
        adoption: clamp(Number(t.adoption ?? 1), 0, 1),
        abs,
        inc,
        series: { abs: s, inc: { benefit: dBenefit, cost: dCost, cf: dCf } }
      };
    });

    // Ranking by incremental NPV (primary), then incremental BCR (secondary), then name.
    const ranked = rows
      .slice()
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;

        const an = Number(a.inc.npv);
        const bn = Number(b.inc.npv);
        if (isFinite(an) && isFinite(bn) && an !== bn) return bn - an;

        const ab = Number(a.inc.bcr);
        const bb = Number(b.inc.bcr);
        if (isFinite(ab) && isFinite(bb) && ab !== bb) return bb - ab;

        return String(a.name).localeCompare(String(b.name));
      })
      .map((r, idx) => ({ ...r, rank: idx + 1 }));

    return {
      control: {
        name: control.name,
        abs: cAbs,
        series: cSeries
      },
      rows: ranked,
      meta: { rate, years, adoptMul, risk }
    };
  }

  // -----------------------------
  // SNAPSHOT RESULTS RENDER (vertical table)
  // -----------------------------
  function classifyDeltaCell(v, betterIsHigher = true) {
    if (!isFinite(v) || v === 0) return "tool2-cell-neutral";
    const good = betterIsHigher ? v > 0 : v < 0;
    return good ? "tool2-cell-good" : "tool2-cell-bad";
  }

  function renderResultsComparison() {
    injectTool2CSS();

    const rate = Number($("#discBase")?.value ?? model.time.discBase) || model.time.discBase;
    const years = Number($("#years")?.value ?? model.time.years) || model.time.years;
    const adoptMul = Number($("#adoptBase")?.value ?? model.adoption.base) ?? model.adoption.base;
    const risk = Number($("#riskBase")?.value ?? model.risk.base) ?? model.risk.base;

    const comp = computeComparison({ rate, years, adoptMul, risk });
    if (!comp.control || !comp.rows.length) {
      setHTML("#resultsMain", `<div class="tool2-help">No treatments available. Upload an Excel file or load the sample template.</div>`);
      return;
    }

    const rows = comp.rows;
    const controlRow = rows.find(r => r.isControl) || rows[0];

    // Headline KPIs (snapshot)
    const nonControl = rows.filter(r => !r.isControl);
    const bestNPV = nonControl.slice().sort((a, b) => (b.inc.npv || -Infinity) - (a.inc.npv || -Infinity))[0] || null;
    const bestBCR = nonControl.slice().sort((a, b) => (b.inc.bcr || -Infinity) - (a.inc.bcr || -Infinity))[0] || null;

    const kpisHTML = `
      <div class="tool2-kpis">
        <div class="tool2-kpi">
          <div class="lbl">Control</div>
          <div class="val">${esc(controlRow.name)}</div>
          <div class="sub">Comparison baseline (not hidden)</div>
        </div>
        <div class="tool2-kpi">
          <div class="lbl">Top by ΔNPV vs control</div>
          <div class="val">${bestNPV ? esc(bestNPV.name) : "n/a"}</div>
          <div class="sub">${bestNPV ? money(bestNPV.inc.npv) : "n/a"}</div>
        </div>
        <div class="tool2-kpi">
          <div class="lbl">Top by ΔBCR vs control</div>
          <div class="val">${bestBCR ? esc(bestBCR.name) : "n/a"}</div>
          <div class="sub">${bestBCR ? fmt(bestBCR.inc.bcr) : "n/a"}</div>
        </div>
        <div class="tool2-kpi">
          <div class="lbl">Scenario</div>
          <div class="val">${years} years, ${fmt(rate)}% discount</div>
          <div class="sub">Adoption ${fmt(adoptMul * 100)}%, risk ${fmt(risk * 100)}%</div>
        </div>
      </div>
    `;

    // Comparison table: indicators as rows, treatments as columns (control included)
    const treatmentCols = rows; // already includes control (ranked with control first)
    const colHeaders = treatmentCols
      .map(r => {
        const tag = r.isControl ? `<span class="tool2-muted">(Control)</span>` : `<span class="tool2-muted">Rank ${r.rank}</span>`;
        return `<th>${esc(r.name)}<div class="tool2-small tool2-muted">${tag}</div></th>`;
      })
      .join("");

    const metricRows = [];

    // Minimal required metrics (absolute + delta rows for clarity and copy/paste)
    const metrics = [
      {
        key: "npv",
        label: "Net present value (NPV)",
        tip:
          "NPV is the present value of benefits minus costs. Positive ΔNPV vs control means the treatment outperforms the control financially under the chosen assumptions.",
        fmt: money,
        betterHigher: true
      },
      {
        key: "pvBenefits",
        label: "Present value of benefits",
        tip:
          "PV benefits discounts future benefits back to today. Higher PV benefits can come from higher yield uplift, higher prices, higher adoption, or lower risk.",
        fmt: money,
        betterHigher: true
      },
      {
        key: "pvCosts",
        label: "Present value of costs",
        tip:
          "PV costs discounts future costs back to today. Higher PV costs can come from higher input costs, labour, services, or capital outlays.",
        fmt: money,
        betterHigher: false
      },
      {
        key: "bcr",
        label: "Benefit–cost ratio (BCR)",
        tip:
          "BCR is PV benefits divided by PV costs. A higher BCR indicates more benefit per dollar of cost. This tool reports ΔBCR vs control as an interpretive aid, not a rule.",
        fmt: x => (isFinite(x) ? fmt(x) : "n/a"),
        betterHigher: true
      },
      {
        key: "roi",
        label: "Return on investment (ROI)",
        tip:
          "ROI is (PV benefits − PV costs) / PV costs. It summarises proportional return. High ROI can occur even when total dollars are modest.",
        fmt: x => (isFinite(x) ? fmt(x) + "%" : "n/a"),
        betterHigher: true
      },
      {
        key: "rank",
        label: "Ranking (by ΔNPV vs control)",
        tip:
          "Ranking sorts treatments by incremental NPV compared to the control. It does not dictate a decision: it is a transparent summary under the selected assumptions.",
        fmt: x => (isFinite(x) ? String(x) : "n/a"),
        betterHigher: false
      }
    ];

    // Helper: metric row builder
    function buildMetricRow(label, tip, getValue, deltaMode, betterHigher, formatFn) {
      const tipIcon = `<span class="tool2-tipicon" tabindex="0" role="img" aria-label="Help" data-tip="${esc(tip)}">i</span>`;
      const headerCell = `<td>${esc(label)} ${tipIcon}<div class="tool2-small tool2-muted">${deltaMode ? "Difference vs control" : "Absolute"}</div></td>`;
      const cells = treatmentCols
        .map(r => {
          const v = getValue(r);
          let cls = "";
          if (deltaMode) {
            cls = classifyDeltaCell(v, betterHigher);
          } else {
            cls = r.isControl ? "tool2-cell-neutral" : "tool2-cell-neutral";
          }
          return `<td class="${cls}">${esc(formatFn(v))}</td>`;
        })
        .join("");
      return `<tr>${headerCell}${cells}</tr>`;
    }

    // For each metric, include Absolute and Δ vs control
    metrics.forEach(m => {
      if (m.key === "rank") {
        // absolute rank row only (delta not meaningful)
        metricRows.push(
          buildMetricRow(
            m.label,
            m.tip,
            r => (r.isControl ? 1 : r.rank),
            false,
            m.betterHigher,
            m.fmt
          )
        );
        return;
      }

      // Absolute
      metricRows.push(
        buildMetricRow(
          m.label,
          m.tip,
          r => r.abs[m.key],
          false,
          m.betterHigher,
          m.fmt
        )
      );

      // Delta vs control (use incremental metrics; control is 0 by construction)
      metricRows.push(
        buildMetricRow(
          m.label,
          m.tip,
          r => (r.isControl ? 0 : r.inc[m.key]),
          true,
          m.betterHigher,
          m.fmt
        )
      );
    });

    const tableHTML = `
      <div class="tool2-table-wrap" id="resultsCompareWrap">
        <table class="tool2-compare" id="resultsCompareTable">
          <thead>
            <tr>
              <th>Economic indicator</th>
              ${colHeaders}
            </tr>
          </thead>
          <tbody>
            ${metricRows.join("")}
          </tbody>
        </table>
      </div>
    `;

    const actionsHTML = `
      <div class="tool2-actions no-print">
        <button id="tool2CopyWord" type="button">Copy table (Word-friendly)</button>
        <button id="tool2CopyTSV" type="button">Copy table (TSV)</button>
        <button id="tool2ExportResultsExcel" type="button">Export results (Excel)</button>
        <button id="tool2PrintResults" type="button">Print / Save as PDF</button>
        <button id="tool2MakeAIPrompt" type="button">Generate AI interpretation prompt</button>
      </div>
    `;

    const detailHTML = `
      <div class="tool2-detail">
        <details>
          <summary>Why do treatments differ? (optional detail)</summary>
          <div style="margin-top:.6rem" class="tool2-small">
            Results depend on yield uplift (and other output deltas), unit values (e.g. $/t), operating costs ($/ha), capital cost (year 0), adoption, risk, horizon, and the discount rate.
            Underperforming treatments remain visible so you can diagnose drivers and explore realistic changes in costs, yields, or prices.
          </div>
        </details>
      </div>
    `;

    setHTML(
      "#resultsMain",
      `
        <div class="tool2-badge"><span class="dot"></span><strong>${esc(TOOL_NAME)}</strong></div>
        ${kpisHTML}
        ${actionsHTML}
        ${tableHTML}
        ${detailHTML}
        <div id="tool2AIPromptBlock" class="tool2-detail"></div>
      `
    );

    bindResultsButtons(comp);
  }

  // Copy helpers (Word-friendly HTML and TSV)
  function tableToTSV(table) {
    const rows = Array.from(table.querySelectorAll("tr"));
    return rows
      .map(r =>
        Array.from(r.querySelectorAll("th,td"))
          .map(c => (c.innerText || "").replace(/\s+\n/g, " ").replace(/\n+/g, " ").trim())
          .join("\t")
      )
      .join("\n");
  }

  async function copyToClipboard({ text, html }) {
    // Prefer rich clipboard if supported
    try {
      if (navigator.clipboard && window.ClipboardItem && html) {
        const item = new ClipboardItem({
          "text/plain": new Blob([text || ""], { type: "text/plain" }),
          "text/html": new Blob([html], { type: "text/html" })
        });
        await navigator.clipboard.write([item]);
        return true;
      }
    } catch (_) {}
    try {
      if (navigator.clipboard) {
        await navigator.clipboard.writeText(text || "");
        return true;
      }
    } catch (_) {}
    return false;
  }

  function bindResultsButtons(comp) {
    const table = $("#resultsCompareTable");
    if (!table) return;

    const btnWord = $("#tool2CopyWord");
    const btnTSV = $("#tool2CopyTSV");
    const btnXLSX = $("#tool2ExportResultsExcel");
    const btnPrint = $("#tool2PrintResults");
    const btnAI = $("#tool2MakeAIPrompt");

    if (btnWord) {
      btnWord.addEventListener("click", async () => {
        const tsv = tableToTSV(table);
        const html = table.outerHTML;
        const ok = await copyToClipboard({ text: tsv, html });
        showToast(ok ? "Copied (Word-friendly)." : "Copy failed. Try the TSV option.");
      });
    }

    if (btnTSV) {
      btnTSV.addEventListener("click", async () => {
        const tsv = tableToTSV(table);
        const ok = await copyToClipboard({ text: tsv });
        showToast(ok ? "Copied (TSV)." : "Copy failed.");
      });
    }

    if (btnXLSX) {
      btnXLSX.addEventListener("click", () => exportResultsExcel(comp));
    }

    if (btnPrint) {
      btnPrint.addEventListener("click", () => {
        // Print results only by scrolling to the table; CSS hides irrelevant items if your HTML uses .no-print
        window.print();
      });
    }

    if (btnAI) {
      btnAI.addEventListener("click", () => {
        renderAIPrompt(comp);
        showToast("AI interpretation prompt generated.");
      });
    }
  }

  // -----------------------------
  // AI PROMPT (flexible, learning + improvement aid)
  // -----------------------------
  function buildAIPrompt(comp) {
    const { meta, rows } = comp;
    const control = rows.find(r => r.isControl) || rows[0];

    // Identify underperformers for guidance (not rules)
    const under = rows
      .filter(r => !r.isControl)
      .filter(r => (isFinite(r.inc.bcr) && r.inc.bcr < 1) || (isFinite(r.inc.npv) && r.inc.npv < 0));

    // Compact treatment summary table for prompt (TSV)
    const header = ["Treatment", "Is control", "Rank (ΔNPV)", "ΔNPV", "ΔPV Benefits", "ΔPV Costs", "ΔBCR", "ΔROI"];
    const lines = [header.join("\t")];
    rows.forEach(r => {
      lines.push(
        [
          r.name,
          r.isControl ? "Yes" : "No",
          r.isControl ? "1" : String(r.rank),
          money(r.isControl ? 0 : r.inc.npv),
          money(r.isControl ? 0 : r.inc.pvBenefits),
          money(r.isControl ? 0 : r.inc.pvCosts),
          isFinite(r.isControl ? 0 : r.inc.bcr) ? fmt(r.inc.bcr) : "n/a",
          isFinite(r.isControl ? 0 : r.inc.roi) ? fmt(r.inc.roi) + "%" : "n/a"
        ].join("\t")
      );
    });

    const improv = under.length
      ? under
          .slice(0, 6)
          .map(r => `- ${r.name}: consider practical levers such as reducing input costs (materials/labour/services), targeting application rates, improving timeliness/operations, improving yield response through agronomic practice, or seeking better output prices/quality. Explain which lever is most plausible given the results table.`)
          .join("\n")
      : "- If any treatment has low ΔBCR or negative ΔNPV, propose realistic improvement levers (cost, yield, price, agronomy) without imposing thresholds or telling the user what to do.";

    const prompt = `
You are helping interpret a farming cost–benefit analysis produced by ${TOOL_NAME}. Your job is to explain results in plain English and help the user learn. Do not dictate a decision and do not impose thresholds. Focus on transparency: why a treatment performs well or poorly and what drives that result.

Context
The table below compares multiple treatments against a Control. All figures reported are differences versus the Control (Δ). Positive ΔNPV means the treatment performs better than the Control financially under the stated assumptions. Negative ΔNPV means it performs worse. ΔPV Costs can be positive (more cost than control) or negative (cost saving). ΔBCR and ΔROI are summary indicators; interpret cautiously and explain limitations.

Scenario assumptions
Time horizon: ${meta.years} years
Discount rate: ${fmt(meta.rate)}%
Adoption multiplier (scenario): ${fmt(meta.adoptMul * 100)}%
Risk (scenario): ${fmt(meta.risk * 100)}%

What I need from you
1) Start with a short “what stands out” paragraph summarising the top performers and the underperformers by ΔNPV and ΔBCR.
2) Explain what each indicator means (NPV, PV benefits, PV costs, BCR, ROI) in a farmer-friendly way.
3) For the top 2–3 treatments, explain the drivers: is performance driven mainly by higher benefits (e.g., yield uplift/value) or by lower costs, and how strong the trade-off is.
4) For the weaker treatments, explain why they underperform and suggest realistic ways performance could improve. This is guidance for reflection, not rules. Examples: reduce costs, improve yield response, improve output prices/quality, refine agronomic practice, or improve targeting/timing.
5) Finish with a neutral note on decision-making: what additional information the user might check (prices, variability, operational constraints, agronomic feasibility) before choosing.

Results table (TSV, differences versus Control)
${lines.join("\n")}

Improvement prompts for low-performing treatments
${improv}
`.trim();

    return prompt;
  }

  function renderAIPrompt(comp) {
    const prompt = buildAIPrompt(comp);
    const block = $("#tool2AIPromptBlock");
    if (!block) return;

    setHTML(
      block,
      `
      <details open>
        <summary>AI-assisted interpretation prompt (copy/paste into Copilot, ChatGPT, or another model)</summary>
        <div class="tool2-help tool2-small">
          Copy the prompt below into your preferred AI tool. The tool does not decide for you. It explains what indicators mean, why a treatment performs well or poorly, and suggests realistic improvement levers for learning.
        </div>
        <div class="tool2-actions no-print">
          <button id="tool2CopyPrompt" type="button">Copy prompt</button>
          <button id="tool2DownloadBrief" type="button">Download prompt (text)</button>
        </div>
        <textarea id="tool2PromptText" style="width:100%;min-height:260px;border-radius:12px;border:1px solid rgba(0,0,0,.15);padding:.7rem;white-space:pre;">${esc(prompt)}</textarea>
      </details>
      `
    );

    const copyBtn = $("#tool2CopyPrompt");
    const dlBtn = $("#tool2DownloadBrief");
    if (copyBtn) {
      copyBtn.addEventListener("click", async () => {
        const ok = await copyToClipboard({ text: prompt });
        showToast(ok ? "Prompt copied." : "Copy failed.");
      });
    }
    if (dlBtn) {
      dlBtn.addEventListener("click", () => {
        downloadFile(`${slug(model.project.name || TOOL_NAME)}_ai_prompt.txt`, prompt, "text/plain");
      });
    }
  }

  // -----------------------------
  // EXCEL-FIRST WORKFLOW
  // -----------------------------
  let parsedExcel = null;

  function hasXLSX() {
    return typeof window !== "undefined" && typeof window.XLSX !== "undefined";
  }

  function slug(s) {
    return (s || TOOL_NAME)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  }

  function downloadFile(filename, text, mime) {
    const blob = new Blob([text], { type: mime || "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 0);
  }

  function buildTemplateWorkbook({ includeSample = false } = {}) {
    const settings = [
      ["Tool", TOOL_NAME],
      ["Project name", model.project.name || ""],
      ["Start year", model.time.startYear],
      ["Time horizon (years)", model.time.years],
      ["Discount rate (base, %)", model.time.discBase],
      ["Adoption multiplier (base, 0-1)", model.adoption.base],
      ["Risk (base, 0-1)", model.risk.base]
    ];

    const outputs = model.outputs.map(o => ({
      output_id: o.id,
      output_name: o.name,
      unit: o.unit,
      unit_value: o.value,
      source: o.source || ""
    }));

    const baseTreats = includeSample ? DEFAULT_FABA_2022_TREATMENTS : model.treatments;

    const treatments = baseTreats.map(t => {
      const obj = {
        treatment_name: t.name,
        is_control: !!t.isControl,
        area_ha: Number(t.area ?? 100),
        adoption: Number(t.adoption ?? 1),
        labour_cost_per_ha: Number(t.labourCost ?? 0),
        materials_cost_per_ha: Number(t.materialsCost ?? 0),
        services_cost_per_ha: Number(t.servicesCost ?? 0),
        capital_cost_year0: Number(t.capitalCost ?? 0),
        notes: t.notes || ""
      };
      // Add delta columns for each output
      model.outputs.forEach(o => {
        obj[`delta_${o.name}`] = (t.deltas && o.id in t.deltas) ? Number(t.deltas[o.id] ?? 0) : 0;
      });
      return obj;
    });

    return { settings, outputs, treatments };
  }

  function downloadExcelTemplate({ includeSample = false } = {}) {
    const wbData = buildTemplateWorkbook({ includeSample });
    const filename = includeSample
      ? `${slug(model.project.name)}_sample_template.xlsx`
      : `${slug(model.project.name)}_template.xlsx`;

    if (hasXLSX()) {
      const XLSX = window.XLSX;

      const wb = XLSX.utils.book_new();
      const wsSettings = XLSX.utils.aoa_to_sheet(wbData.settings);
      const wsOutputs = XLSX.utils.json_to_sheet(wbData.outputs);
      const wsTreatments = XLSX.utils.json_to_sheet(wbData.treatments);

      XLSX.utils.book_append_sheet(wb, wsSettings, "Settings");
      XLSX.utils.book_append_sheet(wb, wsOutputs, "Outputs");
      XLSX.utils.book_append_sheet(wb, wsTreatments, "Treatments");

      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 0);

      showToast(includeSample ? "Sample Excel template downloaded." : "Excel template downloaded.");
    } else {
      // CSV fallback (separate files)
      downloadFile(filename.replace(/\.xlsx$/i, "_Settings.csv"), wbData.settings.map(r => r.join(",")).join("\n"), "text/csv");
      downloadFile(filename.replace(/\.xlsx$/i, "_Outputs.csv"), jsonToCSV(wbData.outputs), "text/csv");
      downloadFile(filename.replace(/\.xlsx$/i, "_Treatments.csv"), jsonToCSV(wbData.treatments), "text/csv");
      showToast("XLSX library not found. Downloaded CSV templates instead.");
    }
  }

  function jsonToCSV(rows) {
    if (!rows || !rows.length) return "";
    const cols = Object.keys(rows[0]);
    const out = [cols.join(",")];
    rows.forEach(r => {
      out.push(cols.map(c => `"${String(r[c] ?? "").replace(/"/g, '""')}"`).join(","));
    });
    return out.join("\n");
  }

  async function parseUploadedExcel(file) {
    if (!file) throw new Error("No file selected.");

    if (hasXLSX()) {
      const XLSX = window.XLSX;
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const sheetNames = wb.SheetNames || [];
      const getSheet = name => wb.Sheets[name] || wb.Sheets[sheetNames.find(n => n.toLowerCase() === name.toLowerCase())];

      const wsSettings = getSheet("Settings");
      const wsOutputs = getSheet("Outputs");
      const wsTreatments = getSheet("Treatments");

      if (!wsOutputs || !wsTreatments) {
        throw new Error("Missing required sheets: Outputs and Treatments.");
      }

      const settingsAOA = wsSettings ? XLSX.utils.sheet_to_json(wsSettings, { header: 1, raw: true }) : [];
      const outputs = XLSX.utils.sheet_to_json(wsOutputs, { raw: true });
      const treatments = XLSX.utils.sheet_to_json(wsTreatments, { raw: true });

      return { settingsAOA, outputs, treatments };
    }

    // CSV fallback: expect single CSV with Treatments columns (user can still load JSON elsewhere)
    const txt = await file.text();
    const rows = parseCSV(txt);
    return { settingsAOA: [], outputs: [], treatments: rows };
  }

  function parseCSV(text) {
    // minimal CSV parser (handles quoted values)
    const lines = String(text || "").split(/\r?\n/).filter(Boolean);
    if (!lines.length) return [];
    const header = splitCSVLine(lines[0]);
    const out = [];
    for (let i = 1; i < lines.length; i++) {
      const vals = splitCSVLine(lines[i]);
      const obj = {};
      header.forEach((h, j) => (obj[h] = vals[j]));
      out.push(obj);
    }
    return out;
  }

  function splitCSVLine(line) {
    const out = [];
    let cur = "";
    let inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') {
        inQ = !inQ;
      } else if (ch === "," && !inQ) {
        out.push(cur);
        cur = "";
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out.map(s => s.trim());
  }

  function validateParsedExcel(parsed) {
    const errs = [];

    // Outputs sheet: require at least output_name + unit_value (or output_id)
    const outputs = parsed.outputs || [];
    if (!outputs.length) errs.push("Outputs sheet has no rows.");

    // Treatments sheet: require treatment_name and is_control at minimum
    const treatments = parsed.treatments || [];
    if (!treatments.length) errs.push("Treatments sheet has no rows.");

    const hasTreatmentName = treatments.length ? ("treatment_name" in treatments[0] || "Treatment" in treatments[0]) : false;
    if (!hasTreatmentName) errs.push("Treatments sheet missing 'treatment_name' column.");

    // Must have exactly one control
    const controlCount = treatments.filter(r => String(r.is_control ?? r.isControl ?? "").toLowerCase() === "true" || String(r.is_control ?? "").toLowerCase() === "yes" || String(r.is_control ?? "").toLowerCase() === "1").length;
    if (controlCount !== 1) errs.push(`Treatments sheet must mark exactly one control (found ${controlCount}).`);

    return errs;
  }

  function commitParsedExcelToModel(parsed) {
    // Settings
    const settings = new Map();
    (parsed.settingsAOA || []).forEach(r => {
      if (!r || r.length < 2) return;
      settings.set(String(r[0]).trim().toLowerCase(), r[1]);
    });

    const years = parseNumber(settings.get("time horizon (years)"));
    const disc = parseNumber(settings.get("discount rate (base, %)"));
    const adopt = parseNumber(settings.get("adoption multiplier (base, 0-1)"));
    const risk = parseNumber(settings.get("risk (base, 0-1)"));
    const startYear = parseNumber(settings.get("start year"));

    if (Number.isFinite(years)) model.time.years = years;
    if (Number.isFinite(disc)) model.time.discBase = disc;
    if (Number.isFinite(adopt)) model.adoption.base = clamp(adopt, 0, 1);
    if (Number.isFinite(risk)) model.risk.base = clamp(risk, 0, 1);
    if (Number.isFinite(startYear)) model.time.startYear = startYear;

    // Outputs
    const outputs = (parsed.outputs || []).map(r => {
      const id = r.output_id ? String(r.output_id) : uid();
      const name = String(r.output_name ?? r.name ?? r.Output ?? "Output").trim();
      const unit = String(r.unit ?? "").trim() || "";
      const v = parseNumber(r.unit_value ?? r.value ?? r.UnitValue);
      return {
        id,
        name,
        unit,
        value: Number.isFinite(v) ? v : 0,
        source: String(r.source ?? "") || "Excel upload"
      };
    }).filter(o => o.name);

    if (outputs.length) model.outputs = outputs;

    // Treatments
    const treatments = (parsed.treatments || []).map(r => {
      const name = String(r.treatment_name ?? r.Treatment ?? r.name ?? "").trim();
      if (!name) return null;

      const isCtrl =
        String(r.is_control ?? r.isControl ?? "").toLowerCase() === "true" ||
        String(r.is_control ?? "").toLowerCase() === "yes" ||
        String(r.is_control ?? "").toLowerCase() === "1";

      const area = parseNumber(r.area_ha ?? r.area ?? 100);
      const adoption = parseNumber(r.adoption ?? 1);

      const labour = parseNumber(r.labour_cost_per_ha ?? r.labourCost ?? 0);
      const materials = parseNumber(r.materials_cost_per_ha ?? r.materialsCost ?? 0);
      const services = parseNumber(r.services_cost_per_ha ?? r.servicesCost ?? 0);
      const capital = parseNumber(r.capital_cost_year0 ?? r.capitalCost ?? 0);

      const t = {
        id: uid(),
        name,
        isControl: !!isCtrl,
        area: Number.isFinite(area) ? area : 0,
        adoption: Number.isFinite(adoption) ? clamp(adoption, 0, 1) : 1,
        deltas: {},
        labourCost: Number.isFinite(labour) ? labour : 0,
        materialsCost: Number.isFinite(materials) ? materials : 0,
        servicesCost: Number.isFinite(services) ? services : 0,
        capitalCost: Number.isFinite(capital) ? capital : 0,
        notes: String(r.notes ?? "") || "Excel upload"
      };

      // Deltas: accept columns named delta_<outputName> or exact output names
      model.outputs.forEach(o => {
        let v = NaN;
        const k1 = `delta_${o.name}`;
        if (k1 in r) v = parseNumber(r[k1]);
        if (!Number.isFinite(v) && o.name in r) v = parseNumber(r[o.name]);
        t.deltas[o.id] = Number.isFinite(v) ? v : 0;
      });

      return t;
    }).filter(Boolean);

    // Enforce exactly one control; if not, attempt fix by first row
    const ctrlIdx = treatments.findIndex(t => t.isControl);
    if (ctrlIdx < 0) {
      treatments[0].isControl = true;
    } else {
      treatments.forEach((t, i) => (t.isControl = i === ctrlIdx));
    }

    model.treatments = treatments;
    ensureTreatmentDeltas();

    // Persist as new default
    saveModel();
  }

  async function handleParseExcel() {
    const input = $("#excelFile") || $("#loadExcelFile") || $("#importFile");
    if (!input || !input.files || !input.files[0]) {
      alert("Please choose an Excel file to upload first.");
      return;
    }
    const file = input.files[0];
    try {
      parsedExcel = await parseUploadedExcel(file);
      const errs = validateParsedExcel(parsedExcel);
      if (errs.length) {
        alert("Excel file issues:\n\n- " + errs.join("\n- "));
        parsedExcel = null;
        return;
      }
      setText("#excelParseStatus", "File parsed and validated. Click ‘Apply to tool’ to update results.");
      showToast("Excel parsed and validated.");
    } catch (e) {
      console.error(e);
      alert(String(e.message || e));
      parsedExcel = null;
    }
  }

  function commitExcelToModel() {
    if (!parsedExcel) {
      alert("No parsed Excel is available. Click ‘Parse’ first.");
      return;
    }
    commitParsedExcelToModel(parsedExcel);
    parsedExcel = null;

    // Re-render all key views
    renderOutputs();
    renderTreatments();
    renderResultsComparison();
    setBasicsFieldsFromModel?.();
    showToast("Excel applied. Results updated.");
  }

  // -----------------------------
  // EXPORT RESULTS (Excel)
  // -----------------------------
  function exportResultsExcel(comp) {
    const { meta, rows } = comp;
    const outRows = rows.map(r => ({
      treatment: r.name,
      is_control: r.isControl ? "Yes" : "No",
      rank_by_delta_npv: r.isControl ? 1 : r.rank,
      delta_npv: r.isControl ? 0 : r.inc.npv,
      delta_pv_benefits: r.isControl ? 0 : r.inc.pvBenefits,
      delta_pv_costs: r.isControl ? 0 : r.inc.pvCosts,
      delta_bcr: r.isControl ? 0 : r.inc.bcr,
      delta_roi_pct: r.isControl ? 0 : r.inc.roi,
      abs_npv: r.abs.npv,
      abs_pv_benefits: r.abs.pvBenefits,
      abs_pv_costs: r.abs.pvCosts,
      abs_bcr: r.abs.bcr,
      abs_roi_pct: r.abs.roi
    }));

    const metaRows = [
      { key: "Tool", value: TOOL_NAME },
      { key: "Project", value: model.project.name || "" },
      { key: "Years", value: meta.years },
      { key: "Discount rate (%)", value: meta.rate },
      { key: "Adoption multiplier", value: meta.adoptMul },
      { key: "Risk", value: meta.risk }
    ];

    if (hasXLSX()) {
      const XLSX = window.XLSX;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(metaRows), "Scenario");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outRows), "Results_Comparison");

      // Inputs for audit
      const outputs = model.outputs.map(o => ({
        output_name: o.name,
        unit: o.unit,
        unit_value: o.value
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outputs), "Inputs_Outputs");

      const treats = model.treatments.map(t => {
        const row = {
          treatment_name: t.name,
          is_control: t.isControl,
          area_ha: t.area,
          adoption: t.adoption,
          labour_cost_per_ha: t.labourCost,
          materials_cost_per_ha: t.materialsCost,
          services_cost_per_ha: t.servicesCost,
          capital_cost_year0: t.capitalCost,
          notes: t.notes || ""
        };
        model.outputs.forEach(o => {
          row[`delta_${o.name}`] = t.deltas?.[o.id] ?? 0;
        });
        return row;
      });
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(treats), "Inputs_Treatments");

      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `${slug(model.project.name)}_results.xlsx`;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        URL.revokeObjectURL(a.href);
        a.remove();
      }, 0);

      showToast("Results exported (Excel).");
    } else {
      downloadFile(`${slug(model.project.name)}_results.csv`, jsonToCSV(outRows), "text/csv");
      showToast("XLSX library not found. Exported CSV instead.");
    }
  }

  // -----------------------------
  // UI: PROJECT / SETTINGS BINDING (keeps your existing fields working)
  // -----------------------------
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

    // Discount schedule inputs if present
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

  // -----------------------------
  // TABS (keep your existing behaviour)
  // -----------------------------
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
    });

    const activeNav =
      document.querySelector("[data-tab].active, [data-tab-target].active, [data-tab-jump].active") ||
      document.querySelector("[data-tab], [data-tab-target], [data-tab-jump]");
    if (activeNav) {
      const target = activeNav.dataset.tab || activeNav.dataset.tabTarget || activeNav.dataset.tabJump;
      if (target) switchTab(target);
    }
  }

  // -----------------------------
  // OUTPUTS RENDER
  // -----------------------------
  function renderOutputs() {
    const host = $("#outputsTable") || $("#outputsBody") || $("#outputs");
    if (!host) return;

    const rows = model.outputs.map((o, idx) => {
      return `
        <tr>
          <td><input data-out-field="name" data-out-id="${esc(o.id)}" value="${esc(o.name)}"/></td>
          <td><input data-out-field="unit" data-out-id="${esc(o.id)}" value="${esc(o.unit || "")}"/></td>
          <td><input data-out-field="value" data-out-id="${esc(o.id)}" value="${esc(o.value)}" /></td>
          <td class="tool2-muted tool2-small">${esc(o.source || "")}</td>
          <td><button type="button" data-out-action="del" data-out-id="${esc(o.id)}">Remove</button></td>
        </tr>
      `;
    }).join("");

    const table = `
      <div class="tool2-help tool2-small">
        Outputs are valued per unit. Treatment deltas represent changes versus the control (e.g. yield uplift t/ha).
      </div>
      <div class="tool2-table-wrap">
        <table class="tool2-compare" style="min-width:900px;width:100%">
          <thead>
            <tr>
              <th>Output</th>
              <th>Unit</th>
              <th>Unit value</th>
              <th>Source</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
      <div class="tool2-actions no-print">
        <button id="addOutput" type="button">Add output</button>
      </div>
    `;
    setHTML(host, table);

    const addBtn = $("#addOutput");
    if (addBtn) {
      addBtn.onclick = () => {
        model.outputs.push({ id: uid(), name: "New output", unit: "", value: 0, source: "User input" });
        ensureTreatmentDeltas();
        saveModel();
        renderOutputs();
        renderTreatments();
        renderResultsComparison();
      };
    }
  }

  // -----------------------------
  // TREATMENTS RENDER
  // - Capital cost ($, year 0) appears before Total cost ($/ha)
  // - Totals update dynamically when cost components change
  // -----------------------------
  function totalOpCostPerHa(t) {
    return (Number(t.labourCost) || 0) + (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0);
  }

  function renderTreatments() {
    const host = $("#treatmentsTable") || $("#treatmentsBody") || $("#treatments");
    if (!host) return;

    const controlTip = "Exactly one treatment must be marked as the control. All results are compared against it.";
    const capTip = "Capital cost is a year-0 lump sum (e.g. once-off equipment purchase). It is not part of operating $/ha totals.";
    const opTip = "Operating costs are per hectare per year (labour, materials, services). Total cost ($/ha) updates automatically.";
    const deltaTip = "Delta values represent differences versus the control (e.g. yield uplift). Control deltas should be 0.";

    const headerDeltaCols = model.outputs
      .map(o => `<th>${esc(o.name)} Δ <span class="tool2-tipicon" tabindex="0" data-tip="${esc(deltaTip)}">i</span></th>`)
      .join("");

    const bodyRows = model.treatments
      .map(t => {
        const opTotal = totalOpCostPerHa(t);
        const deltaCells = model.outputs.map(o => {
          const v = Number(t.deltas?.[o.id] ?? 0) || 0;
          return `<td><input data-trt-delta="1" data-trt-id="${esc(t.id)}" data-out-id="${esc(o.id)}" value="${esc(v)}"/></td>`;
        }).join("");

        return `
          <tr>
            <td><input data-trt-field="name" data-trt-id="${esc(t.id)}" value="${esc(t.name)}"/></td>
            <td style="text-align:center">
              <input type="radio" name="controlRadio" data-trt-field="isControl" data-trt-id="${esc(t.id)}" ${t.isControl ? "checked" : ""}/>
              <span class="tool2-tipicon" tabindex="0" data-tip="${esc(controlTip)}">i</span>
            </td>
            <td><input data-trt-field="area" data-trt-id="${esc(t.id)}" value="${esc(t.area)}"/></td>
            <td><input data-trt-field="adoption" data-trt-id="${esc(t.id)}" value="${esc(t.adoption)}"/></td>

            <td><input data-trt-field="labourCost" data-trt-id="${esc(t.id)}" value="${esc(t.labourCost)}"/></td>
            <td><input data-trt-field="materialsCost" data-trt-id="${esc(t.id)}" value="${esc(t.materialsCost)}"/></td>
            <td><input data-trt-field="servicesCost" data-trt-id="${esc(t.id)}" value="${esc(t.servicesCost)}"/></td>

            <td>
              <input data-trt-field="capitalCost" data-trt-id="${esc(t.id)}" value="${esc(t.capitalCost)}"/>
              <span class="tool2-tipicon" tabindex="0" data-tip="${esc(capTip)}">i</span>
            </td>

            <td class="tool2-cell-neutral"><strong>${money(opTotal)}</strong><div class="tool2-small tool2-muted">Operating ($/ha/yr)</div></td>

            ${deltaCells}

            <td><input data-trt-field="notes" data-trt-id="${esc(t.id)}" value="${esc(t.notes || "")}"/></td>
            <td><button type="button" data-trt-action="del" data-trt-id="${esc(t.id)}">Remove</button></td>
          </tr>
        `;
      })
      .join("");

    const table = `
      <div class="tool2-help tool2-small">
        <div><strong>${TOOL_NAME}</strong> compares each treatment against the control using incremental CBA. Low-performing treatments remain visible to support learning and diagnosis.</div>
        <div style="margin-top:.35rem">
          <span class="tool2-tipicon" tabindex="0" data-tip="${esc(opTip)}">i</span>
          Operating costs update the totals instantly as you edit labour/materials/services.
        </div>
      </div>

      <div class="tool2-table-wrap">
        <table class="tool2-compare" style="min-width:1200px">
          <thead>
            <tr>
              <th>Treatment</th>
              <th>Control</th>
              <th>Area (ha)</th>
              <th>Adoption (0–1)</th>
              <th>Labour ($/ha)</th>
              <th>Materials ($/ha)</th>
              <th>Services ($/ha)</th>
              <th>Capital cost ($, year 0)</th>
              <th>Total cost ($/ha)</th>
              ${headerDeltaCols}
              <th>Notes</th>
              <th></th>
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>

      <div class="tool2-actions no-print">
        <button id="addTreatment" type="button">Add treatment</button>
      </div>
    `;

    setHTML(host, table);

    const addBtn = $("#addTreatment");
    if (addBtn) {
      addBtn.onclick = () => {
        const t = {
          id: uid(),
          name: "New treatment",
          isControl: false,
          area: 100,
          adoption: 1,
          deltas: {},
          labourCost: 0,
          materialsCost: 0,
          servicesCost: 0,
          capitalCost: 0,
          notes: ""
        };
        model.outputs.forEach(o => (t.deltas[o.id] = 0));
        model.treatments.push(t);
        ensureTreatmentDeltas();
        saveModel();
        renderTreatments();
        renderResultsComparison();
      };
    }
  }

  // -----------------------------
  // EVENTS: Inputs update model live
  // -----------------------------
  function bindLiveInputs() {
    // Project/settings fields
    document.addEventListener("input", e => {
      const t = e.target;
      if (!t) return;

      // Discount schedule inputs
      if (t.dataset && t.dataset.discPeriod !== undefined) {
        const idx = +t.dataset.discPeriod;
        const scenario = t.dataset.scenario;
        if (!model.time.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
        const row = model.time.discountSchedule[idx];
        if (row && scenario) {
          const v = parseNumber(t.value);
          if (scenario === "low") row.low = v;
          if (scenario === "base") row.base = v;
          if (scenario === "high") row.high = v;
          saveModel();
          renderResultsComparison();
        }
        return;
      }

      // Outputs
      if (t.matches("[data-out-field]")) {
        const id = t.dataset.outId;
        const field = t.dataset.outField;
        const o = model.outputs.find(x => x.id === id);
        if (!o) return;
        if (field === "value") o.value = parseNumber(t.value) || 0;
        else o[field] = t.value;
        ensureTreatmentDeltas();
        saveModel();
        renderTreatments();
        renderResultsComparison();
        return;
      }

      // Treatments (fields)
      if (t.matches("[data-trt-field]")) {
        const id = t.dataset.trtId;
        const field = t.dataset.trtField;
        const tr = model.treatments.find(x => x.id === id);
        if (!tr) return;

        if (field === "isControl") {
          // handled on change/radio, not here
          return;
        }

        if (field === "area" || field === "adoption") {
          const v = parseNumber(t.value);
          tr[field] = Number.isFinite(v) ? (field === "adoption" ? clamp(v, 0, 1) : v) : tr[field];
        } else if (field === "labourCost" || field === "materialsCost" || field === "servicesCost" || field === "capitalCost") {
          const v = parseNumber(t.value);
          tr[field] = Number.isFinite(v) ? v : 0;
        } else {
          tr[field] = t.value;
        }

        saveModel();
        // Re-render treatments to update Total cost ($/ha) column instantly
        renderTreatments();
        renderResultsComparison();
        return;
      }

      // Treatment deltas
      if (t.matches("[data-trt-delta]")) {
        const trtId = t.dataset.trtId;
        const outId = t.dataset.outId;
        const tr = model.treatments.find(x => x.id === trtId);
        if (!tr) return;
        const v = parseNumber(t.value);
        tr.deltas[outId] = Number.isFinite(v) ? v : 0;
        saveModel();
        renderResultsComparison();
        return;
      }

      // Global scenario params
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

        case "startYear": model.time.startYear = parseNumber(t.value) || model.time.startYear; break;
        case "years": model.time.years = parseNumber(t.value) || model.time.years; break;
        case "discBase": model.time.discBase = parseNumber(t.value) || model.time.discBase; break;
        case "discLow": model.time.discLow = parseNumber(t.value) || model.time.discLow; break;
        case "discHigh": model.time.discHigh = parseNumber(t.value) || model.time.discHigh; break;
        case "mirrFinance": model.time.mirrFinance = parseNumber(t.value) || model.time.mirrFinance; break;
        case "mirrReinvest": model.time.mirrReinvest = parseNumber(t.value) || model.time.mirrReinvest; break;

        case "adoptBase": model.adoption.base = clamp(parseNumber(t.value) || model.adoption.base, 0, 1); break;
        case "adoptLow": model.adoption.low = clamp(parseNumber(t.value) || model.adoption.low, 0, 1); break;
        case "adoptHigh": model.adoption.high = clamp(parseNumber(t.value) || model.adoption.high, 0, 1); break;

        case "riskBase": model.risk.base = clamp(parseNumber(t.value) || model.risk.base, 0, 1); break;
        case "riskLow": model.risk.low = clamp(parseNumber(t.value) || model.risk.low, 0, 1); break;
        case "riskHigh": model.risk.high = clamp(parseNumber(t.value) || model.risk.high, 0, 1); break;

        case "simN": model.sim.n = parseNumber(t.value) || model.sim.n; break;
        case "simVarPct": model.sim.variationPct = parseNumber(t.value) || model.sim.variationPct; break;
        case "simVaryOutputs": model.sim.varyOutputs = String(t.value) === "true"; break;
        case "simVaryTreatCosts": model.sim.varyTreatCosts = String(t.value) === "true"; break;
        case "randSeed": model.sim.seed = t.value ? parseNumber(t.value) : null; break;
      }

      saveModel();
      renderResultsComparison();
    });

    // Control radio change
    document.addEventListener("change", e => {
      const t = e.target;
      if (!t) return;

      if (t.matches("input[type='radio'][name='controlRadio'][data-trt-field='isControl']")) {
        const id = t.dataset.trtId;
        model.treatments.forEach(tr => (tr.isControl = tr.id === id));
        // Enforce control deltas = 0
        const ctrl = model.treatments.find(x => x.isControl);
        if (ctrl) model.outputs.forEach(o => (ctrl.deltas[o.id] = 0));
        saveModel();
        renderTreatments();
        renderResultsComparison();
        showToast("Control updated.");
      }
    });

    // Remove rows
    document.addEventListener("click", e => {
      const btn = e.target.closest("[data-trt-action='del'], [data-out-action='del']");
      if (!btn) return;

      if (btn.dataset.trtAction === "del") {
        const id = btn.dataset.trtId;
        const idx = model.treatments.findIndex(x => x.id === id);
        if (idx >= 0) {
          const wasControl = !!model.treatments[idx].isControl;
          model.treatments.splice(idx, 1);
          if (wasControl && model.treatments.length) model.treatments[0].isControl = true;
          ensureTreatmentDeltas();
          saveModel();
          renderTreatments();
          renderResultsComparison();
          showToast("Treatment removed.");
        }
      }

      if (btn.dataset.outAction === "del") {
        const id = btn.dataset.outId;
        const idx = model.outputs.findIndex(x => x.id === id);
        if (idx >= 0) {
          model.outputs.splice(idx, 1);
          ensureTreatmentDeltas();
          saveModel();
          renderOutputs();
          renderTreatments();
          renderResultsComparison();
          showToast("Output removed.");
        }
      }
    });

    // Main actions if present
    document.addEventListener("click", e => {
      const el = e.target.closest("#recalc, #getResults, [data-action='recalc']");
      if (!el) return;
      e.preventDefault();
      renderResultsComparison();
      showToast("Results recalculated.");
    });
  }

  // -----------------------------
  // EXCEL WORKFLOW BUTTONS (IDs from your existing UI; robust to missing)
  // -----------------------------
  function bindExcelButtons() {
    const parseExcelBtn = $("#parseExcel");
    const importExcelBtn = $("#importExcel");
    const downloadTemplateBtn = $("#downloadTemplate");
    const downloadSampleBtn = $("#downloadSample");

    if (parseExcelBtn) parseExcelBtn.addEventListener("click", e => { e.preventDefault(); handleParseExcel(); });
    if (importExcelBtn) importExcelBtn.addEventListener("click", e => { e.preventDefault(); commitExcelToModel(); });

    if (downloadTemplateBtn) downloadTemplateBtn.addEventListener("click", e => { e.preventDefault(); downloadExcelTemplate({ includeSample: false }); });
    if (downloadSampleBtn) downloadSampleBtn.addEventListener("click", e => { e.preventDefault(); downloadExcelTemplate({ includeSample: true }); });

    const help = $("#excelWorkflowHelp") || $("#excelHelp") || $("#excelInstructions");
    if (help) {
      setHTML(
        help,
        `
          <div class="tool2-help">
            <strong>Excel-first workflow (recommended)</strong><br/>
            1) Download a template (sample or scenario-specific).<br/>
            2) Edit rows in Excel (treatments, costs, and deltas vs control).<br/>
            3) Save the file and upload it here.<br/>
            4) Click Parse, then Apply to tool. Results update automatically and become the new default for analysis.
          </div>
        `
      );
    }
  }

  // -----------------------------
  // OPTIONAL: SIMULATION (kept lightweight; supports “why underperform” exploration)
  // -----------------------------
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

  function runSimulation() {
    const host = $("#simResults") || $("#simulationResults");
    if (!host) {
      showToast("Simulation panel not found in the page.");
      return;
    }

    const rate = Number($("#discBase")?.value ?? model.time.discBase) || model.time.discBase;
    const years = Number($("#years")?.value ?? model.time.years) || model.time.years;
    const adoptMul = Number($("#adoptBase")?.value ?? model.adoption.base) ?? model.adoption.base;
    const risk = Number($("#riskBase")?.value ?? model.risk.base) ?? model.risk.base;

    const N = Math.max(100, Math.floor(Number(model.sim.n) || 1000));
    const varPct = clamp(Number(model.sim.variationPct) || 20, 0, 200) / 100;
    const varyOutputs = !!model.sim.varyOutputs;
    const varyCosts = !!model.sim.varyTreatCosts;

    const rand = rng(model.sim.seed || null);

    const baseComp = computeComparison({ rate, years, adoptMul, risk });
    const treatments = baseComp.rows;

    // Track: probability best by ΔNPV; distribution of ΔNPV for each
    const bestCount = new Map(treatments.map(t => [t.id, 0]));
    const deltaNpvSums = new Map(treatments.map(t => [t.id, 0]));
    const deltaNpvSums2 = new Map(treatments.map(t => [t.id, 0]));
    const bcrAbove1 = new Map(treatments.map(t => [t.id, 0]));

    // Pre-cache original values
    const baseOutputValues = model.outputs.map(o => o.value);
    const baseTreatCosts = model.treatments.map(t => ({
      labourCost: t.labourCost,
      materialsCost: t.materialsCost,
      servicesCost: t.servicesCost,
      capitalCost: t.capitalCost
    }));

    for (let i = 0; i < N; i++) {
      // Perturb outputs
      if (varyOutputs) {
        model.outputs.forEach((o, j) => {
          const u = (rand() * 2 - 1) * varPct;
          o.value = baseOutputValues[j] * (1 + u);
        });
      }

      // Perturb costs
      if (varyCosts) {
        model.treatments.forEach((t, j) => {
          const u1 = (rand() * 2 - 1) * varPct;
          const u2 = (rand() * 2 - 1) * varPct;
          const u3 = (rand() * 2 - 1) * varPct;
          const u4 = (rand() * 2 - 1) * varPct;
          t.labourCost = baseTreatCosts[j].labourCost * (1 + u1);
          t.materialsCost = baseTreatCosts[j].materialsCost * (1 + u2);
          t.servicesCost = baseTreatCosts[j].servicesCost * (1 + u3);
          t.capitalCost = baseTreatCosts[j].capitalCost * (1 + u4);
        });
      }

      const comp = computeComparison({ rate, years, adoptMul, risk });
      const nonCtrl = comp.rows.filter(r => !r.isControl);

      // Best by ΔNPV
      let best = null;
      nonCtrl.forEach(r => {
        if (!best || (isFinite(r.inc.npv) && r.inc.npv > best.inc.npv)) best = r;
      });
      if (best) bestCount.set(best.id, bestCount.get(best.id) + 1);

      // Moments
      comp.rows.forEach(r => {
        const d = r.isControl ? 0 : (r.inc.npv || 0);
        deltaNpvSums.set(r.id, deltaNpvSums.get(r.id) + d);
        deltaNpvSums2.set(r.id, deltaNpvSums2.get(r.id) + d * d);
        const bcr = r.isControl ? 0 : r.inc.bcr;
        if (isFinite(bcr) && bcr > 1) bcrAbove1.set(r.id, bcrAbove1.get(r.id) + 1);
      });
    }

    // Restore originals
    model.outputs.forEach((o, j) => (o.value = baseOutputValues[j]));
    model.treatments.forEach((t, j) => {
      t.labourCost = baseTreatCosts[j].labourCost;
      t.materialsCost = baseTreatCosts[j].materialsCost;
      t.servicesCost = baseTreatCosts[j].servicesCost;
      t.capitalCost = baseTreatCosts[j].capitalCost;
    });

    const summary = treatments.map(r => {
      const mean = deltaNpvSums.get(r.id) / N;
      const m2 = deltaNpvSums2.get(r.id) / N;
      const sd = Math.sqrt(Math.max(m2 - mean * mean, 0));
      return {
        treatment: r.name,
        is_control: r.isControl ? "Yes" : "No",
        prob_best_by_delta_npv: r.isControl ? 0 : bestCount.get(r.id) / N,
        mean_delta_npv: r.isControl ? 0 : mean,
        sd_delta_npv: r.isControl ? 0 : sd,
        prob_delta_bcr_gt_1: r.isControl ? 0 : bcrAbove1.get(r.id) / N
      };
    });

    // Render table
    const rowsHtml = summary
      .map(s => `
        <tr>
          <td>${esc(s.treatment)}</td>
          <td>${esc(s.is_control)}</td>
          <td>${isFinite(s.prob_best_by_delta_npv) ? pct(s.prob_best_by_delta_npv * 100) : "n/a"}</td>
          <td>${money(s.mean_delta_npv)}</td>
          <td>${money(s.sd_delta_npv)}</td>
          <td>${isFinite(s.prob_delta_bcr_gt_1) ? pct(s.prob_delta_bcr_gt_1 * 100) : "n/a"}</td>
        </tr>
      `).join("");

    setHTML(
      host,
      `
      <div class="tool2-help tool2-small">
        Simulation varies unit values and/or costs by ±${fmt(varPct * 100)}% to explore uncertainty.
        It reports the probability a treatment is best by ΔNPV vs control, and how often ΔBCR exceeds 1 (interpret as guidance, not a rule).
      </div>
      <div class="tool2-table-wrap">
        <table class="tool2-compare" style="min-width:1100px">
          <thead>
            <tr>
              <th>Treatment</th>
              <th>Control</th>
              <th>Pr(best by ΔNPV)</th>
              <th>Mean ΔNPV</th>
              <th>SD ΔNPV</th>
              <th>Pr(ΔBCR &gt; 1)</th>
            </tr>
          </thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </div>
      <div class="tool2-actions no-print">
        <button id="tool2ExportSimExcel" type="button">Export simulation summary (Excel)</button>
      </div>
      `
    );

    const btn = $("#tool2ExportSimExcel");
    if (btn) {
      btn.onclick = () => {
        if (hasXLSX()) {
          const XLSX = window.XLSX;
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Simulation_Summary");
          const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
          const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
          const a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = `${slug(model.project.name)}_simulation.xlsx`;
          document.body.appendChild(a);
          a.click();
          setTimeout(() => {
            URL.revokeObjectURL(a.href);
            a.remove();
          }, 0);
          showToast("Simulation exported (Excel).");
        } else {
          downloadFile(`${slug(model.project.name)}_simulation.csv`, jsonToCSV(summary), "text/csv");
          showToast("XLSX library not found. Exported CSV instead.");
        }
      };
    }

    showToast("Simulation complete.");
  }

  function bindSimButton() {
    document.addEventListener("click", e => {
      const el = e.target.closest("#runSim, [data-action='run-sim']");
      if (!el) return;
      e.preventDefault();
      runSimulation();
    });
  }

  // -----------------------------
  // INITIALISE
  // -----------------------------
  function renderAll() {
    injectTool2CSS();
    renderOutputs();
    renderTreatments();
    renderResultsComparison();
  }

  function bindCoreButtons() {
    // PDF export buttons if present
    const exportPdfBtn = $("#exportPdf") || $("#exportPdfFoot");
    if (exportPdfBtn) exportPdfBtn.addEventListener("click", e => { e.preventDefault(); window.print(); });

    // CSV export (fallback if your existing UI expects it)
    const exportCsvBtn = $("#exportCsv") || $("#exportCsvFoot");
    if (exportCsvBtn) exportCsvBtn.addEventListener("click", e => {
      e.preventDefault();
      const comp = computeComparison({
        rate: Number($("#discBase")?.value ?? model.time.discBase) || model.time.discBase,
        years: Number($("#years")?.value ?? model.time.years) || model.time.years,
        adoptMul: Number($("#adoptBase")?.value ?? model.adoption.base) ?? model.adoption.base,
        risk: Number($("#riskBase")?.value ?? model.risk.base) ?? model.risk.base
      });
      const rows = comp.rows.map(r => ({
        treatment: r.name,
        is_control: r.isControl,
        rank: r.isControl ? 1 : r.rank,
        delta_npv: r.isControl ? 0 : r.inc.npv,
        delta_pv_benefits: r.isControl ? 0 : r.inc.pvBenefits,
        delta_pv_costs: r.isControl ? 0 : r.inc.pvCosts,
        delta_bcr: r.isControl ? 0 : r.inc.bcr,
        delta_roi_pct: r.isControl ? 0 : r.inc.roi
      }));
      downloadFile(`${slug(model.project.name)}_results.csv`, jsonToCSV(rows), "text/csv");
      showToast("Results exported (CSV).");
    });

    // Save/load JSON if you keep those buttons
    const saveProjectBtn = $("#saveProject");
    if (saveProjectBtn) {
      saveProjectBtn.addEventListener("click", e => {
        e.preventDefault();
        downloadFile(`${slug(model.project.name)}.json`, JSON.stringify(model, null, 2), "application/json");
        showToast("Project JSON downloaded.");
      });
    }

    const loadFileInput = $("#loadFile");
    const loadProjectBtn = $("#loadProject");
    if (loadProjectBtn && loadFileInput) {
      loadProjectBtn.addEventListener("click", e => {
        e.preventDefault();
        loadFileInput.click();
      });
      loadFileInput.addEventListener("change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          Object.assign(model, obj);
          if (!model.toolName) model.toolName = TOOL_NAME;
          if (!model.time?.discountSchedule) model.time.discountSchedule = JSON.parse(JSON.stringify(DEFAULT_DISCOUNT_SCHEDULE));
          ensureTreatmentDeltas();
          saveModel();
          setBasicsFieldsFromModel();
          renderAll();
          showToast("Project JSON loaded.");
        } catch (err) {
          alert("Invalid JSON file.");
          console.error(err);
        } finally {
          e.target.value = "";
        }
      });
    }
  }

  function ensureToolNameInUI() {
    const titleEls = ["#toolTitle", "#appTitle", "h1[data-app-title]"].map(s => $(s)).filter(Boolean);
    titleEls.forEach(el => (el.textContent = TOOL_NAME));
    document.title = TOOL_NAME;
  }

  // -----------------------------
  // BOOT
  // -----------------------------
  function boot() {
    injectTool2CSS();
    bindTooltips();
    ensureToolNameInUI();

    const loaded = loadModel();
    if (!loaded) {
      seedDefaultTreatmentsIfNeeded();
      ensureTreatmentDeltas();
      saveModel();
    } else {
      // If a user loaded old data without many treatments, still seed fallback.
      seedDefaultTreatmentsIfNeeded();
      ensureTreatmentDeltas();
      saveModel();
    }

    setBasicsFieldsFromModel();
    initTabs();
    bindLiveInputs();
    bindExcelButtons();
    bindSimButton();
    bindCoreButtons();

    // Render
    renderAll();

    // Wire any start button
    const startBtn = $("#startBtn");
    if (startBtn) {
      startBtn.addEventListener("click", e => {
        e.preventDefault();
        switchTab("project");
        showToast("Welcome. Start with the Project tab, then upload your Excel template.");
      });
    }

    showToast(`${TOOL_NAME} ready. Upload Excel to set your default dataset.`);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
