/* app.js (FULL) — Farming CBA Decision Tool 2
   Results matrix vs control + Excel-first workflow + AI prompt + exports + simulation
   Dependencies: xlsx.full.min.js (already loaded in index.html)
*/
(() => {
  "use strict";

  /********************************************************************
   * Constants + utilities
   ********************************************************************/
  const TOOL_NAME = "Farming CBA Decision Tool 2";
  const ORG_NAME = "Newcastle Business School";

  const REQUIRED_SHEETS = [
    "Project",
    "Settings",
    "Outputs",
    "Treatments",
    "TreatmentOutputs",
    "Benefits",
    "Costs",
  ];

  const fmtCurrency = new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: "AUD",
    maximumFractionDigits: 0,
  });
  const fmtCurrency2 = new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: "AUD",
    maximumFractionDigits: 2,
  });
  const fmtNumber2 = new Intl.NumberFormat("en-AU", { maximumFractionDigits: 2 });
  const fmtNumber3 = new Intl.NumberFormat("en-AU", { maximumFractionDigits: 3 });
  const fmtPercent1 = new Intl.NumberFormat("en-AU", { style: "percent", maximumFractionDigits: 1 });

  const clamp = (x, a, b) => Math.min(b, Math.max(a, x));
  const isFiniteNumber = (x) => typeof x === "number" && Number.isFinite(x);

  function uid(prefix = "id") {
    try {
      if (crypto && typeof crypto.randomUUID === "function") return `${prefix}_${crypto.randomUUID()}`;
    } catch (_) {}
    return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
  }

  function safeNum(v, fallback = 0) {
    if (v === null || v === undefined) return fallback;
    if (typeof v === "number") return Number.isFinite(v) ? v : fallback;
    const s = String(v).trim();
    if (s === "") return fallback;
    const x = Number(s);
    return Number.isFinite(x) ? x : fallback;
  }

  function safeStr(v, fallback = "") {
    if (v === null || v === undefined) return fallback;
    return String(v);
  }

  function ynToBool(v, fallback = false) {
    if (typeof v === "boolean") return v;
    const s = String(v ?? "").trim().toLowerCase();
    if (["true", "t", "yes", "y", "1"].includes(s)) return true;
    if (["false", "f", "no", "n", "0"].includes(s)) return false;
    return fallback;
  }

  function downloadBlob(filename, blob) {
    const a = document.createElement("a");
    const url = URL.createObjectURL(blob);
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 800);
  }

  function downloadText(filename, text, mime = "text/plain;charset=utf-8") {
    downloadBlob(filename, new Blob([text], { type: mime }));
  }

  function todayISO() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  function toast(msg, kind = "info", ttl = 3500) {
    const root = document.getElementById("toast-root");
    if (!root) return;

    const el = document.createElement("div");
    el.className = `toast ${kind}`;
    el.setAttribute("role", "status");
    el.innerHTML = `
      <div class="toast-inner">
        <div class="toast-title">${TOOL_NAME}</div>
        <div class="toast-msg">${escapeHtml(msg)}</div>
      </div>
      <button class="toast-x" type="button" aria-label="Dismiss">×</button>
    `;
    root.appendChild(el);

    const kill = () => {
      el.style.opacity = "0";
      el.style.transform = "translateY(-4px)";
      setTimeout(() => el.remove(), 220);
    };
    el.querySelector(".toast-x")?.addEventListener("click", kill);
    setTimeout(kill, ttl);
  }

  function escapeHtml(s) {
    return String(s)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function setLoading(on, text = "Working") {
    const overlay = document.getElementById("appLoading");
    if (!overlay) return;
    if (!on) {
      overlay.setAttribute("data-hidden", "true");
      overlay.style.opacity = "0";
      overlay.style.pointerEvents = "none";
      setTimeout(() => {
        if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
      }, 450);
      return;
    }
    overlay.removeAttribute("data-hidden");
    overlay.style.opacity = "1";
    overlay.style.pointerEvents = "auto";
    const sub = overlay.querySelector(".app-loading-subtitle");
    if (sub) sub.textContent = text;
  }

  function el(tag, attrs = {}, children = []) {
    const node = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === "class") node.className = v;
      else if (k === "html") node.innerHTML = v;
      else if (k.startsWith("on") && typeof v === "function") node.addEventListener(k.slice(2), v);
      else if (v === null || v === undefined) return;
      else node.setAttribute(k, String(v));
    });
    (Array.isArray(children) ? children : [children]).forEach((c) => {
      if (c === null || c === undefined) return;
      if (typeof c === "string") node.appendChild(document.createTextNode(c));
      else node.appendChild(c);
    });
    return node;
  }

  function q(id) {
    return document.getElementById(id);
  }

  function setInputValue(id, value) {
    const node = q(id);
    if (!node) return;
    if (node.type === "checkbox") node.checked = Boolean(value);
    else node.value = value ?? "";
  }

  function getInputValue(id) {
    const node = q(id);
    if (!node) return null;
    if (node.type === "checkbox") return node.checked;
    return node.value;
  }

  function parseYearWithinHorizon(v, startYear, years) {
    const y = safeNum(v, NaN);
    if (!Number.isFinite(y)) return null;
    const minY = startYear;
    const maxY = startYear + years - 1;
    if (y < minY || y > maxY) return null;
    return Math.trunc(y);
  }

  function rangeToIndices(yearFrom, yearTo, startYear, years) {
    const yf = parseYearWithinHorizon(yearFrom, startYear, years);
    const yt = parseYearWithinHorizon(yearTo, startYear, years);
    if (yf === null || yt === null) return null;
    const a = Math.min(yf, yt);
    const b = Math.max(yf, yt);
    const i0 = a - startYear;
    const i1 = b - startYear;
    return { i0, i1 };
  }

  function discountFactor(r, t) {
    return 1 / Math.pow(1 + r, t);
  }

  function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
  }

  /********************************************************************
   * Default sample state (faba bean example)
   ********************************************************************/
  function buildSampleState() {
    const out1 = { id: uid("out"), name: "Yield increase", unit: "t/ha", valuePerUnit: 450, notes: "Farm gate price", source: "Example data" };
    const out2 = { id: uid("out"), name: "Nitrogen saving", unit: "kg N/ha", valuePerUnit: 2.2, notes: "Equivalent fertiliser value", source: "Example data" };

    const t0 = {
      id: uid("trt"),
      name: "Control (business as usual)",
      isControl: true,
      maxAreaHa: 100,
      labourCostPerHa: 0,
      inputCostPerHa: 0,
      notes: "Baseline management",
      source: "Example",
      outputs: [
        { id: uid("to"), outputId: out1.id, qtyPerHa: 0.0, yearFrom: 2026, yearTo: 2030, notes: "" },
        { id: uid("to"), outputId: out2.id, qtyPerHa: 0.0, yearFrom: 2026, yearTo: 2030, notes: "" },
      ],
    };

    const t1 = {
      id: uid("trt"),
      name: "Soil amendment A",
      isControl: false,
      maxAreaHa: 100,
      labourCostPerHa: 18,
      inputCostPerHa: 55,
      notes: "Moderate cost, moderate uplift",
      source: "Example",
      outputs: [
        { id: uid("to"), outputId: out1.id, qtyPerHa: 0.35, yearFrom: 2026, yearTo: 2030, notes: "Trial mean" },
        { id: uid("to"), outputId: out2.id, qtyPerHa: 12, yearFrom: 2026, yearTo: 2030, notes: "Assumed saving" },
      ],
    };

    const t2 = {
      id: uid("trt"),
      name: "Soil amendment B",
      isControl: false,
      maxAreaHa: 100,
      labourCostPerHa: 22,
      inputCostPerHa: 85,
      notes: "Higher cost, higher uplift",
      source: "Example",
      outputs: [
        { id: uid("to"), outputId: out1.id, qtyPerHa: 0.55, yearFrom: 2026, yearTo: 2030, notes: "Trial mean" },
        { id: uid("to"), outputId: out2.id, qtyPerHa: 15, yearFrom: 2026, yearTo: 2030, notes: "Assumed saving" },
      ],
    };

    const b1 = {
      id: uid("ben"),
      name: "Avoided disease loss (project-wide)",
      annualValue: 3000,
      yearFrom: 2027,
      yearTo: 2030,
      notes: "Optional benefit not captured by output rows",
      source: "Example",
    };

    const c1 = {
      id: uid("cst"),
      name: "Training and extension (project-wide)",
      annualValue: 2500,
      yearFrom: 2026,
      yearTo: 2027,
      notes: "Coordination and training",
      source: "Example",
    };

    return {
      version: "2.0.0",
      toolName: TOOL_NAME,
      createdAt: new Date().toISOString(),
      project: {
        projectName: "Faba bean soil amendment trial",
        projectLead: "",
        analystNames: "",
        projectTeam: "",
        organisation: ORG_NAME,
        lastUpdated: todayISO(),
        contactEmail: "",
        contactPhone: "",
        projectSummary: "Example scenario for demonstration. Replace via Excel workflow for your own farm context.",
        projectGoal: "Improve yields and soil condition with practical on-farm treatments.",
        withProject: "Adopt selected amendment(s) on the target area; benefits realised through yield gains and input savings.",
        withoutProject: "Continue current practice (control) with no additional costs or benefits.",
        projectObjectives: "Compare treatments against control using a snapshot results table and exportable indicators.",
        projectActivities: "Trial management, input procurement, application, monitoring.",
        stakeholderGroups: "Farmers, advisors, trial partners.",
      },
      settings: {
        startYear: 2026,
        projectStartYear: 2026,
        years: 5,
        systemType: "single",
        discBase: 7.0,
        discLow: 3.0,
        discHigh: 10.0,
        outputAssumptions: "Values are illustrative. Replace with local prices and trial estimates.",
        mirrFinance: 7.0,
        mirrReinvest: 7.0,
        adoptLow: 0.6,
        adoptBase: 0.85,
        adoptHigh: 1.0,
        allocateProjectItems: true,
        riskLow: 0.05,
        riskBase: 0.12,
        riskHigh: 0.25,
        rTech: 0.05,
        rNonCoop: 0.03,
        rSocio: 0.02,
        rFin: 0.04,
        rMan: 0.03,
        simN: 1500,
        targetBCR: 2,
        randSeed: "",
        simVarPct: 10,
        simVaryOutputs: true,
        simVaryTreatCosts: true,
        simVaryInputCosts: false,
      },
      outputs: [out1, out2],
      treatments: [t0, t1, t2],
      benefits: [b1],
      costs: [c1],
      ui: {
        activeTab: "intro",
        rankMetric: "bcr",
        resultsViewMode: "absolute",
        columnsPerPage: 6,
        showDeltas: true,
        matrixStart: 0,
      },
      cache: {
        lastResults: null,
        lastResultsMeta: null,
      },
      excel: {
        parsed: null,
        parsedSummary: null,
      },
    };
  }

  let STATE = buildSampleState();

  /********************************************************************
   * Core model: build annual cashflows and PV indicators
   ********************************************************************/
  function ensureSingleControl(state) {
    const trts = state.treatments || [];
    if (trts.length === 0) return;
    const controls = trts.filter((t) => Boolean(t.isControl));
    if (controls.length === 0) {
      trts[0].isControl = true;
      toast("No control flagged. The first treatment has been set as the control.", "warn", 4500);
      return;
    }
    if (controls.length > 1) {
      const keep = controls[0].id;
      trts.forEach((t) => (t.isControl = t.id === keep));
      toast("Multiple controls were flagged. Only the first control has been kept.", "warn", 4500);
    }
  }

  function getControl(state) {
    ensureSingleControl(state);
    return (state.treatments || []).find((t) => t.isControl) || null;
  }

  function buildLookups(state) {
    const outById = new Map((state.outputs || []).map((o) => [o.id, o]));
    const trtById = new Map((state.treatments || []).map((t) => [t.id, t]));
    return { outById, trtById };
  }

  function effectiveAreaHa(treatment, adoptionMultiplier) {
    const area = safeNum(treatment.maxAreaHa, 0);
    const adopt = clamp(safeNum(adoptionMultiplier, 1), 0, 1);
    return area * adopt;
  }

  function allocateShareByArea(state, treatId, adoptionMultiplier) {
    const trts = state.treatments || [];
    const total = trts.reduce((s, t) => s + effectiveAreaHa(t, adoptionMultiplier), 0);
    if (total <= 0) return 0;
    const t = trts.find((x) => x.id === treatId);
    if (!t) return 0;
    return effectiveAreaHa(t, adoptionMultiplier) / total;
  }

  function buildAnnualStreamsForTreatment(state, treatment) {
    const s = state.settings;
    const startYear = Math.trunc(safeNum(s.startYear, new Date().getFullYear()));
    const years = Math.max(1, Math.trunc(safeNum(s.years, 5)));
    const adoption = clamp(safeNum(s.adoptBase, 1), 0, 1);
    const risk = clamp(safeNum(s.riskBase, 0), 0, 1);
    const allocate = Boolean(s.allocateProjectItems);

    const { outById } = buildLookups(state);

    const effArea = effectiveAreaHa(treatment, adoption);
    const areaShare = allocate ? allocateShareByArea(state, treatment.id, adoption) : 1;

    const benefit = new Array(years).fill(0);
    const cost = new Array(years).fill(0);

    // Treatment variable costs per ha (labour + inputs), applied each year from projectStartYear onwards
    const projectStartYear = Math.trunc(safeNum(s.projectStartYear, startYear));
    const projectStartIndex = clamp(projectStartYear - startYear, 0, years - 1);

    const labour = safeNum(treatment.labourCostPerHa, 0);
    const inputs = safeNum(treatment.inputCostPerHa, 0);
    const varCostAnnual = (labour + inputs) * effArea;
    for (let t = projectStartIndex; t < years; t++) cost[t] += varCostAnnual;

    // Output-based benefits (qty/ha * value/unit * effArea), within year ranges
    (treatment.outputs || []).forEach((row) => {
      const out = outById.get(row.outputId);
      if (!out) return;
      const r = rangeToIndices(row.yearFrom, row.yearTo, startYear, years);
      if (!r) return;
      const qty = safeNum(row.qtyPerHa, 0);
      const val = safeNum(out.valuePerUnit, 0);
      const annual = qty * val * effArea;
      for (let t = r.i0; t <= r.i1; t++) benefit[t] += annual;
    });

    // Project-wide additional benefits and costs (allocated or applied to all equally)
    (state.benefits || []).forEach((b) => {
      const r = rangeToIndices(b.yearFrom, b.yearTo, startYear, years);
      if (!r) return;
      const annual = safeNum(b.annualValue, 0) * areaShare;
      for (let t = r.i0; t <= r.i1; t++) benefit[t] += annual;
    });

    (state.costs || []).forEach((c) => {
      const r = rangeToIndices(c.yearFrom, c.yearTo, startYear, years);
      if (!r) return;
      const annual = safeNum(c.annualValue, 0) * areaShare;
      for (let t = r.i0; t <= r.i1; t++) cost[t] += annual;
    });

    // Risk reduces realised benefits (not costs)
    for (let t = 0; t < years; t++) benefit[t] = benefit[t] * (1 - risk);

    return { benefit, cost, effArea, adoption, risk, startYear, years };
  }

  function pvFromStream(stream, discRate) {
    const r = clamp(safeNum(discRate, 0), -0.99, 10);
    let pv = 0;
    for (let t = 0; t < stream.length; t++) pv += stream[t] * discountFactor(r, t);
    return pv;
  }

  function computeIndicatorsForTreatment(state, treatment) {
    const s = state.settings;
    const discBasePct = safeNum(s.discBase, 7.0);
    const r = discBasePct / 100;

    const { benefit, cost, effArea, adoption, risk } = buildAnnualStreamsForTreatment(state, treatment);
    const pvB = pvFromStream(benefit, r);
    const pvC = pvFromStream(cost, r);
    const npv = pvB - pvC;

    // Ratios: handle divide-by-zero robustly
    const bcr = pvC > 0 ? pvB / pvC : (pvB > 0 ? Infinity : 0);
    const roi = pvC > 0 ? npv / pvC : (npv > 0 ? Infinity : 0);

    return {
      treatmentId: treatment.id,
      treatmentName: treatment.name,
      isControl: Boolean(treatment.isControl),
      effArea,
      adoption,
      risk,
      pvBenefits: pvB,
      pvCosts: pvC,
      npv,
      bcr,
      roi,
      annualBenefit: benefit,
      annualCost: cost,
    };
  }

  function computeAllResults(state) {
    ensureSingleControl(state);
    const indicators = (state.treatments || []).map((t) => computeIndicatorsForTreatment(state, t));

    const rankMetric = (state.ui?.rankMetric || "bcr").toLowerCase();
    const metricKey = rankMetric === "npv" ? "npv" : rankMetric === "roi" ? "roi" : "bcr";

    const ranked = [...indicators].sort((a, b) => {
      const av = a[metricKey];
      const bv = b[metricKey];
      const aN = Number.isFinite(av) ? av : -Infinity;
      const bN = Number.isFinite(bv) ? bv : -Infinity;
      return bN - aN;
    });

    // Assign ranks (1 best). Stable tie-handling via metric + name.
    const ranks = new Map();
    ranked.forEach((row, i) => ranks.set(row.treatmentId, i + 1));
    indicators.forEach((row) => (row.rank = ranks.get(row.treatmentId) || null));

    const control = indicators.find((x) => x.isControl) || null;

    return {
      generatedAt: new Date().toISOString(),
      metricKey,
      indicators,
      control,
    };
  }

  /********************************************************************
   * Results matrix rendering (control alongside treatments)
   ********************************************************************/
  const MATRIX_ROWS = [
    { key: "npv", label: "Net present value (NPV)", fmt: (x) => fmtCurrency.format(x) },
    { key: "pvBenefits", label: "Present value of benefits (PV benefits)", fmt: (x) => fmtCurrency.format(x) },
    { key: "pvCosts", label: "Present value of costs (PV costs)", fmt: (x) => fmtCurrency.format(x) },
    { key: "bcr", label: "Benefit–cost ratio (BCR)", fmt: (x) => (x === Infinity ? "∞" : fmtNumber3.format(x)) },
    { key: "roi", label: "Return on investment (ROI)", fmt: (x) => (x === Infinity ? "∞" : fmtNumber3.format(x)) },
    { key: "rank", label: "Ranking (selected metric)", fmt: (x) => (x == null ? "n/a" : String(x)) },
  ];

  function buildMatrixData(state, results) {
    const viewMode = state.ui?.resultsViewMode || "absolute";
    const showDeltas = Boolean(state.ui?.showDeltas);
    const control = results.control;

    const cols = results.indicators
      .slice()
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;
        return (a.rank ?? 9999) - (b.rank ?? 9999);
      });

    // In relative mode, show deltas vs control; control column stays but shows baseline marker
    const rows = MATRIX_ROWS.map((r) => {
      const cells = cols.map((c) => {
        if (viewMode === "relative" && control) {
          if (c.isControl) {
            return { value: 0, display: "0", delta: null, isControl: true };
          }
          const base = safeNum(control[r.key], 0);
          const val = safeNum(c[r.key], 0);
          const diff = val - base;
          let display;
          if (r.key === "rank") display = String((c.rank ?? 0) - (control.rank ?? 0));
          else if (r.key === "bcr" || r.key === "roi") display = (diff === Infinity ? "∞" : fmtNumber3.format(diff));
          else display = r.fmt(diff);
          return { value: diff, display, delta: null, isControl: false };
        }

        // Absolute mode
        const val = c[r.key];
        const display = r.key === "bcr" || r.key === "roi"
          ? (val === Infinity ? "∞" : fmtNumber3.format(val))
          : r.fmt(val);

        // Inline delta vs control (optional)
        let delta = null;
        if (showDeltas && control && !c.isControl && r.key !== "rank") {
          const base = safeNum(control[r.key], 0);
          const v = safeNum(c[r.key], 0);
          const diff = v - base;
          if (r.key === "bcr" || r.key === "roi") delta = diff === Infinity ? "Δ ∞" : `Δ ${fmtNumber3.format(diff)}`;
          else delta = `Δ ${fmtCurrency.format(diff)}`;
        }
        if (showDeltas && control && !c.isControl && r.key === "rank") {
          const diff = (c.rank ?? 0) - (control.rank ?? 0);
          delta = `Δ ${diff}`;
        }

        return { value: val, display, delta, isControl: c.isControl };
      });

      return { key: r.key, label: r.label, cells };
    });

    return { cols, rows, viewMode };
  }

  function renderResultsMatrix(state, results) {
    const wrap = q("resultsMatrixScroll");
    if (!wrap) return;

    const ui = state.ui || {};
    const columnsPerPage = Math.max(1, Math.trunc(safeNum(ui.columnsPerPage, 6)));
    const matrixStart = Math.max(0, Math.trunc(safeNum(ui.matrixStart, 0)));

    // Sort columns: control first, then by rank
    const ordered = results.indicators
      .slice()
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;
        return (a.rank ?? 9999) - (b.rank ?? 9999);
      });

    const control = ordered.find((x) => x.isControl) || null;
    const others = ordered.filter((x) => !x.isControl);

    // Control must always appear alongside treatments
    const window = others.slice(matrixStart, matrixStart + columnsPerPage);
    const cols = control ? [control, ...window] : window;

    const data = buildMatrixData(
      { ...state, treatments: state.treatments, ui: state.ui },
      { ...results, indicators: cols, control }
    );

    // Build table
    const table = el("table", { class: "matrix-table", role: "table" });

    const thead = el("thead");
    const trh = el("tr");
    trh.appendChild(el("th", { class: "sticky-col", scope: "col" }, ["Indicator"]));
    data.cols.forEach((c) => {
      const label = c.isControl ? `${c.treatmentName} (Control)` : c.treatmentName;
      const sub = `Rank ${c.rank ?? "n/a"}`;
      trh.appendChild(
        el("th", { scope: "col", class: c.isControl ? "col-control" : "" }, [
          el("div", { class: "col-title" }, [label]),
          el("div", { class: "col-sub muted small" }, [sub]),
        ])
      );
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = el("tbody");
    data.rows.forEach((r) => {
      const tr = el("tr");
      tr.appendChild(el("th", { class: "sticky-col row-label", scope: "row" }, [r.label]));
      r.cells.forEach((cell, idx) => {
        const isCtrl = data.cols[idx]?.isControl;
        const cellClass = isCtrl ? "cell control" : "cell";
        const inner = el("div", { class: "cell-inner" }, [
          el("div", { class: "cell-main" }, [cell.display]),
          cell.delta ? el("div", { class: "cell-delta small muted" }, [cell.delta]) : null,
        ]);
        tr.appendChild(el("td", { class: cellClass }, [inner]));
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    wrap.innerHTML = "";
    wrap.appendChild(table);

    // Window label and prev/next enablement
    const label = q("matrixWindowLabel");
    if (label) {
      const totalOthers = others.length;
      const from = totalOthers === 0 ? 0 : matrixStart + 1;
      const to = Math.min(totalOthers, matrixStart + columnsPerPage);
      label.textContent = `Showing control + treatments ${from} to ${to} of ${totalOthers}`;
    }

    const prev = q("matrixPrev");
    const next = q("matrixNext");
    if (prev) prev.disabled = matrixStart <= 0;
    if (next) next.disabled = matrixStart + columnsPerPage >= others.length;
  }

  function updateSnapshotCards(results) {
    const best = results.indicators.slice().sort((a, b) => (a.rank ?? 9999) - (b.rank ?? 9999))[0] || null;
    const worst = results.indicators.slice().sort((a, b) => (b.rank ?? 0) - (a.rank ?? 0))[0] || null;
    const control = results.control;

    const bestEl = q("snapBest");
    const bestSub = q("snapBestSub");
    const worstEl = q("snapWorst");
    const worstSub = q("snapWorstSub");
    const ctrlEl = q("snapControl");
    const ctrlSub = q("snapControlSub");

    if (bestEl) bestEl.textContent = best ? best.treatmentName : "n/a";
    if (worstEl) worstEl.textContent = worst ? worst.treatmentName : "n/a";
    if (ctrlEl) ctrlEl.textContent = control ? control.treatmentName : "n/a";

    const metricKey = results.metricKey;
    const metricLabel = metricKey === "npv" ? "NPV" : metricKey === "roi" ? "ROI" : "BCR";
    const showMetric = (x) => {
      if (!x) return "n/a";
      const v = x[metricKey];
      if (metricKey === "npv") return fmtCurrency.format(v);
      if (v === Infinity) return "∞";
      return fmtNumber3.format(v);
    };

    if (bestSub) bestSub.textContent = best ? `${metricLabel}: ${showMetric(best)} | Rank ${best.rank}` : "n/a";
    if (worstSub) worstSub.textContent = worst ? `${metricLabel}: ${showMetric(worst)} | Rank ${worst.rank}` : "n/a";
    if (ctrlSub) ctrlSub.textContent = control ? `NPV: ${fmtCurrency.format(control.npv)} | BCR: ${control.bcr === Infinity ? "∞" : fmtNumber3.format(control.bcr)}` : "n/a";
  }

  function updateResultsAssumptions(state, results) {
    const elAss = q("resultsAssumptions");
    if (!elAss) return;

    const s = state.settings;
    const startYear = safeNum(s.startYear, 2026);
    const years = safeNum(s.years, 5);
    const disc = safeNum(s.discBase, 7);
    const adopt = safeNum(s.adoptBase, 1);
    const risk = safeNum(s.riskBase, 0);

    const allocate = Boolean(s.allocateProjectItems);
    const sys = safeStr(s.systemType, "single");

    const control = results.control;

    const lines = [];
    lines.push(`Tool: ${TOOL_NAME}.`);
    lines.push(`Analysis years: ${startYear} to ${startYear + years - 1} (${years} years).`);
    lines.push(`Discount rate (base): ${fmtNumber2.format(disc)}%.`);
    lines.push(`Adoption (base multiplier): ${fmtNumber2.format(adopt)}.`);
    lines.push(`Overall risk (base): ${fmtPercent1.format(risk)} (reduces benefits, not costs).`);
    lines.push(`Project-wide items: ${allocate ? "Allocated by effective area" : "Applied equally to each treatment scenario"}.`);
    lines.push(`System type: ${sys}.`);
    if (control) lines.push(`Control: ${control.treatmentName}.`);
    elAss.textContent = lines.join(" ");
  }

  /********************************************************************
   * CRUD rendering for inputs (Outputs, Treatments, Benefits, Costs)
   ********************************************************************/
  function renderOutputs(state) {
    const root = q("outputsList");
    if (!root) return;
    root.innerHTML = "";

    const items = state.outputs || [];
    if (items.length === 0) {
      root.appendChild(el("div", { class: "small muted" }, ["No outputs yet. Add an output to define monetised benefits."]));
      return;
    }

    items.forEach((o) => {
      const card = el("div", { class: "item-card" });
      card.appendChild(el("div", { class: "item-head" }, [
        el("div", { class: "item-title" }, [o.name || "Untitled output"]),
        el("div", { class: "item-actions" }, [
          el("button", {
            type: "button",
            class: "btn small ghost",
            onClick: () => { removeOutput(o.id); },
          }, ["Remove"]),
        ]),
      ]));

      const grid = el("div", { class: "row-4" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Output name"]),
          el("input", {
            type: "text",
            value: o.name ?? "",
            onInput: (e) => { o.name = e.target.value; scheduleRecalc("Output updated"); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Unit"]),
          el("input", {
            type: "text",
            value: o.unit ?? "",
            onInput: (e) => { o.unit = e.target.value; scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Value per unit (AUD)"]),
          el("input", {
            type: "number",
            step: "0.01",
            value: o.valuePerUnit ?? 0,
            onInput: (e) => { o.valuePerUnit = safeNum(e.target.value, 0); scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Source (optional)"]),
          el("input", {
            type: "text",
            value: o.source ?? "",
            onInput: (e) => { o.source = e.target.value; },
          }),
        ]),
      ]);

      card.appendChild(grid);

      const notes = el("div", { class: "field" }, [
        el("label", {}, ["Notes (optional)"]),
        el("textarea", {
          rows: "2",
          onInput: (e) => { o.notes = e.target.value; },
        }, []),
      ]);
      notes.querySelector("textarea").value = o.notes ?? "";
      card.appendChild(notes);

      root.appendChild(card);
    });
  }

  function renderTreatments(state) {
    const root = q("treatmentsList");
    if (!root) return;
    root.innerHTML = "";

    const treatments = state.treatments || [];
    if (treatments.length === 0) {
      root.appendChild(el("div", { class: "small muted" }, ["No treatments yet. Add a control and one or more treatments."]));
      return;
    }

    ensureSingleControl(state);

    treatments.forEach((t) => {
      const card = el("div", { class: "item-card" });

      const head = el("div", { class: "item-head" }, [
        el("div", { class: "item-title" }, [t.name || "Untitled treatment"]),
        el("div", { class: "item-actions" }, [
          el("button", { type: "button", class: "btn small", onClick: () => addTreatmentOutputRow(t.id) }, ["Add output row"]),
          el("button", {
            type: "button",
            class: "btn small ghost",
            onClick: () => { removeTreatment(t.id); },
          }, ["Remove"]),
        ]),
      ]);

      card.appendChild(head);

      const row = el("div", { class: "row-4" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Treatment name"]),
          el("input", {
            type: "text",
            value: t.name ?? "",
            onInput: (e) => { t.name = e.target.value; scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Control?"]),
          (() => {
            const sel = el("select", {
              onChange: (e) => {
                const yes = e.target.value === "true";
                if (yes) {
                  state.treatments.forEach((x) => (x.isControl = x.id === t.id));
                } else {
                  t.isControl = false;
                }
                ensureSingleControl(state);
                scheduleRecalc("Control updated");
                renderTreatments(state);
              },
            }, [
              el("option", { value: "false" }, ["No"]),
              el("option", { value: "true" }, ["Yes"]),
            ]);
            sel.value = t.isControl ? "true" : "false";
            return sel;
          })(),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Max area (ha)"]),
          el("input", {
            type: "number",
            step: "1",
            min: "0",
            value: t.maxAreaHa ?? 0,
            onInput: (e) => { t.maxAreaHa = safeNum(e.target.value, 0); scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Source (optional)"]),
          el("input", {
            type: "text",
            value: t.source ?? "",
            onInput: (e) => { t.source = e.target.value; },
          }),
        ]),
      ]);
      card.appendChild(row);

      const costRow = el("div", { class: "row-4" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Labour cost per ha (AUD/ha/year)"]),
          el("input", {
            type: "number",
            step: "0.01",
            value: t.labourCostPerHa ?? 0,
            onInput: (e) => { t.labourCostPerHa = safeNum(e.target.value, 0); scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Input cost per ha (AUD/ha/year)"]),
          el("input", {
            type: "number",
            step: "0.01",
            value: t.inputCostPerHa ?? 0,
            onInput: (e) => { t.inputCostPerHa = safeNum(e.target.value, 0); scheduleRecalc(); },
          }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Effective area (base)"]),
          el("div", { class: "metric" }, [
            el("div", { class: "value" }, [fmtNumber2.format(effectiveAreaHa(t, state.settings?.adoptBase ?? 1))]),
            el("div", { class: "small muted" }, ["ha (max area × adoption)"]),
          ]),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Quick note (optional)"]),
          el("input", {
            type: "text",
            value: t.notes ?? "",
            onInput: (e) => { t.notes = e.target.value; },
          }),
        ]),
      ]);
      card.appendChild(costRow);

      // Treatment outputs table
      const outTable = el("div", { class: "table-scroll" });
      const table = el("table", { class: "summary-table" });
      const thead = el("thead", {}, [
        el("tr", {}, [
          el("th", {}, ["Output"]),
          el("th", {}, ["Qty per ha"]),
          el("th", {}, ["From year"]),
          el("th", {}, ["To year"]),
          el("th", {}, ["Notes"]),
          el("th", {}, [""]),
        ]),
      ]);
      table.appendChild(thead);

      const tbody = el("tbody");
      (t.outputs || []).forEach((row) => {
        const tr = el("tr");

        // Output select
        const outSel = el("select", {
          onChange: (e) => { row.outputId = e.target.value; scheduleRecalc(); },
        });
        (state.outputs || []).forEach((o) => {
          outSel.appendChild(el("option", { value: o.id }, [o.name || o.id]));
        });
        outSel.value = row.outputId ?? (state.outputs?.[0]?.id ?? "");

        tr.appendChild(el("td", {}, [outSel]));

        const qty = el("input", {
          type: "number",
          step: "0.01",
          value: row.qtyPerHa ?? 0,
          onInput: (e) => { row.qtyPerHa = safeNum(e.target.value, 0); scheduleRecalc(); },
        });
        tr.appendChild(el("td", {}, [qty]));

        const yf = el("input", {
          type: "number",
          step: "1",
          value: row.yearFrom ?? state.settings.startYear,
          onInput: (e) => { row.yearFrom = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); },
        });
        tr.appendChild(el("td", {}, [yf]));

        const yt = el("input", {
          type: "number",
          step: "1",
          value: row.yearTo ?? (state.settings.startYear + state.settings.years - 1),
          onInput: (e) => { row.yearTo = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); },
        });
        tr.appendChild(el("td", {}, [yt]));

        const note = el("input", {
          type: "text",
          value: row.notes ?? "",
          onInput: (e) => { row.notes = e.target.value; },
        });
        tr.appendChild(el("td", {}, [note]));

        tr.appendChild(el("td", {}, [
          el("button", { type: "button", class: "btn small ghost", onClick: () => removeTreatmentOutputRow(t.id, row.id) }, ["Remove"]),
        ]));

        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      outTable.appendChild(table);

      card.appendChild(el("div", { class: "small muted", style: "margin-top:0.25rem;" }, [
        "Output rows drive benefits (qty per ha × value per unit × effective area). Year ranges must sit within the analysis horizon.",
      ]));
      card.appendChild(outTable);

      root.appendChild(card);
    });
  }

  function renderBenefits(state) {
    const root = q("benefitsList");
    if (!root) return;
    root.innerHTML = "";

    const items = state.benefits || [];
    if (items.length === 0) {
      root.appendChild(el("div", { class: "small muted" }, ["No additional benefits yet. Add if you have benefits not captured by output quantities."]));
      return;
    }

    items.forEach((b) => {
      const card = el("div", { class: "item-card" });
      card.appendChild(el("div", { class: "item-head" }, [
        el("div", { class: "item-title" }, [b.name || "Untitled benefit"]),
        el("div", { class: "item-actions" }, [
          el("button", { type: "button", class: "btn small ghost", onClick: () => removeBenefit(b.id) }, ["Remove"]),
        ]),
      ]));

      card.appendChild(el("div", { class: "row-4" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Benefit name"]),
          el("input", { type: "text", value: b.name ?? "", onInput: (e) => { b.name = e.target.value; } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Annual value (AUD/year)"]),
          el("input", { type: "number", step: "0.01", value: b.annualValue ?? 0, onInput: (e) => { b.annualValue = safeNum(e.target.value, 0); scheduleRecalc(); } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["From year"]),
          el("input", { type: "number", step: "1", value: b.yearFrom ?? state.settings.startYear, onInput: (e) => { b.yearFrom = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["To year"]),
          el("input", { type: "number", step: "1", value: b.yearTo ?? (state.settings.startYear + state.settings.years - 1), onInput: (e) => { b.yearTo = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); } }),
        ]),
      ]));

      const notes = el("div", { class: "row-2" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Notes (optional)"]),
          (() => {
            const ta = el("textarea", { rows: "2", onInput: (e) => { b.notes = e.target.value; } });
            ta.value = b.notes ?? "";
            return ta;
          })(),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Source (optional)"]),
          el("input", { type: "text", value: b.source ?? "", onInput: (e) => { b.source = e.target.value; } }),
        ]),
      ]);
      card.appendChild(notes);

      root.appendChild(card);
    });
  }

  function renderCosts(state) {
    const root = q("costsList");
    if (!root) return;
    root.innerHTML = "";

    const items = state.costs || [];
    if (items.length === 0) {
      root.appendChild(el("div", { class: "small muted" }, ["No project-wide costs yet. Add costs such as coordination, monitoring, training, or capital items."]));
      return;
    }

    items.forEach((c) => {
      const card = el("div", { class: "item-card" });
      card.appendChild(el("div", { class: "item-head" }, [
        el("div", { class: "item-title" }, [c.name || "Untitled cost"]),
        el("div", { class: "item-actions" }, [
          el("button", { type: "button", class: "btn small ghost", onClick: () => removeCost(c.id) }, ["Remove"]),
        ]),
      ]));

      card.appendChild(el("div", { class: "row-4" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Cost name"]),
          el("input", { type: "text", value: c.name ?? "", onInput: (e) => { c.name = e.target.value; } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Annual value (AUD/year)"]),
          el("input", { type: "number", step: "0.01", value: c.annualValue ?? 0, onInput: (e) => { c.annualValue = safeNum(e.target.value, 0); scheduleRecalc(); } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["From year"]),
          el("input", { type: "number", step: "1", value: c.yearFrom ?? state.settings.startYear, onInput: (e) => { c.yearFrom = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); } }),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["To year"]),
          el("input", { type: "number", step: "1", value: c.yearTo ?? (state.settings.startYear + state.settings.years - 1), onInput: (e) => { c.yearTo = safeNum(e.target.value, state.settings.startYear); scheduleRecalc(); } }),
        ]),
      ]));

      const notes = el("div", { class: "row-2" }, [
        el("div", { class: "field" }, [
          el("label", {}, ["Notes (optional)"]),
          (() => {
            const ta = el("textarea", { rows: "2", onInput: (e) => { c.notes = e.target.value; } });
            ta.value = c.notes ?? "";
            return ta;
          })(),
        ]),
        el("div", { class: "field" }, [
          el("label", {}, ["Source (optional)"]),
          el("input", { type: "text", value: c.source ?? "", onInput: (e) => { c.source = e.target.value; } }),
        ]),
      ]);
      card.appendChild(notes);

      root.appendChild(card);
    });
  }

  function renderDatabaseTab(state) {
    const outRoot = q("dbOutputs");
    const trtRoot = q("dbTreatments");
    if (outRoot) outRoot.innerHTML = "";
    if (trtRoot) trtRoot.innerHTML = "";

    if (outRoot) {
      (state.outputs || []).forEach((o) => {
        outRoot.appendChild(el("div", { class: "item-card" }, [
          el("div", { class: "item-head" }, [
            el("div", { class: "item-title" }, [o.name || "Output"]),
            el("div", { class: "small muted" }, [o.unit ? `Unit: ${o.unit}` : ""]),
          ]),
          el("div", { class: "row-2" }, [
            el("div", { class: "field" }, [
              el("label", {}, ["Source"]),
              el("input", { type: "text", value: o.source ?? "", onInput: (e) => { o.source = e.target.value; } }),
            ]),
            el("div", { class: "field" }, [
              el("label", {}, ["Notes"]),
              el("input", { type: "text", value: o.notes ?? "", onInput: (e) => { o.notes = e.target.value; } }),
            ]),
          ]),
        ]));
      });
    }

    if (trtRoot) {
      (state.treatments || []).forEach((t) => {
        trtRoot.appendChild(el("div", { class: "item-card" }, [
          el("div", { class: "item-head" }, [
            el("div", { class: "item-title" }, [t.name || "Treatment"]),
            el("div", { class: "small muted" }, [t.isControl ? "Control" : ""]),
          ]),
          el("div", { class: "row-2" }, [
            el("div", { class: "field" }, [
              el("label", {}, ["Source"]),
              el("input", { type: "text", value: t.source ?? "", onInput: (e) => { t.source = e.target.value; } }),
            ]),
            el("div", { class: "field" }, [
              el("label", {}, ["Notes"]),
              el("input", { type: "text", value: t.notes ?? "", onInput: (e) => { t.notes = e.target.value; } }),
            ]),
          ]),
        ]));
      });
    }
  }

  /********************************************************************
   * CRUD helpers
   ********************************************************************/
  function addOutput() {
    STATE.outputs.push({ id: uid("out"), name: "New output", unit: "", valuePerUnit: 0, notes: "", source: "" });
    renderOutputs(STATE);
    renderTreatments(STATE);
    renderDatabaseTab(STATE);
    scheduleRecalc("Output added");
  }

  function removeOutput(id) {
    const used = (STATE.treatments || []).some((t) => (t.outputs || []).some((r) => r.outputId === id));
    if (used) {
      toast("This output is used by one or more treatment output rows. Remove those rows first.", "warn", 5000);
      return;
    }
    STATE.outputs = (STATE.outputs || []).filter((o) => o.id !== id);
    renderOutputs(STATE);
    renderTreatments(STATE);
    renderDatabaseTab(STATE);
    scheduleRecalc("Output removed");
  }

  function addTreatment() {
    const startYear = Math.trunc(safeNum(STATE.settings.startYear, 2026));
    const endYear = startYear + Math.max(1, Math.trunc(safeNum(STATE.settings.years, 5))) - 1;
    const defaultOutputId = STATE.outputs?.[0]?.id ?? null;

    const t = {
      id: uid("trt"),
      name: "New treatment",
      isControl: false,
      maxAreaHa: 0,
      labourCostPerHa: 0,
      inputCostPerHa: 0,
      notes: "",
      source: "",
      outputs: defaultOutputId
        ? [{ id: uid("to"), outputId: defaultOutputId, qtyPerHa: 0, yearFrom: startYear, yearTo: endYear, notes: "" }]
        : [],
    };
    STATE.treatments.push(t);
    ensureSingleControl(STATE);
    renderTreatments(STATE);
    renderDatabaseTab(STATE);
    scheduleRecalc("Treatment added");
  }

  function removeTreatment(id) {
    const t = (STATE.treatments || []).find((x) => x.id === id);
    const wasControl = Boolean(t?.isControl);
    STATE.treatments = (STATE.treatments || []).filter((x) => x.id !== id);
    if (wasControl) ensureSingleControl(STATE);
    renderTreatments(STATE);
    renderDatabaseTab(STATE);
    scheduleRecalc("Treatment removed");
  }

  function addTreatmentOutputRow(treatmentId) {
    const t = (STATE.treatments || []).find((x) => x.id === treatmentId);
    if (!t) return;
    if (!STATE.outputs || STATE.outputs.length === 0) {
      toast("Add at least one output first (Outputs tab) so treatment output rows can reference it.", "warn", 4500);
      return;
    }
    const startYear = Math.trunc(safeNum(STATE.settings.startYear, 2026));
    const endYear = startYear + Math.max(1, Math.trunc(safeNum(STATE.settings.years, 5))) - 1;

    t.outputs = t.outputs || [];
    t.outputs.push({
      id: uid("to"),
      outputId: STATE.outputs[0].id,
      qtyPerHa: 0,
      yearFrom: startYear,
      yearTo: endYear,
      notes: "",
    });
    renderTreatments(STATE);
    scheduleRecalc("Treatment output row added");
  }

  function removeTreatmentOutputRow(treatmentId, rowId) {
    const t = (STATE.treatments || []).find((x) => x.id === treatmentId);
    if (!t) return;
    t.outputs = (t.outputs || []).filter((r) => r.id !== rowId);
    renderTreatments(STATE);
    scheduleRecalc("Treatment output row removed");
  }

  function addBenefit() {
    const startYear = Math.trunc(safeNum(STATE.settings.startYear, 2026));
    const endYear = startYear + Math.max(1, Math.trunc(safeNum(STATE.settings.years, 5))) - 1;
    STATE.benefits.push({ id: uid("ben"), name: "New benefit", annualValue: 0, yearFrom: startYear, yearTo: endYear, notes: "", source: "" });
    renderBenefits(STATE);
    scheduleRecalc("Benefit added");
  }

  function removeBenefit(id) {
    STATE.benefits = (STATE.benefits || []).filter((b) => b.id !== id);
    renderBenefits(STATE);
    scheduleRecalc("Benefit removed");
  }

  function addCost() {
    const startYear = Math.trunc(safeNum(STATE.settings.startYear, 2026));
    const endYear = startYear + Math.max(1, Math.trunc(safeNum(STATE.settings.years, 5))) - 1;
    STATE.costs.push({ id: uid("cst"), name: "New cost", annualValue: 0, yearFrom: startYear, yearTo: endYear, notes: "", source: "" });
    renderCosts(STATE);
    scheduleRecalc("Cost added");
  }

  function removeCost(id) {
    STATE.costs = (STATE.costs || []).filter((c) => c.id !== id);
    renderCosts(STATE);
    scheduleRecalc("Cost removed");
  }

  /********************************************************************
   * Tab navigation
   ********************************************************************/
  function setActiveTab(tab) {
    const buttons = Array.from(document.querySelectorAll(".tab-link"));
    const panels = Array.from(document.querySelectorAll(".tab-panel"));

    buttons.forEach((b) => {
      const is = b.getAttribute("data-tab") === tab;
      b.classList.toggle("active", is);
      b.setAttribute("aria-selected", is ? "true" : "false");
      b.tabIndex = is ? 0 : -1;
    });

    panels.forEach((p) => {
      const is = p.getAttribute("data-tab-panel") === tab;
      p.classList.toggle("active", is);
      p.classList.toggle("show", is);
      p.setAttribute("aria-hidden", is ? "false" : "true");
    });

    STATE.ui.activeTab = tab;
    // Snapshot-friendly: if results tab, ensure matrix renders immediately
    if (tab === "results") {
      recalcAndRender();
    }
  }

  function wireTabs() {
    document.querySelectorAll(".tab-link").forEach((btn) => {
      btn.addEventListener("click", () => setActiveTab(btn.getAttribute("data-tab")));
    });

    document.querySelectorAll("[data-tab-jump]").forEach((btn) => {
      btn.addEventListener("click", () => setActiveTab(btn.getAttribute("data-tab-jump")));
    });
  }

  /********************************************************************
   * Project metadata and settings binding
   ********************************************************************/
  function bindProjectFields() {
    const p = STATE.project;

    const map = [
      ["projectName", "projectName"],
      ["projectLead", "projectLead"],
      ["analystNames", "analystNames"],
      ["projectTeam", "projectTeam"],
      ["organisation", "organisation"],
      ["lastUpdated", "lastUpdated"],
      ["contactEmail", "contactEmail"],
      ["contactPhone", "contactPhone"],
      ["projectSummary", "projectSummary"],
      ["projectGoal", "projectGoal"],
      ["withProject", "withProject"],
      ["withoutProject", "withoutProject"],
      ["projectObjectives", "projectObjectives"],
      ["projectActivities", "projectActivities"],
      ["stakeholderGroups", "stakeholderGroups"],
    ];

    map.forEach(([id, key]) => setInputValue(id, p[key]));

    map.forEach(([id, key]) => {
      const node = q(id);
      if (!node) return;
      node.addEventListener("input", (e) => {
        p[key] = e.target.value;
      });
    });
  }

  function bindSettingsFields() {
    const s = STATE.settings;

    const bindNum = (id, key, opts = {}) => {
      setInputValue(id, s[key]);
      const node = q(id);
      if (!node) return;
      node.addEventListener("input", (e) => {
        s[key] = safeNum(e.target.value, s[key]);
        if (opts.recalc) scheduleRecalc();
        if (opts.rerender) rerenderAll();
      });
    };

    const bindSel = (id, key, opts = {}) => {
      setInputValue(id, s[key]);
      const node = q(id);
      if (!node) return;
      node.addEventListener("change", (e) => {
        const v = e.target.value;
        s[key] = v;
        if (opts.recalc) scheduleRecalc();
        if (opts.rerender) rerenderAll();
      });
    };

    bindNum("startYear", "startYear", { recalc: true, rerender: true });
    bindNum("projectStartYear", "projectStartYear", { recalc: true });
    bindNum("years", "years", { recalc: true, rerender: true });
    bindSel("systemType", "systemType");

    bindNum("discBase", "discBase", { recalc: true });
    bindNum("discLow", "discLow");
    bindNum("discHigh", "discHigh");
    const oa = q("outputAssumptions");
    if (oa) {
      oa.value = s.outputAssumptions ?? "";
      oa.addEventListener("input", (e) => { s.outputAssumptions = e.target.value; });
    }

    bindNum("mirrFinance", "mirrFinance");
    bindNum("mirrReinvest", "mirrReinvest");

    bindNum("adoptLow", "adoptLow", { recalc: true, rerender: true });
    bindNum("adoptBase", "adoptBase", { recalc: true, rerender: true });
    bindNum("adoptHigh", "adoptHigh", { recalc: true, rerender: true });

    const alloc = q("allocateProjectItems");
    if (alloc) {
      alloc.value = String(Boolean(s.allocateProjectItems));
      alloc.addEventListener("change", (e) => {
        s.allocateProjectItems = e.target.value === "true";
        scheduleRecalc();
      });
    }

    bindNum("riskLow", "riskLow", { recalc: true });
    bindNum("riskBase", "riskBase", { recalc: true });
    bindNum("riskHigh", "riskHigh", { recalc: true });

    bindNum("rTech", "rTech");
    bindNum("rNonCoop", "rNonCoop");
    bindNum("rSocio", "rSocio");
    bindNum("rFin", "rFin");
    bindNum("rMan", "rMan");

    // Simulation fields
    bindNum("simN", "simN");
    bindNum("targetBCR", "targetBCR");
    bindNum("randSeed", "randSeed");
    bindNum("simVarPct", "simVarPct");

    const setBoolSel = (id, key) => {
      const node = q(id);
      if (!node) return;
      node.value = String(Boolean(s[key]));
      node.addEventListener("change", (e) => { s[key] = e.target.value === "true"; });
    };
    setBoolSel("simVaryOutputs", "simVaryOutputs");
    setBoolSel("simVaryTreatCosts", "simVaryTreatCosts");
    setBoolSel("simVaryInputCosts", "simVaryInputCosts");

    const bcrLabel = q("simBcrTargetLabel");
    const bcrIn = q("targetBCR");
    if (bcrLabel && bcrIn) {
      bcrLabel.textContent = safeStr(bcrIn.value || s.targetBCR || 2);
      bcrIn.addEventListener("input", () => { bcrLabel.textContent = bcrIn.value; });
    }

    const combBtn = q("calcCombinedRisk");
    if (combBtn) {
      combBtn.addEventListener("click", () => {
        // Combined risk using multiplicative complement: 1 - Π(1-ri)
        const parts = [
          clamp(safeNum(s.rTech, 0), 0, 1),
          clamp(safeNum(s.rNonCoop, 0), 0, 1),
          clamp(safeNum(s.rSocio, 0), 0, 1),
          clamp(safeNum(s.rFin, 0), 0, 1),
          clamp(safeNum(s.rMan, 0), 0, 1),
        ];
        let prod = 1;
        parts.forEach((ri) => { prod *= (1 - ri); });
        const combined = 1 - prod;
        s.riskBase = combined;
        setInputValue("riskBase", combined);
        const out = q("combinedRiskOut")?.querySelector(".value");
        if (out) out.textContent = fmtPercent1.format(combined);
        toast("Combined base risk calculated and applied to Overall risk base.", "info", 4200);
        scheduleRecalc();
      });
    }
  }

  function bindResultsControls() {
    const rankMetric = q("rankMetric");
    if (rankMetric) {
      rankMetric.value = STATE.ui.rankMetric || "bcr";
      rankMetric.addEventListener("change", (e) => {
        STATE.ui.rankMetric = e.target.value;
        scheduleRecalc();
      });
    }

    const viewMode = q("resultsViewMode");
    if (viewMode) {
      viewMode.value = STATE.ui.resultsViewMode || "absolute";
      viewMode.addEventListener("change", (e) => {
        STATE.ui.resultsViewMode = e.target.value;
        scheduleRecalc();
      });
    }

    const cols = q("columnsPerPage");
    if (cols) {
      cols.value = String(STATE.ui.columnsPerPage || 6);
      cols.addEventListener("change", (e) => {
        STATE.ui.columnsPerPage = safeNum(e.target.value, 6);
        // keep matrixStart bounded
        STATE.ui.matrixStart = 0;
        scheduleRecalc();
      });
    }

    const showD = q("showDeltas");
    if (showD) {
      showD.value = String(Boolean(STATE.ui.showDeltas));
      showD.addEventListener("change", (e) => {
        STATE.ui.showDeltas = e.target.value === "true";
        scheduleRecalc();
      });
    }

    const prev = q("matrixPrev");
    const next = q("matrixNext");
    if (prev) prev.addEventListener("click", () => {
      const step = Math.max(1, Math.trunc(safeNum(STATE.ui.columnsPerPage, 6)));
      STATE.ui.matrixStart = Math.max(0, Math.trunc(safeNum(STATE.ui.matrixStart, 0)) - step);
      recalcAndRender();
    });
    if (next) next.addEventListener("click", () => {
      const step = Math.max(1, Math.trunc(safeNum(STATE.ui.columnsPerPage, 6)));
      STATE.ui.matrixStart = Math.max(0, Math.trunc(safeNum(STATE.ui.matrixStart, 0)) + step);
      recalcAndRender();
    });

    const recalcBtn = q("recalc");
    if (recalcBtn) recalcBtn.addEventListener("click", () => recalcAndRender(true));

    const exportBtn = q("exportResultsXlsx");
    const exportBtnFoot = q("exportResultsXlsxFoot");
    if (exportBtn) exportBtn.addEventListener("click", () => exportResultsExcel());
    if (exportBtnFoot) exportBtnFoot.addEventListener("click", () => exportResultsExcel());

    const copyBtn = q("copyResultsTable");
    if (copyBtn) copyBtn.addEventListener("click", () => copyResultsForWord());

    const pdfBtn = q("exportPdf");
    const pdfBtnFoot = q("exportPdfFoot");
    const doPrint = () => {
      toast("Use your browser print dialog to save as PDF.", "info", 3500);
      window.print();
    };
    if (pdfBtn) pdfBtn.addEventListener("click", doPrint);
    if (pdfBtnFoot) pdfBtnFoot.addEventListener("click", doPrint);
  }

  /********************************************************************
   * Recalc scheduling
   ********************************************************************/
  let recalcTimer = null;

  function scheduleRecalc(msg = null) {
    if (msg) toast(msg, "info", 1800);
    if (recalcTimer) clearTimeout(recalcTimer);
    recalcTimer = setTimeout(() => {
      recalcAndRender(false);
    }, 220);
  }

  function recalcAndRender(forceToast = false) {
    ensureSingleControl(STATE);

    const results = computeAllResults(STATE);
    STATE.cache.lastResults = results;
    STATE.cache.lastResultsMeta = {
      viewMode: STATE.ui.resultsViewMode,
      rankMetric: STATE.ui.rankMetric,
      columnsPerPage: STATE.ui.columnsPerPage,
      showDeltas: STATE.ui.showDeltas,
      matrixStart: STATE.ui.matrixStart,
    };

    updateSnapshotCards(results);
    renderResultsMatrix(STATE, results);
    updateResultsAssumptions(STATE, results);

    if (forceToast) toast("Results updated.", "info", 2200);
  }

  /********************************************************************
   * Excel export: templates + results
   ********************************************************************/
  function sheetFromObjects(rows, headerOrder) {
    const cleanRows = rows.map((r) => {
      const o = {};
      headerOrder.forEach((h) => (o[h] = r[h] ?? ""));
      return o;
    });
    return XLSX.utils.json_to_sheet(cleanRows, { header: headerOrder });
  }

  function buildWorkbookFromState(state, { includeResults = true, blank = false } = {}) {
    const wb = XLSX.utils.book_new();
    const stamp = new Date().toISOString();

    // README
    const readmeLines = [
      `${TOOL_NAME} — Excel workflow`,
      `Generated: ${stamp}`,
      ``,
      `How to use:`,
      `1) Keep sheet names and header columns unchanged.`,
      `2) Edit rows in Outputs, Treatments, TreatmentOutputs, Benefits, Costs.`,
      `3) You may leave IDs blank. The tool will auto-generate IDs on import.`,
      `4) Ensure exactly one treatment has is_control = TRUE.`,
      `5) Year ranges must sit within the analysis horizon (start_year to start_year + years - 1).`,
      ``,
      `This file is designed to be uploaded back into the tool.`,
    ];
    const wsReadme = XLSX.utils.aoa_to_sheet(readmeLines.map((x) => [x]));
    XLSX.utils.book_append_sheet(wb, wsReadme, "README");

    // Project
    const p = state.project || {};
    const projectRows = [{
      tool_name: TOOL_NAME,
      organisation: p.organisation ?? ORG_NAME,
      project_name: blank ? "" : (p.projectName ?? ""),
      project_lead: blank ? "" : (p.projectLead ?? ""),
      analyst_team: blank ? "" : (p.analystNames ?? ""),
      partners_or_team: blank ? "" : (p.projectTeam ?? ""),
      last_updated: blank ? "" : (p.lastUpdated ?? ""),
      contact_email: blank ? "" : (p.contactEmail ?? ""),
      contact_phone: blank ? "" : (p.contactPhone ?? ""),
      short_summary: blank ? "" : (p.projectSummary ?? ""),
      project_goal: blank ? "" : (p.projectGoal ?? ""),
      with_project: blank ? "" : (p.withProject ?? ""),
      without_project: blank ? "" : (p.withoutProject ?? ""),
      objectives: blank ? "" : (p.projectObjectives ?? ""),
      activities: blank ? "" : (p.projectActivities ?? ""),
      stakeholder_groups: blank ? "" : (p.stakeholderGroups ?? ""),
    }];
    const projectHdr = Object.keys(projectRows[0]);
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(projectRows, projectHdr), "Project");

    // Settings
    const s = state.settings || {};
    const settingsRows = [{
      tool_name: TOOL_NAME,
      start_year: blank ? "" : safeNum(s.startYear, 2026),
      project_start_year: blank ? "" : safeNum(s.projectStartYear, safeNum(s.startYear, 2026)),
      years: blank ? "" : safeNum(s.years, 5),
      system_type: blank ? "" : (s.systemType ?? "single"),
      discount_base_percent: blank ? "" : safeNum(s.discBase, 7),
      discount_low_percent: blank ? "" : safeNum(s.discLow, 3),
      discount_high_percent: blank ? "" : safeNum(s.discHigh, 10),
      adoption_low: blank ? "" : safeNum(s.adoptLow, 0.6),
      adoption_base: blank ? "" : safeNum(s.adoptBase, 0.85),
      adoption_high: blank ? "" : safeNum(s.adoptHigh, 1),
      allocate_project_items_by_area: blank ? "" : (Boolean(s.allocateProjectItems) ? "TRUE" : "FALSE"),
      risk_low: blank ? "" : safeNum(s.riskLow, 0.05),
      risk_base: blank ? "" : safeNum(s.riskBase, 0.12),
      risk_high: blank ? "" : safeNum(s.riskHigh, 0.25),
      notes_on_assumptions: blank ? "" : (s.outputAssumptions ?? ""),
    }];
    const settingsHdr = Object.keys(settingsRows[0]);
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(settingsRows, settingsHdr), "Settings");

    // Outputs
    const outputsHdr = ["output_id", "output_name", "unit", "value_per_unit_aud", "notes", "source"];
    const outputsRows = (blank ? [] : (state.outputs || [])).map((o) => ({
      output_id: o.id ?? "",
      output_name: o.name ?? "",
      unit: o.unit ?? "",
      value_per_unit_aud: safeNum(o.valuePerUnit, 0),
      notes: o.notes ?? "",
      source: o.source ?? "",
    }));
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(outputsRows.length ? outputsRows : [outputsHdr.reduce((a, h) => ((a[h] = ""), a), {})], outputsHdr), "Outputs");

    // Treatments
    const trtHdr = ["treatment_id", "treatment_name", "is_control", "max_area_ha", "labour_cost_per_ha_aud", "input_cost_per_ha_aud", "notes", "source"];
    const trtRows = (blank ? [] : (state.treatments || [])).map((t) => ({
      treatment_id: t.id ?? "",
      treatment_name: t.name ?? "",
      is_control: t.isControl ? "TRUE" : "FALSE",
      max_area_ha: safeNum(t.maxAreaHa, 0),
      labour_cost_per_ha_aud: safeNum(t.labourCostPerHa, 0),
      input_cost_per_ha_aud: safeNum(t.inputCostPerHa, 0),
      notes: t.notes ?? "",
      source: t.source ?? "",
    }));
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(trtRows.length ? trtRows : [trtHdr.reduce((a, h) => ((a[h] = ""), a), {})], trtHdr), "Treatments");

    // TreatmentOutputs
    const toHdr = ["row_id", "treatment_id", "output_id", "qty_per_ha", "year_from", "year_to", "notes"];
    const toRows = (blank ? [] : (state.treatments || []).flatMap((t) => (t.outputs || []).map((r) => ({
      row_id: r.id ?? "",
      treatment_id: t.id ?? "",
      output_id: r.outputId ?? "",
      qty_per_ha: safeNum(r.qtyPerHa, 0),
      year_from: safeNum(r.yearFrom, s.startYear),
      year_to: safeNum(r.yearTo, s.startYear),
      notes: r.notes ?? "",
    }))));
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(toRows.length ? toRows : [toHdr.reduce((a, h) => ((a[h] = ""), a), {})], toHdr), "TreatmentOutputs");

    // Benefits
    const benHdr = ["benefit_id", "benefit_name", "annual_value_aud", "year_from", "year_to", "notes", "source"];
    const benRows = (blank ? [] : (state.benefits || [])).map((b) => ({
      benefit_id: b.id ?? "",
      benefit_name: b.name ?? "",
      annual_value_aud: safeNum(b.annualValue, 0),
      year_from: safeNum(b.yearFrom, s.startYear),
      year_to: safeNum(b.yearTo, s.startYear),
      notes: b.notes ?? "",
      source: b.source ?? "",
    }));
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(benRows.length ? benRows : [benHdr.reduce((a, h) => ((a[h] = ""), a), {})], benHdr), "Benefits");

    // Costs
    const costHdr = ["cost_id", "cost_name", "annual_value_aud", "year_from", "year_to", "notes", "source"];
    const costRows = (blank ? [] : (state.costs || [])).map((c) => ({
      cost_id: c.id ?? "",
      cost_name: c.name ?? "",
      annual_value_aud: safeNum(c.annualValue, 0),
      year_from: safeNum(c.yearFrom, s.startYear),
      year_to: safeNum(c.yearTo, s.startYear),
      notes: c.notes ?? "",
      source: c.source ?? "",
    }));
    XLSX.utils.book_append_sheet(wb, sheetFromObjects(costRows.length ? costRows : [costHdr.reduce((a, h) => ((a[h] = ""), a), {})], costHdr), "Costs");

    if (includeResults) {
      const results = computeAllResults(state);
      const ordered = results.indicators
        .slice()
        .sort((a, b) => {
          if (a.isControl && !b.isControl) return -1;
          if (!a.isControl && b.isControl) return 1;
          return (a.rank ?? 9999) - (b.rank ?? 9999);
        });

      const cols = ordered.map((x) => x.isControl ? `${x.treatmentName} (Control)` : x.treatmentName);

      const absAoA = [];
      absAoA.push(["Indicator", ...cols]);
      MATRIX_ROWS.forEach((r) => {
        const row = [r.label];
        ordered.forEach((c) => {
          const v = c[r.key];
          if (r.key === "npv" || r.key === "pvBenefits" || r.key === "pvCosts") row.push(Number(v));
          else if (r.key === "rank") row.push(c.rank ?? "");
          else row.push(v === Infinity ? "INF" : Number(v));
        });
        absAoA.push(row);
      });

      const wsAbs = XLSX.utils.aoa_to_sheet(absAoA);
      XLSX.utils.book_append_sheet(wb, wsAbs, "Results_Absolute");

      const ctrl = results.control;
      const relAoA = [];
      relAoA.push(["Indicator (difference vs control)", ...cols]);
      MATRIX_ROWS.forEach((r) => {
        const row = [r.label];
        ordered.forEach((c) => {
          if (!ctrl) {
            row.push("");
            return;
          }
          if (c.isControl) {
            row.push(0);
            return;
          }
          const base = safeNum(ctrl[r.key], 0);
          const val = safeNum(c[r.key], 0);
          row.push(r.key === "rank" ? (safeNum(c.rank, 0) - safeNum(ctrl.rank, 0)) : (val - base));
        });
        relAoA.push(row);
      });
      const wsRel = XLSX.utils.aoa_to_sheet(relAoA);
      XLSX.utils.book_append_sheet(wb, wsRel, "Results_Relative");
    }

    return wb;
  }

  function writeWorkbook(wb, filename) {
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadBlob(filename, new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  }

  function exportResultsExcel() {
    try {
      const wb = buildWorkbookFromState(STATE, { includeResults: true, blank: false });
      const p = STATE.project?.projectName ? STATE.project.projectName.replace(/[^\w\-]+/g, "_").slice(0, 50) : "Project";
      const filename = `${TOOL_NAME.replace(/\s+/g, "_")}_${p}_Results_${todayISO()}.xlsx`;
      writeWorkbook(wb, filename);
      toast("Results exported to Excel.", "info", 2500);
    } catch (e) {
      console.error(e);
      toast("Could not export Excel. Please try again.", "error", 4500);
    }
  }

  function bindExcelWorkflow() {
    const btnBlank = q("downloadTemplate");
    const btnScenario = q("downloadScenarioTemplate");
    const btnSample = q("downloadSample");
    const btnParse = q("parseExcel");
    const btnImport = q("importExcel");
    const input = q("excelFileInput");
    const status = q("excelParseStatus");

    if (btnBlank) btnBlank.addEventListener("click", () => {
      try {
        const wb = buildWorkbookFromState(STATE, { includeResults: false, blank: true });
        const filename = `${TOOL_NAME.replace(/\s+/g, "_")}_Blank_Template_${todayISO()}.xlsx`;
        writeWorkbook(wb, filename);
        toast("Blank Excel template downloaded.", "info", 2500);
      } catch (e) {
        console.error(e);
        toast("Could not build blank template.", "error", 4500);
      }
    });

    if (btnScenario) btnScenario.addEventListener("click", () => {
      try {
        const wb = buildWorkbookFromState(STATE, { includeResults: true, blank: false });
        const filename = `${TOOL_NAME.replace(/\s+/g, "_")}_Scenario_Template_${todayISO()}.xlsx`;
        writeWorkbook(wb, filename);
        toast("Scenario template downloaded.", "info", 2500);
      } catch (e) {
        console.error(e);
        toast("Could not build scenario template.", "error", 4500);
      }
    });

    if (btnSample) btnSample.addEventListener("click", () => {
      try {
        const sample = buildSampleState();
        const wb = buildWorkbookFromState(sample, { includeResults: true, blank: false });
        const filename = `${TOOL_NAME.replace(/\s+/g, "_")}_Sample_FabaBean_${todayISO()}.xlsx`;
        writeWorkbook(wb, filename);
        toast("Sample Excel file downloaded.", "info", 2500);
      } catch (e) {
        console.error(e);
        toast("Could not build sample file.", "error", 4500);
      }
    });

    if (btnParse && input) {
      btnParse.addEventListener("click", () => input.click());
    }

    if (input) {
      input.addEventListener("change", async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        if (status) status.textContent = `Reading ${file.name} ...`;
        try {
          const arr = await file.arrayBuffer();
          const wb = XLSX.read(arr, { type: "array" });
          const parsed = parseWorkbookToState(wb);

          STATE.excel.parsed = parsed.state;
          STATE.excel.parsedSummary = parsed.summary;

          if (status) status.textContent = parsed.summaryText;
          toast("Excel file parsed. Review the summary, then apply.", "info", 3500);

          if (btnImport) btnImport.classList.remove("ghost");
        } catch (err) {
          console.error(err);
          STATE.excel.parsed = null;
          STATE.excel.parsedSummary = null;
          if (status) status.textContent = `Could not parse Excel file. ${safeStr(err?.message, "")}`.trim();
          toast("Excel parsing failed. Check sheet names and headers.", "error", 5000);
        } finally {
          // allow re-selecting same file
          input.value = "";
        }
      });
    }

    if (btnImport) {
      btnImport.addEventListener("click", () => {
        if (!STATE.excel.parsed) {
          toast("No parsed Excel data to apply yet. Select and parse a file first.", "warn", 4500);
          return;
        }
        try {
          // Apply parsed (keeping UI selections)
          const keepUi = deepClone(STATE.ui);
          const keepCache = deepClone(STATE.cache);
          const applied = deepClone(STATE.excel.parsed);

          applied.ui = { ...keepUi, ...applied.ui };
          applied.cache = keepCache;
          applied.excel = { parsed: null, parsedSummary: null };

          STATE = applied;

          // Rebind fields and rerender
          bindProjectFields();
          bindSettingsFields();

          rerenderAll();
          recalcAndRender(true);

          const st = q("excelParseStatus");
          if (st) st.textContent = "Applied. Your tool has been updated from Excel.";
          toast("Excel inputs applied successfully.", "info", 3500);
        } catch (e) {
          console.error(e);
          toast("Could not apply Excel data. Please check the file and try again.", "error", 5500);
        }
      });
    }
  }

  function getSheetJson(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) return null;
    return XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
  }

  function validateWorkbook(wb) {
    const missing = REQUIRED_SHEETS.filter((s) => !wb.Sheets[s]);
    return { ok: missing.length === 0, missing };
  }

  function parseWorkbookToState(wb) {
    const v = validateWorkbook(wb);
    if (!v.ok) {
      throw new Error(`Missing required sheets: ${v.missing.join(", ")}`);
    }

    const projectRows = getSheetJson(wb, "Project");
    const settingsRows = getSheetJson(wb, "Settings");
    const outputsRows = getSheetJson(wb, "Outputs");
    const trtRows = getSheetJson(wb, "Treatments");
    const toRows = getSheetJson(wb, "TreatmentOutputs");
    const benRows = getSheetJson(wb, "Benefits");
    const costRows = getSheetJson(wb, "Costs");

    const state = buildSampleState(); // base scaffold
    // Keep UI controls and other defaults; overwrite content
    state.outputs = [];
    state.treatments = [];
    state.benefits = [];
    state.costs = [];

    // Project
    const p0 = projectRows?.[0] || {};
    state.project.organisation = safeStr(p0.organisation, ORG_NAME) || ORG_NAME;
    state.project.projectName = safeStr(p0.project_name, "");
    state.project.projectLead = safeStr(p0.project_lead, "");
    state.project.analystNames = safeStr(p0.analyst_team, "");
    state.project.projectTeam = safeStr(p0.partners_or_team, "");
    state.project.lastUpdated = safeStr(p0.last_updated, todayISO());
    state.project.contactEmail = safeStr(p0.contact_email, "");
    state.project.contactPhone = safeStr(p0.contact_phone, "");
    state.project.projectSummary = safeStr(p0.short_summary, "");
    state.project.projectGoal = safeStr(p0.project_goal, "");
    state.project.withProject = safeStr(p0.with_project, "");
    state.project.withoutProject = safeStr(p0.without_project, "");
    state.project.projectObjectives = safeStr(p0.objectives, "");
    state.project.projectActivities = safeStr(p0.activities, "");
    state.project.stakeholderGroups = safeStr(p0.stakeholder_groups, "");

    // Settings
    const s0 = settingsRows?.[0] || {};
    const startYear = safeNum(s0.start_year, 2026);
    const years = Math.max(1, Math.trunc(safeNum(s0.years, 5)));

    state.settings.startYear = Math.trunc(startYear);
    state.settings.projectStartYear = Math.trunc(safeNum(s0.project_start_year, startYear));
    state.settings.years = years;
    state.settings.systemType = safeStr(s0.system_type, "single") || "single";
    state.settings.discBase = safeNum(s0.discount_base_percent, 7);
    state.settings.discLow = safeNum(s0.discount_low_percent, 3);
    state.settings.discHigh = safeNum(s0.discount_high_percent, 10);
    state.settings.adoptLow = clamp(safeNum(s0.adoption_low, 0.6), 0, 1);
    state.settings.adoptBase = clamp(safeNum(s0.adoption_base, 0.85), 0, 1);
    state.settings.adoptHigh = clamp(safeNum(s0.adoption_high, 1), 0, 1);
    state.settings.allocateProjectItems = ynToBool(s0.allocate_project_items_by_area, true);
    state.settings.riskLow = clamp(safeNum(s0.risk_low, 0.05), 0, 1);
    state.settings.riskBase = clamp(safeNum(s0.risk_base, 0.12), 0, 1);
    state.settings.riskHigh = clamp(safeNum(s0.risk_high, 0.25), 0, 1);
    state.settings.outputAssumptions = safeStr(s0.notes_on_assumptions, "");

    // Outputs (allow blank IDs; map by name)
    const outNameToId = new Map();
    (outputsRows || []).forEach((r) => {
      const name = safeStr(r.output_name, "").trim();
      if (!name) return;
      const id = safeStr(r.output_id, "").trim() || uid("out");
      if (outNameToId.has(name.toLowerCase())) return; // keep first
      const o = {
        id,
        name,
        unit: safeStr(r.unit, ""),
        valuePerUnit: safeNum(r.value_per_unit_aud, 0),
        notes: safeStr(r.notes, ""),
        source: safeStr(r.source, ""),
      };
      state.outputs.push(o);
      outNameToId.set(name.toLowerCase(), id);
    });

    // Treatments (allow blank IDs; ensure one control)
    const trtNameToId = new Map();
    (trtRows || []).forEach((r) => {
      const name = safeStr(r.treatment_name, "").trim();
      if (!name) return;
      const id = safeStr(r.treatment_id, "").trim() || uid("trt");
      if (trtNameToId.has(name.toLowerCase())) return; // keep first
      const t = {
        id,
        name,
        isControl: ynToBool(r.is_control, false),
        maxAreaHa: safeNum(r.max_area_ha, 0),
        labourCostPerHa: safeNum(r.labour_cost_per_ha_aud, 0),
        inputCostPerHa: safeNum(r.input_cost_per_ha_aud, 0),
        notes: safeStr(r.notes, ""),
        source: safeStr(r.source, ""),
        outputs: [],
      };
      state.treatments.push(t);
      trtNameToId.set(name.toLowerCase(), id);
    });

    // If user provided no outputs/treatments, keep at least a control scaffold
    if (state.outputs.length === 0) {
      state.outputs.push({ id: uid("out"), name: "Output 1", unit: "", valuePerUnit: 0, notes: "", source: "" });
    }
    if (state.treatments.length === 0) {
      state.treatments.push({
        id: uid("trt"),
        name: "Control",
        isControl: true,
        maxAreaHa: 0,
        labourCostPerHa: 0,
        inputCostPerHa: 0,
        notes: "",
        source: "",
        outputs: [],
      });
    }

    // TreatmentOutputs: map by IDs first, else by names if supplied
    const outIdSet = new Set(state.outputs.map((o) => o.id));
    const trtIdSet = new Set(state.treatments.map((t) => t.id));

    const tryMapOutputId = (row) => {
      const oid = safeStr(row.output_id, "").trim();
      if (oid && outIdSet.has(oid)) return oid;
      const oname = safeStr(row.output_name, "").trim();
      if (oname) return outNameToId.get(oname.toLowerCase()) || null;
      return null;
    };

    const tryMapTreatmentId = (row) => {
      const tid = safeStr(row.treatment_id, "").trim();
      if (tid && trtIdSet.has(tid)) return tid;
      const tname = safeStr(row.treatment_name, "").trim();
      if (tname) return trtNameToId.get(tname.toLowerCase()) || null;
      return null;
    };

    (toRows || []).forEach((r) => {
      const tid = tryMapTreatmentId(r);
      const oid = tryMapOutputId(r);
      if (!tid || !oid) return;

      const rowId = safeStr(r.row_id, "").trim() || uid("to");
      const qty = safeNum(r.qty_per_ha, 0);
      const yf = safeNum(r.year_from, startYear);
      const yt = safeNum(r.year_to, startYear);

      const t = state.treatments.find((x) => x.id === tid);
      if (!t) return;
      t.outputs.push({
        id: rowId,
        outputId: oid,
        qtyPerHa: qty,
        yearFrom: yf,
        yearTo: yt,
        notes: safeStr(r.notes, ""),
      });
    });

    // Benefits
    (benRows || []).forEach((r) => {
      const name = safeStr(r.benefit_name, "").trim();
      if (!name) return;
      state.benefits.push({
        id: safeStr(r.benefit_id, "").trim() || uid("ben"),
        name,
        annualValue: safeNum(r.annual_value_aud, 0),
        yearFrom: safeNum(r.year_from, startYear),
        yearTo: safeNum(r.year_to, startYear),
        notes: safeStr(r.notes, ""),
        source: safeStr(r.source, ""),
      });
    });

    // Costs
    (costRows || []).forEach((r) => {
      const name = safeStr(r.cost_name, "").trim();
      if (!name) return;
      state.costs.push({
        id: safeStr(r.cost_id, "").trim() || uid("cst"),
        name,
        annualValue: safeNum(r.annual_value_aud, 0),
        yearFrom: safeNum(r.year_from, startYear),
        yearTo: safeNum(r.year_to, startYear),
        notes: safeStr(r.notes, ""),
        source: safeStr(r.source, ""),
      });
    });

    // Enforce control uniqueness
    ensureSingleControl(state);

    // Validate year ranges (soft warnings)
    const warnings = [];
    const minY = state.settings.startYear;
    const maxY = state.settings.startYear + state.settings.years - 1;

    const inRange = (y) => Number.isFinite(y) && y >= minY && y <= maxY;
    state.treatments.forEach((t) => {
      (t.outputs || []).forEach((r) => {
        if (!inRange(safeNum(r.yearFrom, NaN)) || !inRange(safeNum(r.yearTo, NaN))) {
          warnings.push(`TreatmentOutputs row has year range outside horizon: ${t.name}`);
        }
      });
    });
    state.benefits.forEach((b) => {
      if (!inRange(safeNum(b.yearFrom, NaN)) || !inRange(safeNum(b.yearTo, NaN))) warnings.push(`Benefit outside horizon: ${b.name}`);
    });
    state.costs.forEach((c) => {
      if (!inRange(safeNum(c.yearFrom, NaN)) || !inRange(safeNum(c.yearTo, NaN))) warnings.push(`Cost outside horizon: ${c.name}`);
    });

    const summary = {
      outputs: state.outputs.length,
      treatments: state.treatments.length,
      treatmentOutputs: state.treatments.reduce((s, t) => s + (t.outputs?.length || 0), 0),
      benefits: state.benefits.length,
      costs: state.costs.length,
      warnings,
    };

    const summaryText = [
      `Parsed successfully.`,
      `Outputs: ${summary.outputs}. Treatments: ${summary.treatments}. Treatment output rows: ${summary.treatmentOutputs}.`,
      `Benefits: ${summary.benefits}. Costs: ${summary.costs}.`,
      ...(warnings.length ? [`Warnings: ${warnings.slice(0, 6).join(" | ")}${warnings.length > 6 ? " | ..." : ""}`] : []),
      `Click "Apply parsed Excel to tool" to update.`,
    ].join(" ");

    return { state, summary, summaryText };
  }

  /********************************************************************
   * Copy results table for Word (HTML + TSV)
   ********************************************************************/
  async function copyResultsForWord() {
    try {
      const results = STATE.cache.lastResults || computeAllResults(STATE);
      const ordered = results.indicators
        .slice()
        .sort((a, b) => {
          if (a.isControl && !b.isControl) return -1;
          if (!a.isControl && b.isControl) return 1;
          return (a.rank ?? 9999) - (b.rank ?? 9999);
        });

      const cols = ordered.map((x) => x.isControl ? `${x.treatmentName} (Control)` : x.treatmentName);

      // Build a clean HTML table (Word-friendly)
      const htmlRows = [];
      htmlRows.push(`<tr><th>Indicator</th>${cols.map((c) => `<th>${escapeHtml(c)}</th>`).join("")}</tr>`);
      MATRIX_ROWS.forEach((r) => {
        const cells = ordered.map((c) => {
          const v = c[r.key];
          if (r.key === "npv" || r.key === "pvBenefits" || r.key === "pvCosts") return fmtCurrency2.format(v);
          if (r.key === "rank") return String(c.rank ?? "n/a");
          if (v === Infinity) return "∞";
          return fmtNumber3.format(v);
        });
        htmlRows.push(`<tr><td>${escapeHtml(r.label)}</td>${cells.map((x) => `<td>${escapeHtml(x)}</td>`).join("")}</tr>`);
      });

      const html = `
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;">
          ${htmlRows.join("")}
        </table>
      `.trim();

      // TSV fallback
      const tsvLines = [];
      tsvLines.push(["Indicator", ...cols].join("\t"));
      MATRIX_ROWS.forEach((r) => {
        const cells = ordered.map((c) => {
          const v = c[r.key];
          if (r.key === "npv" || r.key === "pvBenefits" || r.key === "pvCosts") return fmtCurrency2.format(v);
          if (r.key === "rank") return String(c.rank ?? "n/a");
          if (v === Infinity) return "∞";
          return fmtNumber3.format(v);
        });
        tsvLines.push([r.label, ...cells].join("\t"));
      });
      const tsv = tsvLines.join("\n");

      if (navigator.clipboard && window.ClipboardItem) {
        const item = new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([tsv], { type: "text/plain" }),
        });
        await navigator.clipboard.write([item]);
      } else {
        // Fallback: plain text only
        await navigator.clipboard.writeText(tsv);
      }

      toast("Results table copied. Paste into Word.", "info", 3000);
    } catch (e) {
      console.error(e);
      toast("Could not copy table. Try exporting Excel instead.", "error", 4500);
    }
  }

  /********************************************************************
   * AI prompt builder + exports
   ********************************************************************/
  function buildAiPromptFromCurrentResults() {
    const results = STATE.cache.lastResults || computeAllResults(STATE);
    const control = results.control;

    const audience = q("aiAudience")?.value || "farmer";
    const length = q("aiLength")?.value || "medium";
    const focus = q("aiFocus")?.value || "economic";

    const s = STATE.settings;
    const p = STATE.project;

    const ordered = results.indicators
      .slice()
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;
        return (a.rank ?? 9999) - (b.rank ?? 9999);
      });

    const top = ordered.filter((x) => !x.isControl).slice(0, 3);
    const bottom = ordered.filter((x) => !x.isControl).slice(-3);

    const rowsForPrompt = ordered.map((t) => ({
      treatment: t.isControl ? `${t.treatmentName} (Control)` : t.treatmentName,
      is_control: t.isControl,
      rank: t.rank,
      npv_aud: Math.round(t.npv),
      pv_benefits_aud: Math.round(t.pvBenefits),
      pv_costs_aud: Math.round(t.pvCosts),
      bcr: t.bcr === Infinity ? "INF" : Number(t.bcr.toFixed(4)),
      roi: t.roi === Infinity ? "INF" : Number(t.roi.toFixed(4)),
      effective_area_ha: Number((t.effArea ?? 0).toFixed(2)),
      adoption_base: Number((t.adoption ?? 0).toFixed(3)),
      risk_base: Number((t.risk ?? 0).toFixed(3)),
    }));

    const audienceGuide = {
      farmer: "Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Focus on what drives results and what could be changed.",
      advisor: "Use clear extension and advisory language. Explain indicators briefly and emphasise practical implications and sensitivities.",
      policy: "Use program and policy stakeholder language. Explain indicators and assumptions carefully. Emphasise uncertainty and transparency. Do not prescribe.",
      technical: "Use technically precise language and include brief formulas. Still remain non-prescriptive.",
    }[audience] || "Use clear plain language.";

    const lengthGuide = {
      short: "Aim for a one-page interpretation (around 600 to 900 words).",
      medium: "Aim for a two to three page interpretation (around 1200 to 1800 words).",
      long: "Aim for a longer policy-brief style interpretation (around 2000 to 3000 words) with headings and a short executive summary.",
    }[length] || "Use a moderate length.";

    const focusGuide = {
      economic: "Focus primarily on economic performance and what drives PV benefits and PV costs.",
      risk: "Focus on risk and uncertainty and how risk and adoption assumptions affect results.",
      practical: "Focus on practical actions that could realistically improve low-performing treatments, without dictating decisions.",
      balanced: "Provide a balanced interpretation across performance, assumptions, uncertainty, and practical levers.",
    }[focus] || "Provide a balanced interpretation.";

    const definitions = [
      "Definitions (use these consistently):",
      "NPV = PV benefits − PV costs. Positive NPV indicates net economic gain relative to zero baseline for that treatment scenario.",
      "PV benefits and PV costs are discounted sums over time using the base discount rate.",
      "BCR = PV benefits ÷ PV costs. Values above 1 imply benefits exceed costs in present value terms.",
      "ROI = NPV ÷ PV costs. Interpretable as net gain per dollar of PV cost.",
      "The control is shown alongside treatments for direct comparison. When comparing against control, explain both absolute levels and differences.",
      "Risk reduces realised benefits (not costs) in this tool. Adoption scales effective area.",
    ].join("\n");

    const improvementLevers = [
      "Guidance for low-performing treatments (BCR near or below 1, or low ROI/NPV):",
      "Suggest practical improvement levers as options, not rules. For example:",
      "Reduce variable costs per hectare through input substitution, bulk purchasing, labour efficiency, or better timing.",
      "Increase realised yields or output quantities through agronomic practice changes, better application methods, or improved varietal selection.",
      "Increase output value assumptions by using conservative vs optimistic price scenarios and explaining sensitivity.",
      "Improve adoption feasibility or target the treatment to suitable paddocks rather than full area rollout.",
      "Reduce implementation risk by training, monitoring, staged rollout, and addressing known failure points.",
      "Clarify time profile: benefits may arrive later; consider whether the analysis horizon or ramp-up assumptions are realistic.",
    ].join("\n");

    const nonPrescriptive = [
      "Important constraints:",
      "Do not tell the user what to choose.",
      "Do not impose thresholds as rules.",
      "Explain why a treatment performs well or poorly and what factors drive that result.",
      "Treat this as decision support: show trade-offs, sensitivities, and plausible improvement paths.",
    ].join("\n");

    const context = {
      tool_name: TOOL_NAME,
      project: {
        project_name: p.projectName || "",
        organisation: p.organisation || ORG_NAME,
        last_updated: p.lastUpdated || "",
        short_summary: p.projectSummary || "",
        project_goal: p.projectGoal || "",
      },
      settings: {
        analysis_year_start: safeNum(s.startYear, 2026),
        years: safeNum(s.years, 5),
        discount_rate_base_percent: safeNum(s.discBase, 7),
        adoption_base: safeNum(s.adoptBase, 1),
        risk_base: safeNum(s.riskBase, 0),
        allocate_project_items_by_area: Boolean(s.allocateProjectItems),
      },
      control: control ? {
        name: control.treatmentName,
        npv_aud: Math.round(control.npv),
        bcr: control.bcr === Infinity ? "INF" : Number(control.bcr.toFixed(4)),
        roi: control.roi === Infinity ? "INF" : Number(control.roi.toFixed(4)),
      } : null,
      ranking_metric: results.metricKey,
      results_table: rowsForPrompt,
      highlights: {
        top_treatments: top.map((x) => ({ name: x.treatmentName, rank: x.rank, npv_aud: Math.round(x.npv), bcr: x.bcr === Infinity ? "INF" : Number(x.bcr.toFixed(4)) })),
        bottom_treatments: bottom.map((x) => ({ name: x.treatmentName, rank: x.rank, npv_aud: Math.round(x.npv), bcr: x.bcr === Infinity ? "INF" : Number(x.bcr.toFixed(4)) })),
      },
    };

    const prompt = [
      `You are interpreting results from a farm cost–benefit analysis tool called "${TOOL_NAME}".`,
      `${audienceGuide}`,
      `${lengthGuide}`,
      `${focusGuide}`,
      ``,
      nonPrescriptive,
      ``,
      definitions,
      ``,
      improvementLevers,
      ``,
      "Task: Produce a plain-English interpretation with headings.",
      "Include: (i) a short overview of what the results show; (ii) which treatments perform better or worse and why; (iii) comparison against control; (iv) explanation of key indicators; (v) trade-offs and assumptions; (vi) practical improvement options for low performers (as guidance, not rules).",
      "If values look counter-intuitive, explain plausible drivers (costs, yields, prices, adoption, risk, time profile) rather than assuming error.",
      ``,
      "Data (JSON):",
      JSON.stringify(context, null, 2),
    ].join("\n");

    return prompt;
  }

  function bindAiTab() {
    const buildBtn = q("buildAiPrompt");
    const copyBtn = q("copyAiPrompt");
    const dlBtn = q("exportPromptTxt");
    const ta = q("copilotPreview");

    if (buildBtn && ta) {
      buildBtn.addEventListener("click", () => {
        try {
          const prompt = buildAiPromptFromCurrentResults();
          ta.value = prompt;
          toast("Prompt built from current results.", "info", 2500);
        } catch (e) {
          console.error(e);
          toast("Could not build prompt. Recalculate results and try again.", "error", 5000);
        }
      });
    }

    if (copyBtn && ta) {
      copyBtn.addEventListener("click", async () => {
        try {
          await navigator.clipboard.writeText(ta.value || "");
          toast("Prompt copied. Paste into Copilot or ChatGPT.", "info", 2500);
        } catch (e) {
          console.error(e);
          toast("Could not copy. Select the text and copy manually.", "warn", 4500);
        }
      });
    }

    if (dlBtn && ta) {
      dlBtn.addEventListener("click", () => {
        const p = STATE.project?.projectName ? STATE.project.projectName.replace(/[^\w\-]+/g, "_").slice(0, 50) : "Project";
        downloadText(`${TOOL_NAME.replace(/\s+/g, "_")}_${p}_AI_Prompt_${todayISO()}.txt`, ta.value || "");
        toast("Prompt downloaded.", "info", 2500);
      });
    }

    // Brief downloads (user pastes AI output)
    const paste = q("aiOutputPaste");
    const dlDoc = q("downloadBriefDoc");
    const dlTxt = q("downloadBriefTxt");
    const prn = q("printBriefPdf");

    if (dlTxt && paste) {
      dlTxt.addEventListener("click", () => {
        const text = paste.value || "";
        const p = STATE.project?.projectName ? STATE.project.projectName.replace(/[^\w\-]+/g, "_").slice(0, 50) : "Project";
        downloadText(`${TOOL_NAME.replace(/\s+/g, "_")}_${p}_AI_Brief_${todayISO()}.txt`, text);
        toast("Brief downloaded as text.", "info", 2500);
      });
    }

    if (dlDoc && paste) {
      dlDoc.addEventListener("click", () => {
        const content = (paste.value || "").trim();
        if (!content) {
          toast("Paste an AI-generated interpretation first (optional), then download.", "warn", 4500);
          return;
        }
        const p = STATE.project || {};
        const title = `${TOOL_NAME} — AI interpretation`;
        const html = `
          <!doctype html>
          <html>
          <head>
            <meta charset="utf-8">
            <title>${escapeHtml(title)}</title>
          </head>
          <body style="font-family:Calibri,Arial,sans-serif; font-size:11pt; line-height:1.4;">
            <h1 style="font-size:16pt; margin:0 0 8pt 0;">${escapeHtml(TOOL_NAME)}</h1>
            <div style="margin:0 0 10pt 0;">
              <div><b>Project:</b> ${escapeHtml(p.projectName || "")}</div>
              <div><b>Organisation:</b> ${escapeHtml(p.organisation || ORG_NAME)}</div>
              <div><b>Date:</b> ${escapeHtml(todayISO())}</div>
            </div>
            <hr>
            <pre style="white-space:pre-wrap; font-family:Calibri,Arial,sans-serif; margin-top:10pt;">${escapeHtml(content)}</pre>
          </body>
          </html>
        `.trim();

        const blob = new Blob([html], { type: "application/msword" });
        const proj = p.projectName ? p.projectName.replace(/[^\w\-]+/g, "_").slice(0, 50) : "Project";
        downloadBlob(`${TOOL_NAME.replace(/\s+/g, "_")}_${proj}_AI_Brief_${todayISO()}.doc`, blob);
        toast("Word file downloaded.", "info", 2500);
      });
    }

    if (prn && paste) {
      prn.addEventListener("click", () => {
        const content = (paste.value || "").trim();
        if (!content) {
          toast("Paste an AI-generated interpretation first (optional), then print.", "warn", 4500);
          return;
        }
        const w = window.open("", "_blank", "noopener,noreferrer");
        if (!w) {
          toast("Popup blocked. Allow popups to print.", "warn", 5000);
          return;
        }
        const p = STATE.project || {};
        w.document.write(`
          <!doctype html>
          <html><head><meta charset="utf-8"><title>${escapeHtml(TOOL_NAME)} — Brief</title></head>
          <body style="font-family:Arial,sans-serif; margin:24px;">
            <h1 style="margin:0 0 8px 0;">${escapeHtml(TOOL_NAME)}</h1>
            <div style="margin:0 0 12px 0;">
              <div><b>Project:</b> ${escapeHtml(p.projectName || "")}</div>
              <div><b>Date:</b> ${escapeHtml(todayISO())}</div>
            </div>
            <hr>
            <pre style="white-space:pre-wrap; font-family:Arial,sans-serif;">${escapeHtml(content)}</pre>
          </body></html>
        `);
        w.document.close();
        w.focus();
        setTimeout(() => {
          w.print();
          w.close();
        }, 250);
      });
    }
  }

  /********************************************************************
   * Project save/load (JSON)
   ********************************************************************/
  function bindProjectSaveLoad() {
    const saveBtn = q("saveProject");
    const loadBtn = q("loadProject");
    const loadInput = q("loadFile");
    const startBtn = q("startBtn");
    const startBtn2 = q("startBtn-duplicate");

    if (startBtn) startBtn.addEventListener("click", () => setActiveTab("project"));
    if (startBtn2) startBtn2.addEventListener("click", () => setActiveTab("project"));

    if (saveBtn) {
      saveBtn.addEventListener("click", () => {
        try {
          const payload = {
            ...deepClone(STATE),
            savedAt: new Date().toISOString(),
            toolName: TOOL_NAME,
          };
          const p = STATE.project?.projectName ? STATE.project.projectName.replace(/[^\w\-]+/g, "_").slice(0, 50) : "Project";
          downloadText(`${TOOL_NAME.replace(/\s+/g, "_")}_${p}_${todayISO()}.json`, JSON.stringify(payload, null, 2), "application/json;charset=utf-8");
          toast("Project saved as JSON.", "info", 2500);
        } catch (e) {
          console.error(e);
          toast("Could not save project.", "error", 4500);
        }
      });
    }

    if (loadBtn && loadInput) {
      loadBtn.addEventListener("click", () => loadInput.click());
      loadInput.addEventListener("change", async (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        try {
          const text = await file.text();
          const obj = JSON.parse(text);

          // Minimal validation
          if (!obj || typeof obj !== "object") throw new Error("Invalid JSON.");
          if (!obj.project || !obj.settings) throw new Error("Missing project/settings in file.");

          // Preserve UI defaults if missing
          const keepUi = deepClone(STATE.ui);
          const merged = deepClone(obj);
          merged.ui = { ...keepUi, ...(merged.ui || {}) };
          merged.toolName = TOOL_NAME;

          STATE = merged;
          ensureSingleControl(STATE);

          bindProjectFields();
          bindSettingsFields();
          rerenderAll();
          recalcAndRender(true);

          toast("Project loaded.", "info", 2500);
        } catch (err) {
          console.error(err);
          toast("Could not load project file. Check the JSON format.", "error", 5500);
        } finally {
          loadInput.value = "";
        }
      });
    }
  }

  /********************************************************************
   * Simulation (Monte Carlo) — focused on decision support, not exclusion
   ********************************************************************/
  function seededRng(seed) {
    // Simple LCG for reproducibility if seed provided
    let s = Math.trunc(seed) >>> 0;
    return () => {
      s = (1664525 * s + 1013904223) >>> 0;
      return s / 4294967296;
    };
  }

  function uniform(rng, a, b) {
    return a + (b - a) * rng();
  }

  function varyByPct(rng, x, pct) {
    const p = Math.max(0, pct) / 100;
    const m = uniform(rng, 1 - p, 1 + p);
    return x * m;
  }

  function computeIndicatorsForTreatmentUnderScenario(state, treatment, overrides) {
    const base = deepClone(state);
    const s = base.settings;

    if (overrides.discBase != null) s.discBase = overrides.discBase;
    if (overrides.adoptBase != null) s.adoptBase = overrides.adoptBase;
    if (overrides.riskBase != null) s.riskBase = overrides.riskBase;

    // Optional value perturbations
    if (overrides.outputsValueMap) {
      (base.outputs || []).forEach((o) => {
        if (overrides.outputsValueMap.has(o.id)) o.valuePerUnit = overrides.outputsValueMap.get(o.id);
      });
    }
    if (overrides.treatCostMap) {
      (base.treatments || []).forEach((t) => {
        if (overrides.treatCostMap.has(t.id)) {
          const { labourCostPerHa, inputCostPerHa } = overrides.treatCostMap.get(t.id);
          t.labourCostPerHa = labourCostPerHa;
          t.inputCostPerHa = inputCostPerHa;
        }
      });
    }
    if (overrides.costAnnualMap) {
      (base.costs || []).forEach((c) => {
        if (overrides.costAnnualMap.has(c.id)) c.annualValue = overrides.costAnnualMap.get(c.id);
      });
    }

    const t2 = (base.treatments || []).find((x) => x.id === treatment.id);
    return computeIndicatorsForTreatment(base, t2 || treatment);
  }

  function drawHistogram(canvas, values, bins = 18, label = "") {
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const w = canvas.width;
    const h = canvas.height;

    ctx.clearRect(0, 0, w, h);

    if (!values.length) {
      ctx.fillText("No data", 10, 20);
      return;
    }

    const finite = values.filter((x) => Number.isFinite(x));
    if (!finite.length) {
      ctx.fillText("No finite values", 10, 20);
      return;
    }

    const min = Math.min(...finite);
    const max = Math.max(...finite);
    const span = max - min || 1;

    const counts = new Array(bins).fill(0);
    finite.forEach((v) => {
      const k = Math.min(bins - 1, Math.max(0, Math.floor(((v - min) / span) * bins)));
      counts[k] += 1;
    });

    const maxC = Math.max(...counts) || 1;

    // axes
    ctx.lineWidth = 1;
    ctx.strokeStyle = "#000";
    ctx.beginPath();
    ctx.moveTo(36, 10);
    ctx.lineTo(36, h - 26);
    ctx.lineTo(w - 10, h - 26);
    ctx.stroke();

    // bars (default fill)
    const plotW = (w - 10) - 36;
    const plotH = (h - 26) - 10;
    const barW = plotW / bins;

    for (let i = 0; i < bins; i++) {
      const bh = (counts[i] / maxC) * plotH;
      const x = 36 + i * barW;
      const y = (h - 26) - bh;
      ctx.fillRect(x + 1, y, Math.max(1, barW - 2), bh);
    }

    // labels
    ctx.fillStyle = "#000";
    ctx.font = "12px Arial";
    ctx.fillText(label, 36, 10);
    ctx.font = "10px Arial";
    ctx.fillText(fmtNumber2.format(min), 36, h - 8);
    ctx.fillText(fmtNumber2.format(max), w - 10 - 46, h - 8);
  }

  function bindSimulation() {
    const runBtn = q("runSim");
    if (!runBtn) return;

    runBtn.addEventListener("click", () => {
      try {
        const baseResults = STATE.cache.lastResults || computeAllResults(STATE);

        // Choose the best-ranked non-control treatment for simulation focus (still decision support)
        const focus = baseResults.indicators
          .filter((x) => !x.isControl)
          .slice()
          .sort((a, b) => (a.rank ?? 9999) - (b.rank ?? 9999))[0];

        if (!focus) {
          toast("Add at least one non-control treatment before running simulation.", "warn", 4500);
          return;
        }

        const s = STATE.settings;
        const N = Math.max(100, Math.trunc(safeNum(getInputValue("simN"), s.simN || 1000)));
        const discLow = safeNum(s.discLow, 3) / 100;
        const discHigh = safeNum(s.discHigh, 10) / 100;
        const adoptLow = clamp(safeNum(s.adoptLow, 0.6), 0, 1);
        const adoptHigh = clamp(safeNum(s.adoptHigh, 1), 0, 1);
        const riskLow = clamp(safeNum(s.riskLow, 0.05), 0, 1);
        const riskHigh = clamp(safeNum(s.riskHigh, 0.25), 0, 1);

        const varPct = clamp(safeNum(getInputValue("simVarPct"), s.simVarPct || 0), 0, 100);
        const varyOutputs = ynToBool(getInputValue("simVaryOutputs"), Boolean(s.simVaryOutputs));
        const varyTreatCosts = ynToBool(getInputValue("simVaryTreatCosts"), Boolean(s.simVaryTreatCosts));
        const varyInputCosts = ynToBool(getInputValue("simVaryInputCosts"), Boolean(s.simVaryInputCosts));

        const seedVal = safeStr(getInputValue("randSeed"), "").trim();
        const rng = seedVal ? seededRng(safeNum(seedVal, 12345)) : Math.random;

        const npvs = [];
        const bcrs = [];

        const status = q("simStatus");
        if (status) status.textContent = `Running ${N} simulations for: ${focus.treatmentName}.`;

        const tFocus = (STATE.treatments || []).find((t) => t.id === focus.treatmentId);

        // Pre-store base values for perturbation
        const baseOutputVals = new Map((STATE.outputs || []).map((o) => [o.id, safeNum(o.valuePerUnit, 0)]));
        const baseTreatCosts = new Map((STATE.treatments || []).map((t) => [t.id, { labourCostPerHa: safeNum(t.labourCostPerHa, 0), inputCostPerHa: safeNum(t.inputCostPerHa, 0) }]));
        const baseCostAnnual = new Map((STATE.costs || []).map((c) => [c.id, safeNum(c.annualValue, 0)]));

        // Run
        for (let i = 0; i < N; i++) {
          const disc = uniform(rng, discLow, discHigh) * 100; // percent
          const adopt = uniform(rng, adoptLow, adoptHigh);
          const risk = uniform(rng, riskLow, riskHigh);

          const overrides = { discBase: disc, adoptBase: adopt, riskBase: risk };

          if (varyOutputs) {
            const map = new Map();
            baseOutputVals.forEach((v, k) => map.set(k, varyByPct(rng, v, varPct)));
            overrides.outputsValueMap = map;
          }

          if (varyTreatCosts) {
            const map = new Map();
            baseTreatCosts.forEach((v, k) => {
              map.set(k, {
                labourCostPerHa: varyByPct(rng, v.labourCostPerHa, varPct),
                inputCostPerHa: varyByPct(rng, v.inputCostPerHa, varPct),
              });
            });
            overrides.treatCostMap = map;
          }

          if (varyInputCosts) {
            const map = new Map();
            baseCostAnnual.forEach((v, k) => map.set(k, varyByPct(rng, v, varPct)));
            overrides.costAnnualMap = map;
          }

          const ind = computeIndicatorsForTreatmentUnderScenario(STATE, tFocus, overrides);
          npvs.push(ind.npv);
          bcrs.push(ind.bcr === Infinity ? NaN : ind.bcr);
        }

        // Summary stats
        const finiteNpv = npvs.filter(Number.isFinite);
        const finiteBcr = bcrs.filter(Number.isFinite);

        const stats = (arr) => {
          const x = arr.slice().sort((a, b) => a - b);
          const n = x.length || 1;
          const mean = x.reduce((s, v) => s + v, 0) / n;
          const median = x[Math.floor(n / 2)];
          return { min: x[0], max: x[n - 1], mean, median };
        };

        const npvS = finiteNpv.length ? stats(finiteNpv) : null;
        const bcrS = finiteBcr.length ? stats(finiteBcr) : null;

        // Probabilities
        const prob = (arr, fn) => {
          const x = arr.filter(Number.isFinite);
          if (!x.length) return NaN;
          return x.filter(fn).length / x.length;
        };

        const target = safeNum(getInputValue("targetBCR"), safeNum(s.targetBCR, 2));

        // Write to UI
        const setVal = (id, v, fmtFn) => {
          const node = q(id);
          if (!node) return;
          if (!Number.isFinite(v)) node.textContent = "n/a";
          else node.textContent = fmtFn ? fmtFn(v) : fmtNumber2.format(v);
        };

        if (npvS) {
          setVal("simNpvMin", npvS.min, (x) => fmtCurrency.format(x));
          setVal("simNpvMax", npvS.max, (x) => fmtCurrency.format(x));
          setVal("simNpvMean", npvS.mean, (x) => fmtCurrency.format(x));
          setVal("simNpvMedian", npvS.median, (x) => fmtCurrency.format(x));
          setVal("simNpvProb", prob(finiteNpv, (x) => x > 0), (x) => fmtPercent1.format(x));
        } else {
          ["simNpvMin","simNpvMax","simNpvMean","simNpvMedian","simNpvProb"].forEach((id) => q(id) && (q(id).textContent = "n/a"));
        }

        if (bcrS) {
          setVal("simBcrMin", bcrS.min, (x) => fmtNumber3.format(x));
          setVal("simBcrMax", bcrS.max, (x) => fmtNumber3.format(x));
          setVal("simBcrMean", bcrS.mean, (x) => fmtNumber3.format(x));
          setVal("simBcrMedian", bcrS.median, (x) => fmtNumber3.format(x));
          setVal("simBcrProb1", prob(finiteBcr, (x) => x > 1), (x) => fmtPercent1.format(x));
          setVal("simBcrProbTarget", prob(finiteBcr, (x) => x > target), (x) => fmtPercent1.format(x));
        } else {
          ["simBcrMin","simBcrMax","simBcrMean","simBcrMedian","simBcrProb1","simBcrProbTarget"].forEach((id) => q(id) && (q(id).textContent = "n/a"));
        }

        // Histograms
        drawHistogram(q("histNpv"), finiteNpv, 18, `NPV distribution (${focus.treatmentName})`);
        drawHistogram(q("histBcr"), finiteBcr, 18, `BCR distribution (${focus.treatmentName})`);

        if (status) status.textContent = `Simulation complete for: ${focus.treatmentName}.`;
        toast("Simulation complete. Results are for the best-ranked treatment (base case) to support exploration.", "info", 5000);
      } catch (e) {
        console.error(e);
        const status = q("simStatus");
        if (status) status.textContent = "Simulation failed.";
        toast("Simulation failed. Check inputs and try again.", "error", 6000);
      }
    });
  }

  /********************************************************************
   * Rerender all
   ********************************************************************/
  function rerenderAll() {
    renderOutputs(STATE);
    renderTreatments(STATE);
    renderBenefits(STATE);
    renderCosts(STATE);
    renderDatabaseTab(STATE);

    // Keep results controls in sync
    bindResultsControls();
  }

  /********************************************************************
   * Wire toolbar buttons
   ********************************************************************/
  function wireToolbarButtons() {
    q("addOutput")?.addEventListener("click", addOutput);
    q("addTreatment")?.addEventListener("click", addTreatment);
    q("addBenefit")?.addEventListener("click", addBenefit);
    q("addCost")?.addEventListener("click", addCost);
  }

  /********************************************************************
   * Initialise
   ********************************************************************/
  function init() {
    setLoading(true, "Loading scenario and interface");

    // Tabs
    wireTabs();

    // Toolbar
    wireToolbarButtons();

    // Save/load JSON
    bindProjectSaveLoad();

    // Bind forms
    bindProjectFields();
    bindSettingsFields();
    bindResultsControls();

    // Excel workflow
    bindExcelWorkflow();

    // AI tab
    bindAiTab();

    // Simulation
    bindSimulation();

    // Initial render
    rerenderAll();

    // Results controls sync
    const rm = q("rankMetric");
    if (rm) rm.value = STATE.ui.rankMetric || "bcr";
    const vm = q("resultsViewMode");
    if (vm) vm.value = STATE.ui.resultsViewMode || "absolute";
    const cp = q("columnsPerPage");
    if (cp) cp.value = String(STATE.ui.columnsPerPage || 6);
    const sd = q("showDeltas");
    if (sd) sd.value = String(Boolean(STATE.ui.showDeltas));

    // Active tab
    setActiveTab(STATE.ui.activeTab || "intro");

    // First calc
    recalcAndRender(false);

    // Hide loader
    setTimeout(() => setLoading(false), 250);

    // Accessibility: enter key on tablist
    const tablist = document.querySelector(".tabs-nav");
    if (tablist) {
      tablist.addEventListener("keydown", (e) => {
        if (!["ArrowLeft", "ArrowRight"].includes(e.key)) return;
        const tabs = Array.from(document.querySelectorAll(".tab-link"));
        const idx = tabs.findIndex((t) => t.classList.contains("active"));
        if (idx < 0) return;
        const next = e.key === "ArrowRight" ? (idx + 1) % tabs.length : (idx - 1 + tabs.length) % tabs.length;
        tabs[next].focus();
        setActiveTab(tabs[next].getAttribute("data-tab"));
        e.preventDefault();
      });
    }
  }

  // Boot
  window.addEventListener("DOMContentLoaded", init);
})();
