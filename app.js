// Farming CBA Tool - Newcastle Business School
// Production-grade: import pipeline (upload + paste) for TSV/CSV + data dictionary, validation + commit,
// replicate-specific control baselines, plot-level deltas, missing-safe treatment summaries,
// cost scaling, Data Checks panel, discounted CBA engine with sensitivity grid,
// per-treatment recurrence configuration + scenario save/load to localStorage,
// Results: leaderboard + comparison-to-control grid with deltas/colour cues + filters + narrative,
// Exports: cleaned TSV + summary CSV + sensitivity CSV + workbook (if XLSX available),
// AI Briefing: copy-ready narrative prompt (no bullets, no em dash, no abbreviations) + copy JSON,
// Toasts for all major actions.
//
// IMPORTANT: This script binds only to element IDs. All bindings are guarded (no-ops if an ID is absent).

(() => {
  "use strict";

  // =========================
  // 0) Small utilities
  // =========================
  const uid = () => Math.random().toString(36).slice(2, 10);
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const esc = s =>
    (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const slug = s =>
    (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  const isNum = x => typeof x === "number" && Number.isFinite(x);

  const fmt = n =>
    isFinite(n)
      ? Math.abs(n) >= 1000
        ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
        : n.toLocaleString(undefined, { maximumFractionDigits: 2 })
      : "n/a";
  const money = n => (isFinite(n) ? "$" + fmt(n) : "n/a");
  const ratio = n => (isFinite(n) ? fmt(n) : "n/a");
  const pct = n => (isFinite(n) ? fmt(n) + " per cent" : "n/a");

  const $id = id => document.getElementById(id);
  const nowISO = () => new Date().toISOString();

  function showToast(message) {
    const root = $id("toast-root") || document.body;
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

  function safeOn(el, event, handler, opts) {
    if (!el) return;
    el.addEventListener(event, handler, opts);
  }

  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    // Fallback: temporary textarea
    return new Promise((resolve, reject) => {
      try {
        const ta = document.createElement("textarea");
        ta.value = text;
        ta.setAttribute("readonly", "readonly");
        ta.style.position = "fixed";
        ta.style.top = "-1000px";
        document.body.appendChild(ta);
        ta.select();
        const ok = document.execCommand("copy");
        document.body.removeChild(ta);
        if (ok) resolve();
        else reject(new Error("Copy failed"));
      } catch (e) {
        reject(e);
      }
    });
  }

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

  // =========================
  // 1) Core model
  // =========================
  const DEFAULT_DISCOUNT_SCHEDULE = [
    { label: "2025-2034", from: 2025, to: 2034, low: 2, base: 4, high: 6 },
    { label: "2035-2044", from: 2035, to: 2044, low: 4, base: 7, high: 10 },
    { label: "2045-2054", from: 2045, to: 2054, low: 4, base: 7, high: 10 },
    { label: "2055-2064", from: 2055, to: 2064, low: 3, base: 6, high: 9 },
    { label: "2065-2074", from: 2065, to: 2074, low: 2, base: 5, high: 8 }
  ];

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
        "Identify soil amendment packages that deliver higher yields and acceptable returns after accounting for additional costs.",
      withProject:
        "Growers adopt high-performing amendment packages on relevant soils and use trial evidence to refine practice.",
      withoutProject:
        "Growers continue with baseline practice and do not access detailed economic evidence on soil amendments."
    },

    // Analysis settings (base case)
    config: {
      farmAreaHa: 100,
      startYear: new Date().getFullYear(),
      horizonYears: 10,
      grainPricePerTonne: 450,
      discountRatePct: 7,
      persistence: 1.0, // 1.0 means the yield effect does not decay between applications; lower values decay
      recurrenceMode: "configured", // configured | annual | once
      adoptionMultiplier: 1.0, // applied to the area under the alternative practice
      riskMultiplier: 0.0 // proportion reduction in yield effect; 0 means no reduction
    },

    // Sensitivity grid settings (base + low/high)
    sensitivity: {
      priceLow: 350,
      priceBase: 450,
      priceHigh: 550,
      discLow: 4,
      discBase: 7,
      discHigh: 10,
      persistenceLow: 0.6,
      persistenceBase: 1.0,
      persistenceHigh: 1.0,
      recurrenceModes: ["configured", "annual", "once"]
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

    // Outputs: values used to monetise deltas (per unit)
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t per ha", value: 450, source: "Input Directly" },
      { id: uid(), name: "Screenings", unit: "percentage point", value: -20, source: "Input Directly" },
      { id: uid(), name: "Protein", unit: "percentage point", value: 10, source: "Input Directly" }
    ],

    // Treatments: in this tool, each treatment is an alternative practice to compare against the control.
    // Costs are interpreted as per hectare annual baseline (control) plus incremental application costs (treatment minus control) on application years.
    treatments: [
      {
        id: uid(),
        name: "Control (baseline)",
        area: 100,
        adoption: 1,
        deltas: {}, // deltas are per ha relative to control; for control itself they should be zero
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: true,
        notes: "",
        recurrenceYears: 1 // control assumed annual
      }
    ],

    benefits: [],
    otherCosts: [],
    adoption: { base: 1.0, low: 0.6, high: 1.0 },
    risk: { base: 0.0, low: 0.0, high: 0.3, tech: 0.0, nonCoop: 0.0, socio: 0.0, fin: 0.0, man: 0.0 },

    import: {
      dataset: null,
      dictionary: null,
      meta: null,
      stats: null,
      checks: null,
      lastParsedAt: null,
      lastCommittedAt: null
    },

    ui: {
      resultsFilter: "all", // all | topNpv | topBcr | improvements
      resultsTopN: 5
    }
  };

  function ensureTreatmentDeltas() {
    model.treatments.forEach(t => {
      if (!t.deltas) t.deltas = {};
      model.outputs.forEach(o => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
      if (!("recurrenceYears" in t)) t.recurrenceYears = t.isControl ? 1 : 1;
      if (typeof t.labourCost === "undefined") t.labourCost = 0;
      if (typeof t.materialsCost === "undefined") t.materialsCost = 0;
      if (typeof t.servicesCost === "undefined") t.servicesCost = 0;
      if (typeof t.capitalCost === "undefined") t.capitalCost = 0;
      if (typeof t.adoption !== "number" || !Number.isFinite(t.adoption)) t.adoption = 1;
    });
  }
  ensureTreatmentDeltas();

  function getYieldOutputId() {
    const o = model.outputs.find(x => (x.name || "").toLowerCase().includes("yield"));
    return o ? o.id : model.outputs[0]?.id || null;
  }

  // =========================
  // 2) Parsing: TSV/CSV + Dictionary
  // =========================
  function normaliseHeader(h) {
    return (h ?? "")
      .toString()
      .trim()
      .toLowerCase()
      .replace(/\uFEFF/g, "")
      .replace(/[^\w]+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function detectDelimiter(text) {
    const head = (text || "").slice(0, 4000);
    const lines = head.split(/\r\n|\n|\r/).filter(l => l.trim().length > 0).slice(0, 20);
    const counts = { "\t": 0, ",": 0, ";": 0, "|": 0 };
    for (const line of lines) {
      counts["\t"] += (line.match(/\t/g) || []).length;
      counts[","] += (line.match(/,/g) || []).length;
      counts[";"] += (line.match(/;/g) || []).length;
      counts["|"] += (line.match(/\|/g) || []).length;
    }
    // Prefer tab if any tabs exist, otherwise the max count
    if (counts["\t"] > 0) return "\t";
    let best = ",";
    let bestN = counts[","];
    for (const k of Object.keys(counts)) {
      if (counts[k] > bestN) {
        best = k;
        bestN = counts[k];
      }
    }
    return best;
  }

  function parseDelimited(text, delimiter) {
    const d = delimiter || detectDelimiter(text);
    const rows = [];
    const raw = (text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    let i = 0;
    let field = "";
    let row = [];
    let inQuotes = false;

    function pushField() {
      row.push(field);
      field = "";
    }
    function pushRow() {
      // ignore fully empty rows
      const any = row.some(x => (x ?? "").toString().trim().length > 0);
      if (any) rows.push(row);
      row = [];
    }

    while (i < raw.length) {
      const ch = raw[i];
      if (inQuotes) {
        if (ch === '"') {
          if (raw[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i += 1;
          continue;
        }
        field += ch;
        i += 1;
        continue;
      } else {
        if (ch === '"') {
          inQuotes = true;
          i += 1;
          continue;
        }
        if (ch === d) {
          pushField();
          i += 1;
          continue;
        }
        if (ch === "\n") {
          pushField();
          pushRow();
          i += 1;
          continue;
        }
        field += ch;
        i += 1;
      }
    }
    pushField();
    pushRow();

    if (rows.length === 0) return { delimiter: d, header: [], data: [], warnings: ["No rows detected."] };

    const headerRaw = rows[0].map(h => (h ?? "").toString().trim());
    const header = headerRaw.map(h => h);

    const data = [];
    for (let r = 1; r < rows.length; r++) {
      const obj = {};
      const arr = rows[r];
      for (let c = 0; c < header.length; c++) {
        const key = header[c] || `col_${c + 1}`;
        obj[key] = c < arr.length ? arr[c] : "";
      }
      data.push(obj);
    }

    // Trim trailing completely empty objects
    const cleaned = data.filter(o => Object.values(o).some(v => (v ?? "").toString().trim().length > 0));

    return { delimiter: d, header, data: cleaned, warnings: [] };
  }

  function parseNumber(x) {
    if (x === null || x === undefined) return NaN;
    if (typeof x === "number") return Number.isFinite(x) ? x : NaN;
    const s = String(x).trim();
    if (!s || s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    const cleaned = s.replace(/[\$,]/g, "");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : NaN;
  }

  function parseDictionaryCSV(text) {
    const parsed = parseDelimited(text, ",");
    const header = parsed.header;
    const data = parsed.data;

    const normToIdx = {};
    header.forEach((h, i) => {
      normToIdx[normaliseHeader(h)] = i;
    });

    function getRowVal(rowObj, candidates) {
      for (const c of candidates) {
        const key = header.find(h => normaliseHeader(h) === c);
        if (key && rowObj[key] !== undefined) return rowObj[key];
      }
      return "";
    }

    // Flexible mapping: try common dictionary schemas
    const dict = new Map();
    for (const r of data) {
      const varName =
        (getRowVal(r, ["variable", "var", "field", "name", "column", "question", "item"]) || "").toString().trim();
      if (!varName) continue;

      const label =
        (getRowVal(r, ["label", "title", "display_name", "pretty_name", "variable_label"]) || "").toString().trim();
      const description =
        (getRowVal(r, ["description", "desc", "definition", "notes", "detail"]) || "").toString().trim();
      const unit = (getRowVal(r, ["unit", "units", "measurement_unit"]) || "").toString().trim();
      const type = (getRowVal(r, ["type", "data_type", "format"]) || "").toString().trim();
      const missing = (getRowVal(r, ["missing", "missing_values", "na_values"]) || "").toString().trim();
      const category = (getRowVal(r, ["category", "domain", "group"]) || "").toString().trim();

      dict.set(varName, {
        variable: varName,
        label: label || varName,
        description: description || "",
        unit: unit || "",
        type: type || "",
        missingValues: missing || "",
        category: category || ""
      });
    }

    return { header, rows: data, dict };
  }

  // =========================
  // 3) Validation + derived stats
  // =========================
  function inferKeyColumns(header) {
    const H = header.map(h => ({ raw: h, n: normaliseHeader(h) }));

    const findFirst = patterns => {
      for (const p of patterns) {
        const hit = H.find(x => x.n === p || x.n.includes(p));
        if (hit) return hit.raw;
      }
      return null;
    };

    const replicate = findFirst(["replicate", "rep", "block", "trial_block"]);
    const plot = findFirst(["plot_id", "plot", "plotno", "plot_number", "plotnumber"]);
    const treatment = findFirst(["treatment", "amendment", "trt", "treat", "treatment_name"]);
    const controlFlag = findFirst(["is_control", "control_flag", "control"]);
    const yieldCol = findFirst(["yield_t_ha", "yield", "grain_yield", "yield_tha", "yield_t_per_ha"]);

    // Plot area (ha) helpful for scaling
    const plotArea = findFirst(["plot_area_ha", "area_ha", "plotarea_ha", "plot_area"]);

    // Cost columns: anything that looks like cost, dollars, or per ha cost
    const costCols = H.filter(x => {
      const n = x.n;
      return (
        n.includes("cost") ||
        n.includes("aud") ||
        n.includes("dollars") ||
        n.includes("$/") ||
        n.includes("per_ha") ||
        n.includes("_ha") ||
        n.includes("labour") ||
        n.includes("labor") ||
        n.includes("fert") ||
        n.includes("herbicide") ||
        n.includes("fungicide") ||
        n.includes("insecticide") ||
        n.includes("seed") ||
        n.includes("chemical")
      );
    })
      .map(x => x.raw)
      .filter((v, i, a) => a.indexOf(v) === i);

    return { replicate, plot, treatment, controlFlag, yieldCol, plotArea, costCols };
  }

  function welford() {
    let n = 0;
    let mean = 0;
    let m2 = 0;
    return {
      push(x) {
        if (!Number.isFinite(x)) return;
        n += 1;
        const d = x - mean;
        mean += d / n;
        const d2 = x - mean;
        m2 += d * d2;
      },
      get() {
        const variance = n > 1 ? m2 / (n - 1) : NaN;
        return { n, mean, sd: n > 1 ? Math.sqrt(variance) : NaN, min: NaN, max: NaN };
      },
      getN() {
        return n;
      }
    };
  }

  function computeMinMax(arr) {
    const clean = arr.filter(v => Number.isFinite(v));
    if (!clean.length) return { min: NaN, max: NaN };
    let min = clean[0];
    let max = clean[0];
    for (let i = 1; i < clean.length; i++) {
      if (clean[i] < min) min = clean[i];
      if (clean[i] > max) max = clean[i];
    }
    return { min, max };
  }

  function determineControl(row, keys) {
    // Prefer explicit control flag
    if (keys.controlFlag && row[keys.controlFlag] !== undefined) {
      const v = (row[keys.controlFlag] ?? "").toString().trim().toLowerCase();
      if (v === "1" || v === "true" || v === "yes" || v === "y") return true;
      if (v === "0" || v === "false" || v === "no" || v === "n") return false;
    }
    // Otherwise infer from treatment label
    if (keys.treatment && row[keys.treatment] !== undefined) {
      const t = (row[keys.treatment] ?? "").toString().trim().toLowerCase();
      if (t.includes("control")) return true;
      if (t === "control") return true;
      if (t === "ctl") return true;
    }
    return false;
  }

  function costScalingRule(row, colName, keys, dictMeta) {
    // Cost scaling rule (deterministic):
    // - If the dictionary unit explicitly indicates per ha, keep as per ha.
    // - Else if the column name includes "/ha" or "per ha" or "_ha", treat as per ha.
    // - Else if plot area (ha) exists and is numeric, treat as per plot and convert to per ha by dividing by plot area.
    // - Else keep as-is (assumed per ha).
    const rawVal = row[colName];
    const v = parseNumber(rawVal);
    if (!Number.isFinite(v)) return NaN;

    const unit = (dictMeta?.unit || "").toString().toLowerCase();
    const cn = normaliseHeader(colName);
    const looksPerHa =
      unit.includes("per ha") ||
      unit.includes("/ha") ||
      cn.includes("per_ha") ||
      cn.includes("_ha") ||
      colName.toLowerCase().includes("/ha") ||
      colName.toLowerCase().includes("per ha");

    if (looksPerHa) return v;

    const areaCol = keys.plotArea;
    if (areaCol && row[areaCol] !== undefined) {
      const a = parseNumber(row[areaCol]);
      if (Number.isFinite(a) && a > 0) return v / a;
    }

    return v;
  }

  function validateAndDerive(dataset, dict) {
    const header = dataset.header || [];
    const rows = dataset.data || [];
    const keys = inferKeyColumns(header);

    const meta = {
      delimiter: dataset.delimiter,
      header,
      rowCount: rows.length,
      keys
    };

    // Basic required fields
    const issues = [];
    if (!keys.treatment) issues.push("No treatment column was detected.");
    if (!keys.yieldCol) issues.push("No yield column was detected.");

    // Build per-row derived values
    const derived = new Array(rows.length);
    const repMap = new Map(); // repId -> { controlY: stats, controlC: stats, controls: count, rows: count }
    const treatMap = new Map(); // treatment -> { yields: [], costs: [], repSet, deltas arrays... }

    // First pass: parse replicate ids and control yields/costs per replicate
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const repId = keys.replicate ? (r[keys.replicate] ?? "").toString().trim() : "1";
      const trt = keys.treatment ? (r[keys.treatment] ?? "").toString().trim() : "";
      const isControl = determineControl(r, keys);

      const y = keys.yieldCol ? parseNumber(r[keys.yieldCol]) : NaN;

      // total cost per ha from cost columns (missing-safe sum)
      let totalCostPerHa = 0;
      let anyCost = false;
      for (const c of keys.costCols || []) {
        // Use dictionary if available
        const dictMeta = dict?.dict?.get(c) || dict?.dict?.get(normaliseHeader(c)) || null;
        const cv = costScalingRule(r, c, keys, dictMeta);
        if (Number.isFinite(cv)) {
          totalCostPerHa += cv;
          anyCost = true;
        }
      }
      if (!anyCost) totalCostPerHa = NaN;

      derived[i] = {
        repId: repId || "1",
        treatment: trt,
        isControl,
        yieldTHa: y,
        totalCostPerHa,
        yieldDelta: NaN,
        costDelta: NaN,
        plotId: keys.plot ? (r[keys.plot] ?? "").toString().trim() : "",
        flags: {}
      };

      if (!repMap.has(derived[i].repId)) {
        repMap.set(derived[i].repId, {
          repId: derived[i].repId,
          rows: 0,
          controlRows: 0,
          controlYield: welford(),
          controlCost: welford(),
          controlTreatments: new Set()
        });
      }
      const rep = repMap.get(derived[i].repId);
      rep.rows += 1;

      if (isControl) {
        rep.controlRows += 1;
        rep.controlTreatments.add(trt || "Control");
        rep.controlYield.push(y);
        rep.controlCost.push(totalCostPerHa);
      }
    }

    // Replicate baseline means
    const repBaseline = new Map();
    for (const [repId, rep] of repMap.entries()) {
      const cy = rep.controlYield.get();
      const cc = rep.controlCost.get();
      repBaseline.set(repId, {
        controlYieldMean: Number.isFinite(cy.mean) ? cy.mean : NaN,
        controlCostMean: Number.isFinite(cc.mean) ? cc.mean : NaN,
        controlRows: rep.controlRows,
        controlTreatmentNames: Array.from(rep.controlTreatments)
      });
    }

    // Second pass: compute deltas and treatment summaries
    const treatSummaries = new Map();
    const globalControlYield = welford();
    const globalControlCost = welford();

    for (let i = 0; i < derived.length; i++) {
      const d = derived[i];
      const base = repBaseline.get(d.repId) || { controlYieldMean: NaN, controlCostMean: NaN, controlRows: 0 };

      // Plot-level deltas vs replicate-specific baseline
      if (Number.isFinite(d.yieldTHa) && Number.isFinite(base.controlYieldMean)) d.yieldDelta = d.yieldTHa - base.controlYieldMean;
      if (Number.isFinite(d.totalCostPerHa) && Number.isFinite(base.controlCostMean)) d.costDelta = d.totalCostPerHa - base.controlCostMean;

      if (d.isControl) {
        globalControlYield.push(d.yieldTHa);
        globalControlCost.push(d.totalCostPerHa);
      }

      const tKey = d.treatment || "(missing treatment)";
      if (!treatSummaries.has(tKey)) {
        treatSummaries.set(tKey, {
          treatment: tKey,
          repSet: new Set(),
          isControlEver: false,
          yieldAbs: welford(),
          yieldDelta: welford(),
          costAbs: welford(),
          costDelta: welford(),
          rows: 0,
          missingYield: 0,
          missingCost: 0
        });
      }
      const ts = treatSummaries.get(tKey);
      ts.rows += 1;
      ts.repSet.add(d.repId);
      if (d.isControl) ts.isControlEver = true;
      if (Number.isFinite(d.yieldTHa)) ts.yieldAbs.push(d.yieldTHa);
      else ts.missingYield += 1;
      if (Number.isFinite(d.yieldDelta)) ts.yieldDelta.push(d.yieldDelta);
      if (Number.isFinite(d.totalCostPerHa)) ts.costAbs.push(d.totalCostPerHa);
      else ts.missingCost += 1;
      if (Number.isFinite(d.costDelta)) ts.costDelta.push(d.costDelta);
    }

    // Data checks (counts and summaries)
    const checks = [];

    const missingTreatment = derived.filter(d => !d.treatment || d.treatment === "(missing treatment)").length;
    if (missingTreatment > 0) checks.push({ key: "missing_treatment", label: "Missing treatment label", count: missingTreatment, summary: "Some rows do not identify a treatment." });

    const missingYield = derived.filter(d => !Number.isFinite(d.yieldTHa)).length;
    if (missingYield > 0) checks.push({ key: "missing_yield", label: "Missing or non-numeric yield values", count: missingYield, summary: "Some rows have no usable yield value." });

    const missingRep = keys.replicate ? derived.filter(d => !d.repId || d.repId === "1").length : 0;
    if (!keys.replicate) {
      checks.push({ key: "no_replicate_col", label: "No replicate column detected", count: 1, summary: "Replicate-specific baselines use a single pooled replicate." });
    } else if (missingRep > 0) {
      checks.push({ key: "missing_replicate", label: "Missing replicate identifiers", count: missingRep, summary: "Some rows have an empty replicate value." });
    }

    // Replicate control presence checks
    let repsNoControl = 0;
    let repsMultiControlNames = 0;
    const repProblems = [];
    for (const [repId, b] of repBaseline.entries()) {
      if (!b.controlRows || b.controlRows <= 0) {
        repsNoControl += 1;
        repProblems.push(`Replicate ${repId} has no control rows.`);
      } else if ((b.controlTreatmentNames || []).length > 1) {
        repsMultiControlNames += 1;
        repProblems.push(`Replicate ${repId} has multiple control labels: ${b.controlTreatmentNames.join(", ")}.`);
      }
    }
    if (repsNoControl > 0) checks.push({ key: "rep_no_control", label: "Replicates with no control rows", count: repsNoControl, summary: repProblems.slice(0, 6).join(" ") });
    if (repsMultiControlNames > 0) checks.push({ key: "rep_multi_control", label: "Replicates with multiple control labels", count: repsMultiControlNames, summary: repProblems.slice(0, 6).join(" ") });

    const negYield = derived.filter(d => Number.isFinite(d.yieldTHa) && d.yieldTHa < 0).length;
    if (negYield > 0) checks.push({ key: "negative_yield", label: "Negative yield values", count: negYield, summary: "Some rows have negative yield values." });

    const negCost = derived.filter(d => Number.isFinite(d.totalCostPerHa) && d.totalCostPerHa < 0).length;
    if (negCost > 0) checks.push({ key: "negative_cost", label: "Negative cost values", count: negCost, summary: "Some rows have negative cost values." });

    // Yield outliers (simple z-score within replicate, if possible)
    let outliers = 0;
    for (const [repId, rep] of repMap.entries()) {
      const ys = derived.filter(d => d.repId === repId).map(d => d.yieldTHa).filter(Number.isFinite);
      if (ys.length < 6) continue;
      const mean = ys.reduce((a, b) => a + b, 0) / ys.length;
      const sd = Math.sqrt(ys.reduce((a, b) => a + (b - mean) * (b - mean), 0) / (ys.length - 1));
      if (!Number.isFinite(sd) || sd === 0) continue;
      for (const d of derived) {
        if (d.repId !== repId) continue;
        if (!Number.isFinite(d.yieldTHa)) continue;
        const z = (d.yieldTHa - mean) / sd;
        if (Math.abs(z) > 3) outliers += 1;
      }
    }
    if (outliers > 0) checks.push({ key: "yield_outliers", label: "Potential yield outliers", count: outliers, summary: "Some yields are more than three standard deviations from the replicate mean." });

    // Global control mean (fallback)
    const globalControlYieldStats = globalControlYield.get();
    const globalControlCostStats = globalControlCost.get();

    const stats = {
      repBaseline,
      derived,
      treatSummaries,
      globalControl: {
        meanYieldTHa: Number.isFinite(globalControlYieldStats.mean) ? globalControlYieldStats.mean : NaN,
        meanCostPerHa: Number.isFinite(globalControlCostStats.mean) ? globalControlCostStats.mean : NaN
      }
    };

    return { meta, issues, stats, checks };
  }

  // =========================
  // 4) Commit: build treatments + deltas from imported dataset
  // =========================
  function commitImportedToModel() {
    const ds = model.import.dataset;
    const dict = model.import.dictionary;
    if (!ds || !ds.header || !ds.data) {
      showToast("No parsed dataset to commit.");
      return;
    }

    const vd = validateAndDerive(ds, dict);
    model.import.meta = vd.meta;
    model.import.stats = vd.stats;
    model.import.checks = vd.checks;

    // Require treatment and yield
    if (vd.issues && vd.issues.length) {
      renderDataChecks();
      showToast("Dataset has blocking issues. Please review Data Checks.");
      return;
    }

    // Identify control label and control means
    const treatSummaries = vd.stats.treatSummaries;
    const controlCandidates = Array.from(treatSummaries.values()).filter(t => t.isControlEver);
    let controlName = null;
    if (controlCandidates.length) {
      // Prefer a label that explicitly contains "control"
      const explicit = controlCandidates.find(x => (x.treatment || "").toLowerCase().includes("control"));
      controlName = (explicit || controlCandidates[0]).treatment;
    } else {
      // Fallback: any treatment containing "control"
      const any = Array.from(treatSummaries.keys()).find(k => (k || "").toLowerCase().includes("control"));
      controlName = any || null;
    }

    const yieldId = getYieldOutputId();
    if (!yieldId) {
      showToast("No yield output is defined. Please add an output named Grain yield.");
      return;
    }

    // Ensure output value matches base price (keep them aligned)
    const yieldOutput = model.outputs.find(o => o.id === yieldId);
    if (yieldOutput) yieldOutput.value = Number(model.config.grainPricePerTonne) || yieldOutput.value || 0;

    // Build treatments list from summaries
    const allTreatments = Array.from(treatSummaries.values())
      .filter(t => t.treatment && t.treatment !== "(missing treatment)")
      .sort((a, b) => a.treatment.localeCompare(b.treatment));

    // Control stats
    let controlMeanYield = vd.stats.globalControl.meanYieldTHa;
    let controlMeanCost = vd.stats.globalControl.meanCostPerHa;

    // If global control is missing, attempt from controlName summary
    if ((!Number.isFinite(controlMeanYield) || !Number.isFinite(controlMeanCost)) && controlName) {
      const cs = treatSummaries.get(controlName);
      if (cs) {
        const y = cs.yieldAbs.get();
        const c = cs.costAbs.get();
        if (!Number.isFinite(controlMeanYield)) controlMeanYield = y.mean;
        if (!Number.isFinite(controlMeanCost)) controlMeanCost = c.mean;
      }
    }

    if (!Number.isFinite(controlMeanYield)) controlMeanYield = 0;
    if (!Number.isFinite(controlMeanCost)) controlMeanCost = 0;

    // Rebuild model.treatments
    const newTreatments = [];

    for (const ts of allTreatments) {
      const yAbs = ts.yieldAbs.get();
      const yDel = ts.yieldDelta.get();
      const cAbs = ts.costAbs.get();
      const cDel = ts.costDelta.get();

      const meanYieldAbs = Number.isFinite(yAbs.mean) ? yAbs.mean : NaN;
      const meanCostAbs = Number.isFinite(cAbs.mean) ? cAbs.mean : NaN;

      const inferredIsControl =
        controlName ? ts.treatment === controlName : (ts.treatment || "").toLowerCase().includes("control");

      // Prefer replicate-delta mean, else absolute minus global control
      const meanYieldDelta = Number.isFinite(yDel.mean)
        ? yDel.mean
        : Number.isFinite(meanYieldAbs)
          ? meanYieldAbs - controlMeanYield
          : 0;

      const meanCostDelta = Number.isFinite(cDel.mean)
        ? cDel.mean
        : Number.isFinite(meanCostAbs)
          ? meanCostAbs - controlMeanCost
          : 0;

      // Store absolute baseline costs on control; for treatments store absolute as control + delta if absolute missing
      const absCostPerHa =
        inferredIsControl
          ? (Number.isFinite(meanCostAbs) ? meanCostAbs : controlMeanCost)
          : (Number.isFinite(meanCostAbs) ? meanCostAbs : controlMeanCost + meanCostDelta);

      // For the editable model, we store absolute baseline annual costs for the practice in the three buckets.
      // If the imported file does not provide category splits, we store total cost in Materials by default.
      const labour = 0;
      const services = 0;
      const materials = Number.isFinite(absCostPerHa) ? absCostPerHa : 0;

      const t = {
        id: uid(),
        name: ts.treatment,
        area: Number(model.config.farmAreaHa) || 0,
        adoption: 1,
        deltas: {},
        labourCost: labour,
        materialsCost: materials,
        servicesCost: services,
        capitalCost: 0,
        constrained: true,
        source: "Farm Trials",
        isControl: !!inferredIsControl,
        notes: "",
        recurrenceYears: inferredIsControl ? 1 : 1,
        meta: {
          nPlots: ts.rows,
          nReplicates: ts.repSet.size,
          meanYieldTHa: Number.isFinite(meanYieldAbs) ? meanYieldAbs : controlMeanYield + meanYieldDelta,
          meanYieldDeltaTHa: meanYieldDelta,
          meanCostPerHa: absCostPerHa,
          meanCostDeltaPerHa: meanCostDelta,
          sdYieldTHa: yAbs.sd,
          sdYieldDeltaTHa: yDel.sd,
          sdCostPerHa: cAbs.sd,
          sdCostDeltaPerHa: cDel.sd,
          missingYield: ts.missingYield,
          missingCost: ts.missingCost
        }
      };

      // Initialise deltas for outputs
      model.outputs.forEach(o => {
        t.deltas[o.id] = 0;
      });
      // Set yield delta in the yield output
      t.deltas[yieldId] = inferredIsControl ? 0 : (Number.isFinite(meanYieldDelta) ? meanYieldDelta : 0);

      newTreatments.push(t);
    }

    // Ensure exactly one control is set
    let controlCount = newTreatments.filter(x => x.isControl).length;
    if (controlCount === 0) {
      if (newTreatments.length) {
        newTreatments[0].isControl = true;
        showToast("No control was detected. The first treatment was set as the control.");
      }
    } else if (controlCount > 1) {
      // Keep the first explicit control
      const idx = newTreatments.findIndex(x => (x.name || "").toLowerCase().includes("control"));
      newTreatments.forEach((x, i) => (x.isControl = i === (idx >= 0 ? idx : 0)));
      showToast("Multiple controls were detected. A single control was retained.");
    }

    model.treatments = newTreatments;
    ensureTreatmentDeltas();

    // Align analysis settings with time settings
    model.time.startYear = model.config.startYear;
    model.time.years = model.config.horizonYears;
    model.time.discBase = model.config.discountRatePct;
    model.outputs.find(o => o.id === yieldId).value = model.config.grainPricePerTonne;

    model.import.lastCommittedAt = nowISO();

    // Render updates
    renderTreatments();
    renderRecurrenceConfig();
    renderDataChecks();
    calcAndRenderAllOutputs();
    renderResults();
    renderAiBriefingPreview();

    showToast("Dataset committed. Treatments, baselines, deltas, and results were updated.");
  }

  // =========================
  // 5) Results engine: discounted CBA per treatment vs control
  // =========================
  function annuityFactor(N, rPct) {
    const r = rPct / 100;
    return r === 0 ? N : (1 - Math.pow(1 + r, -N)) / r;
  }

  function presentValue(series, ratePct) {
    let pv = 0;
    for (let t = 0; t < series.length; t++) pv += series[t] / Math.pow(1 + ratePct / 100, t);
    return pv;
  }

  function getControlTreatment() {
    let c = model.treatments.find(t => t.isControl);
    if (!c && model.treatments.length) {
      model.treatments[0].isControl = true;
      c = model.treatments[0];
      showToast("A control was not set. The first treatment is now the control.");
    }
    return c || null;
  }

  function getPerHaAnnualCost(t) {
    return (Number(t.materialsCost) || 0) + (Number(t.servicesCost) || 0) + (Number(t.labourCost) || 0);
  }

  function getRecurrenceYears(t, mode) {
    if (mode === "annual") return 1;
    if (mode === "once") return Number(model.config.horizonYears) + 999; // effectively once
    const r = Math.max(1, Math.round(Number(t.recurrenceYears) || 1));
    return r;
  }

  function buildApplicationYears(horizonYears, recurrenceYears) {
    // Application in year 1, then every recurrenceYears (year indices)
    const yrs = [];
    for (let y = 1; y <= horizonYears; y += recurrenceYears) yrs.push(y);
    return yrs;
  }

  function computeTreatmentSeries(t, scenario, control) {
    const N = Math.max(1, Math.round(Number(scenario.horizonYears)));
    const discount = Number(scenario.discountRatePct) || 0;
    const price = Number(scenario.grainPricePerTonne) || 0;
    const persistence = clamp(Number(scenario.persistence), 0, 1);
    const adoptionMul = clamp(Number(scenario.adoptionMultiplier), 0, 1);
    const riskMul = clamp(Number(scenario.riskMultiplier), 0, 1);
    const recurrenceMode = scenario.recurrenceMode || "configured";

    const areaTotal = Number(scenario.farmAreaHa) || Number(t.area) || 0;
    const adoptedArea = areaTotal * adoptionMul;

    // Use yield delta and cost delta relative to control
    const yieldId = getYieldOutputId();
    const deltaYield = yieldId ? Number(t.deltas[yieldId]) || 0 : 0;

    const controlCostPerHa = getPerHaAnnualCost(control);
    const treatCostPerHa = getPerHaAnnualCost(t);
    const deltaCostPerHa = treatCostPerHa - controlCostPerHa;

    // Control yield level: from imported stats if available, else infer as zero baseline.
    const controlYieldTHa =
      Number(model.import?.stats?.globalControl?.meanYieldTHa) ||
      Number(control?.meta?.meanYieldTHa) ||
      0;

    const recurrenceYears = t.isControl ? 1 : getRecurrenceYears(t, recurrenceMode);
    const applicationYears = t.isControl ? buildApplicationYears(N, 1) : buildApplicationYears(N, recurrenceYears);

    // Additional outputs: monetised deltas (non-yield)
    const otherDeltaValuePerHa = (() => {
      let v = 0;
      for (const o of model.outputs) {
        if (o.id === yieldId) continue;
        const d = Number(t.deltas[o.id]) || 0;
        const val = Number(o.value) || 0;
        v += d * val;
      }
      return v;
    })();

    // Series are indexed 0..N
    const benefits = new Array(N + 1).fill(0);
    const costs = new Array(N + 1).fill(0);
    const cashflow = new Array(N + 1).fill(0);

    // Year 0 capital: incremental capital (treatment minus control); always on adopted area
    const capDelta = (Number(t.capitalCost) || 0) - (Number(control.capitalCost) || 0);
    if (!t.isControl && capDelta !== 0) costs[0] += capDelta;

    // Baseline control annual revenue and baseline control annual cost apply to full area
    for (let year = 1; year <= N; year++) {
      // Control revenue on full area
      const baseRevenue = controlYieldTHa * price * areaTotal;

      // Treatment effect on adopted area: decays between applications and resets on application year
      let lastApp = 1;
      for (let k = 0; k < applicationYears.length; k++) {
        if (applicationYears[k] <= year) lastApp = applicationYears[k];
        else break;
      }
      const yearsSince = Math.max(0, year - lastApp);
      const effectiveDeltaYield = t.isControl ? 0 : deltaYield * Math.pow(persistence, yearsSince);
      const deltaRevenue = effectiveDeltaYield * price * adoptedArea;

      // Additional monetised deltas on adopted area
      const otherDeltaRevenue = otherDeltaValuePerHa * adoptedArea;

      // Global benefits (if used) are treated as project-wide and apply to all alternatives equally, so they are added to all.
      // They can be edited in the Benefits tab if present in the interface.
      const extraGlobalBenefits = 0;

      benefits[year] = baseRevenue + deltaRevenue + otherDeltaRevenue + extraGlobalBenefits;

      // Costs: baseline control annual cost on full area
      const baseCost = controlCostPerHa * areaTotal;

      // Incremental application cost on adopted area only in application years
      const isAppYear = applicationYears.includes(year);
      const appCost = (!t.isControl && isAppYear) ? deltaCostPerHa * adoptedArea : 0;

      // Global other costs (if present) apply equally to all alternatives
      const extraGlobalCosts = 0;

      costs[year] = baseCost + appCost + extraGlobalCosts;

      cashflow[year] = benefits[year] - costs[year];
    }
    cashflow[0] = benefits[0] - costs[0];

    const pvBenefits = presentValue(benefits, discount);
    const pvCosts = presentValue(costs, discount);
    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? (npv / pvCosts) * 100 : NaN;

    return {
      treatmentId: t.id,
      treatmentName: t.name,
      isControl: !!t.isControl,
      horizonYears: N,
      discountRatePct: discount,
      pricePerTonne: price,
      persistence,
      recurrenceMode,
      recurrenceYears,
      adoptionMultiplier: adoptionMul,
      riskMultiplier: riskMul,
      areaHa: areaTotal,
      adoptedAreaHa: adoptedArea,
      controlYieldTHa,
      deltaYieldTHa: deltaYield,
      controlCostPerHa,
      treatCostPerHa,
      deltaCostPerHa,
      applicationYears,
      benefits,
      costs,
      cashflow,
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roi
    };
  }

  function computeBaseCaseScenario() {
    return {
      farmAreaHa: Number(model.config.farmAreaHa) || 0,
      startYear: Number(model.config.startYear) || new Date().getFullYear(),
      horizonYears: Math.max(1, Math.round(Number(model.config.horizonYears) || 10)),
      grainPricePerTonne: Number(model.config.grainPricePerTonne) || 0,
      discountRatePct: Number(model.config.discountRatePct) || 0,
      persistence: clamp(Number(model.config.persistence), 0, 1),
      recurrenceMode: model.config.recurrenceMode || "configured",
      adoptionMultiplier: clamp(Number(model.config.adoptionMultiplier), 0, 1),
      riskMultiplier: clamp(Number(model.config.riskMultiplier), 0, 1)
    };
  }

  function computeAllTreatmentResults(scenario) {
    const control = getControlTreatment();
    if (!control) return { control: null, rows: [] };

    const results = model.treatments.map(t => computeTreatmentSeries(t, scenario, control));

    // Rank by NPV descending, excluding control for ranking
    const ranked = results
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const rankMap = new Map();
    ranked.forEach((r, idx) => rankMap.set(r.treatmentId, idx + 1));
    results.forEach(r => {
      r.rank = r.isControl ? null : (rankMap.get(r.treatmentId) || null);
    });

    // Identify best by NPV and best by BCR
    const bestNpv = ranked.length ? ranked[0] : null;
    const bestBcr = results
      .filter(r => !r.isControl && Number.isFinite(r.bcr))
      .slice()
      .sort((a, b) => b.bcr - a.bcr)[0] || null;

    return { control: results.find(r => r.isControl) || null, rows: results, bestNpv, bestBcr };
  }

  // =========================
  // 6) Sensitivity grid
  // =========================
  function buildSensitivityGridLong() {
    const s = model.sensitivity;
    const prices = [Number(s.priceLow), Number(s.priceBase), Number(s.priceHigh)].filter(Number.isFinite);
    const discs = [Number(s.discLow), Number(s.discBase), Number(s.discHigh)].filter(Number.isFinite);
    const pers = [Number(s.persistenceLow), Number(s.persistenceBase), Number(s.persistenceHigh)].filter(Number.isFinite);
    const recModes = Array.isArray(s.recurrenceModes) && s.recurrenceModes.length ? s.recurrenceModes : ["configured", "annual", "once"];

    const base = computeBaseCaseScenario();
    const out = [];

    let scenarioId = 0;
    for (const p of prices) {
      for (const d of discs) {
        for (const pe of pers) {
          for (const rm of recModes) {
            scenarioId += 1;
            const sc = {
              ...base,
              grainPricePerTonne: p,
              discountRatePct: d,
              persistence: clamp(pe, 0, 1),
              recurrenceMode: rm
            };
            const res = computeAllTreatmentResults(sc);
            for (const r of res.rows) {
              out.push({
                scenario_id: scenarioId,
                grain_price_per_tonne: p,
                discount_rate_pct: d,
                persistence: clamp(pe, 0, 1),
                recurrence_mode: rm,
                treatment: r.treatmentName,
                is_control: r.isControl ? 1 : 0,
                rank_by_npv: r.rank ?? "",
                pv_benefits: r.pvBenefits,
                pv_costs: r.pvCosts,
                net_present_value: r.npv,
                benefit_cost_ratio: r.bcr,
                return_on_investment_percent: r.roi
              });
            }
          }
        }
      }
    }
    return out;
  }

  function exportSensitivityCsv() {
    const grid = buildSensitivityGridLong();
    if (!grid.length) {
      showToast("Sensitivity grid is empty.");
      return;
    }
    const header = Object.keys(grid[0]);
    const csv = [
      header.join(","),
      ...grid.map(r => header.map(k => (r[k] === null || r[k] === undefined ? "" : String(r[k]).replace(/"/g, '""'))).join(","))
    ].join("\r\n");
    downloadFile(`${slug(model.project.name)}_sensitivity_grid.csv`, csv, "text/csv");
    showToast("Sensitivity grid CSV downloaded.");
  }

  // =========================
  // 7) Cleaned dataset export
  // =========================
  function buildCleanedDatasetTsv() {
    const ds = model.import.dataset;
    const stats = model.import.stats;
    if (!ds || !stats || !stats.derived) return "";
    const keys = model.import.meta?.keys || inferKeyColumns(ds.header || []);
    const derived = stats.derived;

    // Standardised columns + include original treatment and replicate labels
    const header = [
      "plot_id",
      "replicate_id",
      "treatment",
      "is_control",
      "yield_t_per_ha",
      "yield_delta_vs_control_t_per_ha",
      "total_cost_per_ha",
      "cost_delta_vs_control_per_ha"
    ];

    const lines = [header.join("\t")];

    for (const d of derived) {
      lines.push([
        d.plotId || "",
        d.repId || "",
        d.treatment || "",
        d.isControl ? "1" : "0",
        Number.isFinite(d.yieldTHa) ? d.yieldTHa : "",
        Number.isFinite(d.yieldDelta) ? d.yieldDelta : "",
        Number.isFinite(d.totalCostPerHa) ? d.totalCostPerHa : "",
        Number.isFinite(d.costDelta) ? d.costDelta : ""
      ].join("\t"));
    }
    return lines.join("\n");
  }

  function exportCleanedDatasetTsv() {
    const tsv = buildCleanedDatasetTsv();
    if (!tsv) {
      showToast("No cleaned dataset is available to export.");
      return;
    }
    downloadFile(`${slug(model.project.name)}_cleaned_dataset.tsv`, tsv, "text/tab-separated-values");
    showToast("Cleaned dataset TSV downloaded.");
  }

  // =========================
  // 8) Treatment summary export
  // =========================
  function exportTreatmentSummaryCsv() {
    const scenario = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(scenario);
    if (!res.rows.length) {
      showToast("No treatment results to export.");
      return;
    }

    const control = res.control;
    const rows = res.rows.map(r => {
      const t = model.treatments.find(x => x.id === r.treatmentId);
      const meta = t && t.meta ? t.meta : {};
      const deltaNpv = control ? r.npv - control.npv : NaN;
      const deltaPvCost = control ? r.pvCosts - control.pvCosts : NaN;

      return {
        treatment: r.treatmentName,
        is_control: r.isControl ? 1 : 0,
        rank_by_npv: r.rank ?? "",
        n_plots: meta.nPlots ?? "",
        n_replicates: meta.nReplicates ?? "",
        mean_yield_t_per_ha: meta.meanYieldTHa ?? "",
        mean_yield_delta_t_per_ha: meta.meanYieldDeltaTHa ?? "",
        mean_cost_per_ha: meta.meanCostPerHa ?? "",
        mean_cost_delta_per_ha: meta.meanCostDeltaPerHa ?? "",
        pv_benefits: r.pvBenefits,
        pv_costs: r.pvCosts,
        net_present_value: r.npv,
        benefit_cost_ratio: r.bcr,
        return_on_investment_percent: r.roi,
        delta_net_present_value_vs_control: deltaNpv,
        delta_present_value_costs_vs_control: deltaPvCost
      };
    });

    const header = Object.keys(rows[0]);
    const csv = [
      header.join(","),
      ...rows.map(o => header.map(k => (o[k] === null || o[k] === undefined ? "" : String(o[k]).replace(/"/g, '""'))).join(","))
    ].join("\r\n");

    downloadFile(`${slug(model.project.name)}_treatment_summary.csv`, csv, "text/csv");
    showToast("Treatment summary CSV downloaded.");
  }

  // =========================
  // 9) Workbook export (optional)
  // =========================
  function exportWorkbookIfPossible() {
    if (typeof XLSX === "undefined") {
      showToast("Workbook export requires the XLSX library.");
      return;
    }

    const cleanedTsv = buildCleanedDatasetTsv();
    const scenario = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(scenario);
    const sens = buildSensitivityGridLong();

    const wb = XLSX.utils.book_new();

    // Cleaned dataset sheet
    if (cleanedTsv) {
      const lines = cleanedTsv.split("\n").map(l => l.split("\t"));
      const sh = XLSX.utils.aoa_to_sheet(lines);
      XLSX.utils.book_append_sheet(wb, sh, "CleanedData");
    }

    // Treatment results sheet
    if (res.rows.length) {
      const table = [
        [
          "Treatment",
          "Is control",
          "Rank by net present value",
          "Present value of benefits",
          "Present value of costs",
          "Net present value",
          "Benefit cost ratio",
          "Return on investment percent"
        ],
        ...res.rows.map(r => [
          r.treatmentName,
          r.isControl ? 1 : 0,
          r.rank ?? "",
          r.pvBenefits,
          r.pvCosts,
          r.npv,
          r.bcr,
          r.roi
        ])
      ];
      const sh = XLSX.utils.aoa_to_sheet(table);
      XLSX.utils.book_append_sheet(wb, sh, "Results");
    }

    // Sensitivity sheet
    if (sens.length) {
      const header = Object.keys(sens[0]);
      const table = [header, ...sens.map(o => header.map(k => o[k]))];
      const sh = XLSX.utils.aoa_to_sheet(table);
      XLSX.utils.book_append_sheet(wb, sh, "Sensitivity");
    }

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${slug(model.project.name)}_workbook.xlsx`, out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    showToast("Workbook downloaded.");
  }

  // =========================
  // 10) Rendering: Data checks + recurrence config + results + AI briefing
  // =========================
  function renderDataChecks() {
    const root = $id("dataChecksList") || $id("dataChecks");
    if (!root) return;

    const checks = model.import.checks || [];
    const issues = (model.import.meta && model.import.dataset && model.import.stats) ? [] : [];
    const blocking = [];
    if (model.import.dataset && model.import.meta) {
      const vd = validateAndDerive(model.import.dataset, model.import.dictionary);
      if (vd.issues && vd.issues.length) blocking.push(...vd.issues);
    }

    root.innerHTML = "";

    if (blocking.length) {
      const box = document.createElement("div");
      box.className = "item";
      box.innerHTML = `<h4>Blocking issues</h4><div class="small muted">${esc(blocking.join(" "))}</div>`;
      root.appendChild(box);
    }

    if (!checks.length && !blocking.length) {
      const p = document.createElement("p");
      p.className = "small muted";
      p.textContent = "No data checks are available yet. Parse and commit a dataset to populate checks.";
      root.appendChild(p);
      return;
    }

    for (const c of checks) {
      const el = document.createElement("div");
      el.className = "item";
      el.innerHTML = `
        <div class="row-2">
          <div class="field">
            <label>${esc(c.label)}</label>
            <div class="metric"><div class="value">${fmt(c.count)}</div></div>
          </div>
          <div class="field">
            <label>Summary</label>
            <div class="small muted">${esc(c.summary || "")}</div>
          </div>
        </div>
      `;
      root.appendChild(el);
    }
  }

  function renderRecurrenceConfig() {
    const root = $id("recurrenceConfig") || $id("recurrenceTable") || $id("configRecurrence");
    if (!root) return;

    const rows = model.treatments.slice().sort((a, b) => (a.isControl ? -1 : 1) - (b.isControl ? -1 : 1) || a.name.localeCompare(b.name));
    root.innerHTML = "";

    // Build a simple editable table
    const tbl = document.createElement("table");
    tbl.className = "summary-table recurrence-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Treatment</th>
          <th>Control</th>
          <th>Application frequency in years</th>
          <th>Notes</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            t => `
          <tr>
            <td>${esc(t.name)}</td>
            <td>${t.isControl ? "Yes" : "No"}</td>
            <td>
              <input type="number" min="1" step="1" value="${Math.max(1, Math.round(Number(t.recurrenceYears) || 1))}" data-rec-years="${t.id}" ${t.isControl ? "readonly" : ""} />
            </td>
            <td class="small muted">${esc(t.notes || "")}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    `;
    root.appendChild(tbl);

    // Bind
    tbl.addEventListener("input", e => {
      const id = e.target && e.target.dataset ? e.target.dataset.recYears : null;
      if (!id) return;
      const t = model.treatments.find(x => x.id === id);
      if (!t || t.isControl) return;
      const v = Math.max(1, Math.round(Number(e.target.value) || 1));
      t.recurrenceYears = v;
      renderResults();
      renderAiBriefingPreview();
      showToast("Recurrence settings updated.");
    });
  }

  function renderLeaderboard(res) {
    const root = $id("resultsLeaderboard") || $id("leaderboard");
    if (!root) return;
    root.innerHTML = "";

    const rows = res.rows
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const topN = Math.max(1, Math.round(Number(model.ui.resultsTopN) || 5));
    const filtered = model.ui.resultsFilter === "topNpv" ? rows.slice(0, topN)
      : model.ui.resultsFilter === "topBcr"
        ? rows.slice().sort((a, b) => (Number.isFinite(b.bcr) ? b.bcr : -Infinity) - (Number.isFinite(a.bcr) ? a.bcr : -Infinity)).slice(0, topN)
        : model.ui.resultsFilter === "improvements"
          ? rows.filter(r => Number.isFinite(r.npv) && res.control && Number.isFinite(res.control.npv) ? r.npv > res.control.npv : r.npv > 0)
          : rows;

    const tbl = document.createElement("table");
    tbl.className = "summary-table leaderboard-table";
    tbl.innerHTML = `
      <thead>
        <tr>
          <th>Rank</th>
          <th>Treatment</th>
          <th>Net present value</th>
          <th>Benefit cost ratio</th>
          <th>Present value of benefits</th>
          <th>Present value of costs</th>
        </tr>
      </thead>
      <tbody>
        ${
          filtered.length
            ? filtered
                .map(
                  r => `
          <tr>
            <td>${r.rank ?? ""}</td>
            <td>${esc(r.treatmentName)}</td>
            <td class="${Number.isFinite(r.npv) && r.npv >= 0 ? "positive" : "negative"}">${money(r.npv)}</td>
            <td>${Number.isFinite(r.bcr) ? ratio(r.bcr) : "n/a"}</td>
            <td>${money(r.pvBenefits)}</td>
            <td>${money(r.pvCosts)}</td>
          </tr>
        `
                )
                .join("")
            : `<tr><td colspan="6" class="small muted">No treatments are available to rank.</td></tr>`
        }
      </tbody>
    `;
    root.appendChild(tbl);
  }

  function colourClassForDelta(indicatorKey, deltaValue) {
    if (!Number.isFinite(deltaValue)) return "";
    // For PV costs, lower is better, so improvement is negative delta
    if (indicatorKey === "pvCosts") return deltaValue <= 0 ? "positive" : "negative";
    // For all others, higher is better
    return deltaValue >= 0 ? "positive" : "negative";
  }

  function renderComparisonGrid(res) {
    const root = $id("comparisonGrid") || $id("comparisonToControl") || $id("resultsComparison");
    if (!root) return;
    root.innerHTML = "";

    const control = res.control;
    if (!control) {
      root.innerHTML = `<div class="small muted">No control is set. Please set a control in the Treatments tab.</div>`;
      return;
    }

    const treatments = res.rows
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (a.rank ?? 1e9) - (b.rank ?? 1e9));

    const topN = Math.max(1, Math.round(Number(model.ui.resultsTopN) || 5));
    let viewTreatments = treatments;

    if (model.ui.resultsFilter === "topNpv") viewTreatments = treatments.slice(0, topN);
    else if (model.ui.resultsFilter === "topBcr") {
      viewTreatments = treatments
        .slice()
        .sort((a, b) => (Number.isFinite(b.bcr) ? b.bcr : -Infinity) - (Number.isFinite(a.bcr) ? a.bcr : -Infinity))
        .slice(0, topN);
    } else if (model.ui.resultsFilter === "improvements") {
      viewTreatments = treatments.filter(r => Number.isFinite(r.npv) && Number.isFinite(control.npv) ? r.npv > control.npv : r.npv > 0);
    }

    const indicators = [
      { key: "pvBenefits", label: "Present value of benefits" },
      { key: "pvCosts", label: "Present value of costs" },
      { key: "npv", label: "Net present value" },
      { key: "bcr", label: "Benefit cost ratio" },
      { key: "roi", label: "Return on investment" },
      { key: "rank", label: "Rank" },
      { key: "deltaNpv", label: "Change in net present value versus control" },
      { key: "deltaPvCost", label: "Change in present value of costs versus control" }
    ];

    function getVal(r, key) {
      if (key === "rank") return r.isControl ? "" : (r.rank ?? "");
      if (key === "roi") return r.roi;
      if (key === "deltaNpv") return r.npv - control.npv;
      if (key === "deltaPvCost") return r.pvCosts - control.pvCosts;
      return r[key];
    }

    function formatVal(key, v) {
      if (key === "bcr") return Number.isFinite(v) ? ratio(v) : "n/a";
      if (key === "roi") return Number.isFinite(v) ? pct(v) : "n/a";
      if (key === "rank") return v === "" ? "" : String(v);
      if (key === "pvBenefits" || key === "pvCosts" || key === "npv" || key === "deltaNpv" || key === "deltaPvCost") return money(v);
      return Number.isFinite(v) ? fmt(v) : "n/a";
    }

    function formatDeltaCell(indKey, vAbs, vCtrl) {
      // absolute + percent where meaningful (control non-zero and indicator is money-like)
      const absTxt = formatVal(indKey, vAbs);
      if (!Number.isFinite(vAbs) || !Number.isFinite(vCtrl) || vCtrl === 0) return absTxt;
      const pctVal = (vAbs / vCtrl) * 100;
      if (indKey === "bcr" || indKey === "roi" || indKey === "rank") return absTxt;
      return `${absTxt}\n(${fmt(pctVal)} per cent)`;
    }

    // Build table with sticky first column intention handled by CSS; JS produces semantic structure.
    const tbl = document.createElement("table");
    tbl.className = "summary-table comparison-grid";

    const headCells = [];
    headCells.push(`<th class="sticky-col">Indicator</th>`);
    headCells.push(`<th class="sticky-head control-col">Control (baseline)</th>`);

    for (const t of viewTreatments) {
      headCells.push(`<th class="sticky-head">${esc(t.treatmentName)}</th>`);
      headCells.push(`<th class="sticky-head delta-col">Δ versus control</th>`);
    }

    const bodyRows = indicators
      .map(ind => {
        const controlVal = getVal(control, ind.key);
        const ctrlTxt = formatVal(ind.key, controlVal);

        const cells = [];
        cells.push(`<td class="sticky-col">${esc(ind.label)}</td>`);
        cells.push(`<td class="control-col">${ctrlTxt}</td>`);

        for (const t of viewTreatments) {
          const v = getVal(t, ind.key);
          const delta = ind.key === "deltaNpv" || ind.key === "deltaPvCost"
            ? v
            : (Number.isFinite(v) && Number.isFinite(controlVal) ? v - controlVal : NaN);

          const tTxt = formatVal(ind.key, v);
          const deltaTxt = formatDeltaCell(ind.key, delta, controlVal);

          const deltaClass = colourClassForDelta(ind.key === "pvCosts" ? "pvCosts" : ind.key, delta);
          const deltaCell = `<td class="${deltaClass} delta-col" style="white-space:pre-line">${esc(deltaTxt)}</td>`;
          cells.push(`<td>${esc(tTxt)}</td>`);
          cells.push(deltaCell);
        }

        return `<tr>${cells.join("")}</tr>`;
      })
      .join("");

    tbl.innerHTML = `
      <thead><tr>${headCells.join("")}</tr></thead>
      <tbody>${bodyRows}</tbody>
    `;

    root.appendChild(tbl);
  }

  function renderResultsNarrative(res) {
    const root = $id("resultsNarrative") || $id("whatThisMeans") || $id("resultsWhatThisMeans");
    if (!root) return;

    const control = res.control;
    if (!control) {
      root.textContent = "A control is not set. Please select a control in the Treatments tab.";
      return;
    }

    const bestNpv = res.bestNpv;
    const bestBcr = res.bestBcr;

    const s = computeBaseCaseScenario();
    const adoptionPct = (s.adoptionMultiplier * 100);
    const riskPct = (s.riskMultiplier * 100);

    const paragraphs = [];

    paragraphs.push(
      `This results view compares each alternative practice against the control baseline over ${s.horizonYears} years on ${fmt(s.farmAreaHa)} hectares. It uses a grain price of ${money(s.grainPricePerTonne)} per tonne and a discount rate of ${fmt(s.discountRatePct)} per cent. The yield effect decays between applications according to the persistence setting of ${fmt(s.persistence)}.`
    );

    paragraphs.push(
      `Adoption is applied as a coverage share of the area for the alternative practice, set to ${fmt(adoptionPct)} per cent in the base case. Risk is applied as a reduction to the yield effect, set to ${fmt(riskPct)} per cent in the base case.`
    );

    if (bestNpv) {
      paragraphs.push(
        `On net present value, the strongest performer is ${bestNpv.treatmentName}. Its net present value is ${money(bestNpv.npv)}, compared with ${money(control.npv)} for the control. The change in present value of costs versus control is ${money(bestNpv.pvCosts - control.pvCosts)}.`
      );
    }

    if (bestBcr) {
      paragraphs.push(
        `On benefit cost ratio, the strongest performer is ${bestBcr.treatmentName}. Its benefit cost ratio is ${ratio(bestBcr.bcr)}, and its net present value is ${money(bestBcr.npv)}.`
      );
    }

    paragraphs.push(
      `When a treatment improves net present value, this can happen by lifting present value of benefits, lowering present value of costs, or both. The delta columns show these changes directly versus the control. For present value of costs, a negative change is an improvement because it means lower costs.`
    );

    root.textContent = paragraphs.join("\n\n");
  }

  function renderResultsFilters() {
    const root = $id("resultsFilters") || $id("comparisonFilters");
    if (!root) return;

    // Expect buttons with IDs if present; otherwise render nothing.
    const btnTopNpv = $id("filterTopNpv");
    const btnTopBcr = $id("filterTopBcr");
    const btnImprove = $id("filterImprovements");
    const btnAll = $id("filterAll");

    const set = f => {
      model.ui.resultsFilter = f;
      renderResults();
      renderAiBriefingPreview();
      showToast("Results filter updated.");
    };

    safeOn(btnTopNpv, "click", e => { e.preventDefault(); set("topNpv"); });
    safeOn(btnTopBcr, "click", e => { e.preventDefault(); set("topBcr"); });
    safeOn(btnImprove, "click", e => { e.preventDefault(); set("improvements"); });
    safeOn(btnAll, "click", e => { e.preventDefault(); set("all"); });
  }

  function renderResults() {
    const scenario = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(scenario);
    renderLeaderboard(res);
    renderComparisonGrid(res);
    renderResultsNarrative(res);
  }

  // =========================
  // 11) AI briefing
  // =========================
  function buildResultsJson() {
    const s = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(s);

    const compact = {
      tool_name: "Farming CBA Decision Tool",
      generated_at_iso: nowISO(),
      project: {
        name: model.project.name,
        organisation: model.project.organisation,
        summary: model.project.summary,
        goal: model.project.goal
      },
      analysis_settings: {
        farm_area_hectares: s.farmAreaHa,
        horizon_years: s.horizonYears,
        start_year: s.startYear,
        grain_price_per_tonne: s.grainPricePerTonne,
        discount_rate_percent: s.discountRatePct,
        persistence: s.persistence,
        recurrence_mode: s.recurrenceMode,
        adoption_multiplier: s.adoptionMultiplier,
        risk_multiplier: s.riskMultiplier
      },
      data_summary: {
        rows_imported: model.import.meta?.rowCount ?? "",
        replicate_column: model.import.meta?.keys?.replicate ?? "",
        treatment_column: model.import.meta?.keys?.treatment ?? "",
        yield_column: model.import.meta?.keys?.yieldCol ?? "",
        cost_columns_detected: (model.import.meta?.keys?.costCols || []).length
      },
      data_checks: (model.import.checks || []).map(c => ({ label: c.label, count: c.count, summary: c.summary })),
      treatments: res.rows.map(r => ({
        name: r.treatmentName,
        is_control: r.isControl,
        rank_by_net_present_value: r.rank ?? null,
        present_value_of_benefits: r.pvBenefits,
        present_value_of_costs: r.pvCosts,
        net_present_value: r.npv,
        benefit_cost_ratio: r.bcr,
        return_on_investment_percent: r.roi,
        application_frequency_years: r.recurrenceYears,
        adoption_area_hectares: r.adoptedAreaHa,
        yield_delta_t_per_ha: r.deltaYieldTHa,
        cost_delta_per_ha: r.deltaCostPerHa
      }))
    };

    return compact;
  }

  function buildAiPromptText() {
    const s = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(s);
    const control = res.control;

    const treatmentsRanked = res.rows
      .filter(r => !r.isControl)
      .slice()
      .sort((a, b) => (Number.isFinite(b.npv) ? b.npv : -Infinity) - (Number.isFinite(a.npv) ? a.npv : -Infinity));

    const top = treatmentsRanked.slice(0, 5);

    const adoptionText = `${fmt(s.adoptionMultiplier * 100)} per cent`;
    const riskText = `${fmt(s.riskMultiplier * 100)} per cent`;

    const lines = [];

    lines.push(
      `Write a clear narrative briefing for a farm decision maker using the results and data checks provided in the accompanying JSON results object. Use full sentences and paragraphs only. Do not use bullet points. Do not use em dash characters. Do not use abbreviations.`
    );

    lines.push(
      `Context: The project is titled ${model.project.name}. The analysis compares each soil amendment treatment against the control baseline over ${s.horizonYears} years on ${fmt(s.farmAreaHa)} hectares. The grain price is ${money(s.grainPricePerTonne)} per tonne. The discount rate is ${fmt(s.discountRatePct)} per cent. The persistence setting is ${fmt(s.persistence)}. The recurrence mode is ${s.recurrenceMode}. The adoption multiplier is ${adoptionText}. The risk multiplier is ${riskText}.`
    );

    if (control) {
      lines.push(
        `Start by explaining what present value of benefits, present value of costs, net present value, benefit cost ratio, and return on investment mean in plain language. Then describe the control baseline values and what the control represents. The control present value of benefits is ${money(control.pvBenefits)} and the control present value of costs is ${money(control.pvCosts)} with net present value of ${money(control.npv)}.`
      );
    }

    if (top.length) {
      const t0 = top[0];
      lines.push(
        `Then summarise the leading treatments on net present value. The highest net present value treatment is ${t0.treatmentName} with net present value of ${money(t0.npv)}. Its present value of benefits is ${money(t0.pvBenefits)} and its present value of costs is ${money(t0.pvCosts)}. Explain whether it wins mainly by higher benefits or lower costs relative to the control, and highlight its application frequency in years and its implied yield and cost deltas per hectare.`
      );
      if (top.length > 1) {
        const t1 = top[1];
        lines.push(
          `Also discuss the second ranked treatment ${t1.treatmentName} with net present value of ${money(t1.npv)} and explain how it differs from the leading treatment in benefits, costs, and application frequency.`
        );
      }
      const bestBcr = res.bestBcr;
      if (bestBcr) {
        lines.push(
          `Identify the treatment with the strongest benefit cost ratio. This treatment is ${bestBcr.treatmentName} with benefit cost ratio of ${ratio(bestBcr.bcr)} and net present value of ${money(bestBcr.npv)}. Explain why a strong benefit cost ratio can occur even when net present value is not the maximum.`
        );
      }
    }

    lines.push(
      `Include a short section on data quality. Use the data checks triggers and counts to describe any issues such as missing yield values, missing treatment labels, replicates with no control rows, or potential outliers, and explain how these issues could affect confidence in the economic conclusions.`
    );

    lines.push(
      `Include a short section on sensitivity. Explain how changing grain price, discount rate, persistence, and recurrence can change the ordering of treatments. Use the sensitivity grid outputs in the JSON to explain the direction of change, focusing on what is most likely to alter conclusions.`
    );

    lines.push(
      `End with practical guidance on what inputs would most improve confidence in the results, such as clearer separation of baseline annual costs versus one off amendment costs, more seasons of yield data, or better recording of plot area and cost units. Do not recommend a single treatment. Present trade offs and uncertainty.`
    );

    return lines.join("\n\n");
  }

  function renderAiBriefingPreview() {
    const promptEl = $id("aiPrompt") || $id("aiBriefingPrompt") || $id("copilotPrompt");
    if (promptEl) promptEl.value = buildAiPromptText();

    const jsonEl = $id("resultsJson") || $id("copyJsonBox") || $id("copilotPreview");
    if (jsonEl) jsonEl.value = JSON.stringify(buildResultsJson(), null, 2);
  }

  // =========================
  // 12) Import pipeline: upload + paste for dataset and dictionary
  // =========================
  async function parseDatasetText(text, sourceLabel) {
    const parsed = parseDelimited(text, null);
    model.import.dataset = parsed;
    model.import.lastParsedAt = nowISO();
    showToast(`Dataset parsed from ${sourceLabel}.`);
    // Update inferred meta and checks preview
    const vd = validateAndDerive(parsed, model.import.dictionary);
    model.import.meta = vd.meta;
    model.import.stats = vd.stats;
    model.import.checks = vd.checks;
    renderDataChecks();
    renderAiBriefingPreview();
    return parsed;
  }

  async function parseDictionaryText(text, sourceLabel) {
    const dict = parseDictionaryCSV(text);
    model.import.dictionary = dict;
    showToast(`Dictionary parsed from ${sourceLabel}.`);
    // Refresh checks preview with dictionary context
    if (model.import.dataset) {
      const vd = validateAndDerive(model.import.dataset, dict);
      model.import.meta = vd.meta;
      model.import.stats = vd.stats;
      model.import.checks = vd.checks;
      renderDataChecks();
      renderAiBriefingPreview();
    }
    return dict;
  }

  function bindImportControls() {
    // Dataset controls
    const dataFile = $id("dataFile");
    const dataPaste = $id("dataPaste");
    const parseDataBtn = $id("parseData");
    const parseDataPasteBtn = $id("parseDataPaste");
    const commitDataBtn = $id("commitData");

    // Dictionary controls
    const dictFile = $id("dictFile");
    const dictPaste = $id("dictPaste");
    const parseDictBtn = $id("parseDict");
    const parseDictPasteBtn = $id("parseDictPaste");
    const commitBtn = $id("commitImport");

    // Backward compatible IDs (if present)
    const parseExcelBtn = $id("parseExcel");
    const importExcelBtn = $id("importExcel");

    // Parse dataset from file
    safeOn(parseDataBtn, "click", async e => {
      e.preventDefault();
      if (!dataFile || !dataFile.files || !dataFile.files[0]) {
        showToast("Please choose a dataset file to parse.");
        return;
      }
      const file = dataFile.files[0];
      const text = await file.text();
      await parseDatasetText(text, `file ${file.name}`);
    });

    // Parse dataset from paste
    safeOn(parseDataPasteBtn, "click", async e => {
      e.preventDefault();
      const text = (dataPaste && dataPaste.value) ? dataPaste.value : "";
      if (!text.trim()) {
        showToast("Please paste dataset text to parse.");
        return;
      }
      await parseDatasetText(text, "pasted text");
    });

    // Parse dictionary from file
    safeOn(parseDictBtn, "click", async e => {
      e.preventDefault();
      if (!dictFile || !dictFile.files || !dictFile.files[0]) {
        showToast("Please choose a dictionary file to parse.");
        return;
      }
      const file = dictFile.files[0];
      const text = await file.text();
      await parseDictionaryText(text, `file ${file.name}`);
    });

    // Parse dictionary from paste
    safeOn(parseDictPasteBtn, "click", async e => {
      e.preventDefault();
      const text = (dictPaste && dictPaste.value) ? dictPaste.value : "";
      if (!text.trim()) {
        showToast("Please paste dictionary text to parse.");
        return;
      }
      await parseDictionaryText(text, "pasted text");
    });

    // Commit buttons
    safeOn(commitDataBtn, "click", e => {
      e.preventDefault();
      commitImportedToModel();
    });
    safeOn(commitBtn, "click", e => {
      e.preventDefault();
      commitImportedToModel();
    });

    // Backward compatible Excel parse/commit buttons re-used for dataset parse/commit if dataset controls are absent
    safeOn(parseExcelBtn, "click", async e => {
      if (parseDataBtn || parseDataPasteBtn) return; // dataset UI exists, do not hijack
      e.preventDefault();

      // If a file input with id "dataFile" exists, use it; else create a picker
      if (dataFile) {
        if (!dataFile.files || !dataFile.files[0]) {
          showToast("Please choose a dataset file to parse.");
          return;
        }
        const file = dataFile.files[0];
        await parseDatasetText(await file.text(), `file ${file.name}`);
        return;
      }

      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".tsv,.csv,.txt";
      input.style.display = "none";
      document.body.appendChild(input);
      input.addEventListener("change", async ev => {
        const file = ev.target.files && ev.target.files[0];
        document.body.removeChild(input);
        if (!file) return;
        await parseDatasetText(await file.text(), `file ${file.name}`);
      });
      input.click();
    });

    safeOn(importExcelBtn, "click", e => {
      if (commitDataBtn || commitBtn) return; // commit UI exists
      e.preventDefault();
      commitImportedToModel();
    });
  }

  // =========================
  // 13) Scenario save/load (localStorage)
  // =========================
  const scenarioStorageKey = () => `farming_cba_scenarios::${slug(model.project.name)}`;

  function getScenarioStore() {
    try {
      const raw = localStorage.getItem(scenarioStorageKey());
      if (!raw) return { version: 1, scenarios: [] };
      const obj = JSON.parse(raw);
      if (!obj || !Array.isArray(obj.scenarios)) return { version: 1, scenarios: [] };
      return obj;
    } catch {
      return { version: 1, scenarios: [] };
    }
  }

  function setScenarioStore(store) {
    localStorage.setItem(scenarioStorageKey(), JSON.stringify(store));
  }

  function collectScenarioFromModel(name) {
    const s = computeBaseCaseScenario();
    const treatmentRec = {};
    model.treatments.forEach(t => {
      treatmentRec[t.id] = { recurrenceYears: Math.max(1, Math.round(Number(t.recurrenceYears) || 1)), isControl: !!t.isControl, name: t.name };
    });

    return {
      name: name || `Scenario ${new Date().toISOString().slice(0, 10)}`,
      saved_at_iso: nowISO(),
      config: { ...model.config },
      sensitivity: { ...model.sensitivity },
      treatment_recurrence: treatmentRec
    };
  }

  function applyScenarioToModel(scn) {
    if (!scn) return;

    if (scn.config && typeof scn.config === "object") {
      // Only apply known keys
      for (const k of Object.keys(model.config)) {
        if (k in scn.config) model.config[k] = scn.config[k];
      }
    }
    if (scn.sensitivity && typeof scn.sensitivity === "object") {
      for (const k of Object.keys(model.sensitivity)) {
        if (k in scn.sensitivity) model.sensitivity[k] = scn.sensitivity[k];
      }
    }
    if (scn.treatment_recurrence && typeof scn.treatment_recurrence === "object") {
      // Restore recurrence and control flags by id where possible
      const map = scn.treatment_recurrence;
      let foundControl = false;
      model.treatments.forEach(t => {
        if (map[t.id]) {
          const r = map[t.id];
          if (!t.isControl) t.recurrenceYears = Math.max(1, Math.round(Number(r.recurrenceYears) || 1));
          if (r.isControl) foundControl = true;
        }
      });
      if (foundControl) {
        model.treatments.forEach(t => (t.isControl = false));
        // Set control by stored flag, else keep first
        const ctrlId = Object.keys(map).find(id => map[id] && map[id].isControl);
        const ctrl = model.treatments.find(t => t.id === ctrlId);
        if (ctrl) ctrl.isControl = true;
      }
    }

    // Align time and output value
    model.time.startYear = model.config.startYear;
    model.time.years = model.config.horizonYears;
    model.time.discBase = model.config.discountRatePct;
    const yieldId = getYieldOutputId();
    const yieldOut = model.outputs.find(o => o.id === yieldId);
    if (yieldOut) yieldOut.value = model.config.grainPricePerTonne;

    // Re-render
    setBasicsFieldsFromModel();
    renderRecurrenceConfig();
    renderResults();
    renderAiBriefingPreview();

    showToast(`Scenario loaded: ${scn.name}.`);
  }

  function renderScenarioList() {
    const sel = $id("scenarioSelect");
    if (!sel) return;
    const store = getScenarioStore();
    sel.innerHTML = "";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = "Select a saved scenario";
    sel.appendChild(opt0);

    for (const s of store.scenarios) {
      const opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = s.name;
      sel.appendChild(opt);
    }
  }

  function bindScenarioControls() {
    const nameEl = $id("scenarioName");
    const saveBtn = $id("scenarioSave");
    const loadBtn = $id("scenarioLoad");
    const delBtn = $id("scenarioDelete");
    const sel = $id("scenarioSelect");

    safeOn(saveBtn, "click", e => {
      e.preventDefault();
      const nm = (nameEl && nameEl.value ? nameEl.value : "").trim() || `Scenario ${new Date().toISOString().slice(0, 10)} ${new Date().toISOString().slice(11, 19)}`;
      const store = getScenarioStore();
      const existingIdx = store.scenarios.findIndex(s => s.name === nm);
      const scn = collectScenarioFromModel(nm);
      if (existingIdx >= 0) store.scenarios[existingIdx] = scn;
      else store.scenarios.push(scn);
      setScenarioStore(store);
      renderScenarioList();
      showToast("Scenario saved to local storage.");
    });

    safeOn(loadBtn, "click", e => {
      e.preventDefault();
      const store = getScenarioStore();
      const nm = sel ? sel.value : "";
      const scn = store.scenarios.find(s => s.name === nm);
      if (!scn) {
        showToast("Please select a saved scenario to load.");
        return;
      }
      applyScenarioToModel(scn);
    });

    safeOn(delBtn, "click", e => {
      e.preventDefault();
      const nm = sel ? sel.value : "";
      if (!nm) {
        showToast("Please select a scenario to delete.");
        return;
      }
      const store = getScenarioStore();
      const before = store.scenarios.length;
      store.scenarios = store.scenarios.filter(s => s.name !== nm);
      setScenarioStore(store);
      renderScenarioList();
      if (before !== store.scenarios.length) showToast("Scenario deleted from local storage.");
      else showToast("Scenario not found.");
    });

    safeOn(sel, "change", () => {
      // no toast on selection
    });
  }

  // =========================
  // 14) Bind base settings fields (existing IDs)
  // =========================
  function setBasicsFieldsFromModel() {
    // Project fields
    if ($id("projectName")) $id("projectName").value = model.project.name || "";
    if ($id("projectLead")) $id("projectLead").value = model.project.lead || "";
    if ($id("analystNames")) $id("analystNames").value = model.project.analysts || "";
    if ($id("projectTeam")) $id("projectTeam").value = model.project.team || "";
    if ($id("projectSummary")) $id("projectSummary").value = model.project.summary || "";
    if ($id("projectObjectives")) $id("projectObjectives").value = model.project.objectives || "";
    if ($id("projectActivities")) $id("projectActivities").value = model.project.activities || "";
    if ($id("stakeholderGroups")) $id("stakeholderGroups").value = model.project.stakeholders || "";
    if ($id("lastUpdated")) $id("lastUpdated").value = model.project.lastUpdated || "";
    if ($id("projectGoal")) $id("projectGoal").value = model.project.goal || "";
    if ($id("withProject")) $id("withProject").value = model.project.withProject || "";
    if ($id("withoutProject")) $id("withoutProject").value = model.project.withoutProject || "";
    if ($id("organisation")) $id("organisation").value = model.project.organisation || "";
    if ($id("contactEmail")) $id("contactEmail").value = model.project.contactEmail || "";
    if ($id("contactPhone")) $id("contactPhone").value = model.project.contactPhone || "";

    // Time/risk fields (existing)
    if ($id("startYear")) $id("startYear").value = model.time.startYear;
    if ($id("years")) $id("years").value = model.time.years;
    if ($id("discBase")) $id("discBase").value = model.time.discBase;
    if ($id("discLow")) $id("discLow").value = model.time.discLow;
    if ($id("discHigh")) $id("discHigh").value = model.time.discHigh;

    // Analysis config (if IDs exist)
    if ($id("farmAreaHa")) $id("farmAreaHa").value = model.config.farmAreaHa;
    if ($id("grainPrice")) $id("grainPrice").value = model.config.grainPricePerTonne;
    if ($id("discountRate")) $id("discountRate").value = model.config.discountRatePct;
    if ($id("persistence")) $id("persistence").value = model.config.persistence;
    if ($id("recurrenceMode")) $id("recurrenceMode").value = model.config.recurrenceMode;
    if ($id("adoptionMultiplier")) $id("adoptionMultiplier").value = model.config.adoptionMultiplier;
    if ($id("riskMultiplier")) $id("riskMultiplier").value = model.config.riskMultiplier;

    // Sensitivity config (if IDs exist)
    if ($id("priceLow")) $id("priceLow").value = model.sensitivity.priceLow;
    if ($id("priceBase")) $id("priceBase").value = model.sensitivity.priceBase;
    if ($id("priceHigh")) $id("priceHigh").value = model.sensitivity.priceHigh;
    if ($id("sensDiscLow")) $id("sensDiscLow").value = model.sensitivity.discLow;
    if ($id("sensDiscBase")) $id("sensDiscBase").value = model.sensitivity.discBase;
    if ($id("sensDiscHigh")) $id("sensDiscHigh").value = model.sensitivity.discHigh;
    if ($id("persistLow")) $id("persistLow").value = model.sensitivity.persistenceLow;
    if ($id("persistBase")) $id("persistBase").value = model.sensitivity.persistenceBase;
    if ($id("persistHigh")) $id("persistHigh").value = model.sensitivity.persistenceHigh;
  }

  function bindBasics() {
    document.addEventListener("input", e => {
      const t = e.target;
      if (!t || !t.id) return;

      const id = t.id;

      // Project
      if (id === "projectName") model.project.name = t.value;
      else if (id === "projectLead") model.project.lead = t.value;
      else if (id === "analystNames") model.project.analysts = t.value;
      else if (id === "projectTeam") model.project.team = t.value;
      else if (id === "projectSummary") model.project.summary = t.value;
      else if (id === "projectObjectives") model.project.objectives = t.value;
      else if (id === "projectActivities") model.project.activities = t.value;
      else if (id === "stakeholderGroups") model.project.stakeholders = t.value;
      else if (id === "lastUpdated") model.project.lastUpdated = t.value;
      else if (id === "projectGoal") model.project.goal = t.value;
      else if (id === "withProject") model.project.withProject = t.value;
      else if (id === "withoutProject") model.project.withoutProject = t.value;
      else if (id === "organisation") model.project.organisation = t.value;
      else if (id === "contactEmail") model.project.contactEmail = t.value;
      else if (id === "contactPhone") model.project.contactPhone = t.value;

      // Time
      else if (id === "startYear") {
        model.time.startYear = +t.value;
        model.config.startYear = +t.value;
      } else if (id === "years") {
        model.time.years = +t.value;
        model.config.horizonYears = +t.value;
      } else if (id === "discBase") {
        model.time.discBase = +t.value;
        model.config.discountRatePct = +t.value;
        model.sensitivity.discBase = +t.value;
      } else if (id === "discLow") {
        model.time.discLow = +t.value;
        model.sensitivity.discLow = +t.value;
      } else if (id === "discHigh") {
        model.time.discHigh = +t.value;
        model.sensitivity.discHigh = +t.value;
      }

      // Analysis config (optional)
      else if (id === "farmAreaHa") model.config.farmAreaHa = +t.value;
      else if (id === "grainPrice") {
        model.config.grainPricePerTonne = +t.value;
        model.sensitivity.priceBase = +t.value;
        // keep yield output value aligned
        const yid = getYieldOutputId();
        const yo = model.outputs.find(o => o.id === yid);
        if (yo) yo.value = +t.value;
      } else if (id === "discountRate") model.config.discountRatePct = +t.value;
      else if (id === "persistence") model.config.persistence = +t.value;
      else if (id === "recurrenceMode") model.config.recurrenceMode = t.value;
      else if (id === "adoptionMultiplier") model.config.adoptionMultiplier = +t.value;
      else if (id === "riskMultiplier") model.config.riskMultiplier = +t.value;

      // Sensitivity inputs (optional)
      else if (id === "priceLow") model.sensitivity.priceLow = +t.value;
      else if (id === "priceBase") model.sensitivity.priceBase = +t.value;
      else if (id === "priceHigh") model.sensitivity.priceHigh = +t.value;
      else if (id === "sensDiscLow") model.sensitivity.discLow = +t.value;
      else if (id === "sensDiscBase") model.sensitivity.discBase = +t.value;
      else if (id === "sensDiscHigh") model.sensitivity.discHigh = +t.value;
      else if (id === "persistLow") model.sensitivity.persistenceLow = +t.value;
      else if (id === "persistBase") model.sensitivity.persistenceBase = +t.value;
      else if (id === "persistHigh") model.sensitivity.persistenceHigh = +t.value;
      else return;

      // Update results on changes
      renderResults();
      renderAiBriefingPreview();
    });

    // Buttons (exports and sensitivity)
    safeOn($id("exportCleanTsv"), "click", e => { e.preventDefault(); exportCleanedDatasetTsv(); });
    safeOn($id("exportSummaryCsv"), "click", e => { e.preventDefault(); exportTreatmentSummaryCsv(); });
    safeOn($id("exportSensitivityCsv"), "click", e => { e.preventDefault(); exportSensitivityCsv(); });
    safeOn($id("exportWorkbook"), "click", e => { e.preventDefault(); exportWorkbookIfPossible(); });

    // AI briefing actions
    safeOn($id("copyAiPrompt"), "click", e => {
      e.preventDefault();
      const prompt = buildAiPromptText();
      copyToClipboard(prompt)
        .then(() => showToast("AI briefing prompt copied."))
        .catch(() => showToast("Copy failed. Please copy the prompt manually."));
    });

    safeOn($id("copyResultsJson"), "click", e => {
      e.preventDefault();
      const json = JSON.stringify(buildResultsJson(), null, 2);
      copyToClipboard(json)
        .then(() => showToast("Results JSON copied."))
        .catch(() => showToast("Copy failed. Please copy the JSON manually."));
    });

    // Sensitivity run button (optional)
    safeOn($id("runSensitivity"), "click", e => {
      e.preventDefault();
      // This tool computes sensitivity on-demand for export and AI; for UI, re-render narrative as confirmation.
      renderAiBriefingPreview();
      showToast("Sensitivity grid updated for the current settings.");
    });

    // Scenario save/load controls
    bindScenarioControls();
    renderScenarioList();

    // Results filters
    renderResultsFilters();
  }

  // =========================
  // 15) Existing renderers for outputs and treatments editor
  // =========================
  function renderOutputs() {
    const root = $id("outputsList");
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
          <div class="field"><label>Value ($ per unit)</label><input type="number" step="0.01" value="${Number(o.value) || 0}" data-k="value" data-id="${o.id}" /></div>
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

      // Keep grain price aligned if this is yield output and grainPrice control exists
      if ((o.name || "").toLowerCase().includes("yield")) {
        model.config.grainPricePerTonne = Number(o.value) || model.config.grainPricePerTonne;
        if ($id("grainPrice")) $id("grainPrice").value = model.config.grainPricePerTonne;
      }

      ensureTreatmentDeltas();
      renderTreatments();
      renderResults();
      renderAiBriefingPreview();
    };

    root.onclick = e => {
      const id = e.target.dataset.delOutput;
      if (!id) return;
      if (!confirm("Remove this output metric?")) return;
      model.outputs = model.outputs.filter(o => o.id !== id);
      model.treatments.forEach(t => delete t.deltas[id]);
      ensureTreatmentDeltas();
      renderOutputs();
      renderTreatments();
      renderResults();
      renderAiBriefingPreview();
      showToast("Output metric removed.");
    };
  }

  function renderTreatments() {
    const root = $id("treatmentsList");
    if (!root) return;
    root.innerHTML = "";
    const list = model.treatments.slice();

    for (const t of list) {
      const perHaCost = getPerHaAnnualCost(t);
      const el = document.createElement("div");
      el.className = "item";

      el.innerHTML = `
        <h4>Treatment: ${esc(t.name)}</h4>
        <div class="row">
          <div class="field"><label>Name</label><input value="${esc(t.name)}" data-tk="name" data-id="${t.id}" /></div>
          <div class="field"><label>Area (ha)</label><input type="number" step="0.01" value="${Number(t.area) || 0}" data-tk="area" data-id="${t.id}" /></div>
          <div class="field"><label>Control vs treatment</label>
            <select data-tk="isControl" data-id="${t.id}">
              <option value="treatment" ${!t.isControl ? "selected" : ""}>Treatment</option>
              <option value="control" ${t.isControl ? "selected" : ""}>Control</option>
            </select>
          </div>
          <div class="field"><label>Application frequency (years)</label>
            <input type="number" min="1" step="1" value="${Math.max(1, Math.round(Number(t.recurrenceYears) || 1))}" data-tk="recurrenceYears" data-id="${t.id}" ${t.isControl ? "readonly" : ""} />
          </div>
          <div class="field"><label>&nbsp;</label><button class="btn small danger" data-del-treatment="${t.id}">Remove</button></div>
        </div>

        <div class="row-6">
          <div class="field"><label>Materials cost ($ per ha per year)</label><input type="number" step="0.01" value="${Number(t.materialsCost) || 0}" data-tk="materialsCost" data-id="${t.id}" /></div>
          <div class="field"><label>Services cost ($ per ha per year)</label><input type="number" step="0.01" value="${Number(t.servicesCost) || 0}" data-tk="servicesCost" data-id="${t.id}" /></div>
          <div class="field"><label>Labour cost ($ per ha per year)</label><input type="number" step="0.01" value="${Number(t.labourCost) || 0}" data-tk="labourCost" data-id="${t.id}" /></div>
          <div class="field"><label>Total annual cost ($ per ha)</label><input type="number" step="0.01" value="${Number(perHaCost) || 0}" readonly data-total-cost="${t.id}" /></div>
          <div class="field"><label>Capital cost ($ year 0)</label><input type="number" step="0.01" value="${Number(t.capitalCost) || 0}" data-tk="capitalCost" data-id="${t.id}" /></div>
          <div class="field"><label>Source</label>
            <select data-tk="source" data-id="${t.id}">
              ${["Farm Trials","Plant Farm","ABARES","GRDC","Input Directly"]
                .map(s => `<option ${s === t.source ? "selected" : ""}>${s}</option>`)
                .join("")}
            </select>
          </div>
        </div>

        <div class="field">
          <label>Notes</label>
          <textarea data-tk="notes" data-id="${t.id}" rows="2">${esc(t.notes || "")}</textarea>
        </div>

        <h5>Output deltas relative to control (per ha)</h5>
        <div class="row">
          ${model.outputs
            .map(
              o => `
            <div class="field">
              <label>${esc(o.name)} (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${Number(t.deltas[o.id]) || 0}" data-td="${o.id}" data-id="${t.id}" ${t.isControl ? "readonly" : ""} />
            </div>
          `
            )
            .join("")}
        </div>

        ${
          t.meta
            ? `
          <div class="row-6">
            <div class="field"><label>Plots</label><div class="metric"><div class="value">${fmt(t.meta.nPlots ?? "")}</div></div></div>
            <div class="field"><label>Replicates</label><div class="metric"><div class="value">${fmt(t.meta.nReplicates ?? "")}</div></div></div>
            <div class="field"><label>Mean yield (t per ha)</label><div class="metric"><div class="value">${isFinite(t.meta.meanYieldTHa) ? fmt(t.meta.meanYieldTHa) : "n/a"}</div></div></div>
            <div class="field"><label>Mean yield delta (t per ha)</label><div class="metric"><div class="value">${isFinite(t.meta.meanYieldDeltaTHa) ? fmt(t.meta.meanYieldDeltaTHa) : "n/a"}</div></div></div>
            <div class="field"><label>Mean cost (per ha)</label><div class="metric"><div class="value">${isFinite(t.meta.meanCostPerHa) ? money(t.meta.meanCostPerHa) : "n/a"}</div></div></div>
            <div class="field"><label>Mean cost delta (per ha)</label><div class="metric"><div class="value">${isFinite(t.meta.meanCostDeltaPerHa) ? money(t.meta.meanCostDeltaPerHa) : "n/a"}</div></div></div>
          </div>
        `
            : ""
        }

        <div class="kv"><small class="muted">id:</small> <code>${t.id}</code></div>
      `;

      root.appendChild(el);
    }

    root.oninput = e => {
      const id = e.target.dataset.id;
      if (!id) return;
      const t = model.treatments.find(x => x.id === id);
      if (!t) return;

      const tk = e.target.dataset.tk;
      if (tk) {
        if (tk === "isControl") {
          const wantControl = e.target.value === "control";
          model.treatments.forEach(tt => (tt.isControl = false));
          t.isControl = wantControl;
          if (t.isControl) t.recurrenceYears = 1;
          ensureTreatmentDeltas();
          renderTreatments();
          renderRecurrenceConfig();
          renderResults();
          renderAiBriefingPreview();
          showToast(`Control set to ${t.name}.`);
          return;
        }
        if (tk === "name" || tk === "source" || tk === "notes") t[tk] = e.target.value;
        else if (tk === "recurrenceYears") {
          if (!t.isControl) t.recurrenceYears = Math.max(1, Math.round(Number(e.target.value) || 1));
        } else t[tk] = +e.target.value;

        if (tk === "materialsCost" || tk === "servicesCost" || tk === "labourCost") {
          const box = e.target.closest(".item");
          if (box) {
            const mats = Number(box.querySelector(`input[data-tk="materialsCost"][data-id="${id}"]`)?.value || 0);
            const serv = Number(box.querySelector(`input[data-tk="servicesCost"][data-id="${id}"]`)?.value || 0);
            const lab = Number(box.querySelector(`input[data-tk="labourCost"][data-id="${id}"]`)?.value || 0);
            const totalField = box.querySelector(`input[data-total-cost="${id}"]`);
            if (totalField) totalField.value = mats + serv + lab;
          }
        }
      }

      const td = e.target.dataset.td;
      if (td) {
        if (!t.isControl) t.deltas[td] = +e.target.value;
      }

      renderRecurrenceConfig();
      renderResults();
      renderAiBriefingPreview();
    };

    root.onclick = e => {
      const delId = e.target.dataset.delTreatment;
      if (!delId) return;
      if (!confirm("Remove this treatment?")) return;
      model.treatments = model.treatments.filter(x => x.id !== delId);
      if (!model.treatments.some(x => x.isControl) && model.treatments.length) model.treatments[0].isControl = true;
      ensureTreatmentDeltas();
      renderTreatments();
      renderRecurrenceConfig();
      renderResults();
      renderAiBriefingPreview();
      showToast("Treatment removed.");
    };
  }

  // =========================
  // 16) Legacy summary cards: populate if IDs exist
  // =========================
  function calcAndRenderAllOutputs() {
    // Populate existing summary cards with the best NPV treatment compared with control
    const pvB = $id("pvBenefits");
    const pvC = $id("pvCosts");
    const npvEl = $id("npv");
    const bcrEl = $id("bcr");
    const roiEl = $id("roi");

    const scenario = computeBaseCaseScenario();
    const res = computeAllTreatmentResults(scenario);
    const control = res.control;
    const best = res.bestNpv;

    if (!pvB && !pvC && !npvEl && !bcrEl && !roiEl) return;

    if (best) {
      if (pvB) pvB.textContent = money(best.pvBenefits);
      if (pvC) pvC.textContent = money(best.pvCosts);
      if (npvEl) {
        npvEl.textContent = money(best.npv);
        npvEl.className = "value " + (best.npv >= 0 ? "positive" : "negative");
      }
      if (bcrEl) bcrEl.textContent = Number.isFinite(best.bcr) ? ratio(best.bcr) : "n/a";
      if (roiEl) roiEl.textContent = Number.isFinite(best.roi) ? pct(best.roi) : "n/a";
    } else if (control) {
      if (pvB) pvB.textContent = money(control.pvBenefits);
      if (pvC) pvC.textContent = money(control.pvCosts);
      if (npvEl) {
        npvEl.textContent = money(control.npv);
        npvEl.className = "value " + (control.npv >= 0 ? "positive" : "negative");
      }
      if (bcrEl) bcrEl.textContent = Number.isFinite(control.bcr) ? ratio(control.bcr) : "n/a";
      if (roiEl) roiEl.textContent = Number.isFinite(control.roi) ? pct(control.roi) : "n/a";
    }
  }

  // =========================
  // 17) Exports: hook existing buttons if present
  // =========================
  function bindExportButtons() {
    safeOn($id("exportCsv"), "click", e => { e.preventDefault(); exportTreatmentSummaryCsv(); });
    safeOn($id("exportCsvFoot"), "click", e => { e.preventDefault(); exportTreatmentSummaryCsv(); });
    safeOn($id("exportPdf"), "click", e => { e.preventDefault(); window.print(); showToast("Print dialog opened."); });
    safeOn($id("exportPdfFoot"), "click", e => { e.preventDefault(); window.print(); showToast("Print dialog opened."); });
  }

  // =========================
  // 18) Tabs (existing behaviour)
  // =========================
  function switchTab(target) {
    if (!target) return;
    const navEls = Array.from(document.querySelectorAll("[data-tab],[data-tab-target],[data-tab-jump]"));
    navEls.forEach(el => {
      const key = el.dataset.tab || el.dataset.tabTarget || el.dataset.tabJump;
      const isActive = key === target;
      el.classList.toggle("active", isActive);
      el.setAttribute("aria-selected", isActive ? "true" : "false");
    });

    const panels = Array.from(document.querySelectorAll(".tab-panel"));
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
      if (target) switchTab(target);
      return;
    }

    const firstPanel = document.querySelector(".tab-panel");
    if (firstPanel) {
      const key = firstPanel.dataset.tabPanel || (firstPanel.id ? firstPanel.id.replace(/^tab-/, "") : "");
      if (key) switchTab(key);
    }
  }

  // =========================
  // 19) Buttons that must exist in older layouts
  // =========================
  function initAddButtons() {
    safeOn($id("addOutput"), "click", e => {
      e.preventDefault();
      const id = uid();
      model.outputs.push({ id, name: "Custom output", unit: "unit", value: 0, source: "Input Directly" });
      model.treatments.forEach(t => (t.deltas[id] = 0));
      renderOutputs();
      renderTreatments();
      renderResults();
      renderAiBriefingPreview();
      showToast("Output metric added.");
    });

    safeOn($id("addTreatment"), "click", e => {
      e.preventDefault();
      if (model.treatments.length >= 64) {
        alert("Maximum of 64 treatments reached.");
        return;
      }
      const t = {
        id: uid(),
        name: "New treatment",
        area: Number(model.config.farmAreaHa) || 0,
        adoption: 1,
        deltas: {},
        labourCost: 0,
        materialsCost: 0,
        servicesCost: 0,
        capitalCost: 0,
        constrained: true,
        source: "Input Directly",
        isControl: false,
        notes: "",
        recurrenceYears: 1
      };
      model.outputs.forEach(o => (t.deltas[o.id] = 0));
      model.treatments.push(t);
      ensureTreatmentDeltas();
      renderTreatments();
      renderRecurrenceConfig();
      renderResults();
      renderAiBriefingPreview();
      showToast("Treatment added.");
    });

    safeOn($id("startBtn"), "click", e => {
      e.preventDefault();
      switchTab("project");
      showToast("Welcome. Start with the Project tab.");
    });
  }

  // =========================
  // 20) Save/load project JSON (existing IDs)
  // =========================
  function bindProjectSaveLoad() {
    safeOn($id("saveProject"), "click", e => {
      e.preventDefault();
      const data = JSON.stringify(model, null, 2);
      downloadFile(`cba_${slug(model.project.name)}.json`, data, "application/json");
      showToast("Project JSON downloaded.");
    });

    const loadBtn = $id("loadProject");
    const loadFile = $id("loadFile");
    if (loadBtn && loadFile) {
      safeOn(loadBtn, "click", e => {
        e.preventDefault();
        loadFile.click();
      });

      safeOn(loadFile, "change", async e => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const text = await file.text();
        try {
          const obj = JSON.parse(text);
          // Shallow merge into model, preserving expected shapes
          for (const k of Object.keys(obj)) model[k] = obj[k];

          ensureTreatmentDeltas();
          setBasicsFieldsFromModel();
          renderOutputs();
          renderTreatments();
          renderRecurrenceConfig();
          renderDataChecks();
          renderScenarioList();
          renderResults();
          renderAiBriefingPreview();
          showToast("Project JSON loaded.");
        } catch (err) {
          console.error(err);
          alert("Invalid JSON file.");
          showToast("Invalid project JSON.");
        } finally {
          e.target.value = "";
        }
      });
    }
  }

  // =========================
  // 21) Default dataset from embedded text blocks (optional IDs)
  // =========================
  function tryLoadEmbeddedDefaults() {
    const defaultData = $id("defaultDatasetText");
    const defaultDict = $id("defaultDictionaryText");
    const dataText = defaultData ? (defaultData.value || defaultData.textContent || "") : "";
    const dictText = defaultDict ? (defaultDict.value || defaultDict.textContent || "") : "";

    let did = false;

    if (dictText && dictText.trim().length > 0) {
      try {
        parseDictionaryText(dictText, "embedded dictionary");
        did = true;
      } catch (e) {
        console.error(e);
      }
    }
    if (dataText && dataText.trim().length > 0) {
      parseDatasetText(dataText, "embedded dataset").then(() => {
        // auto-commit if embedded defaults exist
        commitImportedToModel();
      });
      did = true;
    }
    if (did) showToast("Embedded defaults loaded.");
  }

  // =========================
  // 22) Bootstrap
  // =========================
  document.addEventListener("DOMContentLoaded", () => {
    // Align sensitivity base values with analysis base values
    model.sensitivity.priceBase = model.config.grainPricePerTonne;
    model.sensitivity.discBase = model.config.discountRatePct;

    ensureTreatmentDeltas();
    initTabs();
    bindBasics();
    bindExportButtons();
    bindImportControls();
    bindProjectSaveLoad();
    initAddButtons();

    setBasicsFieldsFromModel();
    renderOutputs();
    renderTreatments();
    renderRecurrenceConfig();
    renderScenarioList();
    renderDataChecks();
    renderResults();
    renderAiBriefingPreview();
    calcAndRenderAllOutputs();

    // Optional embedded defaults
    tryLoadEmbeddedDefaults();
  });
})();
