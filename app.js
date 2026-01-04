// app.js
(() => {
  "use strict";

  const state = {
    raw: { columns: [], rows: [], delimiter: "\t" },
    mapping: {
      replicate: null,
      treatment: null,
      controlFlag: null,
      yield: null,
      variableCost: null,
      capitalCost: null
    },
    enrichedRows: [],
    replicates: [],
    treatments: [],
    controlsFound: 0,
    treatmentSummaries: {}, // key -> summary
    cbaConfig: {
      pricePerTonne: 400,
      horizonYears: 10,
      discountRate: 0.04
    },
    results: {
      metricsByTreatment: {}, // key -> metrics
      controlTreatmentKey: null,
      ranking: [],
      filter: "all",
      lastSensitivity: []
    },
    chart: null
  };

  // ---------- Utilities ----------

  const normaliseName = (name) =>
    String(name || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");

  const toNumber = (v) => {
    if (v === null || v === undefined) return null;
    const s = String(v).trim();
    if (!s || s === "?") return null;
    const n = parseFloat(s.replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  };

  const isTruthyControl = (v) => {
    const s = String(v).trim().toLowerCase();
    return s === "true" || s === "t" || s === "1" || s === "yes" || s === "y";
  };

  const avg = (arr) => {
    const nums = arr.filter((x) => Number.isFinite(x));
    if (!nums.length) return null;
    const sum = nums.reduce((a, b) => a + b, 0);
    return sum / nums.length;
  };

  const createEl = (tag, props = {}, children = []) => {
    const el = document.createElement(tag);
    Object.entries(props).forEach(([k, v]) => {
      if (k === "class") el.className = v;
      else if (k === "text") el.textContent = v;
      else if (k === "html") el.innerHTML = v;
      else el.setAttribute(k, v);
    });
    (Array.isArray(children) ? children : [children]).forEach((child) => {
      if (child === null || child === undefined) return;
      if (typeof child === "string") el.appendChild(document.createTextNode(child));
      else el.appendChild(child);
    });
    return el;
  };

  const formatCurrency = (v) => {
    if (!Number.isFinite(v)) return "";
    const abs = Math.abs(v);
    const decimals = abs >= 1000 ? 0 : 2;
    return "$" + v.toFixed(decimals);
  };

  const formatNumber = (v, decimals = 2) => {
    if (!Number.isFinite(v)) return "";
    return v.toFixed(decimals);
  };

  const annuityFactor = (r, n) => {
    if (r <= 0) return n;
    return (1 - Math.pow(1 + r, -n)) / r;
  };

  const notify = (message) => {
    const region = document.getElementById("toastRegion");
    if (region) region.textContent = message;
    const container = document.getElementById("toastContainer");
    if (!container) return;
    const toast = createEl("div", { class: "toast" });
    const msg = createEl("div", { class: "toast-message", text: message });
    const closeBtn = createEl("button", { class: "toast-close", type: "button" }, "×");
    closeBtn.addEventListener("click", () => {
      container.removeChild(toast);
    });
    toast.appendChild(msg);
    toast.appendChild(closeBtn);
    container.appendChild(toast);
    setTimeout(() => {
      if (toast.parentNode === container) container.removeChild(toast);
    }, 5000);
  };

  // ---------- Parsing ----------

  const detectDelimiter = (headerLine) => {
    if (headerLine.includes("\t")) return "\t";
    if (headerLine.includes(",")) return ",";
    return "\t";
  };

  const parseDelimited = (text) => {
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length < 2) {
      return { columns: [], rows: [], delimiter: "\t" };
    }
    const delimiter = detectDelimiter(lines[0]);
    const split = (line) => line.split(delimiter);
    const header = split(lines[0]);
    const columns = header.map((name) => ({
      name,
      norm: normaliseName(name)
    }));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const parts = split(lines[i]);
      const row = {};
      columns.forEach((col, idx) => {
        row[col.name] = parts[idx] !== undefined ? parts[idx] : "";
      });
      rows.push(row);
    }
    return { columns, rows, delimiter };
  };

  // Column mapping with support for replicate_id, total_cost_per_ha_raw, cost_amendment_input_per_ha_raw
  const inferMapping = () => {
    const norms = state.raw.columns.map((c) => c.norm);
    const byNorm = {};
    state.raw.columns.forEach((c, idx) => {
      byNorm[c.norm] = c.name;
    });

    const findCol = (candidates) => {
      for (const cand of candidates) {
        const norm = normaliseName(cand);
        if (byNorm[norm]) return byNorm[norm];
      }
      return null;
    };

    state.mapping.replicate = findCol(["replicate_id", "replicate", "rep", "rep_no"]);
    state.mapping.treatment = findCol([
      "treatment_name",
      "amendment_name",
      "practice_change_label",
      "treatment"
    ]);
    state.mapping.controlFlag = findCol(["is_control", "control_flag", "control"]);
    state.mapping.yield = findCol(["yield_t_ha", "grain_yield_t_ha", "yield"]);
    // raw total cost and amendment input cost
    const totalCostCol = findCol([
      "total_cost_per_ha_raw",
      "total_cost_per_ha",
      "total_cost",
      "variable_cost_per_ha",
      "variable_cost"
    ]);
    const amendCostCol = findCol([
      "cost_amendment_input_per_ha_raw",
      "cost_amendment_input_per_ha",
      "amendment_cost_per_ha"
    ]);
    state.mapping.variableCost = totalCostCol || findCol(["variable_cost_per_ha", "variable_cost"]);
    state.mapping.capitalCost = amendCostCol || findCol(["capital_cost_per_ha", "capital_cost"]);
  };

  const enrichRows = () => {
    const m = state.mapping;
    const rows = [];
    const replicateSet = new Set();
    const treatmentSet = new Set();
    let controlsFound = 0;

    const controlByReplicate = new Map();

    // First pass: basic enrichment
    state.raw.rows.forEach((row, index) => {
      const rep = m.replicate ? String(row[m.replicate] ?? "").trim() : "";
      const repKey = rep || "ALL";
      const treatment = m.treatment ? String(row[m.treatment] ?? "").trim() : "";
      const isControl = m.controlFlag ? isTruthyControl(row[m.controlFlag]) : false;
      const y = m.yield ? toNumber(row[m.yield]) : null;

      let vCostRaw = m.variableCost ? toNumber(row[m.variableCost]) : null;
      let cCostRaw = m.capitalCost ? toNumber(row[m.capitalCost]) : null;

      // Split total cost into variable and capital where possible
      let varCost = null;
      let capCost = null;
      if (vCostRaw !== null && cCostRaw !== null && cCostRaw >= 0 && cCostRaw <= vCostRaw) {
        varCost = vCostRaw - cCostRaw;
        capCost = cCostRaw;
      } else if (vCostRaw !== null) {
        varCost = vCostRaw;
        capCost = cCostRaw !== null ? cCostRaw : 0;
      } else if (cCostRaw !== null) {
        varCost = 0;
        capCost = cCostRaw;
      }

      const totalCost = (varCost || 0) + (capCost || 0);

      const enriched = {
        __rowIndex: index,
        __replicateId: repKey,
        __treatment: treatment,
        __isControl: isControl,
        __yield: y,
        __varCost: varCost,
        __capCost: capCost,
        __totalCost: totalCost,
        raw: row
      };

      rows.push(enriched);

      if (repKey) replicateSet.add(repKey);
      if (treatment) treatmentSet.add(treatment);
      if (isControl) {
        controlsFound += 1;
        const arr = controlByReplicate.get(repKey) || [];
        arr.push(enriched);
        controlByReplicate.set(repKey, arr);
      }
    });

    // Compute replicate-level control baselines
    const baselineByReplicate = new Map();
    controlByReplicate.forEach((list, repKey) => {
      const yields = list.map((r) => r.__yield).filter((v) => Number.isFinite(v));
      const costs = list.map((r) => r.__totalCost).filter((v) => Number.isFinite(v));
      baselineByReplicate.set(repKey, {
        yield: avg(yields),
        totalCost: avg(costs)
      });
    });

    // Fallback: if no per-replicate control, but there is an ALL baseline
    const allBaseline = baselineByReplicate.get("ALL") || null;

    // Second pass: attach deltas vs control within replicate
    rows.forEach((r) => {
      const base = baselineByReplicate.get(r.__replicateId) || allBaseline || null;
      if (base) {
        r.__controlYield = base.yield;
        r.__controlTotalCost = base.totalCost;
        r.__deltaYield = Number.isFinite(r.__yield) && Number.isFinite(base.yield)
          ? r.__yield - base.yield
          : null;
        r.__deltaCost = Number.isFinite(r.__totalCost) && Number.isFinite(base.totalCost)
          ? r.__totalCost - base.totalCost
          : null;
      } else {
        r.__controlYield = null;
        r.__controlTotalCost = null;
        r.__deltaYield = null;
        r.__deltaCost = null;
      }
    });

    state.enrichedRows = rows;
    state.replicates = Array.from(replicateSet).sort();
    state.treatments = Array.from(treatmentSet).sort();
    state.controlsFound = controlsFound;
  };

  const buildTreatmentSummaries = () => {
    const summaries = {};
    state.treatments.forEach((t) => {
      summaries[t] = {
        name: t,
        replicateIds: new Set(),
        yields: [],
        varCosts: [],
        capCosts: [],
        totalCosts: [],
        deltaYields: [],
        deltaCosts: [],
        isControlTreatment: false
      };
    });

    state.enrichedRows.forEach((r) => {
      const key = r.__treatment || "";
      if (!key || !summaries[key]) return;
      const s = summaries[key];
      s.replicateIds.add(r.__replicateId);
      if (Number.isFinite(r.__yield)) s.yields.push(r.__yield);
      if (Number.isFinite(r.__varCost)) s.varCosts.push(r.__varCost);
      if (Number.isFinite(r.__capCost)) s.capCosts.push(r.__capCost);
      if (Number.isFinite(r.__totalCost)) s.totalCosts.push(r.__totalCost);
      if (Number.isFinite(r.__deltaYield)) s.deltaYields.push(r.__deltaYield);
      if (Number.isFinite(r.__deltaCost)) s.deltaCosts.push(r.__deltaCost);
      if (r.__isControl) s.isControlTreatment = true;
    });

    Object.values(summaries).forEach((s) => {
      s.meanYield = avg(s.yields);
      s.meanVarCost = avg(s.varCosts);
      s.meanCapCost = avg(s.capCosts);
      s.meanTotalCost = avg(s.totalCosts);
      s.meanDeltaYield = avg(s.deltaYields);
      s.meanDeltaCost = avg(s.deltaCosts);
      s.replicateCount = s.replicateIds.size;
    });

    state.treatmentSummaries = summaries;
  };

  // ---------- CBA ----------

  const computeCbaMetrics = () => {
    const cfg = state.cbaConfig;
    const metricsByTreatment = {};
    let controlTreatmentKey = null;

    // heuristics: treatment(s) with any isControlTreatment true
    const candidates = Object.values(state.treatmentSummaries).filter(
      (s) => s.isControlTreatment
    );
    if (candidates.length) {
      controlTreatmentKey = candidates[0].name;
    }

    // fallback: if none marked, try "Control"
    if (!controlTreatmentKey) {
      if (state.treatmentSummaries["Control"]) {
        controlTreatmentKey = "Control";
      } else {
        const maybe = Object.keys(state.treatmentSummaries).find((k) =>
          normaliseName(k).includes("control")
        );
        controlTreatmentKey = maybe || null;
      }
    }

    const a = annuityFactor(cfg.discountRate, cfg.horizonYears);

    Object.values(state.treatmentSummaries).forEach((s) => {
      const y = s.meanYield;
      const v = s.meanVarCost || 0;
      const c = s.meanCapCost || 0;
      const pvB = Number.isFinite(y) ? y * cfg.pricePerTonne * a : null;
      const pvVar = v * a;
      const pvCap = c;
      const pvC = (Number.isFinite(pvVar) ? pvVar : 0) + (Number.isFinite(pvCap) ? pvCap : 0);
      const npv = Number.isFinite(pvB) ? pvB - pvC : null;
      const bcr = pvC > 0 && Number.isFinite(pvB) ? pvB / pvC : null;
      const roi = pvC > 0 && Number.isFinite(npv) ? npv / pvC : null;

      metricsByTreatment[s.name] = {
        name: s.name,
        isControl: s.isControlTreatment || s.name === controlTreatmentKey,
        replicateCount: s.replicateCount,
        meanYield: y,
        meanVarCost: s.meanVarCost,
        meanCapCost: s.meanCapCost,
        meanTotalCost: s.meanTotalCost,
        pvB,
        pvVar,
        pvCap,
        pvC,
        npv,
        bcr,
        roi,
        meanDeltaYield: s.meanDeltaYield,
        meanDeltaCost: s.meanDeltaCost
      };
    });

    // ranking by NPV
    const ranking = Object.values(metricsByTreatment).sort((a, b) => {
      const an = Number.isFinite(a.npv) ? a.npv : -Infinity;
      const bn = Number.isFinite(b.npv) ? b.npv : -Infinity;
      return bn - an;
    });
    ranking.forEach((m, idx) => {
      m.rank = idx + 1;
    });

    // deltas vs control
    if (controlTreatmentKey && metricsByTreatment[controlTreatmentKey]) {
      const ctrl = metricsByTreatment[controlTreatmentKey];
      Object.values(metricsByTreatment).forEach((m) => {
        if (!Number.isFinite(m.npv) || !Number.isFinite(ctrl.npv)) {
          m.deltaNPV = null;
        } else {
          m.deltaNPV = m.npv - ctrl.npv;
        }
        if (!Number.isFinite(m.pvC) || !Number.isFinite(ctrl.pvC)) {
          m.deltaPVCost = null;
        } else {
          m.deltaPVCost = m.pvC - ctrl.pvC;
        }
      });
    }

    state.results.metricsByTreatment = metricsByTreatment;
    state.results.controlTreatmentKey = controlTreatmentKey;
    state.results.ranking = ranking;
  };

  // ---------- Rendering ----------

  const updateHeaderStatus = () => {
    const rows = state.raw.rows.length;
    const tCount = state.treatments.length;
    const rCount = state.replicates.length;
    const cCount = state.controlsFound;

    const datasetLabel = document.getElementById("datasetSummaryLabel");
    const controlLabel = document.getElementById("controlSummaryLabel");
    const sr = document.getElementById("statusRowsLoaded");
    const st = document.getElementById("statusTreatments");
    const sp = document.getElementById("statusReplicates");
    const sc = document.getElementById("statusControlsFound");

    if (datasetLabel) {
      datasetLabel.textContent = rows
        ? `${rows} rows · ${tCount} treatments · ${rCount} replicates`
        : "No data loaded";
    }
    if (controlLabel) {
      controlLabel.textContent = state.results.controlTreatmentKey
        ? state.results.controlTreatmentKey
        : "Not detected";
    }
    if (sr) sr.textContent = String(rows);
    if (st) st.textContent = String(tCount);
    if (sp) sp.textContent = String(rCount);
    if (sc) sc.textContent = String(cCount);
  };

  const renderMappingSummary = () => {
    const table = document.getElementById("mappingSummary");
    if (!table) return;
    const tbody = table.querySelector("tbody");
    tbody.innerHTML = "";

    const addRow = (label, colName) => {
      const tr = document.createElement("tr");
      const td1 = createEl("td", {}, label);
      const td2 = document.createElement("td");
      if (colName) {
        td2.textContent = colName;
      } else {
        td2.textContent = "Not detected";
        td2.classList.add("muted");
      }
      tr.appendChild(td1);
      tr.appendChild(td2);
      tbody.appendChild(tr);
    };

    addRow("Replicate", state.mapping.replicate);
    addRow("Treatment name", state.mapping.treatment);
    addRow("Control flag", state.mapping.controlFlag);
    addRow("Yield (t/ha)", state.mapping.yield);
    addRow("Variable cost (per ha)", state.mapping.variableCost);
    addRow("Capital cost (per ha, year 0)", state.mapping.capitalCost);
  };

  const renderDataChecks = () => {
    const ul = document.getElementById("dataChecksList");
    if (!ul) return;
    ul.innerHTML = "";

    const add = (text, type) => {
      const li = document.createElement("li");
      li.textContent = text;
      li.classList.add(type === "ok" ? "ok" : "warn");
      ul.appendChild(li);
    };

    if (!state.raw.rows.length) {
      add("No data loaded.", "warn");
      return;
    }

    if (state.mapping.replicate) add("Replicate column detected.", "ok");
    else add("Replicate column not detected. Using ALL as a single replicate.", "warn");

    if (state.mapping.treatment) add("Treatment name column detected.", "ok");
    else add("Treatment name column not detected.", "warn");

    if (state.mapping.controlFlag) {
      if (state.controlsFound > 0) {
        add(`Control flag detected and ${state.controlsFound} control rows found.`, "ok");
      } else {
        add("Control flag detected, but no rows flagged as control.", "warn");
      }
    } else {
      add("Control flag column not detected. Control baseline may be ambiguous.", "warn");
    }

    if (state.mapping.yield) {
      const yields = state.enrichedRows.map((r) => r.__yield).filter((v) => Number.isFinite(v));
      if (yields.length) {
        add(`Yield column detected with ${yields.length} numeric values.`, "ok");
      } else {
        add("Yield column detected, but values are not numeric.", "warn");
      }
    } else {
      add("Yield column not detected.", "warn");
    }

    if (state.mapping.variableCost) {
      add("Variable or total cost column detected.", "ok");
    } else {
      add("No variable or total cost column detected.", "warn");
    }

    if (state.mapping.capitalCost) {
      add("Capital cost column detected (using amendment input cost per ha where applicable).", "ok");
    } else {
      add("No explicit capital cost column detected. Capital cost is treated as zero.", "warn");
    }
  };

  const renderDataPreview = () => {
    const table = document.getElementById("dataPreviewTable");
    if (!table) return;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!state.raw.columns.length || !state.raw.rows.length) return;

    const headerRow = document.createElement("tr");
    state.raw.columns.forEach((c) => {
      headerRow.appendChild(createEl("th", {}, c.name));
    });
    thead.appendChild(headerRow);

    const maxRows = Math.min(10, state.raw.rows.length);
    for (let i = 0; i < maxRows; i++) {
      const rawRow = state.raw.rows[i];
      const tr = document.createElement("tr");
      state.raw.columns.forEach((c) => {
        tr.appendChild(createEl("td", {}, String(rawRow[c.name] ?? "")));
      });
      tbody.appendChild(tr);
    }
  };

  const renderLeaderboard = () => {
    const table = document.getElementById("leaderboardTable");
    if (!table) return;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!state.results.ranking.length) return;

    const headerRow = document.createElement("tr");
    ["Rank", "Treatment", "Control", "Replicates", "Yield (t/ha)", "PV benefits", "PV costs", "NPV", "BCR", "ROI"].forEach(
      (h) => headerRow.appendChild(createEl("th", {}, h))
    );
    thead.appendChild(headerRow);

    const bestNPV = state.results.ranking[0]?.npv ?? null;

    state.results.ranking.forEach((m) => {
      const tr = document.createElement("tr");
      if (m.isControl) tr.classList.add("highlight-control");
      if (bestNPV !== null && m.npv === bestNPV) tr.classList.add("highlight-best");

      tr.appendChild(createEl("td", {}, String(m.rank)));
      const nameTd = document.createElement("td");
      nameTd.textContent = m.name;
      if (m.isControl) {
        const badge = createEl("span", { class: "results-badge-control" }, "Control");
        nameTd.appendChild(document.createTextNode(" "));
        nameTd.appendChild(badge);
      } else if (m.rank === 1) {
        const badge = createEl("span", { class: "results-badge-best" }, "Top");
        nameTd.appendChild(document.createTextNode(" "));
        nameTd.appendChild(badge);
      }
      tr.appendChild(nameTd);

      tr.appendChild(createEl("td", {}, m.isControl ? "Yes" : ""));
      tr.appendChild(createEl("td", {}, String(m.replicateCount || 0)));
      tr.appendChild(createEl("td", {}, formatNumber(m.meanYield, 3)));

      tr.appendChild(createEl("td", {}, formatCurrency(m.pvB)));
      tr.appendChild(createEl("td", {}, formatCurrency(m.pvC)));
      tr.appendChild(createEl("td", {}, formatCurrency(m.npv)));
      tr.appendChild(createEl("td", {}, formatNumber(m.bcr, 2)));
      tr.appendChild(createEl("td", {}, formatNumber(m.roi, 2)));

      tbody.appendChild(tr);
    });
  };

  const getFilteredTreatments = () => {
    const metrics = state.results.metricsByTreatment;
    const ranking = state.results.ranking.slice();

    if (!ranking.length) return [];

    const controlKey = state.results.controlTreatmentKey;
    const controlMetrics = controlKey ? metrics[controlKey] : null;

    if (state.results.filter === "top5") {
      return ranking.slice(0, 5);
    }
    if (state.results.filter === "bcr") {
      const byBcr = ranking
        .filter((m) => Number.isFinite(m.bcr))
        .sort((a, b) => b.bcr - a.bcr)
        .slice(0, 5);
      return byBcr;
    }
    if (state.results.filter === "better" && controlMetrics) {
      return ranking.filter((m) => Number.isFinite(m.deltaNPV) && m.deltaNPV > 0);
    }
    return ranking;
  };

  const renderComparisonTable = () => {
    const table = document.getElementById("comparisonTable");
    if (!table) return;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!state.results.ranking.length) return;

    const metrics = state.results.metricsByTreatment;
    const controlKey = state.results.controlTreatmentKey;
    const controlMetrics = controlKey ? metrics[controlKey] : null;
    const subset = getFilteredTreatments();

    if (!controlMetrics) return;

    // Header
    const trHead = document.createElement("tr");
    trHead.appendChild(createEl("th", { class: "indicator-col" }, "Indicator"));
    trHead.appendChild(createEl("th", { class: "control-col" }, `Control (${controlMetrics.name})`));
    subset.forEach((m) => {
      if (m.name === controlKey) return;
      trHead.appendChild(createEl("th", { class: "treatment-col" }, m.name));
    });
    thead.appendChild(trHead);

    const indicators = [
      {
        key: "pvB",
        label: "PV benefits per ha",
        formatter: formatCurrency
      },
      {
        key: "pvC",
        label: "PV costs per ha",
        formatter: formatCurrency
      },
      {
        key: "npv",
        label: "Net present value per ha",
        formatter: formatCurrency
      },
      {
        key: "bcr",
        label: "Benefit cost ratio",
        formatter: (v) => formatNumber(v, 2)
      },
      {
        key: "roi",
        label: "Return on investment",
        formatter: (v) => formatNumber(v, 2)
      },
      {
        key: "deltaNPV",
        label: "Change in NPV vs control",
        formatter: formatCurrency
      },
      {
        key: "deltaPVCost",
        label: "Change in PV costs vs control",
        formatter: formatCurrency
      }
    ];

    indicators.forEach((ind) => {
      const tr = document.createElement("tr");
      const th = createEl("th", { class: "indicator-col" }, ind.label);
      tr.appendChild(th);

      const ctrlVal = controlMetrics[ind.key];
      const ctrlTd = document.createElement("td");
      ctrlTd.textContent = ind.formatter(ctrlVal);
      tr.appendChild(ctrlTd);

      subset.forEach((m) => {
        if (m.name === controlKey) return;
        const val = m[ind.key];
        const td = document.createElement("td");
        td.textContent = ind.formatter(val);
        if (ind.key === "deltaNPV" && Number.isFinite(val)) {
          if (val > 0) td.classList.add("positive");
          if (val < 0) td.classList.add("negative");
        }
        if (ind.key === "deltaPVCost" && Number.isFinite(val)) {
          if (val < 0) td.classList.add("positive");
          if (val > 0) td.classList.add("negative");
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
  };

  const updateChart = () => {
    const canvas = document.getElementById("resultsChart");
    if (!canvas) return;
    const subset = getFilteredTreatments();
    if (!subset.length) {
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
      return;
    }

    const labels = subset.map((m) => m.name);
    const npvData = subset.map((m) => m.npv || 0);
    const pvBData = subset.map((m) => m.pvB || 0);
    const pvCData = subset.map((m) => m.pvC || 0);

    if (state.chart) {
      state.chart.destroy();
    }

    state.chart = new Chart(canvas.getContext("2d"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "NPV per ha", data: npvData },
          { label: "PV benefits per ha", data: pvBData },
          { label: "PV costs per ha", data: pvCData }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: "top" },
          tooltip: { mode: "index", intersect: false }
        },
        scales: {
          x: { stacked: false },
          y: { stacked: false, beginAtZero: true }
        }
      }
    });
  };

  const updateJsonExports = () => {
    const out = {
      cbaConfig: state.cbaConfig,
      controlTreatment: state.results.controlTreatmentKey,
      metricsByTreatment: state.results.metricsByTreatment
    };
    const jsonArea = document.getElementById("jsonResultsTextarea");
    if (jsonArea) {
      jsonArea.value = JSON.stringify(out, null, 2);
    }
  };

  const updateAiBriefing = () => {
    const metrics = state.results.metricsByTreatment;
    const ranking = state.results.ranking;
    if (!Object.keys(metrics).length || !ranking.length) {
      const area = document.getElementById("aiBriefingTextarea");
      if (area) area.value = "";
      return;
    }
    const cfg = state.cbaConfig;
    const controlKey = state.results.controlTreatmentKey;
    const controlMetrics = controlKey ? metrics[controlKey] : null;

    const top = ranking[0];
    const topName = top ? top.name : "";
    const topNPV = top ? top.npv : null;
    const textLines = [];

    textLines.push(
      "You are preparing a narrative summary of a faba beans soil management trial that has been analysed using a partial budget and cost benefit framework."
    );
    textLines.push(
      "Each treatment is evaluated per hectare, relative to a control that represents business as usual practice."
    );
    textLines.push(
      "The analysis uses the following economic assumptions: a grain price of " +
        formatCurrency(cfg.pricePerTonne) +
        " per tonne, a time horizon of " +
        String(cfg.horizonYears) +
        " years, and an annual discount rate of " +
        formatNumber(cfg.discountRate * 100, 1) +
        " percent."
    );

    if (controlMetrics) {
      textLines.push(
        "Under these assumptions the control treatment " +
          controlMetrics.name +
          " has an estimated net present value per hectare of " +
          formatCurrency(controlMetrics.npv) +
          ", with present value benefits of " +
          formatCurrency(controlMetrics.pvB) +
          " and present value costs of " +
          formatCurrency(controlMetrics.pvC) +
          "."
      );
    }

    if (top && controlMetrics && top.name !== controlMetrics.name) {
      textLines.push(
        "The top ranked treatment on net present value is " +
          topName +
          " with an estimated net present value per hectare of " +
          formatCurrency(topNPV) +
          "."
      );
      if (Number.isFinite(top.deltaNPV)) {
        textLines.push(
          "Compared with the control this treatment changes net present value by " +
            formatCurrency(top.deltaNPV) +
            " per hectare."
        );
      }
      if (Number.isFinite(top.deltaPVCost)) {
        textLines.push(
          "The change in present value of costs relative to the control is " +
            formatCurrency(top.deltaPVCost) +
            " per hectare."
        );
      }
    }

    textLines.push(
      "In your response, describe clearly which treatments appear most attractive economically, how sensitive these conclusions might be to the price and discount assumptions, and how large the gains or losses are for farmers per hectare."
    );

    const jsonBlock = {
      cbaConfig: state.cbaConfig,
      controlTreatment: state.results.controlTreatmentKey,
      treatments: state.results.ranking.map((m) => ({
        name: m.name,
        isControl: m.isControl,
        rank: m.rank,
        replicateCount: m.replicateCount,
        meanYield: m.meanYield,
        meanVarCost: m.meanVarCost,
        meanCapCost: m.meanCapCost,
        meanTotalCost: m.meanTotalCost,
        pvBenefits: m.pvB,
        pvCosts: m.pvC,
        npv: m.npv,
        bcr: m.bcr,
        roi: m.roi,
        deltaNPV: m.deltaNPV,
        deltaPVCost: m.deltaPVCost
      }))
    };

    textLines.push(
      "At the end of your response use the following JSON object to anchor any quantitative statements and tables:"
    );
    textLines.push(JSON.stringify(jsonBlock));

    const area = document.getElementById("aiBriefingTextarea");
    if (area) area.value = textLines.join("\n\n");
  };

  // ---------- Sensitivity ----------

  const runSensitivityGrid = () => {
    const metrics = state.results.metricsByTreatment;
    if (!Object.keys(metrics).length) {
      notify("Run the main analysis before sensitivity.");
      return;
    }
    const stepPrice = toNumber(document.getElementById("sensPriceStepInput").value) || 50;
    const stepRate =
      (toNumber(document.getElementById("sensRateStepInput").value) || 1) / 100;

    const basePrice = state.cbaConfig.pricePerTonne;
    const baseRate = state.cbaConfig.discountRate;

    const priceLevels = [basePrice - stepPrice, basePrice, basePrice + stepPrice].filter(
      (x, idx, arr) => x > 0 && arr.indexOf(x) === idx
    );
    const rateLevels = [baseRate - stepRate, baseRate, baseRate + stepRate].filter(
      (x, idx, arr) => x >= 0 && arr.indexOf(x) === idx
    );

    const summaries = state.treatmentSummaries;
    const results = [];

    priceLevels.forEach((p) => {
      rateLevels.forEach((r) => {
        const a = annuityFactor(r, state.cbaConfig.horizonYears);
        const localMetrics = Object.values(summaries).map((s) => {
          const y = s.meanYield;
          const v = s.meanVarCost || 0;
          const c = s.meanCapCost || 0;
          const pvB = Number.isFinite(y) ? y * p * a : null;
          const pvVar = v * a;
          const pvCap = c;
          const pvC = (Number.isFinite(pvVar) ? pvVar : 0) + (Number.isFinite(pvCap) ? pvCap : 0);
          const npv = Number.isFinite(pvB) ? pvB - pvC : null;
          return { name: s.name, npv };
        });
        const best = localMetrics
          .filter((m) => Number.isFinite(m.npv))
          .sort((a, b) => b.npv - a.npv)[0];
        if (best) {
          results.push({
            price: p,
            rate: r,
            bestTreatment: best.name,
            npv: best.npv
          });
        }
      });
    });

    state.results.lastSensitivity = results;
    renderSensitivityTable();
    notify("Sensitivity grid updated.");
  };

  const renderSensitivityTable = () => {
    const table = document.getElementById("sensitivityTable");
    if (!table) return;
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const results = state.results.lastSensitivity || [];
    if (!results.length) return;

    const trHead = document.createElement("tr");
    ["Price per tonne", "Discount rate", "Top treatment", "NPV per ha"].forEach((h) =>
      trHead.appendChild(createEl("th", {}, h))
    );
    thead.appendChild(trHead);

    results.forEach((r) => {
      const tr = document.createElement("tr");
      tr.appendChild(createEl("td", {}, formatCurrency(r.price)));
      tr.appendChild(
        createEl("td", {}, formatNumber(r.rate * 100, 1) + " %")
      );
      tr.appendChild(createEl("td", {}, r.bestTreatment));
      tr.appendChild(createEl("td", {}, formatCurrency(r.npv)));
      tbody.appendChild(tr);
    });
  };

  // ---------- Exports ----------

  const buildTsv = () => {
    if (!state.raw.columns.length || !state.raw.rows.length) return "";
    const header = state.raw.columns.map((c) => c.name).join("\t");
    const lines = [header];
    state.raw.rows.forEach((row) => {
      const cells = state.raw.columns.map((c) => String(row[c.name] ?? ""));
      lines.push(cells.join("\t"));
    });
    return lines.join("\n");
  };

  const buildSummaryCsv = () => {
    const metrics = state.results.metricsByTreatment;
    const cols = [
      "treatment_name",
      "is_control",
      "rank",
      "replicates",
      "mean_yield_t_ha",
      "mean_variable_cost_per_ha",
      "mean_capital_cost_per_ha",
      "mean_total_cost_per_ha",
      "pv_benefits_per_ha",
      "pv_costs_per_ha",
      "npv_per_ha",
      "bcr",
      "roi",
      "delta_npv_vs_control",
      "delta_pv_costs_vs_control"
    ];
    const lines = [cols.join(",")];
    state.results.ranking.forEach((m) => {
      const cells = [
        '"' + m.name.replace(/"/g, '""') + '"',
        m.isControl ? "1" : "0",
        m.rank,
        m.replicateCount || "",
        m.meanYield ?? "",
        m.meanVarCost ?? "",
        m.meanCapCost ?? "",
        m.meanTotalCost ?? "",
        m.pvB ?? "",
        m.pvC ?? "",
        m.npv ?? "",
        m.bcr ?? "",
        m.roi ?? "",
        m.deltaNPV ?? "",
        m.deltaPVCost ?? ""
      ];
      lines.push(cells.join(","));
    });
    return lines.join("\n");
  };

  const buildComparisonCsv = () => {
    const metrics = state.results.metricsByTreatment;
    const ranking = getFilteredTreatments();
    if (!ranking.length) return "";
    const controlKey = state.results.controlTreatmentKey;
    if (!controlKey) return "";

    const ctrl = metrics[controlKey];

    const indicators = [
      { key: "pvB", label: "PV benefits per ha" },
      { key: "pvC", label: "PV costs per ha" },
      { key: "npv", label: "Net present value per ha" },
      { key: "bcr", label: "Benefit cost ratio" },
      { key: "roi", label: "Return on investment" },
      { key: "deltaNPV", label: "Change in NPV vs control" },
      { key: "deltaPVCost", label: "Change in PV costs vs control" }
    ];

    const headers = ["indicator", ctrl.name];
    ranking.forEach((m) => {
      if (m.name === controlKey) return;
      headers.push(m.name);
    });
    const lines = [headers.join(",")];

    indicators.forEach((ind) => {
      const row = [ind.label, metrics[controlKey][ind.key] ?? ""];
      ranking.forEach((m) => {
        if (m.name === controlKey) return;
        row.push(m[ind.key] ?? "");
      });
      lines.push(row.join(","));
    });

    return lines.join("\n");
  };

  const downloadTextFile = (filename, text) => {
    const blob = new Blob(["\uFEFF" + text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const copyToClipboard = (text) => {
    if (!navigator.clipboard) {
      const temp = document.createElement("textarea");
      temp.value = text;
      document.body.appendChild(temp);
      temp.select();
      document.execCommand("copy");
      document.body.removeChild(temp);
      return;
    }
    navigator.clipboard.writeText(text).catch(() => {});
  };

  // ---------- Main flow ----------

  const recomputeAll = () => {
    if (!state.raw.rows.length) return;
    inferMapping();
    enrichRows();
    buildTreatmentSummaries();
    computeCbaMetrics();
    renderMappingSummary();
    renderDataChecks();
    renderDataPreview();
    renderLeaderboard();
    renderComparisonTable();
    updateChart();
    updateHeaderStatus();
    updateJsonExports();
    updateAiBriefing();
  };

  const loadFromParsed = (parsed) => {
    state.raw.columns = parsed.columns;
    state.raw.rows = parsed.rows;
    state.raw.delimiter = parsed.delimiter;
    recomputeAll();
    notify("Dataset loaded and processed.");
  };

  const handleFileLoad = () => {
    const input = document.getElementById("fileInput");
    if (!input || !input.files || !input.files.length) {
      notify("Select a TSV or CSV file first.");
      return;
    }
    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const text = String(e.target.result || "");
      const parsed = parseDelimited(text);
      if (!parsed.rows.length) {
        notify("The file did not contain any usable rows.");
        return;
      }
      loadFromParsed(parsed);
    };
    reader.readAsText(file);
  };

  const handlePasteLoad = () => {
    const area = document.getElementById("pasteInput");
    if (!area) return;
    const text = area.value || "";
    if (!text.trim()) {
      notify("Paste the table content first.");
      return;
    }
    const parsed = parseDelimited(text);
    if (!parsed.rows.length) {
      notify("The pasted text did not contain any usable rows.");
      return;
    }
    loadFromParsed(parsed);
  };

  const clearData = () => {
    state.raw = { columns: [], rows: [], delimiter: "\t" };
    state.enrichedRows = [];
    state.treatmentSummaries = {};
    state.results.metricsByTreatment = {};
    state.results.controlTreatmentKey = null;
    state.results.ranking = [];
    state.results.lastSensitivity = [];
    state.treatments = [];
    state.replicates = [];
    state.controlsFound = 0;

    renderMappingSummary();
    renderDataChecks();
    renderDataPreview();
    renderLeaderboard();
    renderComparisonTable();
    renderSensitivityTable();
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }
    updateHeaderStatus();
    updateJsonExports();
    updateAiBriefing();
    notify("All data cleared.");
  };

  const applyConfig = () => {
    const priceEl = document.getElementById("pricePerTonneInput");
    const horizonEl = document.getElementById("horizonYearsInput");
    const rateEl = document.getElementById("discountRateInput");

    const price = toNumber(priceEl.value) || 0;
    const horizon = Math.max(1, Math.round(toNumber(horizonEl.value) || 10));
    const rate = Math.max(0, (toNumber(rateEl.value) || 4) / 100);

    state.cbaConfig.pricePerTonne = price;
    state.cbaConfig.horizonYears = horizon;
    state.cbaConfig.discountRate = rate;

    if (!state.raw.rows.length) {
      notify("Settings updated. Load data to see results.");
      return;
    }
    computeCbaMetrics();
    renderLeaderboard();
    renderComparisonTable();
    updateChart();
    updateHeaderStatus();
    updateJsonExports();
    updateAiBriefing();
    notify("CBA settings applied.");
  };

  const setupTabs = () => {
    const buttons = Array.from(document.querySelectorAll(".tab-button"));
    const panels = Array.from(document.querySelectorAll(".tab-panel"));
    buttons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const target = btn.getAttribute("data-tab");
        if (!target) return;
        buttons.forEach((b) => b.classList.remove("active"));
        panels.forEach((p) => p.classList.remove("active"));
        btn.classList.add("active");
        const panel = document.getElementById(target);
        if (panel) panel.classList.add("active");
      });
    });
  };

  const setupResultsFilters = () => {
    const buttons = Array.from(
      document.querySelectorAll("[data-results-filter]")
    );
    buttons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const f = btn.getAttribute("data-results-filter");
        state.results.filter = f || "all";
        renderComparisonTable();
        updateChart();
      });
    });
  };

  const setupExports = () => {
    const btnTsv = document.getElementById("btnExportCleanTsv");
    const btnSummary = document.getElementById("btnExportSummaryCsv");
    const btnComp = document.getElementById("btnExportComparisonCsv");
    const btnCopyJson = document.getElementById("btnCopyResultsJson");
    const btnCopyAi = document.getElementById("btnCopyAiBriefing");

    if (btnTsv) {
      btnTsv.addEventListener("click", () => {
        if (!state.raw.rows.length) {
          notify("No dataset to export.");
          return;
        }
        const tsv = buildTsv();
        downloadTextFile("faba_beans_trial_clean.tsv", tsv);
        notify("Cleaned dataset exported.");
      });
    }

    if (btnSummary) {
      btnSummary.addEventListener("click", () => {
        if (!state.results.ranking.length) {
          notify("Run the analysis before exporting the summary.");
          return;
        }
        const csv = buildSummaryCsv();
        downloadTextFile("faba_beans_treatment_summary.csv", csv);
        notify("Treatment summary exported.");
      });
    }

    if (btnComp) {
      btnComp.addEventListener("click", () => {
        if (!state.results.ranking.length) {
          notify("Run the analysis before exporting the comparison table.");
          return;
        }
        const csv = buildComparisonCsv();
        if (!csv) {
          notify("Comparison table is not available.");
          return;
        }
        downloadTextFile("faba_beans_comparison_to_control.csv", csv);
        notify("Comparison to control exported.");
      });
    }

    if (btnCopyJson) {
      btnCopyJson.addEventListener("click", () => {
        const area = document.getElementById("jsonResultsTextarea");
        const text = area ? area.value : "";
        if (!text.trim()) {
          notify("No JSON results available to copy.");
          return;
        }
        copyToClipboard(text);
        notify("Results JSON copied to clipboard.");
      });
    }

    if (btnCopyAi) {
      btnCopyAi.addEventListener("click", () => {
        const area = document.getElementById("aiBriefingTextarea");
        const text = area ? area.value : "";
        if (!text.trim()) {
          notify("No AI briefing text available to copy.");
          return;
        }
        copyToClipboard(text);
        notify("AI briefing text copied to clipboard.");
      });
    }
  };

  const setupMainHandlers = () => {
    const btnLoadFile = document.getElementById("btnLoadFile");
    const btnLoadPaste = document.getElementById("btnLoadPaste");
    const btnClearData = document.getElementById("btnClearData");
    const btnApplyConfig = document.getElementById("btnApplyConfig");
    const btnRunSensitivity = document.getElementById("btnRunSensitivity");

    if (btnLoadFile) btnLoadFile.addEventListener("click", handleFileLoad);
    if (btnLoadPaste) btnLoadPaste.addEventListener("click", handlePasteLoad);
    if (btnClearData) btnClearData.addEventListener("click", clearData);
    if (btnApplyConfig) btnApplyConfig.addEventListener("click", applyConfig);
    if (btnRunSensitivity) btnRunSensitivity.addEventListener("click", runSensitivityGrid);
  };

  document.addEventListener("DOMContentLoaded", () => {
    setupTabs();
    setupResultsFilters();
    setupExports();
    setupMainHandlers();
    renderMappingSummary();
    renderDataChecks();
    updateHeaderStatus();
  });
})();
