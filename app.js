/* app.js */
(() => {
  "use strict";

  // ---------- DOM helpers ----------
  const $ = (id) => document.getElementById(id);
  const q = (sel, root = document) => root.querySelector(sel);
  const qa = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const safe = (fn) => {
    try { return fn(); } catch (e) {
      console.error(e);
      setStatus(`Error: ${e?.message || e}`);
      toast(`Error: ${e?.message || e}`);
      return null;
    }
  };

  // ---------- UI ----------
  const setStatus = (msg) => { const el = $("statusBar"); if (el) el.textContent = msg; };
  const setDataState = (msg, ok = true) => {
    const t = $("dataStateText");
    const dot = q("#dataStateChip .dot");
    if (t) t.textContent = msg;
    if (dot) dot.style.background = ok ? "var(--accent2)" : "var(--accent)";
  };

  let toastTimer = null;
  const toast = (msg) => {
    const el = $("toast");
    if (!el) return;
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, 2400);
  };

  const show = (el) => { if (el) el.hidden = false; };
  const hide = (el) => { if (el) el.hidden = true; };

  // ---------- App state ----------
  const DEFAULT_CONFIG = {
    discountRatePct: 5.0,
    horizonYears: 10,
    currency: "AUD"
  };

  const MAPPING = [
    { sheet: "Config", column: "discount_rate_pct", desc: "Discount rate in percent (e.g., 5)" },
    { sheet: "Config", column: "horizon_years", desc: "Time horizon in years (e.g., 10)" },
    { sheet: "Config", column: "currency", desc: "Currency code (AUD, USD, etc.)" },
    { sheet: "Treatments", column: "treatment", desc: "Treatment name (include 'Control')" },
    { sheet: "Treatments", column: "type", desc: "Optional: 'control' or 'treatment'" },
    { sheet: "Cashflows", column: "scenario", desc: "Scenario label (e.g., Base)" },
    { sheet: "Cashflows", column: "year", desc: "Year index (0..horizon)" },
    { sheet: "Cashflows", column: "treatment", desc: "Treatment name (must match Treatments sheet)" },
    { sheet: "Cashflows", column: "benefit", desc: "Annual benefit in currency units" },
    { sheet: "Cashflows", column: "cost", desc: "Annual cost in currency units" }
  ];

  // Built-in sample data (keeps the tool usable immediately)
  const SAMPLE = (() => {
    const treatments = [
      { treatment: "Control", type: "control" },
      { treatment: "Treatment A (Soil amendment)", type: "treatment" },
      { treatment: "Treatment B (Improved grazing)", type: "treatment" },
      { treatment: "Treatment C (Water infrastructure)", type: "treatment" }
    ];

    const scenarios = ["Base", "Dry year", "Wet year"];
    const horizon = DEFAULT_CONFIG.horizonYears;

    // Deterministic cashflow generator (no randomness)
    const rows = [];
    for (const sc of scenarios) {
      for (let year = 0; year <= horizon; year++) {
        for (const tr of treatments) {
          const name = tr.treatment;

          // control baseline
          let benefit = 22000 + year * 600;
          let cost = 9000 + year * 250;

          // scenario modifiers
          if (sc === "Dry year") { benefit *= 0.85; cost *= 1.08; }
          if (sc === "Wet year") { benefit *= 1.08; cost *= 0.98; }

          // treatment modifiers
          if (name.includes("Soil")) {
            // higher upfront cost, then benefit lift
            cost += (year === 0 ? 18000 : 2500);
            benefit += (year >= 2 ? 6000 + year * 200 : 1200);
          } else if (name.includes("grazing")) {
            cost += (year === 0 ? 8000 : 1800);
            benefit += (year >= 1 ? 4500 + year * 250 : 800);
          } else if (name.includes("Water")) {
            cost += (year === 0 ? 26000 : 3200);
            benefit += (year >= 3 ? 7800 + year * 300 : 1000);
          }

          // Ensure non-negative
          benefit = Math.max(0, Math.round(benefit));
          cost = Math.max(0, Math.round(cost));

          rows.push({
            scenario: sc,
            year,
            treatment: name,
            benefit,
            cost
          });
        }
      }
    }

    return {
      config: { ...DEFAULT_CONFIG, horizonYears: horizon, currency: DEFAULT_CONFIG.currency },
      treatments,
      cashflows: rows
    };
  })();

  const state = {
    config: { ...DEFAULT_CONFIG },
    treatments: [...SAMPLE.treatments],
    cashflows: [...SAMPLE.cashflows],
    scenario: "Base",
    lastResults: null
  };

  // ---------- Core calculations ----------
  const discountFactor = (r, t) => 1 / Math.pow(1 + r, t);

  const fmtMoney = (x) => {
    if (!Number.isFinite(x)) return "—";
    const cur = state.config.currency || "AUD";
    try {
      return new Intl.NumberFormat(undefined, { style: "currency", currency: cur, maximumFractionDigits: 0 }).format(x);
    } catch {
      // Fallback if browser doesn't like currency code
      return `${cur} ${Math.round(x).toLocaleString()}`;
    }
  };

  const fmtNum = (x, digits = 2) => {
    if (!Number.isFinite(x)) return "—";
    return x.toFixed(digits);
  };

  const getScenarios = () => {
    const set = new Set(state.cashflows.map(r => String(r.scenario || "Base")));
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  };

  const ensureScenario = () => {
    const scenarios = getScenarios();
    if (!scenarios.includes(state.scenario)) state.scenario = scenarios[0] || "Base";
  };

  const computeResults = () => {
    ensureScenario();
    const r = (state.config.discountRatePct ?? DEFAULT_CONFIG.discountRatePct) / 100;
    const horizon = Math.max(1, parseInt(state.config.horizonYears ?? DEFAULT_CONFIG.horizonYears, 10));
    const sc = state.scenario;

    // treatments list, ensure control exists
    const treatments = state.treatments.map(t => String(t.treatment)).filter(Boolean);
    const controlName =
      state.treatments.find(t => String(t.type).toLowerCase() === "control")?.treatment
      || treatments.find(n => n.toLowerCase() === "control")
      || treatments[0]
      || "Control";

    const byTr = new Map();
    for (const name of treatments) {
      byTr.set(name, { pvB: 0, pvC: 0, pvNet: 0, yearly: [] });
    }

    const relevant = state.cashflows
      .filter(row => String(row.scenario || "Base") === sc)
      .filter(row => Number.isFinite(+row.year))
      .filter(row => +row.year >= 0 && +row.year <= horizon);

    // index for missing years -> treat as 0
    const key = (treatment, year) => `${treatment}__${year}`;
    const cfIndex = new Map();
    for (const row of relevant) {
      const tr = String(row.treatment || "");
      if (!byTr.has(tr)) continue;
      cfIndex.set(key(tr, +row.year), {
        benefit: +row.benefit || 0,
        cost: +row.cost || 0
      });
    }

    for (const name of treatments) {
      const acc = byTr.get(name);
      for (let year = 0; year <= horizon; year++) {
        const cf = cfIndex.get(key(name, year)) || { benefit: 0, cost: 0 };
        const df = discountFactor(r, year);
        const pvB = cf.benefit * df;
        const pvC = cf.cost * df;
        acc.pvB += pvB;
        acc.pvC += pvC;
        acc.pvNet += (pvB - pvC);
        acc.yearly.push({ year, ...cf, df, pvB, pvC, pvNet: pvB - pvC });
      }
    }

    // metrics
    const metrics = [
      { key: "npv", label: "Net present value (NPV)", type: "money" },
      { key: "pvB", label: "Present value of benefits (PV benefits)", type: "money" },
      { key: "pvC", label: "Present value of costs (PV costs)", type: "money" },
      { key: "bcr", label: "Benefit–cost ratio (BCR)", type: "ratio" },
      { key: "roi", label: "Return on investment (ROI)", type: "ratio" },
      { key: "rank", label: "Ranking (by NPV)", type: "int" }
    ];

    const rows = treatments.map(name => {
      const acc = byTr.get(name);
      const pvB = acc.pvB;
      const pvC = acc.pvC;
      const npv = acc.pvNet;
      const bcr = pvC > 0 ? pvB / pvC : NaN;
      const roi = pvC > 0 ? npv / pvC : NaN;
      return { name, pvB, pvC, npv, bcr, roi };
    });

    // ranking by NPV (descending), include control alongside
    const sorted = [...rows].sort((a, b) => (b.npv - a.npv));
    const rankMap = new Map(sorted.map((r, i) => [r.name, i + 1]));
    for (const r0 of rows) r0.rank = rankMap.get(r0.name);

    // deltas vs control
    const controlRow = rows.find(r0 => r0.name === controlName) || rows[0];
    for (const r0 of rows) r0.deltaNpvVsControl = r0.npv - (controlRow?.npv ?? 0);

    return {
      scenario: sc,
      discountRatePct: r * 100,
      horizonYears: horizon,
      treatments,
      controlName: controlName,
      metrics,
      rows,
      byTr
    };
  };

  // ---------- Rendering ----------
  const renderScenarioSelect = () => {
    const sel = $("scenarioSelect");
    if (!sel) return;
    const scenarios = getScenarios();
    ensureScenario();

    sel.innerHTML = "";
    for (const sc of scenarios) {
      const opt = document.createElement("option");
      opt.value = sc;
      opt.textContent = sc;
      if (sc === state.scenario) opt.selected = true;
      sel.appendChild(opt);
    }
  };

  const renderMappingTable = () => {
    const tb = q("#mappingTable tbody");
    if (!tb) return;
    tb.innerHTML = "";
    for (const m of MAPPING) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${m.sheet}</td><td><code>${m.column}</code></td><td>${m.desc}</td>`;
      tb.appendChild(tr);
    }
  };

  const renderPreview = (validation = null) => {
    const tb = q("#dataPreviewTable tbody");
    if (!tb) return;
    const scenarios = getScenarios();
    const tCount = state.treatments.length;
    const cfCount = state.cashflows.length;

    const control =
      state.treatments.find(t => String(t.type).toLowerCase() === "control")?.treatment
      || state.treatments.find(t => String(t.treatment).toLowerCase() === "control")?.treatment
      || state.treatments[0]?.treatment
      || "Control";

    const rows = [
      ["Treatments", String(tCount), "Includes Control alongside treatments"],
      ["Control name", String(control), "Used for ΔNPV in Details"],
      ["Scenarios", scenarios.join(", "), "Use sidebar selector"],
      ["Cashflow rows", String(cfCount), "Scenario × year × treatment rows"],
      ["Discount rate (%)", String(state.config.discountRatePct), "Editable in Settings"],
      ["Horizon (years)", String(state.config.horizonYears), "Editable in Settings"],
      ["Currency", String(state.config.currency), "Used for formatting"]
    ];

    if (validation?.warnings?.length) {
      rows.push(["Validation warnings", String(validation.warnings.length), "See messages above"]);
    }
    if (validation?.errors?.length) {
      rows.push(["Validation errors", String(validation.errors.length), "Fix and re-upload"]);
    }

    tb.innerHTML = "";
    for (const [f, v, n] of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${f}</td><td>${v}</td><td class="muted">${n}</td>`;
      tb.appendChild(tr);
    }
  };

  const renderAssumptions = () => {
    const tb = q("#assumptionsTable tbody");
    if (!tb) return;

    const scenarios = getScenarios();
    const years = state.cashflows
      .filter(r => String(r.scenario || "Base") === state.scenario)
      .map(r => +r.year)
      .filter(n => Number.isFinite(n));

    const minY = years.length ? Math.min(...years) : "—";
    const maxY = years.length ? Math.max(...years) : "—";

    const rows = [
      ["Ranking rule", "NPV (descending)", "Ranking includes control alongside treatments"],
      ["Discounting", "End-of-year cashflows", "Year 0 is undiscounted (df=1)"],
      ["Scenario list", scenarios.join(", "), "From Cashflows sheet"],
      ["Years present in selected scenario", `${minY} .. ${maxY}`, "Filtered by horizon in Results"]
    ];

    tb.innerHTML = "";
    for (const [a, v, n] of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${a}</td><td>${v}</td><td class="muted">${n}</td>`;
      tb.appendChild(tr);
    }
  };

  const renderConfigInputs = () => {
    const d = $("discountRateInput");
    const h = $("timeHorizonInput");
    const c = $("currencySelect");
    if (d) d.value = String(state.config.discountRatePct ?? DEFAULT_CONFIG.discountRatePct);
    if (h) h.value = String(state.config.horizonYears ?? DEFAULT_CONFIG.horizonYears);
    if (c) c.value = String(state.config.currency ?? DEFAULT_CONFIG.currency);

    const smin = $("sensDiscountMin");
    const smax = $("sensDiscountMax");
    if (smin && (smin.value === "")) smin.value = "0";
    if (smax && (smax.value === "")) smax.value = "10";
  };

  const renderResults = (res) => {
    // labels
    const scL = $("scenarioLabel");
    const dL = $("discountLabel");
    const hL = $("horizonLabel");
    if (scL) scL.textContent = res.scenario;
    if (dL) dL.textContent = `${fmtNum(res.discountRatePct, 1)}%`;
    if (hL) hL.textContent = `${res.horizonYears} years`;

    // header
    const head = $("resultsHeaderRow");
    if (!head) return;
    head.innerHTML = `<th class="sticky-col">Indicator</th>`;
    for (const t of res.treatments) {
      const th = document.createElement("th");
      th.textContent = t;
      head.appendChild(th);
    }

    // body
    const body = $("resultsBody");
    if (!body) return;
    body.innerHTML = "";

    const rowByName = new Map(res.rows.map(r => [r.name, r]));

    const metricValue = (metricKey, row) => {
      if (!row) return "—";
      if (metricKey === "npv") return fmtMoney(row.npv);
      if (metricKey === "pvB") return fmtMoney(row.pvB);
      if (metricKey === "pvC") return fmtMoney(row.pvC);
      if (metricKey === "bcr") return fmtNum(row.bcr, 2);
      if (metricKey === "roi") return fmtNum(row.roi, 2);
      if (metricKey === "rank") return Number.isFinite(row.rank) ? String(row.rank) : "—";
      return "—";
    };

    for (const m of res.metrics) {
      const tr = document.createElement("tr");
      const first = document.createElement("td");
      first.className = "sticky-col";
      first.textContent = m.label;
      tr.appendChild(first);

      for (const t of res.treatments) {
        const td = document.createElement("td");
        const r0 = rowByName.get(t);
        td.textContent = metricValue(m.key, r0);
        tr.appendChild(td);
      }
      body.appendChild(tr);
    }

    // details
    const detailsTb = q("#treatmentBreakdownTable tbody");
    const driversBox = $("driversBox");
    if (detailsTb) {
      detailsTb.innerHTML = "";
      const control = res.rows.find(r => r.name === res.controlName) || res.rows[0];
      const sorted = [...res.rows].sort((a, b) => (b.npv - a.npv));
      for (const r0 of sorted) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${r0.name}${r0.name === res.controlName ? " (Control)" : ""}</td>
          <td>${fmtMoney(r0.pvB)}</td>
          <td>${fmtMoney(r0.pvC)}</td>
          <td>${fmtMoney(r0.npv)}</td>
          <td>${fmtMoney(r0.npv - (control?.npv ?? 0))}</td>
        `;
        detailsTb.appendChild(tr);
      }
    }
    if (driversBox) {
      const sorted = [...res.rows].sort((a, b) => (b.npv - a.npv));
      const best = sorted[0];
      const worst = sorted[sorted.length - 1];
      const control = res.rows.find(r => r.name === res.controlName) || res.rows[0];

      driversBox.innerHTML = `
        <strong>What is driving differences?</strong><br/>
        Higher NPV is driven by higher PV benefits, lower PV costs, or earlier benefits (less discounting).<br/><br/>
        <strong>Best (by NPV):</strong> ${best?.name || "—"} with NPV ${fmtMoney(best?.npv)} (Δ vs control ${fmtMoney((best?.npv ?? 0) - (control?.npv ?? 0))}).<br/>
        <strong>Lowest (by NPV):</strong> ${worst?.name || "—"} with NPV ${fmtMoney(worst?.npv)}.
      `;
    }

    state.lastResults = res;
  };

  const renderAll = (validation = null) => {
    renderScenarioSelect();
    renderMappingTable();
    renderPreview(validation);
    renderAssumptions();
    renderConfigInputs();
    const res = computeResults();
    renderResults(res);
  };

  // ---------- Tabs ----------
  const activateTab = (tabName) => {
    const btns = qa(".tab-btn");
    const panels = qa(".tab-panel");

    for (const b of btns) b.classList.toggle("is-active", b.dataset.tab === tabName);
    for (const p of panels) p.classList.toggle("is-active", p.getAttribute("data-tab-panel") === tabName);

    const main = $("main");
    if (main) main.focus();
  };

  // ---------- Clipboard / export ----------
  const tableToTSV = (tableEl) => {
    const rows = qa("tr", tableEl);
    const out = [];
    for (const r of rows) {
      const cells = qa("th,td", r).map(c => (c.textContent ?? "").replace(/\s+/g, " ").trim());
      out.push(cells.join("\t"));
    }
    return out.join("\n");
  };

  const copyResults = async () => safe(async () => {
    const table = $("resultsTable");
    if (!table) return;
    const tsv = tableToTSV(table);
    await navigator.clipboard.writeText(tsv);
    toast("Copied.");
    setStatus("Copied results table to clipboard.");
  });

  const logExport = (msg) => {
    const el = $("exportLog");
    if (!el) return;
    const time = new Date().toLocaleTimeString();
    el.textContent = `[${time}] ${msg}\n` + (el.textContent || "");
  };

  const exportExcelResults = () => safe(() => {
    const res = state.lastResults || computeResults();

    // Prepare a tidy vertical table for Excel: Indicator rows, treatment columns
    const header = ["Indicator", ...res.treatments];
    const table = [header];

    const rowByName = new Map(res.rows.map(r => [r.name, r]));
    const metricValueRaw = (metricKey, row) => {
      if (!row) return null;
      if (metricKey === "npv") return row.npv;
      if (metricKey === "pvB") return row.pvB;
      if (metricKey === "pvC") return row.pvC;
      if (metricKey === "bcr") return row.bcr;
      if (metricKey === "roi") return row.roi;
      if (metricKey === "rank") return row.rank;
      return null;
    };

    for (const m of res.metrics) {
      const r = [m.label];
      for (const t of res.treatments) {
        const row = rowByName.get(t);
        r.push(metricValueRaw(m.key, row));
      }
      table.push(r);
    }

    // Also include a long-form sheet for auditing
    const audit = [["Scenario","Year","Treatment","Benefit","Cost"]];
    const horizon = res.horizonYears;
    for (const row of state.cashflows) {
      if (String(row.scenario || "Base") !== res.scenario) continue;
      const y = +row.year;
      if (!Number.isFinite(y) || y < 0 || y > horizon) continue;
      audit.push([row.scenario, y, row.treatment, +row.benefit || 0, +row.cost || 0]);
    }

    if (typeof XLSX === "undefined") {
      // Fallback: CSV (results only)
      const csv = table.map(r => r.map(v => (v === null || v === undefined) ? "" : String(v)).join(",")).join("\n");
      downloadBlob(csv, `Farming_CBA_Results_${slug(res.scenario)}.csv`, "text/csv;charset=utf-8");
      toast("Exported CSV (XLSX not available).");
      logExport("Exported CSV (SheetJS not available).");
      return;
    }

    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet(table);
    const ws2 = XLSX.utils.aoa_to_sheet(audit);

    XLSX.utils.book_append_sheet(wb, ws1, "Results");
    XLSX.utils.book_append_sheet(wb, ws2, "Cashflows (selected)");

    // Add a small config sheet
    const cfg = [
      ["scenario", res.scenario],
      ["discount_rate_pct", res.discountRatePct],
      ["horizon_years", res.horizonYears],
      ["currency", state.config.currency]
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(cfg);
    XLSX.utils.book_append_sheet(wb, ws3, "Config");

    XLSX.writeFile(wb, `Farming_CBA_Results_${slug(res.scenario)}.xlsx`);
    toast("Exported Excel.");
    logExport("Exported Excel results workbook.");
  });

  const exportWordHtml = () => safe(() => {
    const res = state.lastResults || computeResults();
    const table = $("resultsTable");
    if (!table) return;

    const html = `<!doctype html>
<html><head><meta charset="utf-8">
<title>Farming CBA Results</title>
<style>
  body{font-family: Arial, sans-serif; margin:24px; color:#111;}
  h1{font-size:18px; margin:0 0 8px;}
  p{margin:0 0 16px; color:#444;}
  table{border-collapse:collapse; width:100%; font-size:12px;}
  th,td{border:1px solid #ccc; padding:8px; text-align:left;}
  th{background:#f3f3f3;}
</style>
</head><body>
<h1>Farming CBA Decision Tool 2 — Results</h1>
<p>Scenario: ${escapeHtml(res.scenario)} | Discount: ${res.discountRatePct.toFixed(1)}% | Horizon: ${res.horizonYears} years | Currency: ${escapeHtml(state.config.currency)}</p>
${table.outerHTML}
</body></html>`;

    downloadBlob(html, `Farming_CBA_Results_${slug(res.scenario)}.html`, "text/html;charset=utf-8");
    toast("Downloaded Word-friendly HTML.");
    logExport("Downloaded Word-friendly HTML.");
  });

  // ---------- Excel template / upload ----------
  const slug = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const downloadBlob = (data, filename, mime) => {
    const blob = new Blob([data], { type: mime });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 500);
  };

  const buildTemplateWorkbook = (scenario = null) => {
    const sc = scenario || state.scenario;
    const cfg = [
      ["discount_rate_pct", state.config.discountRatePct],
      ["horizon_years", state.config.horizonYears],
      ["currency", state.config.currency]
    ];

    const tr = [["treatment", "type"]];
    for (const t of state.treatments) tr.push([t.treatment, t.type || "treatment"]);

    const cf = [["scenario","year","treatment","benefit","cost"]];
    const horizon = state.config.horizonYears;
    const treatNames = state.treatments.map(t => t.treatment);

    // Use current in-memory cashflows if available for chosen scenario, else write zeros.
    const idx = new Map();
    for (const r of state.cashflows) {
      if (String(r.scenario || "Base") !== sc) continue;
      idx.set(`${r.year}__${r.treatment}`, r);
    }

    for (let year = 0; year <= horizon; year++) {
      for (const tName of treatNames) {
        const r0 = idx.get(`${year}__${tName}`);
        cf.push([sc, year, tName, r0 ? r0.benefit : 0, r0 ? r0.cost : 0]);
      }
    }

    return { cfg, tr, cf };
  };

  const downloadTemplate = (scenario = null) => safe(() => {
    const { cfg, tr, cf } = buildTemplateWorkbook(scenario);

    if (typeof XLSX === "undefined") {
      // Fallback: 3 CSV files
      downloadBlob(cfg.map(r => r.join(",")).join("\n"), "Config_template.csv", "text/csv;charset=utf-8");
      downloadBlob(tr.map(r => r.join(",")).join("\n"), "Treatments_template.csv", "text/csv;charset=utf-8");
      downloadBlob(cf.map(r => r.join(",")).join("\n"), "Cashflows_template.csv", "text/csv;charset=utf-8");
      toast("Downloaded CSV templates (XLSX not available).");
      setStatus("Downloaded CSV templates.");
      return;
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cfg), "Config");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(tr), "Treatments");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cf), "Cashflows");
    XLSX.writeFile(wb, `Farming_CBA_Template_${slug(scenario || state.scenario)}.xlsx`);
    toast("Downloaded Excel template.");
    setStatus("Downloaded Excel template.");
  });

  const normaliseHeader = (h) => String(h || "").trim().toLowerCase().replace(/\s+/g, "_");
  const sheetToObjects = (ws) => {
    const rows = (typeof XLSX !== "undefined")
      ? XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false, defval: "" })
      : [];

    if (!rows.length) return [];
    const header = rows[0].map(normaliseHeader);
    const out = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every(v => String(v).trim() === "")) continue;
      const obj = {};
      for (let j = 0; j < header.length; j++) obj[header[j]] = r[j];
      out.push(obj);
    }
    return out;
  };

  const validateParsed = (parsed) => {
    const errors = [];
    const warnings = [];

    if (!parsed.treatments.length) errors.push("No treatments found (Treatments sheet).");
    if (!parsed.cashflows.length) errors.push("No cashflows found (Cashflows sheet).");

    const names = parsed.treatments.map(t => String(t.treatment || "").trim()).filter(Boolean);
    const set = new Set(names);

    if (!set.size) errors.push("Treatments sheet: 'treatment' column appears empty.");
    const hasControl = parsed.treatments.some(t => String(t.type || "").toLowerCase() === "control")
      || names.some(n => n.toLowerCase() === "control");
    if (!hasControl) warnings.push("No explicit control found. The first treatment will be treated as control.");

    // ensure cashflow treatment names exist
    const bad = [];
    for (const r of parsed.cashflows) {
      const t = String(r.treatment || "").trim();
      if (!t) continue;
      if (!set.has(t)) bad.push(t);
    }
    if (bad.length) warnings.push(`Some cashflow rows reference treatments not in Treatments sheet (e.g., ${bad.slice(0,3).join(", ")}).`);

    // numeric checks
    const badYears = parsed.cashflows.filter(r => !Number.isFinite(+r.year)).length;
    if (badYears) warnings.push(`Some cashflow rows have non-numeric years (${badYears} rows).`);

    return { errors, warnings };
  };

  const applyParsed = (parsed) => {
    // config
    const d = +parsed.config.discount_rate_pct;
    const h = +parsed.config.horizon_years;
    const c = String(parsed.config.currency || state.config.currency || DEFAULT_CONFIG.currency).trim() || DEFAULT_CONFIG.currency;

    state.config.discountRatePct = Number.isFinite(d) ? d : state.config.discountRatePct;
    state.config.horizonYears = Number.isFinite(h) ? Math.max(1, Math.round(h)) : state.config.horizonYears;
    state.config.currency = c;

    // treatments
    const t = parsed.treatments
      .map(r => ({
        treatment: String(r.treatment || "").trim(),
        type: String(r.type || "treatment").trim()
      }))
      .filter(r => r.treatment);

    state.treatments = t.length ? t : [...state.treatments];

    // cashflows
    state.cashflows = parsed.cashflows
      .map(r => ({
        scenario: String(r.scenario || "Base").trim() || "Base",
        year: Number.isFinite(+r.year) ? Math.round(+r.year) : NaN,
        treatment: String(r.treatment || "").trim(),
        benefit: Number.isFinite(+r.benefit) ? +r.benefit : 0,
        cost: Number.isFinite(+r.cost) ? +r.cost : 0
      }))
      .filter(r => r.treatment);

    ensureScenario();
  };

  const parseWorkbook = (wb) => {
    const sheetNames = wb.SheetNames || [];

    const getSheet = (name) => {
      const found = sheetNames.find(s => s.toLowerCase() === name.toLowerCase());
      return found ? wb.Sheets[found] : null;
    };

    const wsCfg = getSheet("Config");
    const wsTr = getSheet("Treatments");
    const wsCf = getSheet("Cashflows");

    const config = { discount_rate_pct: state.config.discountRatePct, horizon_years: state.config.horizonYears, currency: state.config.currency };

    if (wsCfg) {
      const rows = XLSX.utils.sheet_to_json(wsCfg, { header: 1, blankrows: false, defval: "" });
      for (const r of rows) {
        const k = normaliseHeader(r[0]);
        const v = r[1];
        if (!k) continue;
        if (k === "discount_rate_pct") config.discount_rate_pct = v;
        if (k === "horizon_years") config.horizon_years = v;
        if (k === "currency") config.currency = v;
      }
    }

    const treatments = wsTr ? sheetToObjects(wsTr) : [];
    const cashflows = wsCf ? sheetToObjects(wsCf) : [];
    return { config, treatments, cashflows };
  };

  const uploadExcelFromInput = (inputEl, summaryEl, okEl, warnEl) => safe(async () => {
    if (!inputEl || !inputEl.files || !inputEl.files[0]) {
      toast("Choose a file first.");
      setStatus("No file selected.");
      return;
    }
    if (typeof XLSX === "undefined") {
      toast("Excel import needs SheetJS (XLSX).");
      setStatus("Excel import unavailable (XLSX not loaded).");
      return;
    }

    const file = inputEl.files[0];
    setStatus(`Reading ${file.name}...`);
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const parsed = parseWorkbook(wb);
    const validation = validateParsed(parsed);

    if (summaryEl) {
      if (validation.errors.length || validation.warnings.length) {
        summaryEl.hidden = false;
        summaryEl.className = "callout " + (validation.errors.length ? "callout--warning" : "callout--info");
        summaryEl.innerHTML =
          (validation.errors.length ? `<strong>Errors</strong><br/>${validation.errors.map(e => `• ${escapeHtml(e)}`).join("<br/>")}<br/><br/>` : "") +
          (validation.warnings.length ? `<strong>Warnings</strong><br/>${validation.warnings.map(w => `• ${escapeHtml(w)}`).join("<br/>")}` : "");
      } else {
        summaryEl.hidden = false;
        summaryEl.className = "callout callout--success";
        summaryEl.textContent = "Uploaded and parsed successfully.";
      }
    }

    if (warnEl) {
      if (validation.warnings.length) {
        warnEl.hidden = false;
        warnEl.textContent = validation.warnings.join(" ");
      } else {
        warnEl.hidden = true;
        warnEl.textContent = "";
      }
    }
    if (okEl) {
      if (!validation.errors.length) {
        okEl.hidden = false;
        okEl.textContent = "Excel loaded successfully.";
      } else {
        okEl.hidden = true;
        okEl.textContent = "";
      }
    }

    if (validation.errors.length) {
      setDataState("Upload has errors", false);
      setStatus("Upload has errors. Fix and re-upload.");
      toast("Upload has errors.");
      renderAll(validation);
      return;
    }

    applyParsed(parsed);
    setDataState("Excel loaded", true);
    setStatus("Excel loaded. Results updated.");
    toast("Excel loaded.");
    renderAll(validation);
  });

  // ---------- Modal ----------
  const openModal = (id) => {
    const el = $(id);
    if (!el) return;
    el.hidden = false;
    el.setAttribute("aria-hidden", "false");
  };
  const closeModal = (id) => {
    const el = $(id);
    if (!el) return;
    el.hidden = true;
    el.setAttribute("aria-hidden", "true");
  };

  // ---------- Utils ----------
  const escapeHtml = (s) => String(s ?? "")
    .replaceAll("&","&amp;").replaceAll("<","&lt;").replaceAll(">","&gt;")
    .replaceAll('"',"&quot;").replaceAll("'","&#039;");

  // ---------- Sensitivity ----------
  const runSensitivity = () => safe(() => {
    const minEl = $("sensDiscountMin");
    const maxEl = $("sensDiscountMax");
    const warn = $("sensWarn");
    const tb = q("#sensTable tbody");

    if (!minEl || !maxEl || !tb) return;

    const dMin = +minEl.value;
    const dMax = +maxEl.value;

    if (!Number.isFinite(dMin) || !Number.isFinite(dMax) || dMin < 0 || dMax < 0 || dMax < dMin) {
      if (warn) { warn.hidden = false; warn.textContent = "Please enter a valid min/max (max ≥ min, both ≥ 0)."; }
      return;
    }
    if (warn) { warn.hidden = true; warn.textContent = ""; }

    const original = state.config.discountRatePct;
    const steps = 11;
    const step = (dMax - dMin) / (steps - 1 || 1);

    tb.innerHTML = "";

    for (let i = 0; i < steps; i++) {
      const d = dMin + step * i;
      state.config.discountRatePct = d;
      const res = computeResults();

      const sorted = [...res.rows].sort((a, b) => (b.npv - a.npv));
      const best = sorted[0];
      const control = res.rows.find(r => r.name === res.controlName) || res.rows[0];

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${fmtNum(d, 2)}</td>
        <td>${best?.name || "—"}</td>
        <td>${fmtMoney(control?.npv ?? NaN)}</td>
        <td>${fmtMoney(best?.npv ?? NaN)}</td>
        <td>${fmtMoney((best?.npv ?? 0) - (control?.npv ?? 0))}</td>
      `;
      tb.appendChild(tr);
    }

    state.config.discountRatePct = original;
    renderAll();
    setStatus("Sensitivity run completed.");
    toast("Sensitivity completed.");
  });

  // ---------- Buttons wiring ----------
  const onClick = (id, fn) => {
    const el = $(id);
    if (!el) return;
    el.addEventListener("click", (e) => { e.preventDefault(); fn(); });
  };

  const init = () => safe(() => {
    // Tabs (event delegation)
    const nav = $("navTabs");
    if (nav) {
      nav.addEventListener("click", (e) => {
        const btn = e.target.closest(".tab-btn");
        if (!btn) return;
        const tab = btn.dataset.tab;
        if (!tab) return;
        activateTab(tab);
      });
    }

    // Scenario selector
    const scenarioSel = $("scenarioSelect");
    if (scenarioSel) {
      scenarioSel.addEventListener("change", () => {
        state.scenario = scenarioSel.value;
        renderAll();
        setStatus(`Scenario set to ${state.scenario}.`);
      });
    }

    // Navigation shortcuts
    onClick("goResultsBtn", () => activateTab("results"));
    onClick("goDataBtn", () => activateTab("data"));

    // Template downloads
    onClick("downloadTemplateBtn", () => downloadTemplate(null));
    onClick("downloadTemplateBtn2", () => downloadTemplate(null));
    onClick("downloadScenarioTemplateBtn", () => downloadTemplate(state.scenario));

    // Upload triggers
    const uploadBtn = $("uploadExcelBtn");
    const uploadIn = $("uploadExcelInput");
    if (uploadBtn && uploadIn) {
      uploadBtn.addEventListener("click", () => {
        if (!uploadIn.files || !uploadIn.files[0]) uploadIn.click();
        else uploadExcelFromInput(uploadIn, $("validationSummary"), null, null);
      });
      uploadIn.addEventListener("change", () => uploadExcelFromInput(uploadIn, $("validationSummary"), null, null));
    }

    const uploadBtn2 = $("uploadExcelBtn2");
    const uploadIn2 = $("uploadExcelInput2");
    if (uploadBtn2 && uploadIn2) {
      uploadBtn2.addEventListener("click", () => {
        if (!uploadIn2.files || !uploadIn2.files[0]) uploadIn2.click();
        else uploadExcelFromInput(uploadIn2, null, $("dataOk"), $("dataWarnings"));
      });
      uploadIn2.addEventListener("change", () => uploadExcelFromInput(uploadIn2, null, $("dataOk"), $("dataWarnings")));
    }

    // Settings apply/reset
    onClick("applyConfigBtn", () => {
      const d = +($("discountRateInput")?.value ?? state.config.discountRatePct);
      const h = +($("timeHorizonInput")?.value ?? state.config.horizonYears);
      const c = String($("currencySelect")?.value ?? state.config.currency);

      if (Number.isFinite(d)) state.config.discountRatePct = d;
      if (Number.isFinite(h)) state.config.horizonYears = Math.max(1, Math.round(h));
      if (c) state.config.currency = c;

      renderAll();
      setStatus("Settings applied.");
      toast("Settings applied.");
    });

    onClick("resetConfigBtn", () => {
      state.config = { ...DEFAULT_CONFIG };
      renderAll();
      setStatus("Settings reset.");
      toast("Settings reset.");
    });

    // Assumptions edit (simple inline prompt; keeps everything functional)
    onClick("editAssumptionsBtn", () => {
      const msg = "Assumptions are derived from data + settings. Edit Treatments/Cashflows in Excel, then upload.";
      const w = $("assumptionWarnings");
      if (w) { w.hidden = false; w.textContent = msg; }
      toast("Assumptions are data-driven.");
    });
    onClick("restoreAssumptionsBtn", () => {
      const w = $("assumptionWarnings");
      if (w) { w.hidden = true; w.textContent = ""; }
      toast("Defaults restored.");
    });

    // Recalc / reset
    onClick("recalcBtn", () => {
      renderAll();
      setStatus("Recalculated.");
      toast("Recalculated.");
    });

    onClick("resetBtn", () => {
      state.config = { ...DEFAULT_CONFIG };
      state.treatments = [...SAMPLE.treatments];
      state.cashflows = [...SAMPLE.cashflows];
      state.scenario = "Base";
      state.lastResults = null;
      renderAll();
      setDataState("Reset to sample data", true);
      setStatus("Reset to sample data.");
      toast("Reset.");
    });

    // Results details toggle
    onClick("toggleDetailsBtn", () => {
      const card = $("detailsCard");
      if (!card) return;
      const willShow = card.hidden === true;
      card.hidden = !willShow;
      const btn = $("toggleDetailsBtn");
      if (btn) btn.textContent = willShow ? "Hide details" : "Show details";
    });

    // Copy/export buttons (all wired)
    ["copyResultsBtnTop","copyResultsBtn","copyResultsBtn2"].forEach(id => onClick(id, () => copyResults()));
    ["exportExcelBtnTop","exportExcelBtn","exportExcelBtn3"].forEach(id => onClick(id, () => exportExcelResults()));
    onClick("exportWordBtn", () => exportWordHtml());

    // Sensitivity
    onClick("runSensitivityBtn", () => runSensitivity());

    // Mapping modal
    onClick("openMappingBtn", () => openModal("mappingModal"));
    document.addEventListener("click", (e) => {
      const closeId = e.target?.getAttribute?.("data-close-modal");
      if (closeId) closeModal(closeId);
    });

    // Initial render
    setDataState("Loaded (sample data)", true);
    renderAll();
    setStatus("Ready.");
  });

  // Ensure bindings happen even if defer ordering changes
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
