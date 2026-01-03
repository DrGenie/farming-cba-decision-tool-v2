// app.js
/* Farming CBA Decision Aid (vanilla JS)
   - Fully functional tabs
   - CRUD costs/benefits lines
   - NPV, BCR, ROI, payback
   - Streams table + charts (SVG)
   - One-way sensitivity + export
   - Import/export JSON + localStorage persistence
*/

const APP_KEY = "farming_cba_v1_1";
const APP_VERSION = "1.1";

const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

function uid(prefix = "id") {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function clamp(n, lo, hi) {
  const x = Number.isFinite(n) ? n : lo;
  return Math.max(lo, Math.min(hi, x));
}

function safeNum(v, fallback = 0) {
  const n = Number(v);
  return Number.isFinite(n) ? n : fallback;
}

function round2(x) {
  const n = safeNum(x, 0);
  return Math.round(n * 100) / 100;
}

function escapeHtml(s) {
  return String(s)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function downloadText(filename, text, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function formatCurrency(amount, currency = "AUD") {
  const n = safeNum(amount, 0);
  try {
    return new Intl.NumberFormat(undefined, {
      style: "currency",
      currency,
      maximumFractionDigits: 0,
    }).format(n);
  } catch {
    return `${currency} ${Math.round(n).toLocaleString()}`;
  }
}

function formatNumber(amount, dp = 2) {
  const n = safeNum(amount, 0);
  return n.toLocaleString(undefined, { maximumFractionDigits: dp });
}

function setStatus(msg) {
  const el = $("#headerStatus");
  if (!el) return;
  el.textContent = msg;
  clearTimeout(setStatus._t);
  setStatus._t = setTimeout(() => (el.textContent = "Ready"), 1600);
}

/* ---------- Modal + Toast ---------- */

const Modal = (() => {
  const backdrop = $("#modalBackdrop");
  const modal = $("#modal");
  const titleEl = $("#modalTitle");
  const bodyEl = $("#modalBody");
  const footEl = $("#modalFoot");
  const btnClose = $("#btnModalClose");
  const btnCancel = $("#btnModalCancel");
  const btnOk = $("#btnModalOk");

  let onOk = null;
  let onCancel = null;

  function open({ title = "Modal", bodyHtml = "", okText = "OK", cancelText = "Cancel", showCancel = true, showOk = true, danger = false, onOkCb = null, onCancelCb = null } = {}) {
    titleEl.textContent = title;
    bodyEl.innerHTML = bodyHtml;
    btnOk.textContent = okText;
    btnCancel.textContent = cancelText;

    btnCancel.style.display = showCancel ? "" : "none";
    btnOk.style.display = showOk ? "" : "none";

    btnOk.classList.toggle("btn-primary", !danger);
    btnOk.classList.toggle("btn-ghost", danger);
    btnOk.style.borderColor = danger ? "rgba(252,165,165,0.35)" : "";

    onOk = onOkCb;
    onCancel = onCancelCb;

    backdrop.hidden = false;
    modal.hidden = false;

    setTimeout(() => btnOk.focus(), 0);
  }

  function close() {
    backdrop.hidden = true;
    modal.hidden = true;
    onOk = null;
    onCancel = null;
  }

  function wire() {
    btnClose.addEventListener("click", () => {
      if (onCancel) onCancel();
      close();
    });
    btnCancel.addEventListener("click", () => {
      if (onCancel) onCancel();
      close();
    });
    btnOk.addEventListener("click", () => {
      const cb = onOk;
      close();
      if (cb) cb();
    });

    backdrop.addEventListener("click", () => {
      if (onCancel) onCancel();
      close();
    });

    document.addEventListener("keydown", (e) => {
      if (modal.hidden) return;
      if (e.key === "Escape") {
        if (onCancel) onCancel();
        close();
      }
    });
  }

  return { open, close, wire };
})();

function toast(title, body, ms = 2600) {
  const wrap = $("#toasts");
  if (!wrap) return;
  const el = document.createElement("div");
  el.className = "toast";
  el.innerHTML = `<div class="toast-title">${escapeHtml(title)}</div><div class="toast-body">${escapeHtml(body)}</div>`;
  wrap.appendChild(el);
  setTimeout(() => {
    el.style.opacity = "0";
    el.style.transform = "translateY(4px)";
    setTimeout(() => el.remove(), 240);
  }, ms);
}

/* ---------- State ---------- */

function defaultState() {
  return {
    version: APP_VERSION,
    meta: {
      projectName: "",
      region: "",
      owner: "",
      currency: "AUD",
    },
    settings: {
      scenario: "baseline",
      discountRate: 7,
      horizon: 10,
      startYear: 2026,
      inflation: 0,
    },
    baseline: {
      area: 0,
      yield: 0,
      price: 0,
      adoption: 100,
    },
    costs: [],
    benefits: [],
    lastSensitivity: null,
    ui: {
      tab: "intro",
      sidebarCollapsed: false,
    },
  };
}

let state = defaultState();

function loadState() {
  const raw = localStorage.getItem(APP_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return;
    state = mergeState(defaultState(), parsed);
  } catch {
    // ignore
  }
}

function mergeState(base, incoming) {
  const out = structuredClone(base);
  const assign = (obj, src) => {
    if (!src || typeof src !== "object") return;
    for (const k of Object.keys(src)) {
      const v = src[k];
      if (v && typeof v === "object" && !Array.isArray(v)) {
        if (!obj[k] || typeof obj[k] !== "object" || Array.isArray(obj[k])) obj[k] = {};
        assign(obj[k], v);
      } else {
        obj[k] = v;
      }
    }
  };
  assign(out, incoming);
  if (!Array.isArray(out.costs)) out.costs = [];
  if (!Array.isArray(out.benefits)) out.benefits = [];
  return out;
}

function saveState() {
  localStorage.setItem(APP_KEY, JSON.stringify(state));
}

/* ---------- Tabs ---------- */

function setActiveTab(tabKey) {
  const nav = $("#navTabs");
  const panels = $("#panels");
  if (!nav || !panels) return;

  const key = tabKey || "intro";
  state.ui.tab = key;

  $$(".nav-item", nav).forEach((btn) => {
    const is = btn.dataset.tab === key;
    btn.classList.toggle("is-active", is);
    btn.setAttribute("aria-selected", is ? "true" : "false");
  });

  $$(".panel", panels).forEach((p) => {
    const is = p.dataset.panel === key;
    p.classList.toggle("is-active", is);
  });

  const crumb = $("#crumbCurrent");
  if (crumb) {
    const name = key.charAt(0).toUpperCase() + key.slice(1);
    crumb.textContent = name === "Appendix" ? "Technical Appendix" : name;
  }

  if (key === "results") {
    renderAllResults();
  } else if (key === "sensitivity") {
    renderSensitivityFromState();
  }

  saveState();
}

/* ---------- Year handling ---------- */

function normaliseYearToAbs(inputYear) {
  const y = safeNum(inputYear, NaN);
  const startYear = safeNum(state.settings.startYear, 2026);
  const horizon = clamp(safeNum(state.settings.horizon, 10), 1, 50);

  // If looks like an offset (0..horizon-1), convert
  if (Number.isFinite(y) && y >= 0 && y <= horizon - 1) return startYear + Math.floor(y);

  // Else treat as absolute year
  if (Number.isFinite(y)) return Math.floor(y);

  return startYear;
}

function absYearToIndex(absYear) {
  const startYear = safeNum(state.settings.startYear, 2026);
  return Math.floor(absYear) - startYear;
}

/* ---------- Streams + Metrics ---------- */

function computeStreams(customState = null) {
  const s = customState || state;

  const currency = s.meta.currency || "AUD";
  const startYear = clamp(safeNum(s.settings.startYear, 2026), 1990, 2100);
  const horizon = clamp(safeNum(s.settings.horizon, 10), 1, 50);
  const r = clamp(safeNum(s.settings.discountRate, 7), 0, 30) / 100;
  const infl = safeNum(s.settings.inflation, 0) / 100;

  const adoption = clamp(safeNum(s.baseline.adoption, 100), 0, 100) / 100;

  const years = [];
  const benefits = new Array(horizon).fill(0);
  const costs = new Array(horizon).fill(0);

  for (let t = 0; t < horizon; t++) {
    years.push(startYear + t);
  }

  function addLines(lines, targetArray) {
    for (const line of lines) {
      const amt = safeNum(line.amount, 0);
      const dur = clamp(safeNum(line.duration, 1), 1, 1000);
      const absStart = normaliseYearToAbsWithState(line.year, s);
      const startIdx = absYearToIndexWithState(absStart, s);
      for (let k = 0; k < dur; k++) {
        const idx = startIdx + k;
        if (idx < 0 || idx >= horizon) continue;
        // Inflate amounts over time (optional)
        const escalator = infl !== 0 ? Math.pow(1 + infl, idx) : 1;
        targetArray[idx] += amt * escalator;
      }
    }
  }

  addLines(s.benefits || [], benefits);
  addLines(s.costs || [], costs);

  // Apply adoption scaling to both streams
  for (let t = 0; t < horizon; t++) {
    benefits[t] *= adoption;
    costs[t] *= adoption;
  }

  const net = benefits.map((b, i) => b - costs[i]);

  const discountFactor = [];
  const discountedNet = [];
  const discountedBenefits = [];
  const discountedCosts = [];

  for (let t = 0; t < horizon; t++) {
    const df = 1 / Math.pow(1 + r, t);
    discountFactor.push(df);
    discountedBenefits.push(benefits[t] * df);
    discountedCosts.push(costs[t] * df);
    discountedNet.push(net[t] * df);
  }

  const pvBenefits = discountedBenefits.reduce((a, b) => a + b, 0);
  const pvCosts = discountedCosts.reduce((a, b) => a + b, 0);
  const npv = discountedNet.reduce((a, b) => a + b, 0);
  const bcr = pvCosts > 0 ? pvBenefits / pvCosts : null;
  const roi = pvCosts > 0 ? (pvBenefits - pvCosts) / pvCosts : null;

  // Payback (discounted cumulative net)
  let cum = 0;
  let paybackIdx = null;
  for (let t = 0; t < horizon; t++) {
    cum += discountedNet[t];
    if (cum >= 0) {
      paybackIdx = t;
      break;
    }
  }

  return {
    currency,
    startYear,
    horizon,
    r,
    infl,
    adoption,
    years,
    benefits,
    costs,
    net,
    discountFactor,
    discountedNet,
    pvBenefits,
    pvCosts,
    npv,
    bcr,
    roi,
    paybackIdx,
  };
}

function normaliseYearToAbsWithState(inputYear, s) {
  const y = safeNum(inputYear, NaN);
  const startYear = safeNum(s.settings.startYear, 2026);
  const horizon = clamp(safeNum(s.settings.horizon, 10), 1, 50);

  if (Number.isFinite(y) && y >= 0 && y <= horizon - 1) return startYear + Math.floor(y);
  if (Number.isFinite(y)) return Math.floor(y);
  return startYear;
}

function absYearToIndexWithState(absYear, s) {
  const startYear = safeNum(s.settings.startYear, 2026);
  return Math.floor(absYear) - startYear;
}

/* ---------- Rendering ---------- */

function syncHeaderPills() {
  const scenLabel = ({
    baseline: "Baseline",
    optimistic: "Optimistic",
    pessimistic: "Pessimistic",
  })[state.settings.scenario] || "Baseline";

  $("#pillScenario").textContent = `Scenario: ${scenLabel}`;
  $("#pillCurrency").textContent = state.meta.currency || "AUD";
  $("#currencySuffix").textContent = state.meta.currency || "AUD";
}

function renderInputsToUI() {
  // Meta
  $("#projName").value = state.meta.projectName || "";
  $("#projRegion").value = state.meta.region || "";
  $("#projOwner").value = state.meta.owner || "";
  $("#projCurrency").value = state.meta.currency || "AUD";

  // Settings
  $("#selScenario").value = state.settings.scenario || "baseline";
  $("#selDiscount").value = safeNum(state.settings.discountRate, 7);
  $("#selHorizon").value = safeNum(state.settings.horizon, 10);

  $("#inpDiscount").value = safeNum(state.settings.discountRate, 7);
  $("#inpHorizon").value = safeNum(state.settings.horizon, 10);
  $("#inpStartYear").value = safeNum(state.settings.startYear, 2026);
  $("#inpInflation").value = safeNum(state.settings.inflation, 0);

  // Baseline
  $("#inpArea").value = safeNum(state.baseline.area, 0);
  $("#inpYield").value = safeNum(state.baseline.yield, 0);
  $("#inpPrice").value = safeNum(state.baseline.price, 0);
  $("#inpAdoption").value = safeNum(state.baseline.adoption, 100);

  syncHeaderPills();
}

function pullMetaFromUI() {
  state.meta.projectName = $("#projName").value.trim();
  state.meta.region = $("#projRegion").value.trim();
  state.meta.owner = $("#projOwner").value.trim();
  state.meta.currency = $("#projCurrency").value || "AUD";
  syncHeaderPills();
  saveState();
}

function pullInputsFromUI() {
  state.settings.discountRate = clamp(safeNum($("#inpDiscount").value, 7), 0, 30);
  state.settings.horizon = clamp(safeNum($("#inpHorizon").value, 10), 1, 50);
  state.settings.startYear = clamp(safeNum($("#inpStartYear").value, 2026), 1990, 2100);
  state.settings.inflation = clamp(safeNum($("#inpInflation").value, 0), -5, 30);

  state.baseline.area = Math.max(0, safeNum($("#inpArea").value, 0));
  state.baseline.yield = Math.max(0, safeNum($("#inpYield").value, 0));
  state.baseline.price = Math.max(0, safeNum($("#inpPrice").value, 0));
  state.baseline.adoption = clamp(safeNum($("#inpAdoption").value, 100), 0, 100);

  // Keep sidebar quick settings aligned
  $("#selDiscount").value = state.settings.discountRate;
  $("#selHorizon").value = state.settings.horizon;

  saveState();
  syncHeaderPills();
}

function applyQuickSettings() {
  state.settings.scenario = $("#selScenario").value || "baseline";
  state.settings.discountRate = clamp(safeNum($("#selDiscount").value, 7), 0, 30);
  state.settings.horizon = clamp(safeNum($("#selHorizon").value, 10), 1, 50);

  // Mirror into inputs tab
  $("#inpDiscount").value = state.settings.discountRate;
  $("#inpHorizon").value = state.settings.horizon;

  // Scenario presets (non-destructive multipliers)
  // Baseline: no change
  // Optimistic: benefits +10%, costs -5%
  // Pessimistic: benefits -10%, costs +5%
  // Implemented at calculation time by adjusting streams rather than mutating stored lines.
  saveState();
  syncHeaderPills();
  renderAllResults();
}

function scenarioAdjustStreams(streams) {
  const scen = state.settings.scenario || "baseline";
  let bMul = 1;
  let cMul = 1;
  if (scen === "optimistic") { bMul = 1.10; cMul = 0.95; }
  if (scen === "pessimistic") { bMul = 0.90; cMul = 1.05; }

  const out = structuredClone(streams);
  out.benefits = out.benefits.map((v) => v * bMul);
  out.costs = out.costs.map((v) => v * cMul);
  out.net = out.benefits.map((v, i) => v - out.costs[i]);

  const r = out.r;
  out.discountFactor = out.discountFactor.map((_, t) => 1 / Math.pow(1 + r, t));
  out.discountedNet = out.net.map((v, t) => v * out.discountFactor[t]);
  out.pvBenefits = out.benefits.reduce((acc, v, t) => acc + v * out.discountFactor[t], 0);
  out.pvCosts = out.costs.reduce((acc, v, t) => acc + v * out.discountFactor[t], 0);
  out.npv = out.discountedNet.reduce((a, b) => a + b, 0);
  out.bcr = out.pvCosts > 0 ? out.pvBenefits / out.pvCosts : null;
  out.roi = out.pvCosts > 0 ? (out.pvBenefits - out.pvCosts) / out.pvCosts : null;

  let cum = 0;
  out.paybackIdx = null;
  for (let t = 0; t < out.horizon; t++) {
    cum += out.discountedNet[t];
    if (cum >= 0) { out.paybackIdx = t; break; }
  }
  return out;
}

function renderKpis(streams) {
  const c = state.meta.currency || "AUD";

  $("#kpiNPV").textContent = formatCurrency(streams.npv, c);
  $("#kpiBCR").textContent = streams.bcr == null ? "—" : formatNumber(streams.bcr, 2);
  $("#kpiROI").textContent = streams.roi == null ? "—" : `${formatNumber(streams.roi * 100, 1)}%`;
  $("#kpiPayback").textContent =
    streams.paybackIdx == null ? "—" : `${streams.paybackIdx + 1}`;

  $("#sumBenefits").textContent = formatCurrency(streams.pvBenefits, c);
  $("#sumCosts").textContent = formatCurrency(streams.pvCosts, c);
  $("#sumNPV").textContent = formatCurrency(streams.npv, c);
  $("#sumBCR").textContent = streams.bcr == null ? "—" : formatNumber(streams.bcr, 2);
  $("#sumROI").textContent = streams.roi == null ? "—" : `${formatNumber(streams.roi * 100, 1)}%`;
  $("#sumPayback").textContent =
    streams.paybackIdx == null ? "No payback within horizon" : `${streams.paybackIdx + 1} year(s)`;

  $("#kpiNPVNote").textContent = `Discounted at ${formatNumber(state.settings.discountRate, 1)}%`;
  $("#kpiPaybackNote").textContent = streams.paybackIdx == null ? "Within horizon" : "Years";
}

function renderStreamsTable(streams) {
  const tbody = $("#tblStreams tbody");
  tbody.innerHTML = "";

  const c = state.meta.currency || "AUD";

  for (let i = 0; i < streams.horizon; i++) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="num">${streams.years[i]}</td>
      <td class="num">${formatCurrency(streams.benefits[i], c)}</td>
      <td class="num">${formatCurrency(streams.costs[i], c)}</td>
      <td class="num">${formatCurrency(streams.net[i], c)}</td>
      <td class="num">${formatNumber(streams.discountFactor[i], 4)}</td>
      <td class="num">${formatCurrency(streams.discountedNet[i], c)}</td>
    `;
    tbody.appendChild(tr);
  }
}

function polylinePath(xs, ys) {
  if (xs.length === 0) return "";
  const pts = xs.map((x, i) => `${x},${ys[i]}`).join(" ");
  return pts;
}

function renderLineChart(container, series, labels, opts = {}) {
  const w = 1000;
  const h = 240;
  const pad = 30;

  const xN = series[0].values.length;
  const xMin = 0;
  const xMax = Math.max(1, xN - 1);

  // Compute y-range across all series
  let yMin = Infinity, yMax = -Infinity;
  for (const s of series) {
    for (const v of s.values) {
      if (!Number.isFinite(v)) continue;
      yMin = Math.min(yMin, v);
      yMax = Math.max(yMax, v);
    }
  }
  if (!Number.isFinite(yMin) || !Number.isFinite(yMax)) { yMin = 0; yMax = 1; }
  if (yMin === yMax) { yMin -= 1; yMax += 1; }

  const xScale = (i) => pad + (i - xMin) * ((w - 2 * pad) / (xMax - xMin || 1));
  const yScale = (v) => pad + (yMax - v) * ((h - 2 * pad) / (yMax - yMin || 1));

  const gridY = 4;
  const gridLines = [];
  for (let g = 0; g <= gridY; g++) {
    const val = yMin + (g * (yMax - yMin)) / gridY;
    const y = yScale(val);
    gridLines.push({ y, val });
  }

  const paths = series.map((s, idx) => {
    const xs = s.values.map((_, i) => xScale(i));
    const ys = s.values.map((v) => yScale(v));
    return { ...s, xs, ys, pts: polylinePath(xs, ys), idx };
  });

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", opts.ariaLabel || "Chart");

  // background grid
  for (const gl of gridLines) {
    const line = document.createElementNS(svg.namespaceURI, "line");
    line.setAttribute("x1", pad);
    line.setAttribute("x2", w - pad);
    line.setAttribute("y1", gl.y);
    line.setAttribute("y2", gl.y);
    line.setAttribute("stroke", "rgba(255,255,255,0.10)");
    line.setAttribute("stroke-width", "1");
    svg.appendChild(line);

    const t = document.createElementNS(svg.namespaceURI, "text");
    t.setAttribute("x", 8);
    t.setAttribute("y", gl.y + 4);
    t.setAttribute("fill", "rgba(255,255,255,0.55)");
    t.setAttribute("font-size", "11");
    t.textContent = formatNumber(gl.val, 0);
    svg.appendChild(t);
  }

  // x labels (first, middle, last)
  const xLabelIdx = [0, Math.floor((xN - 1) / 2), xN - 1].filter((v, i, a) => a.indexOf(v) === i);
  for (const i of xLabelIdx) {
    const t = document.createElementNS(svg.namespaceURI, "text");
    t.setAttribute("x", xScale(i));
    t.setAttribute("y", h - 10);
    t.setAttribute("text-anchor", "middle");
    t.setAttribute("fill", "rgba(255,255,255,0.60)");
    t.setAttribute("font-size", "11");
    t.textContent = labels[i] ?? `${i}`;
    svg.appendChild(t);
  }

  // series lines
  const strokes = [
    "rgba(110,231,255,0.95)",
    "rgba(167,139,250,0.92)",
    "rgba(134,239,172,0.85)",
    "rgba(255,255,255,0.80)",
  ];

  for (const p of paths) {
    const pl = document.createElementNS(svg.namespaceURI, "polyline");
    pl.setAttribute("points", p.pts);
    pl.setAttribute("fill", "none");
    pl.setAttribute("stroke", strokes[p.idx % strokes.length]);
    pl.setAttribute("stroke-width", "2.5");
    pl.setAttribute("stroke-linecap", "round");
    pl.setAttribute("stroke-linejoin", "round");
    svg.appendChild(pl);
  }

  container.innerHTML = "";
  container.appendChild(svg);

  const legend = document.createElement("div");
  legend.className = "chart-legend";
  series.forEach((s, idx) => {
    const item = document.createElement("div");
    item.className = "legend-item";
    const sw = document.createElement("span");
    sw.className = "legend-swatch";
    sw.style.background = strokes[idx % strokes.length];
    item.appendChild(sw);
    const tx = document.createElement("span");
    tx.textContent = s.name;
    item.appendChild(tx);
    legend.appendChild(item);
  });
  container.appendChild(legend);
}

function renderStreamsChart(streams) {
  const el = $("#chartStreams");
  if (!el) return;

  renderLineChart(
    el,
    [
      { name: "Benefits", values: streams.benefits },
      { name: "Costs", values: streams.costs },
      { name: "Net", values: streams.net },
    ],
    streams.years.map(String),
    { ariaLabel: "Benefits, costs, and net over time" }
  );
}

function renderCostsTable() {
  const tbody = $("#tblCosts tbody");
  tbody.innerHTML = "";
  const lines = state.costs || [];
  for (const line of lines) tbody.appendChild(renderLineRow(line, "costs"));
}

function renderBenefitsTable() {
  const tbody = $("#tblBenefits tbody");
  tbody.innerHTML = "";
  const lines = state.benefits || [];
  for (const line of lines) tbody.appendChild(renderLineRow(line, "benefits"));
}

function renderLineRow(line, kind) {
  const tr = document.createElement("tr");
  tr.dataset.kind = kind;
  tr.dataset.id = line.id;

  const isCosts = kind === "costs";
  const types = isCosts
    ? ["Capex", "Opex", "Other"]
    : ["Market", "Cost saving", "Environmental", "Other"];

  tr.innerHTML = `
    <td>
      <input class="cell-input" data-field="name" type="text" value="${escapeHtml(line.name || "")}" placeholder="${isCosts ? "e.g., Upfront equipment" : "e.g., Yield improvement"}" />
    </td>
    <td>
      <select class="cell-select" data-field="type">
        ${types.map((t) => `<option value="${escapeHtml(t)}"${(line.type || types[0]) === t ? " selected" : ""}>${escapeHtml(t)}</option>`).join("")}
      </select>
    </td>
    <td class="num">
      <input class="cell-input num" data-field="year" type="number" step="1" value="${escapeHtml(line.year ?? 0)}" />
    </td>
    <td class="num">
      <input class="cell-input num" data-field="amount" type="number" step="0.01" value="${escapeHtml(line.amount ?? 0)}" />
    </td>
    <td class="num">
      <input class="cell-input num" data-field="duration" type="number" min="1" step="1" value="${escapeHtml(line.duration ?? 1)}" />
    </td>
    <td class="num">
      <button class="btn btn-ghost btn-sm" data-action="remove">Remove</button>
    </td>
  `;
  return tr;
}

function renderAllResults() {
  const baseStreams = computeStreams();
  const streams = scenarioAdjustStreams(baseStreams);
  renderKpis(streams);
  renderStreamsTable(streams);
  renderStreamsChart(streams);
  syncHeaderPills();
}

/* ---------- Exports ---------- */

function exportStateJson() {
  const payload = structuredClone(state);
  payload.exportedAt = new Date().toISOString();
  const fname = `farming_cba_${(state.meta.projectName || "project").replaceAll(/\s+/g, "_").slice(0, 24)}.json`;
  downloadText(fname, JSON.stringify(payload, null, 2), "application/json");
  toast("Exported", "Project JSON downloaded.");
}

function importStateJsonFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || "{}"));
      state = mergeState(defaultState(), parsed);
      saveState();
      applySidebarCollapsedState();
      renderInputsToUI();
      renderCostsTable();
      renderBenefitsTable();
      renderAllResults();
      renderSensitivityFromState();
      setActiveTab(state.ui.tab || "intro");
      toast("Imported", "Project loaded successfully.");
    } catch {
      toast("Import failed", "Invalid JSON file.");
    }
  };
  reader.readAsText(file);
}

function exportResultsCsv() {
  const baseStreams = computeStreams();
  const streams = scenarioAdjustStreams(baseStreams);
  const c = state.meta.currency || "AUD";

  const summaryLines = [
    ["Metric", "Value"],
    ["Currency", c],
    ["Scenario", state.settings.scenario],
    ["Discount rate (%)", formatNumber(state.settings.discountRate, 2)],
    ["Horizon (years)", String(state.settings.horizon)],
    ["Start year", String(state.settings.startYear)],
    ["Inflation (%)", formatNumber(state.settings.inflation, 2)],
    ["Adoption (%)", formatNumber(state.baseline.adoption, 0)],
    ["PV Benefits", round2(streams.pvBenefits)],
    ["PV Costs", round2(streams.pvCosts)],
    ["NPV", round2(streams.npv)],
    ["BCR", streams.bcr == null ? "" : round2(streams.bcr)],
    ["ROI", streams.roi == null ? "" : round2(streams.roi)],
    ["Payback (years)", streams.paybackIdx == null ? "" : String(streams.paybackIdx + 1)],
  ];

  const streamHeader = ["Year", "Benefits", "Costs", "Net", "DiscountFactor", "DiscountedNet"];
  const streamRows = streams.years.map((y, i) => [
    y,
    round2(streams.benefits[i]),
    round2(streams.costs[i]),
    round2(streams.net[i]),
    round2(streams.discountFactor[i]),
    round2(streams.discountedNet[i]),
  ]);

  const csv = []
    .concat([["Summary"], ...summaryLines, [""], ["Streams"], streamHeader, ...streamRows])
    .map((row) => row.map((x) => {
      const s = String(x ?? "");
      return /[,"\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
    }).join(","))
    .join("\n");

  const fname = `farming_cba_results_${Date.now()}.csv`;
  downloadText(fname, csv, "text/csv");
  toast("Exported", "Results CSV downloaded.");
}

function exportSensitivityCsv() {
  const sens = state.lastSensitivity;
  if (!sens || !Array.isArray(sens.rows) || sens.rows.length === 0) {
    toast("Nothing to export", "Run sensitivity first.");
    return;
  }

  const header = ["Value", "NPV", "BCR", "ROI"];
  const rows = sens.rows.map((r) => [r.value, round2(r.npv), r.bcr == null ? "" : round2(r.bcr), r.roi == null ? "" : round2(r.roi)]);
  const csv = [header, ...rows]
    .map((row) => row.map((x) => {
      const s = String(x ?? "");
      return /[,"\n]/.test(s) ? `"${s.replaceAll('"', '""')}"` : s;
    }).join(","))
    .join("\n");

  const fname = `farming_cba_sensitivity_${sens.param}_${Date.now()}.csv`;
  downloadText(fname, csv, "text/csv");
  toast("Exported", "Sensitivity CSV downloaded.");
}

/* ---------- Sensitivity ---------- */

function setSensitivityDefaultsForParam(param) {
  const low = $("#sensLow");
  const high = $("#sensHigh");
  const steps = $("#sensSteps");
  const help = $("#sensHelp");

  if (param === "discount") {
    low.value = "5";
    high.value = "10";
    steps.value = "11";
    help.textContent = "Discount rate is interpreted as a percentage (e.g., 5..10).";
    return;
  }

  if (param === "adoption") {
    low.value = "60";
    high.value = "100";
    steps.value = "9";
    help.textContent = "Adoption is interpreted as a percentage (0..100).";
    return;
  }

  if (param === "benefits" || param === "costs") {
    low.value = "80";
    high.value = "120";
    steps.value = "9";
    help.textContent = "Scale is interpreted as percent of baseline (e.g., 80..120 means -20%..+20%).";
    return;
  }
}

function linspace(a, b, n) {
  const out = [];
  if (n <= 1) return [a];
  const step = (b - a) / (n - 1);
  for (let i = 0; i < n; i++) out.push(a + i * step);
  return out;
}

function runSensitivity() {
  const param = $("#sensParam").value || "discount";
  const low = safeNum($("#sensLow").value, 0);
  const high = safeNum($("#sensHigh").value, 0);
  const steps = clamp(safeNum($("#sensSteps").value, 11), 3, 50);

  const values = linspace(low, high, steps).map((v) => round2(v));
  const rows = [];

  for (const v of values) {
    const s = structuredClone(state);

    if (param === "discount") {
      s.settings.discountRate = clamp(v, 0, 30);
    } else if (param === "adoption") {
      s.baseline.adoption = clamp(v, 0, 100);
    } else if (param === "benefits") {
      const mul = v / 100;
      s.benefits = (s.benefits || []).map((x) => ({ ...x, amount: safeNum(x.amount, 0) * mul }));
    } else if (param === "costs") {
      const mul = v / 100;
      s.costs = (s.costs || []).map((x) => ({ ...x, amount: safeNum(x.amount, 0) * mul }));
    }

    const base = computeStreams(s);
    const adj = scenarioAdjustStreams(base);

    rows.push({
      value: v,
      npv: adj.npv,
      bcr: adj.bcr,
      roi: adj.roi,
    });
  }

  state.lastSensitivity = { param, low, high, steps, rows };
  saveState();

  renderSensitivityFromState();
  toast("Sensitivity complete", `Computed ${rows.length} points.`);
  setStatus("Sensitivity updated");
}

function renderSensitivityFromState() {
  const tbody = $("#tblSensitivity tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const sens = state.lastSensitivity;
  const c = state.meta.currency || "AUD";
  if (!sens || !Array.isArray(sens.rows) || sens.rows.length === 0) {
    $("#chartSensitivity").innerHTML = `<div class="tiny muted">Run sensitivity to see results.</div>`;
    return;
  }

  for (const r of sens.rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="num">${formatNumber(r.value, 2)}</td>
      <td class="num">${formatCurrency(r.npv, c)}</td>
      <td class="num">${r.bcr == null ? "—" : formatNumber(r.bcr, 2)}</td>
      <td class="num">${r.roi == null ? "—" : `${formatNumber(r.roi * 100, 1)}%`}</td>
    `;
    tbody.appendChild(tr);
  }

  // Chart: value (x) vs NPV (y)
  const labels = sens.rows.map((r) => String(r.value));
  const npvVals = sens.rows.map((r) => r.npv);

  const el = $("#chartSensitivity");
  renderLineChart(
    el,
    [{ name: "NPV", values: npvVals }],
    labels,
    { ariaLabel: "Sensitivity of NPV to chosen parameter" }
  );
}

/* ---------- Line CRUD ---------- */

function addLine(kind) {
  const isCosts = kind === "costs";
  const line = {
    id: uid(isCosts ? "cost" : "ben"),
    name: "",
    type: isCosts ? "Capex" : "Market",
    year: 0,        // can be offset or absolute
    amount: 0,
    duration: 1,
  };
  state[kind].push(line);
  saveState();
  if (isCosts) renderCostsTable(); else renderBenefitsTable();
  renderAllResults();
  setStatus("Line added");
}

function clearLines(kind) {
  const isCosts = kind === "costs";
  Modal.open({
    title: `Clear ${isCosts ? "costs" : "benefits"}?`,
    bodyHtml: `<p class="p">This will remove all ${isCosts ? "cost" : "benefit"} lines. This cannot be undone.</p>`,
    okText: "Clear",
    cancelText: "Cancel",
    danger: true,
    onOkCb: () => {
      state[kind] = [];
      saveState();
      if (isCosts) renderCostsTable(); else renderBenefitsTable();
      renderAllResults();
      toast("Cleared", `${isCosts ? "Costs" : "Benefits"} cleared.`);
    },
  });
}

function updateLine(kind, id, field, value) {
  const arr = state[kind];
  const idx = arr.findIndex((x) => x.id === id);
  if (idx === -1) return;

  const line = arr[idx];
  if (field === "name" || field === "type") {
    line[field] = String(value ?? "");
  } else if (field === "year") {
    line.year = safeNum(value, 0);
  } else if (field === "amount") {
    line.amount = safeNum(value, 0);
  } else if (field === "duration") {
    line.duration = clamp(safeNum(value, 1), 1, 1000);
  }
  saveState();
  renderAllResults();
}

function removeLine(kind, id) {
  const arr = state[kind];
  const idx = arr.findIndex((x) => x.id === id);
  if (idx === -1) return;
  arr.splice(idx, 1);
  saveState();
  if (kind === "costs") renderCostsTable(); else renderBenefitsTable();
  renderAllResults();
}

/* ---------- Sidebar ---------- */

function applySidebarCollapsedState() {
  const sidebar = $("#sidebar");
  const btn = $("#btnNav");
  if (!sidebar || !btn) return;
  sidebar.classList.toggle("is-collapsed", !!state.ui.sidebarCollapsed);
  btn.setAttribute("aria-expanded", state.ui.sidebarCollapsed ? "false" : "true");
}

function toggleSidebar() {
  state.ui.sidebarCollapsed = !state.ui.sidebarCollapsed;
  applySidebarCollapsedState();
  saveState();
}

/* ---------- Help / Tour ---------- */

function openHelp() {
  Modal.open({
    title: "Help",
    bodyHtml: `
      <p class="p"><strong>Tabs</strong> are on the left. Start at Inputs, then add Costs and Benefits.</p>
      <p class="p"><strong>Year</strong> in cost/benefit tables can be either an offset (0..horizon-1) or an absolute year (e.g., 2027).</p>
      <p class="p"><strong>Duration</strong> is the number of years the annual amount repeats starting at the given year.</p>
      <p class="p"><strong>Scenarios</strong> apply simple multipliers (Optimistic: benefits +10%, costs −5%; Pessimistic: benefits −10%, costs +5%) without changing your stored lines.</p>
      <p class="p"><strong>Export</strong> downloads the project as JSON. Import restores it.</p>
    `,
    okText: "Close",
    showCancel: false,
  });
}

function openTour() {
  Modal.open({
    title: "Guided tour",
    bodyHtml: `
      <p class="p">1) Open <strong>Inputs</strong>, set start year, horizon, discount rate, and adoption, then click <strong>Save inputs</strong>.</p>
      <p class="p">2) Open <strong>Costs</strong>, click <strong>Add cost</strong>, enter an amount and duration. Repeat for each cost item.</p>
      <p class="p">3) Open <strong>Benefits</strong>, click <strong>Add benefit</strong>, enter annual benefits and durations.</p>
      <p class="p">4) Open <strong>Results</strong> to see KPIs, streams, and export a CSV table.</p>
      <p class="p">5) Open <strong>Sensitivity</strong> to stress-test discount rate, adoption, or scale benefits/costs.</p>
    `,
    okText: "Got it",
    showCancel: false,
  });
}

/* ---------- Example ---------- */

function loadExample() {
  state.meta.projectName = "Example: Practice change";
  state.meta.region = "NSW";
  state.meta.owner = "Farm business";
  state.meta.currency = "AUD";

  state.settings.startYear = 2026;
  state.settings.horizon = 10;
  state.settings.discountRate = 7;
  state.settings.inflation = 0;
  state.settings.scenario = "baseline";

  state.baseline.area = 500;
  state.baseline.yield = 3.2;
  state.baseline.price = 350;
  state.baseline.adoption = 100;

  state.costs = [
    { id: uid("cost"), name: "Upfront equipment", type: "Capex", year: 0, amount: 75000, duration: 1 },
    { id: uid("cost"), name: "Training & implementation", type: "Opex", year: 0, amount: 12000, duration: 2 },
    { id: uid("cost"), name: "Ongoing maintenance", type: "Opex", year: 1, amount: 6000, duration: 9 },
  ];

  state.benefits = [
    { id: uid("ben"), name: "Yield uplift", type: "Market", year: 1, amount: 45000, duration: 9 },
    { id: uid("ben"), name: "Input cost savings", type: "Cost saving", year: 1, amount: 15000, duration: 9 },
    { id: uid("ben"), name: "Environmental co-benefit (monetised)", type: "Environmental", year: 2, amount: 8000, duration: 8 },
  ];

  saveState();
  renderInputsToUI();
  renderCostsTable();
  renderBenefitsTable();
  renderAllResults();
  renderSensitivityFromState();
  setActiveTab("results");
  toast("Example loaded", "Example inputs, costs, and benefits inserted.");
}

/* ---------- Wiring ---------- */

function wireTabs() {
  $("#navTabs").addEventListener("click", (e) => {
    const btn = e.target.closest(".nav-item");
    if (!btn) return;
    const tab = btn.dataset.tab;
    if (!tab) return;
    setActiveTab(tab);
  });
}

function wireQuickSettings() {
  $("#selScenario").addEventListener("change", applyQuickSettings);
  $("#selDiscount").addEventListener("input", applyQuickSettings);
  $("#selHorizon").addEventListener("input", applyQuickSettings);
}

function wireMetaInputs() {
  $("#btnSaveMeta").addEventListener("click", () => {
    pullMetaFromUI();
    renderAllResults();
    toast("Saved", "Project metadata saved.");
    setStatus("Saved");
  });

  $("#projCurrency").addEventListener("change", () => {
    pullMetaFromUI();
    renderAllResults();
  });

  $("#btnLoadExampleIntro").addEventListener("click", loadExample);
}

function wireInputsTab() {
  $("#btnSaveInputs").addEventListener("click", () => {
    pullInputsFromUI();
    renderAllResults();
    toast("Saved", "Inputs saved and results updated.");
    setStatus("Saved");
  });

  $("#btnLoadExample").addEventListener("click", loadExample);

  // live preview updates
  ["inpDiscount", "inpHorizon", "inpStartYear", "inpInflation", "inpAdoption"].forEach((id) => {
    $(`#${id}`).addEventListener("input", () => {
      pullInputsFromUI();
      renderAllResults();
    });
  });
}

function wireTables() {
  $("#tblCosts").addEventListener("input", (e) => {
    const el = e.target;
    const tr = el.closest("tr");
    if (!tr) return;
    const kind = tr.dataset.kind;
    const id = tr.dataset.id;
    const field = el.dataset.field;
    if (!kind || !id || !field) return;
    updateLine(kind, id, field, el.value);
  });

  $("#tblBenefits").addEventListener("input", (e) => {
    const el = e.target;
    const tr = el.closest("tr");
    if (!tr) return;
    const kind = tr.dataset.kind;
    const id = tr.dataset.id;
    const field = el.dataset.field;
    if (!kind || !id || !field) return;
    updateLine(kind, id, field, el.value);
  });

  $("#tblCosts").addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;
    if (btn.dataset.action !== "remove") return;
    const tr = btn.closest("tr");
    if (!tr) return;
    removeLine("costs", tr.dataset.id);
  });

  $("#tblBenefits").addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;
    if (btn.dataset.action !== "remove") return;
    const tr = btn.closest("tr");
    if (!tr) return;
    removeLine("benefits", tr.dataset.id);
  });
}

function wireActions() {
  $("#btnNav").addEventListener("click", toggleSidebar);

  $("#btnAddCost").addEventListener("click", () => addLine("costs"));
  $("#btnClearCosts").addEventListener("click", () => clearLines("costs"));

  $("#btnAddBenefit").addEventListener("click", () => addLine("benefits"));
  $("#btnClearBenefits").addEventListener("click", () => clearLines("benefits"));

  $("#btnRecalc").addEventListener("click", () => {
    renderAllResults();
    toast("Updated", "Results recalculated.");
    setStatus("Updated");
  });

  $("#btnExport").addEventListener("click", exportStateJson);
  $("#btnExportResults").addEventListener("click", exportResultsCsv);
  $("#btnExportSensitivity").addEventListener("click", exportSensitivityCsv);

  $("#btnImport").addEventListener("click", () => $("#fileImport").click());
  $("#fileImport").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!file) return;
    importStateJsonFile(file);
  });

  $("#btnReset").addEventListener("click", () => {
    Modal.open({
      title: "Reset tool?",
      bodyHtml: `<p class="p">This will clear the saved project from this browser and restore defaults.</p>`,
      okText: "Reset",
      cancelText: "Cancel",
      danger: true,
      onOkCb: () => {
        localStorage.removeItem(APP_KEY);
        state = defaultState();
        saveState();
        applySidebarCollapsedState();
        renderInputsToUI();
        renderCostsTable();
        renderBenefitsTable();
        renderAllResults();
        renderSensitivityFromState();
        setActiveTab("intro");
        toast("Reset", "Defaults restored.");
      },
    });
  });

  $("#btnHelp").addEventListener("click", openHelp);
  $("#btnTour").addEventListener("click", openTour);
}

function wireSensitivity() {
  $("#sensParam").addEventListener("change", (e) => {
    setSensitivityDefaultsForParam(e.target.value);
  });

  $("#btnRunSensitivity").addEventListener("click", runSensitivity);
}

function wireGlobal() {
  window.addEventListener("beforeunload", () => saveState());

  // close sidebar on mobile when clicking outside
  document.addEventListener("click", (e) => {
    const sidebar = $("#sidebar");
    if (!sidebar) return;
    const isMobile = window.matchMedia("(max-width: 920px)").matches;
    if (!isMobile) return;
    if (!state.ui.sidebarCollapsed && !sidebar.contains(e.target) && !$("#btnNav").contains(e.target)) {
      state.ui.sidebarCollapsed = true;
      applySidebarCollapsedState();
      saveState();
    }
  });
}

/* ---------- Init ---------- */

function init() {
  Modal.wire();

  loadState();

  // If first load, ensure ui defaults are set
  state.ui = state.ui || { tab: "intro", sidebarCollapsed: false };
  if (!state.ui.tab) state.ui.tab = "intro";

  $("#appVersion").textContent = APP_VERSION;

  applySidebarCollapsedState();
  renderInputsToUI();
  renderCostsTable();
  renderBenefitsTable();
  renderAllResults();
  renderSensitivityFromState();

  wireTabs();
  wireQuickSettings();
  wireMetaInputs();
  wireInputsTab();
  wireTables();
  wireActions();
  wireSensitivity();
  wireGlobal();

  setActiveTab(state.ui.tab);

  // Ensure sensitivity UI defaults match selected param
  setSensitivityDefaultsForParam($("#sensParam").value || "discount");

  setStatus("Loaded");
}

init();
