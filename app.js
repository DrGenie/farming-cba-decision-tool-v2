// app.js
/* Fully functional Farming CBA Decision Aid (vanilla JS + SheetJS):
   Existing features kept and extended (not replaced):
   - Active tabs (Import, Mapping, Assumptions, Results, Cashflows, Sensitivity, Export, Help)
   - Robust Excel read (auto header detection + dedupe headers)
   - Flexible mapping for treatment/yield/cost
   - Per-treatment overrides: adoption, yield multiplier, cost multiplier, cost timing
   - Full CBA: PV benefits, PV costs, NPV, BCR, ROI, ranking; control included in vertical table
   - Cashflow tables + simple canvas charts
   - Sensitivity (one-way) + XLSX export
   - Clean Excel export (multiple sheets)

   NEW (improvements):
   - Introduction tab support (optional in HTML): renders audience-friendly overview text blocks
   - Additional benefits/costs (adders) configurable and applied to cashflows + PV
   - Simulations (Monte Carlo) for uncertainty on key drivers, outputs distributions + probs
   - Copilot/AI integration tab (no external calls): builds a structured JSON prompt for policy brief + tables, copy/download
   - Info tooltips (data-tooltip attribute): accessible hover/focus tooltips
   - Better accessibility: keyboard navigation for tabs (arrow keys, Home/End), reduced JS hard-fail if new tab DOM not present

   ASCII-only.
*/

/* =========================
   Embedded Lockhart sample
   ========================= */
const EMBEDDED_LOCKHART_ROWS = [
  {"Plot":1,"Rep":1,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.029229293617021,"Protein":23.2,"|":17945.488764568763},
  {"Plot":2,"Rep":1,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.539273035489362,"Protein":23.6,"|":24884.884058984914},
  {"Plot":3,"Rep":1,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.54757540287234,"Protein":23.7,"|":18463.88888888889},
  {"Plot":4,"Rep":1,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.37207183687234,"Protein":24.7,"|":1012.5633802816902},
  {"Plot":5,"Rep":1,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.667165176319149,"Protein":23.9,"|":912.1951219512195},
  {"Plot":6,"Rep":1,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.199593337872341,"Protein":23.3,"|":976.8292682926829},
  {"Plot":7,"Rep":1,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.614808249489362,"Protein":23.4,"|":1033.6585365853657},
  {"Plot":8,"Rep":1,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.958392845957447,"Protein":23.1,"|":782.9268292682927},
  {"Plot":9,"Rep":1,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.280082541361702,"Protein":23.7,"|":1059.349593495935},
  {"Plot":10,"Rep":1,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.839204480808511,"Protein":23.8,"|":908.9430894308944},
  {"Plot":11,"Rep":1,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.142857142857143,"Protein":23.3,"|":790.650406504065},
  {"Plot":12,"Rep":1,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.421177186276596,"Protein":23.4,"|":748.780487804878},
  {"Plot":13,"Rep":1,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.35011252212766,"Protein":23.7,"|":1024.5934959349595},
  {"Plot":14,"Rep":1,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.35667234112766,"Protein":23.2,"|":923.1707317073171},
  {"Plot":15,"Rep":1,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.269241385957447,"Protein":23.9,"|":1062.439024390244},
  {"Plot":16,"Rep":1,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.626622926382979,"Protein":23.5,"|":694.6341463414634},
  {"Plot":17,"Rep":2,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":8.16927218306383,"Protein":23.5,"|":1059.349593495935},
  {"Plot":18,"Rep":2,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.467401828978723,"Protein":23.3,"|":748.780487804878},
  {"Plot":19,"Rep":2,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.751529506382978,"Protein":23.1,"|":923.1707317073171},
  {"Plot":20,"Rep":2,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.771685019574468,"Protein":23.7,"|":1024.5934959349595},
  {"Plot":21,"Rep":2,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.373296617021276,"Protein":23.5,"|":976.8292682926829},
  {"Plot":22,"Rep":2,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.794740369574468,"Protein":23.4,"|":908.9430894308944},
  {"Plot":23,"Rep":2,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.830657711489362,"Protein":23.2,"|":790.650406504065},
  {"Plot":24,"Rep":2,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":7.082754630638297,"Protein":23.3,"|":1033.6585365853657},
  {"Plot":25,"Rep":2,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.722016223404255,"Protein":23.6,"|":1062.439024390244},
  {"Plot":26,"Rep":2,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.417663186382979,"Protein":23.2,"|":912.1951219512195},
  {"Plot":27,"Rep":2,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":7.268999694468085,"Protein":23.1,"|":782.9268292682927},
  {"Plot":28,"Rep":2,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.640707778085106,"Protein":23.5,"|":1012.5633802816902},
  {"Plot":29,"Rep":2,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.601502030425532,"Protein":23.1,"|":18463.88888888889},
  {"Plot":30,"Rep":2,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.842683569361702,"Protein":23.5,"|":17787.777777777777},
  {"Plot":31,"Rep":2,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.006188607446808,"Protein":23.2,"|":17848.538011695906},
  {"Plot":32,"Rep":2,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.482897900425532,"Protein":23.2,"|":694.6341463414634},
  {"Plot":33,"Rep":3,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.322008111702128,"Protein":23.4,"|":1024.5934959349595},
  {"Plot":34,"Rep":3,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.048594062978723,"Protein":23.5,"|":923.1707317073171},
  {"Plot":35,"Rep":3,"Trt":1,"Amendment":"Lime only","Practice Change":"Crop 1","Yield t/ha":7.468294498723404,"Protein":23.5,"|":790.650406504065},
  {"Plot":36,"Rep":3,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":7.208078594468085,"Protein":23.5,"|":748.780487804878},
  {"Plot":37,"Rep":3,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.82302918893617,"Protein":23.5,"|":1059.349593495935},
  {"Plot":38,"Rep":3,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.866323363404255,"Protein":23.7,"|":1033.6585365853657},
  {"Plot":39,"Rep":3,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice Change":"Crop 1","Yield t/ha":7.176539666595744,"Protein":23.4,"|":976.8292682926829},
  {"Plot":40,"Rep":3,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":7.064219107659574,"Protein":23.4,"|":782.9268292682927},
  {"Plot":41,"Rep":3,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice Change":"Crop 1","Yield t/ha":6.519162115744681,"Protein":23.4,"|":17945.488764568763},
  {"Plot":42,"Rep":3,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.05956678787234,"Protein":23.7,"|":1062.439024390244},
  {"Plot":43,"Rep":3,"Trt":6,"Amendment":"Deep OM (CP1)","Practice Change":"Crop 1","Yield t/ha":6.396029749361702,"Protein":23.7,"|":17787.777777777777},
  {"Plot":44,"Rep":3,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice Change":"Crop 1","Yield t/ha":6.268999694468085,"Protein":23.6,"|":1012.5633802816902},
  {"Plot":45,"Rep":3,"Trt":2,"Amendment":"Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.120229763404255,"Protein":23.7,"|":912.1951219512195},
  {"Plot":46,"Rep":3,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice Change":"Crop 1","Yield t/ha":7.462494905957446,"Protein":23.4,"|":908.9430894308944},
  {"Plot":47,"Rep":3,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice Change":"Crop 1","Yield t/ha":6.435858070212766,"Protein":23.6,"|":18463.88888888889},
  {"Plot":48,"Rep":3,"Trt":0,"Amendment":"Control","Practice Change":"Crop 1","Yield t/ha":7.269686206808511,"Protein":23.5,"|":694.6341463414634}
];

/* =========================
   State
   ========================= */
const state = {
  source: null, // {kind:'excel'|'embedded', name:string}
  raw: [],
  columns: [],
  map: {
    treatment: null,
    baseline: null,
    yield: null,
    cost: null,
    optional1: null
  },
  assumptions: {
    areaHa: 100,
    horizonYears: 10,
    discountRatePct: 7,
    priceYear1: 450,
    priceGrowthPct: 0,
    yieldScale: 1,
    costScale: 1,

    benefitStartYear: 1,
    effectDurationYears: 5,
    decay: "linear", // none|linear|exp
    halfLifeYears: 2,

    costTimingDefault: "y1_only" // y1_only|y0_only|annual_duration|annual_horizon|split_50_50
  },

  // treatmentOverrides[name] = {enabled, adoption, yieldMult, costMult, costTiming}
  treatmentOverrides: {},

  // Additional benefits/costs (adders) applied on top of incremental yield/cost.
  // Each adder contributes to benefits or costs streams depending on `kind`.
  // amountBasis:
  //   - "per_ha": amount * (area*adoption) unless per_ha view
  //   - "total": amount is already total (whole-farm)
  //   - "per_t": amount * incremental_yield(t/ha) * (area*adoption) * benefitFactor * (optionally use baseline yield? we use incremental yield)
  // timing:
  //   - "y0_only" | "y1_only" | "annual_horizon" | "annual_duration" | "custom"
  // appliesTo: "all" or array of treatment names
  adders: {
    items: [
      // Seed with examples (editable in UI; safe if UI not present)
      // {id:"B1", kind:"benefit", label:"Soil carbon credit", amount:15, amountBasis:"per_ha", timing:"annual_duration", startYear:1, endYear:5, growthPct:0, appliesTo:"all", enabled:false, notes:"Example only"},
      // {id:"C1", kind:"cost", label:"Extra labour/monitoring", amount:8, amountBasis:"per_ha", timing:"annual_duration", startYear:1, endYear:5, growthPct:0, appliesTo:"all", enabled:false, notes:"Example only"}
    ]
  },

  // Copilot / AI prompt builder preferences
  copilot: {
    audience: "mixed", // farmer|policy|research|mixed
    tone: "plain",     // plain|formal|technical
    length: "long",    // short|medium|long
    includeTables: true,
    includeAssumptions: true,
    includeCashflows: false,
    includeSensitivity: true,
    includeSimulations: true
  },

  // Monte Carlo simulation settings (uncertainty ranges around current inputs)
  sim: {
    enabled: true,
    draws: 500,
    seed: 12345,
    // ranges are symmetric +/- around base
    pricePct: 0.15,       // +/- 15%
    yieldMultPct: 0.15,   // +/- 15% around treatment yieldMult (applies to all treatments)
    costMultPct: 0.15,    // +/- 15% around treatment costMult (applies to all treatments)
    discountPts: 0.02,    // +/- 2 percentage points
    durationYears: 2,     // +/- 2 years
    distribution: "triangular" // triangular|uniform
  },

  ui: {
    activeTab: "import",
    rankBy: "npv",
    view: "whole_farm",
    cashflowTreatment: null
  },

  cache: {
    trialSummary: null,        // per-treatment means and deltas
    cbaPerTreatment: null,     // per-treatment cba results + cashflows
    lastSensitivity: null,
    lastSim: null              // latest sim summary payload
  }
};

/* =========================
   DOM helpers
   ========================= */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));
const byId = (id) => document.getElementById(id);
function on(el, ev, fn){ if (el) el.addEventListener(ev, fn); }

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

function toNum(v){
  if (v === null || v === undefined) return NaN;
  if (typeof v === "number") return v;
  if (typeof v === "string"){
    const s = v.replace(/[, ]+/g, "").replace(/^\$/,"");
    const n = Number(s);
    return Number.isFinite(n) ? n : NaN;
  }
  return NaN;
}

function fmtNumber(x, digits=2){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(undefined, {maximumFractionDigits: digits, minimumFractionDigits: digits});
}

function fmtInt(x){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(undefined, {maximumFractionDigits: 0});
}

function fmtMoney(x, digits=0){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  const abs = Math.abs(n);
  const s = abs.toLocaleString(undefined, {minimumFractionDigits: digits, maximumFractionDigits: digits});
  return (n < 0 ? "-$" : "$") + s;
}

function toast(msg, ms=2200){
  const t = byId("toast");
  const tt = byId("toastText");
  if (!t || !tt) return;
  tt.textContent = msg;
  t.hidden = false;
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { t.hidden = true; }, ms);
}

/* =========================
   Tabs + accessibility
   ========================= */
function setActiveTab(tab){
  state.ui.activeTab = tab;
  $$(".tab").forEach(btn => {
    const onTab = btn.dataset.tab === tab;
    btn.classList.toggle("is-active", onTab);
    btn.setAttribute("aria-selected", onTab ? "true" : "false");
    if (onTab) btn.focus({preventScroll:true});
  });
  $$(".panel").forEach(p => p.classList.toggle("is-active", p.dataset.panel === tab));
}

function bindTabKeyboardNav(){
  const tablist = $(".tabs");
  if (!tablist) return;
  on(tablist, "keydown", (e) => {
    const tabs = $$(".tab");
    if (!tabs.length) return;
    const current = document.activeElement;
    const idx = tabs.indexOf(current);
    if (idx < 0) return;

    if (e.key === "ArrowRight" || e.key === "ArrowLeft" || e.key === "Home" || e.key === "End"){
      e.preventDefault();
      let next = idx;
      if (e.key === "ArrowRight") next = (idx + 1) % tabs.length;
      if (e.key === "ArrowLeft") next = (idx - 1 + tabs.length) % tabs.length;
      if (e.key === "Home") next = 0;
      if (e.key === "End") next = tabs.length - 1;
      tabs[next].focus();
    }
    if (e.key === "Enter" || e.key === " "){
      const btn = document.activeElement;
      if (btn && btn.classList.contains("tab")) btn.click();
    }
  });
}

/* =========================
   Tooltips (data-tooltip)
   Accessible: hover/focus shows tooltip; Esc closes
   ========================= */
const tooltip = {
  el: null,
  activeFor: null
};

function ensureTooltipEl(){
  if (tooltip.el) return tooltip.el;
  const div = document.createElement("div");
  div.id = "appTooltip";
  div.setAttribute("role","tooltip");
  div.style.position = "fixed";
  div.style.zIndex = "9999";
  div.style.maxWidth = "360px";
  div.style.padding = "10px 12px";
  div.style.borderRadius = "14px";
  div.style.border = "1px solid rgba(255,255,255,0.14)";
  div.style.background = "rgba(11,16,32,0.92)";
  div.style.backdropFilter = "blur(12px)";
  div.style.color = "rgba(255,255,255,0.88)";
  div.style.fontSize = "13px";
  div.style.lineHeight = "1.35";
  div.style.boxShadow = "0 10px 30px rgba(0,0,0,0.35)";
  div.style.display = "none";
  document.body.appendChild(div);
  tooltip.el = div;
  return div;
}

function showTooltipFor(el){
  const text = el.getAttribute("data-tooltip");
  if (!text) return;
  const t = ensureTooltipEl();
  tooltip.activeFor = el;
  t.textContent = text;
  t.style.display = "block";

  const r = el.getBoundingClientRect();
  const pad = 10;
  const tw = Math.min(360, window.innerWidth - 2*pad);
  t.style.maxWidth = tw + "px";

  // initial position (below, aligned left)
  let left = Math.min(window.innerWidth - pad - tw, Math.max(pad, r.left));
  let top = r.bottom + 10;
  t.style.left = left + "px";
  t.style.top = top + "px";

  // if offscreen bottom, move above
  const tr = t.getBoundingClientRect();
  if (tr.bottom > window.innerHeight - pad){
    top = Math.max(pad, r.top - tr.height - 10);
    t.style.top = top + "px";
  }

  // aria-describedby
  const tipId = "appTooltip";
  el.setAttribute("aria-describedby", tipId);
}

function hideTooltip(){
  const t = tooltip.el;
  if (t) t.style.display = "none";
  if (tooltip.activeFor){
    tooltip.activeFor.removeAttribute("aria-describedby");
  }
  tooltip.activeFor = null;
}

function bindTooltips(){
  // Attach to any element with data-tooltip (existing + future dynamic)
  const attach = (el) => {
    if (el._tooltipBound) return;
    el._tooltipBound = true;
    on(el, "mouseenter", () => showTooltipFor(el));
    on(el, "mouseleave", hideTooltip);
    on(el, "focus", () => showTooltipFor(el));
    on(el, "blur", hideTooltip);
  };
  $$("[data-tooltip]").forEach(attach);

  on(document, "keydown", (e) => {
    if (e.key === "Escape") hideTooltip();
  });
}

/* =========================
   Excel import
   ========================= */
function isEmptyCell(v){
  return v === null || v === undefined || (typeof v === "string" && v.trim() === "");
}

function densestRowIndex(rows2d, maxScan=60){
  const lim = Math.min(rows2d.length, maxScan);
  let best = {idx: 0, score: -1};
  for (let i=0; i<lim; i++){
    const row = rows2d[i] || [];
    const score = row.reduce((acc, v) => acc + (isEmptyCell(v) ? 0 : 1), 0);
    if (score > best.score){
      best = {idx: i, score};
    }
  }
  return best.idx;
}

function dedupeHeaders(headers){
  const seen = new Map();
  return headers.map(h => {
    const name = (h === null || h === undefined) ? "" : String(h).trim();
    const base = name === "" ? "col" : name;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}__${count}`;
  });
}

function sheetToRows2D(wb){
  const firstSheet = wb.SheetNames[0];
  const ws = wb.Sheets[firstSheet];
  return XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null});
}

function rows2DToObjects(rows2d){
  const headerIdx = densestRowIndex(rows2d, 60);
  const headersRaw = rows2d[headerIdx] || [];
  const headers = dedupeHeaders(headersRaw);

  const out = [];
  for (let r=headerIdx+1; r<rows2d.length; r++){
    const row = rows2d[r] || [];
    const nonEmpty = row.reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0);
    if (nonEmpty === 0) continue;

    const obj = {};
    for (let c=0; c<headers.length; c++){
      obj[headers[c]] = (c < row.length) ? row[c] : null;
    }

    // Drop very sparse/footer-ish rows
    const filled = Object.values(obj).reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0);
    if (filled < Math.max(3, Math.floor(headers.length*0.03))) continue;

    out.push(obj);
  }

  return {rows: out, columns: headers};
}

function unique(arr){
  const s = new Set(arr);
  return Array.from(s);
}

function percentile(arr, p){
  if (!arr.length) return NaN;
  const a = arr.slice().sort((x,y)=>x-y);
  const idx = (p/100)*(a.length-1);
  const lo = Math.floor(idx);
  const hi = Math.ceil(idx);
  if (lo === hi) return a[lo];
  const t = idx - lo;
  return a[lo]*(1-t) + a[hi]*t;
}

function detectDefaultMappings(rows, columns){
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : []);

  const treatmentKey =
    (cols.includes("Amendment") ? "Amendment" : null) ||
    (cols.includes("Treatment") ? "Treatment" : null) ||
    (cols.includes("Trt") ? "Trt" : null) ||
    (cols.find(c => /treat|amend/i.test(c)) || null);

  let yieldKey = null;
  if (cols.includes("Yield t/ha")) yieldKey = "Yield t/ha";
  else yieldKey = cols.find(c => /^yield\b/i.test(c)) || null;

  let costKey = cols.includes("|") ? "|" : null;
  if (!costKey){
    const candidates = cols.filter(c => /cost|\$|\/ha|per ha|ha\b/i.test(c));
    const scored = candidates.map(c => {
      const nums = rows.map(r => toNum(r[c])).filter(v => Number.isFinite(v));
      const coverage = nums.length / Math.max(1, rows.length);
      const p50 = nums.length ? percentile(nums, 50) : -Infinity;
      return {c, coverage, p50};
    }).filter(x => x.coverage >= 0.5);
    scored.sort((a,b) => (b.p50 - a.p50) || (b.coverage - a.coverage));
    costKey = scored[0]?.c || null;
  }

  const optional1 =
    (cols.includes("Protein") ? "Protein" : null) ||
    (cols.find(c => /protein|quality/i.test(c)) || null);

  let baseline = null;
  if (treatmentKey){
    const groups = unique(rows.map(r => String(r[treatmentKey] ?? "").trim()).filter(Boolean));
    baseline = groups.find(g => g.toLowerCase() === "control") || groups[0] || null;
  }

  return {treatmentKey, yieldKey, costKey, optional1, baseline};
}

async function importExcelArrayBuffer(buf, name="workbook.xlsx"){
  if (!window.XLSX) throw new Error("XLSX library missing. Check index.html include.");
  const wb = XLSX.read(buf, {type:"array"});
  const rows2d = sheetToRows2D(wb);
  const {rows, columns} = rows2DToObjects(rows2d);

  state.raw = rows;
  state.columns = columns;
  state.source = {kind:"excel", name};

  const d = detectDefaultMappings(rows, columns);
  state.map.treatment = d.treatmentKey;
  state.map.yield = d.yieldKey;
  state.map.cost = d.costKey;
  state.map.optional1 = d.optional1;
  state.map.baseline = d.baseline;

  initTreatmentOverrides();
  invalidateCache();
  toast(`Imported: ${name}`);
  renderAll();
}

function loadEmbedded(){
  state.raw = EMBEDDED_LOCKHART_ROWS.map(r => ({...r}));
  state.columns = dedupeHeaders(Object.keys(state.raw[0] || {}));
  state.source = {kind:"embedded", name:"Lockhart (embedded sample)"};

  const d = detectDefaultMappings(state.raw, state.columns);
  state.map.treatment = d.treatmentKey;
  state.map.yield = d.yieldKey;
  state.map.cost = d.costKey;
  state.map.optional1 = d.optional1;
  state.map.baseline = d.baseline;

  initTreatmentOverrides();
  invalidateCache();
  toast("Loaded embedded sample");
  renderAll();
}

async function tryLoadBundled(){
  const candidates = ["./data.xlsx", "./Data for Lockhart-FA-031225 (1).xlsx"];
  for (const url of candidates){
    try{
      const res = await fetch(url);
      if (!res.ok) continue;
      const buf = await res.arrayBuffer();
      await importExcelArrayBuffer(buf, url.split("/").pop());
      return true;
    } catch(e){
      // keep trying
    }
  }
  return false;
}

/* =========================
   Treatment overrides
   ========================= */
function initTreatmentOverrides(){
  state.treatmentOverrides = {};
  const tkey = state.map.treatment;
  if (!tkey || !state.raw.length) return;

  const names = unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean)).sort((a,b)=>a.localeCompare(b));
  for (const n of names){
    state.treatmentOverrides[n] = {
      enabled: true,
      adoption: 1,
      yieldMult: 1,
      costMult: 1,
      costTiming: "default"
    };
  }
  if (state.map.baseline && state.treatmentOverrides[state.map.baseline]){
    state.treatmentOverrides[state.map.baseline].enabled = true;
    state.treatmentOverrides[state.map.baseline].adoption = 1;
  }
}

function invalidateCache(){
  state.cache.trialSummary = null;
  state.cache.cbaPerTreatment = null;
}

/* =========================
   Core computations
   ========================= */
function mean(arr){
  const xs = (arr || []).filter(v => Number.isFinite(v));
  if (!xs.length) return NaN;
  return xs.reduce((a,b)=>a+b,0)/xs.length;
}

function clamp01(x){
  if (!Number.isFinite(x)) return 0;
  return Math.max(0, Math.min(1, x));
}

function getTreatmentNames(){
  const tkey = state.map.treatment;
  if (!tkey || !state.raw.length) return [];
  return unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean)).sort((a,b)=>a.localeCompare(b));
}

function buildTrialSummary(){
  if (state.cache.trialSummary) return state.cache.trialSummary;

  const rows = state.raw;
  const tkey = state.map.treatment;
  const ykey = state.map.yield;
  const ckey = state.map.cost;
  if (!rows.length || !tkey || !ykey || !ckey) return [];

  const yScale = Number(state.assumptions.yieldScale) || 1;
  const cScale = Number(state.assumptions.costScale) || 1;

  const buckets = new Map();
  for (const r of rows){
    const name = String(r[tkey] ?? "").trim();
    if (!name) continue;
    const y = toNum(r[ykey]) * yScale;
    const c = toNum(r[ckey]) * cScale;
    if (!buckets.has(name)) buckets.set(name, {name, n:0, yields:[], costs:[]});
    const b = buckets.get(name);
    b.n += 1;
    if (Number.isFinite(y)) b.yields.push(y);
    if (Number.isFinite(c)) b.costs.push(c);
  }

  const out = [];
  for (const b of buckets.values()){
    out.push({
      name: b.name,
      n: b.n,
      yield_mean: mean(b.yields),
      cost_mean: mean(b.costs)
    });
  }
  out.sort((a,b)=>a.name.localeCompare(b.name));

  const baseName = state.map.baseline;
  const base = out.find(x => x.name === baseName) || null;

  for (const r of out){
    r.delta_yield = base ? (r.yield_mean - base.yield_mean) : NaN;
    r.delta_cost  = base ? (r.cost_mean  - base.cost_mean)  : NaN;
  }

  state.cache.trialSummary = out;
  return out;
}

function priceForYear(year){
  const p1 = Number(state.assumptions.priceYear1) || 0;
  const g = (Number(state.assumptions.priceGrowthPct) || 0) / 100;
  return p1 * Math.pow(1 + g, Math.max(0, year - 1));
}

function discountFactor(t){
  const r = (Number(state.assumptions.discountRatePct) || 0) / 100;
  return 1 / Math.pow(1 + r, t);
}

function benefitFactorByYear(year){
  const start = Math.max(1, Math.floor(Number(state.assumptions.benefitStartYear) || 1));
  const D = Math.max(1, Math.floor(Number(state.assumptions.effectDurationYears) || 1));
  const decay = state.assumptions.decay || "linear";

  const idx = year - start;
  if (idx < 0) return 0;
  if (idx >= D) return 0;

  if (decay === "none") return 1;
  if (decay === "linear"){
    if (D === 1) return 1;
    return Math.max(0, 1 - (idx / (D - 1)));
  }
  if (decay === "exp"){
    const hl = Math.max(0.1, Number(state.assumptions.halfLifeYears) || 2);
    const lambda = Math.log(2) / hl;
    return Math.exp(-lambda * idx);
  }
  return 1;
}

function resolveCostTiming(name){
  const ov = state.treatmentOverrides[name];
  if (!ov) return state.assumptions.costTimingDefault;
  const ct = ov.costTiming || "default";
  if (ct === "default") return state.assumptions.costTimingDefault;
  return ct;
}

/* =========================
   Additional benefits/costs (adders)
   ========================= */
function normAdder(a){
  // Ensure all fields exist and are safe
  const out = {
    id: String(a?.id ?? "").trim() || ("A" + Math.random().toString(16).slice(2,8).toUpperCase()),
    kind: (a?.kind === "cost") ? "cost" : "benefit",
    label: String(a?.label ?? "Untitled").trim(),
    enabled: (a?.enabled === undefined) ? true : !!a.enabled,
    amount: Number(a?.amount ?? 0),
    amountBasis: (a?.amountBasis === "total" || a?.amountBasis === "per_t") ? a.amountBasis : "per_ha",
    timing: (a?.timing || "annual_duration"),
    startYear: Math.max(0, Math.floor(Number(a?.startYear ?? 1))),
    endYear: Math.max(0, Math.floor(Number(a?.endYear ?? state.assumptions.effectDurationYears))),
    growthPct: Number(a?.growthPct ?? 0),
    appliesTo: a?.appliesTo ?? "all", // "all" or array of names
    notes: String(a?.notes ?? "").trim()
  };
  return out;
}

function adderAppliesTo(adder, treatmentName){
  if (!adder.enabled) return false;
  const at = adder.appliesTo;
  if (at === "all") return true;
  if (Array.isArray(at)) return at.includes(treatmentName);
  return true;
}

function applyAdderToStreams(adder, treatmentName, streams){
  // streams: {benefits[], costs[], scale, perHaView, incYieldPerHa, horizonT}
  if (!adderAppliesTo(adder, treatmentName)) return;

  const T = streams.horizonT;
  const g = (Number(adder.growthPct) || 0) / 100;

  // Determine timing years
  let years = [];
  const timing = adder.timing;

  if (timing === "y0_only") years = [0];
  else if (timing === "y1_only") years = [1];
  else if (timing === "annual_horizon") years = Array.from({length: T}, (_,i)=>i+1);
  else if (timing === "annual_duration") {
    // use benefitFactor window; but for adders we use start/end explicit with defaults
    const start = Math.max(0, Math.floor(Number(adder.startYear)));
    const end = Math.max(start, Math.floor(Number(adder.endYear)));
    years = [];
    for (let y=Math.max(0,start); y<=Math.min(T,end); y++){
      years.push(y);
    }
    // If user set startYear=1, it will include 1..end
  } else if (timing === "custom") {
    const start = Math.max(0, Math.floor(Number(adder.startYear)));
    const end = Math.max(start, Math.floor(Number(adder.endYear)));
    years = [];
    for (let y=start; y<=Math.min(T,end); y++) years.push(y);
  } else {
    // fallback
    years = Array.from({length: T}, (_,i)=>i+1);
  }

  // Amount scaling
  for (const y of years){
    let amt = Number(adder.amount) || 0;

    // growth applies from year 1 onward for recurring; if year 0, apply no growth
    if (y >= 1) amt = amt * Math.pow(1 + g, y - 1);

    let v = 0;
    if (adder.amountBasis === "per_ha"){
      v = amt * streams.scale;
    } else if (adder.amountBasis === "total"){
      v = (streams.perHaView ? (amt / Math.max(1, (streams.scale || 1))) : amt);
      // If in per-ha view, best-effort convert: total/scale. If scale=1 (per-ha already), stays total.
    } else if (adder.amountBasis === "per_t"){
      // apply to incremental yield (t/ha) with the same benefitFactor window used for yield effects
      // if user wants per_t independent of benefitFactor they can set annual_horizon and amountBasis per_ha/total
      const incY = streams.incYieldPerHa;
      const bf = (y === 0) ? 0 : benefitFactorByYear(y);
      v = amt * incY * bf * streams.scale;
    }

    if (adder.kind === "benefit") streams.benefits[y] += v;
    else streams.costs[y] += v;
  }
}

/* =========================
   Cashflows + CBA
   ========================= */
function buildCashflowsForTreatment(treatmentName){
  const summary = buildTrialSummary();
  const baseName = state.map.baseline;
  const base = summary.find(x => x.name === baseName);
  const tr = summary.find(x => x.name === treatmentName);
  if (!base || !tr) return null;

  const ov = state.treatmentOverrides[treatmentName] || {enabled:true, adoption:1, yieldMult:1, costMult:1};
  const enabled = !!ov.enabled;
  const adoption = clamp01(Number(ov.adoption));
  const yMult = Number(ov.yieldMult) || 1;
  const cMult = Number(ov.costMult) || 1;

  const area = Number(state.assumptions.areaHa) || 0;
  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));

  // Incremental deltas (per ha) vs baseline
  const dy_perha_raw = (tr.delta_yield || 0);
  const dc_perha_raw = (tr.delta_cost  || 0);

  const dy_perha = dy_perha_raw * yMult;
  const dc_perha = dc_perha_raw * cMult;

  const perHaView = (state.ui.view === "per_ha");
  const scale = perHaView ? 1 : (area * adoption);

  const benefits = Array(T+1).fill(0); // year 0..T
  const costs = Array(T+1).fill(0);

  // Base yield-driven benefits
  for (let y=1; y<=T; y++){
    const bf = benefitFactorByYear(y);
    const p = priceForYear(y);
    const b = dy_perha * p * bf;
    benefits[y] = b * scale;
  }

  // Incremental costs (from dataset delta costs) with timing
  const timing = resolveCostTiming(treatmentName);
  if (timing === "y1_only"){
    costs[1] = dc_perha * scale;
  } else if (timing === "y0_only"){
    costs[0] = dc_perha * scale;
  } else if (timing === "annual_duration"){
    for (let y=1; y<=T; y++){
      const bf = benefitFactorByYear(y);
      if (bf > 0) costs[y] = dc_perha * scale;
    }
  } else if (timing === "annual_horizon"){
    for (let y=1; y<=T; y++) costs[y] = dc_perha * scale;
  } else if (timing === "split_50_50"){
    costs[0] = 0.5 * dc_perha * scale;
    costs[1] += 0.5 * dc_perha * scale;
  } else {
    costs[1] = dc_perha * scale;
  }

  // Apply additional adders (benefits/costs) on top
  const streams = {
    benefits,
    costs,
    scale,
    perHaView,
    incYieldPerHa: dy_perha,
    horizonT: T
  };
  for (const a of state.adders.items.map(normAdder)){
    applyAdderToStreams(a, treatmentName, streams);
  }

  // If disabled, zero-out in whole-farm view
  if (!enabled && !perHaView){
    for (let y=0; y<=T; y++){
      benefits[y] = 0;
      costs[y] = 0;
    }
  }

  const net = benefits.map((b,i)=>b - costs[i]);

  // PV
  let pvB = 0, pvC = 0, npv = 0;
  for (let t=0; t<=T; t++){
    const df = discountFactor(t);
    pvB += benefits[t]*df;
    pvC += costs[t]*df;
    npv += net[t]*df;
  }
  const bcr = (pvC === 0) ? (pvB === 0 ? NaN : Infinity) : (pvB / pvC);
  const roi = (pvC === 0) ? NaN : (npv / pvC);

  return {
    name: treatmentName,
    baseline: baseName,
    enabled,
    adoption,
    yieldMult: yMult,
    costMult: cMult,
    timing,
    dy_perha,
    dc_perha,
    // adders are embedded inside benefits/costs already
    benefits,
    costs,
    net,
    pvBenefits: pvB,
    pvCosts: pvC,
    npv,
    bcr,
    roi
  };
}

function buildCbaAllTreatments(){
  if (state.cache.cbaPerTreatment) return state.cache.cbaPerTreatment;

  const names = getTreatmentNames();
  const base = state.map.baseline;
  if (!base || !names.length) return [];

  const ordered = [base, ...names.filter(n => n !== base)];
  const out = [];
  for (const n of ordered){
    const cf = buildCashflowsForTreatment(n);
    if (cf) out.push(cf);
  }
  state.cache.cbaPerTreatment = out;
  return out;
}

/* =========================
   Simulations (Monte Carlo)
   ========================= */
function mulberry32(seed){
  let a = seed >>> 0;
  return function(){
    a |= 0;
    a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function sampleTriangular(rng, min, mode, max){
  const u = rng();
  const c = (mode - min) / (max - min);
  if (u < c){
    return min + Math.sqrt(u * (max - min) * (mode - min));
  }
  return max - Math.sqrt((1 - u) * (max - min) * (max - mode));
}

function sampleUniform(rng, min, max){
  return min + (max - min) * rng();
}

function runMonteCarlo(){
  const names = getTreatmentNames();
  const base = state.map.baseline;
  if (!names.length || !base) return null;

  const draws = Math.max(50, Math.floor(Number(state.sim.draws) || 500));
  const rng = mulberry32(Number(state.sim.seed) || 12345);

  const baseAssump = JSON.parse(JSON.stringify(state.assumptions));
  const baseOv = JSON.parse(JSON.stringify(state.treatmentOverrides));

  const dist = state.sim.distribution || "triangular";
  const drawVal = (min, mode, max) => (dist === "uniform" ? sampleUniform(rng, min, max) : sampleTriangular(rng, min, mode, max));

  // storage: treatment -> array of metric values
  const store = {};
  for (const n of names) store[n] = {npv:[], bcr:[], roi:[]};
  const bestCount = {}; for (const n of names) bestCount[n] = 0;

  // Evaluate quickly by swapping assumptions/overrides in-place then restoring
  const savedAssump = state.assumptions;
  const savedOv = state.treatmentOverrides;

  try{
    for (let d=0; d<draws; d++){
      // sample drivers
      const price0 = baseAssump.priceYear1;
      const price = drawVal(price0*(1-state.sim.pricePct), price0, price0*(1+state.sim.pricePct));

      const disc0 = baseAssump.discountRatePct;
      const disc = drawVal(Math.max(0, disc0 - state.sim.discountPts*100), disc0, disc0 + state.sim.discountPts*100);

      const dur0 = baseAssump.effectDurationYears;
      const dur = Math.max(1, Math.round(drawVal(Math.max(1, dur0 - state.sim.durationYears), dur0, dur0 + state.sim.durationYears)));

      // treatment mult shocks applied multiplicatively around existing multipliers
      const yShock = drawVal(1 - state.sim.yieldMultPct, 1, 1 + state.sim.yieldMultPct);
      const cShock = drawVal(1 - state.sim.costMultPct, 1, 1 + state.sim.costMultPct);

      state.assumptions = JSON.parse(JSON.stringify(baseAssump));
      state.treatmentOverrides = JSON.parse(JSON.stringify(baseOv));

      state.assumptions.priceYear1 = price;
      state.assumptions.discountRatePct = disc;
      state.assumptions.effectDurationYears = dur;

      // Apply shocks to all treatments (baseline included, but baseline dy/dc is zero so it will not matter)
      for (const n of Object.keys(state.treatmentOverrides)){
        state.treatmentOverrides[n].yieldMult = (Number(baseOv[n]?.yieldMult) || 1) * yShock;
        state.treatmentOverrides[n].costMult = (Number(baseOv[n]?.costMult) || 1) * cShock;
      }

      invalidateCache();
      const cba = buildCbaAllTreatments();

      // track best by NPV (highest finite)
      let bestN = null;
      let bestV = -Infinity;

      for (const t of cba){
        const npv = Number(t.npv);
        const bcr = Number(t.bcr);
        const roi = Number(t.roi);

        if (Number.isFinite(npv)) store[t.name].npv.push(npv);
        if (Number.isFinite(bcr)) store[t.name].bcr.push(bcr);
        if (Number.isFinite(roi)) store[t.name].roi.push(roi);

        if (Number.isFinite(npv) && npv > bestV){
          bestV = npv;
          bestN = t.name;
        }
      }
      if (bestN) bestCount[bestN] += 1;
    }
  } finally {
    // restore
    state.assumptions = savedAssump;
    state.treatmentOverrides = savedOv;
    invalidateCache();
  }

  // summaries
  const summaryRows = [];
  for (const n of names){
    const xs = store[n].npv.slice().sort((a,b)=>a-b);
    const ys = store[n].bcr.slice().sort((a,b)=>a-b);
    const zs = store[n].roi.slice().sort((a,b)=>a-b);

    const p = (arr, q) => percentile(arr, q);

    const npvMean = xs.length ? xs.reduce((a,b)=>a+b,0)/xs.length : NaN;
    const bcrMean = ys.length ? ys.reduce((a,b)=>a+b,0)/ys.length : NaN;
    const roiMean = zs.length ? zs.reduce((a,b)=>a+b,0)/zs.length : NaN;

    const probNpvPos = xs.length ? (xs.filter(v=>v>0).length / xs.length) : NaN;
    const probBcrGt1 = ys.length ? (ys.filter(v=>v>1).length / ys.length) : NaN;
    const probBest = bestCount[n] / draws;

    summaryRows.push({
      treatment: n,
      isBaseline: (n === base),
      draws: draws,
      npv_mean: npvMean,
      npv_p10: p(xs,10),
      npv_p50: p(xs,50),
      npv_p90: p(xs,90),
      prob_npv_gt_0: probNpvPos,
      bcr_mean: bcrMean,
      bcr_p10: p(ys,10),
      bcr_p50: p(ys,50),
      bcr_p90: p(ys,90),
      prob_bcr_gt_1: probBcrGt1,
      roi_mean: roiMean,
      roi_p10: p(zs,10),
      roi_p50: p(zs,50),
      roi_p90: p(zs,90),
      prob_best_by_npv: probBest
    });
  }

  // rank by prob best
  summaryRows.sort((a,b) => (b.prob_best_by_npv - a.prob_best_by_npv) || (b.npv_p50 - a.npv_p50));

  const payload = {
    settings: JSON.parse(JSON.stringify(state.sim)),
    baseAssumptions: JSON.parse(JSON.stringify(state.assumptions)),
    summary: summaryRows
  };

  state.cache.lastSim = payload;
  return payload;
}

/* =========================
   Copilot / AI prompt builder
   ========================= */
function buildCopilotPayload(){
  const cba = buildCbaAllTreatments();
  const trial = buildTrialSummary();
  const sens = state.cache.lastSensitivity;
  const sim = state.cache.lastSim;

  const metric = state.ui.rankBy;
  const audience = state.copilot.audience;
  const tone = state.copilot.tone;
  const length = state.copilot.length;

  // Build vertical table structure explicitly for AI to reproduce in Word
  const scored = cba.map(x => ({name:x.name, v: (metric==="npv"?x.npv:(metric==="bcr"?x.bcr:x.roi))}))
    .sort((a,b)=> {
      const aBad = !Number.isFinite(a.v);
      const bBad = !Number.isFinite(b.v);
      if (aBad && bBad) return a.name.localeCompare(b.name);
      if (aBad) return 1;
      if (bBad) return -1;
      return (b.v - a.v);
    });
  const rankMap = {};
  for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

  const verticalTable = [
    {indicator:"PV benefits", values:Object.fromEntries(cba.map(t=>[t.name, t.pvBenefits]))},
    {indicator:"PV costs", values:Object.fromEntries(cba.map(t=>[t.name, t.pvCosts]))},
    {indicator:"NPV", values:Object.fromEntries(cba.map(t=>[t.name, t.npv]))},
    {indicator:"BCR", values:Object.fromEntries(cba.map(t=>[t.name, t.bcr]))},
    {indicator:"ROI", values:Object.fromEntries(cba.map(t=>[t.name, t.roi]))},
    {indicator:`Rank (by ${metric.toUpperCase()})`, values:Object.fromEntries(cba.map(t=>[t.name, rankMap[t.name]]))}
  ];

  // Include adders (additional benefits/costs) for transparency
  const adders = state.adders.items.map(normAdder);

  // Instructions (human-ready): no external calls; just a structured prompt
  const instructions = {
    task: "Write a decision-support policy brief and produce clean tables that can be pasted into Word.",
    audiences: {
      farmer: "Plain language, focus on what drives dollars and what can be changed on-farm. Avoid jargon.",
      policy: "Policy-relevant framing, implications, implementation considerations, and budget/benefit narrative.",
      research: "Transparent methods, assumptions, limitations, and reproducibility details.",
      mixed: "Layered narrative: farmer-facing summary first, then policy and technical appendices."
    },
    style: {
      tone,
      length,
      constraints: [
        "Do not tell the reader what to choose; present trade-offs.",
        "Explain NPV, PV benefits/costs, BCR, ROI in practical terms.",
        "Always compare treatments against the baseline/control and state which baseline is used.",
        "If negative NPV or low BCR occurs, suggest plausible improvement levers (cost reduction, yield gain, price, adoption, timing) without prescribing."
      ]
    },
    outputs: {
      mustInclude: [
        "Executive summary (short, non-technical)",
        "Methods and assumptions summary",
        "Results section interpreting vertical CBA table",
        "Drivers of results (yield delta, cost delta, additional adders, discounting, duration/decay)",
        "Uncertainty (sensitivity and simulation if provided)",
        "Implementation notes (practicalities, monitoring, data needs)",
        "Appendix: copy-ready tables"
      ],
      tables: [
        "Vertical CBA table (indicators as rows; treatments as columns; include baseline)",
        "Trial means and deltas vs baseline",
        "If simulation provided: NPV distribution summary (P10/P50/P90 and probabilities)"
      ]
    }
  };

  const payload = {
    meta: {
      generatedAt: new Date().toISOString(),
      dataset: state.source?.name || "unknown",
      baseline: state.map.baseline,
      view: state.ui.view,
      rankBy: metric
    },
    assumptions: state.copilot.includeAssumptions ? state.assumptions : undefined,
    mappings: state.map,
    adders,
    trialSummary: trial,
    cba: cba.map(t => ({
      name: t.name,
      pvBenefits: t.pvBenefits,
      pvCosts: t.pvCosts,
      npv: t.npv,
      bcr: t.bcr,
      roi: t.roi,
      dy_perha: t.dy_perha,
      dc_perha: t.dc_perha,
      adoption: t.adoption,
      enabled: t.enabled,
      costTiming: t.timing
    })),
    verticalTable,
    sensitivity: state.copilot.includeSensitivity ? sens : undefined,
    simulation: state.copilot.includeSimulations ? sim : undefined,
    cashflows: state.copilot.includeCashflows ? cba.map(t => ({name:t.name, benefits:t.benefits, costs:t.costs, net:t.net})) : undefined,
    instructions
  };

  return payload;
}

async function copyTextToClipboard(text){
  try{
    await navigator.clipboard.writeText(text);
    toast("Copied to clipboard");
  } catch(e){
    // fallback
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    ta.remove();
    toast("Copied to clipboard");
  }
}

/* =========================
   Rendering
   ========================= */
function renderAll(){
  renderDatasetPill();
  renderImportKPIs();
  renderColumns();
  renderPreviewTable();

  renderMappingSelectors();
  renderTreatmentsConfigTable();

  renderAssumptionsControls();
  renderSanity();

  renderResults();
  renderCashflows();
  renderSensitivity();
  renderHalfLifeVisibility();

  // New tabs (render only if DOM exists)
  renderIntro();
  renderAdders();
  renderSimulations();
  renderCopilot();

  // tooltips
  bindTooltips();
}

function renderDatasetPill(){
  const pill = byId("datasetPill");
  if (!pill) return;
  if (!state.raw.length){
    pill.textContent = "No data loaded";
    pill.style.borderColor = "rgba(255,255,255,0.12)";
    return;
  }
  const src = state.source?.name || "Dataset";
  pill.textContent = `${src} · ${state.raw.length} rows`;
  pill.style.borderColor = "rgba(125,211,252,0.35)";
}

function renderImportKPIs(){
  const rowsEl = byId("kpiRows");
  const trtEl = byId("kpiTreatments");
  const baseEl = byId("kpiBaseline");
  const keysEl = byId("kpiKeys");
  if (!rowsEl || !trtEl || !baseEl || !keysEl) return;

  rowsEl.textContent = fmtInt(state.raw.length);
  const tkey = state.map.treatment;
  const treatments = tkey ? getTreatmentNames() : [];
  trtEl.textContent = fmtInt(treatments.length);
  baseEl.textContent = state.map.baseline || "—";
  keysEl.textContent = (state.map.yield && state.map.cost) ? `${state.map.yield} | ${state.map.cost}` : "—";
}

function renderColumns(){
  const box = byId("columnsChips");
  if (!box) return;
  box.innerHTML = "";
  if (!state.columns.length){
    box.innerHTML = `<div class="muted">No columns yet.</div>`;
    return;
  }
  const frag = document.createDocumentFragment();
  for (const c of state.columns){
    const d = document.createElement("div");
    d.className = "chip";
    d.textContent = c;
    frag.appendChild(d);
  }
  box.appendChild(frag);
}

function renderPreviewTable(){
  const wrap = byId("previewTableWrap");
  const note = byId("previewNote");
  if (!wrap || !note) return;

  if (!state.raw.length){
    note.hidden = false;
    wrap.hidden = true;
    wrap.innerHTML = "";
    return;
  }
  note.hidden = true;
  wrap.hidden = false;

  const cols = (state.columns || Object.keys(state.raw[0] || {})).slice(0, 10);
  const rows = state.raw.slice(0, 8);

  let html = `<table><thead><tr>`;
  for (const c of cols) html += `<th>${escapeHtml(c)}</th>`;
  html += `</tr></thead><tbody>`;
  for (const r of rows){
    html += `<tr>`;
    for (const c of cols){
      const v = r[c];
      html += `<td class="${typeof v === "number" ? "mono" : ""}">${escapeHtml(v)}</td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

function fillSelect(sel, options, value, allowEmpty=false, emptyLabel="—"){
  if (!sel) return;
  sel.innerHTML = "";
  if (allowEmpty){
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = emptyLabel;
    sel.appendChild(opt0);
  }
  for (const o of options){
    const opt = document.createElement("option");
    opt.value = o;
    opt.textContent = o;
    if (o === value) opt.selected = true;
    sel.appendChild(opt);
  }
  sel.disabled = options.length === 0;
}

function renderMappingSelectors(){
  const cols = state.columns || [];
  fillSelect(byId("mapTreatment"), cols, state.map.treatment, false);
  fillSelect(byId("mapYield"), cols, state.map.yield, true, "—");
  fillSelect(byId("mapCost"), cols, state.map.cost, true, "—");
  fillSelect(byId("mapOptional1"), cols, state.map.optional1 || "", true, "—");

  const tnames = getTreatmentNames();
  fillSelect(byId("mapBaseline"), tnames, state.map.baseline, false);

  fillSelect(byId("cashflowTreatment"), tnames, state.ui.cashflowTreatment || state.map.baseline || "", false);
  fillSelect(byId("sensTreatment"), tnames, byId("sensTreatment")?.value || state.map.baseline || "", false);
}

function renderTreatmentsConfigTable(){
  const wrap = byId("treatmentsConfigWrap");
  if (!wrap) return;
  const tnames = getTreatmentNames();
  if (!tnames.length || !state.map.baseline){
    wrap.innerHTML = `<div class="muted">Load data and set a treatment mapping first.</div>`;
    return;
  }

  const costTimingOptions = [
    {v:"default", label:"Default (from Assumptions)"},
    {v:"y1_only", label:"Year 1 only"},
    {v:"y0_only", label:"Year 0 only"},
    {v:"annual_duration", label:"Annual during effect duration"},
    {v:"annual_horizon", label:"Annual for full horizon"},
    {v:"split_50_50", label:"Split 50% Year 0, 50% Year 1"}
  ];

  let html = `<table>
    <thead><tr>
      <th data-tooltip="Treatment label used to group rows and compute averages.">Treatment</th>
      <th data-tooltip="Disable removes whole-farm impacts for this treatment (per-ha view still shows per-ha).">Enable</th>
      <th data-tooltip="Share of farm area that adopts this option. Whole-farm totals scale by area x adoption.">Adoption (0-1)</th>
      <th data-tooltip="Multiplier applied to incremental yield difference vs baseline.">Yield mult</th>
      <th data-tooltip="Multiplier applied to incremental cost difference vs baseline.">Cost mult</th>
      <th data-tooltip="When incremental costs occur over time (setup vs annual).">Cost timing</th>
    </tr></thead>
    <tbody>`;

  for (const name of tnames){
    const ov = state.treatmentOverrides[name] || {enabled:true, adoption:1, yieldMult:1, costMult:1, costTiming:"default"};
    const isBase = name === state.map.baseline;
    const badge = isBase ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(name)}${badge}</td>
      <td class="mono">
        <input type="checkbox" data-ov="enabled" data-name="${escapeHtml(name)}" ${ov.enabled ? "checked":""} />
      </td>
      <td><input class="tinput" type="number" step="0.05" min="0" max="1" data-ov="adoption" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.adoption))}" /></td>
      <td><input class="tinput" type="number" step="0.05" data-ov="yieldMult" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.yieldMult))}" /></td>
      <td><input class="tinput" type="number" step="0.05" data-ov="costMult" data-name="${escapeHtml(name)}" value="${escapeHtml(String(ov.costMult))}" /></td>
      <td>
        <select class="tselect" data-ov="costTiming" data-name="${escapeHtml(name)}">
          ${costTimingOptions.map(o => `<option value="${o.v}" ${o.v===ov.costTiming ? "selected":""}>${escapeHtml(o.label)}</option>`).join("")}
        </select>
      </td>
    </tr>`;
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;

  wrap.querySelectorAll("input[data-ov], select[data-ov]").forEach(el => {
    on(el, "change", (e) => {
      const name = e.target.getAttribute("data-name");
      const key = e.target.getAttribute("data-ov");
      if (!state.treatmentOverrides[name]) state.treatmentOverrides[name] = {enabled:true, adoption:1, yieldMult:1, costMult:1, costTiming:"default"};

      if (e.target.type === "checkbox"){
        state.treatmentOverrides[name][key] = e.target.checked;
      } else if (key === "costTiming"){
        state.treatmentOverrides[name][key] = e.target.value;
      } else {
        state.treatmentOverrides[name][key] = Number(e.target.value);
      }

      if (name === state.map.baseline){
        state.treatmentOverrides[name].enabled = true;
      }

      invalidateCache();
      renderResults();
      renderCashflows();
      renderCopilot(); // prompt updates
    });
  });
}

function renderAssumptionsControls(){
  const a = state.assumptions;
  if (byId("assumpArea")) byId("assumpArea").value = a.areaHa;
  if (byId("assumpHorizon")) byId("assumpHorizon").value = a.horizonYears;
  if (byId("assumpDiscount")) byId("assumpDiscount").value = a.discountRatePct;
  if (byId("assumpPrice")) byId("assumpPrice").value = a.priceYear1;
  if (byId("assumpPriceGrowth")) byId("assumpPriceGrowth").value = a.priceGrowthPct;
  if (byId("assumpYieldScale")) byId("assumpYieldScale").value = String(a.yieldScale);
  if (byId("assumpCostScale")) byId("assumpCostScale").value = String(a.costScale);

  if (byId("assumpBenefitStart")) byId("assumpBenefitStart").value = a.benefitStartYear;
  if (byId("assumpDuration")) byId("assumpDuration").value = a.effectDurationYears;
  if (byId("assumpDecay")) byId("assumpDecay").value = a.decay;
  if (byId("assumpHalfLife")) byId("assumpHalfLife").value = a.halfLifeYears;
  if (byId("assumpCostTimingDefault")) byId("assumpCostTimingDefault").value = a.costTimingDefault;

  // Simulation controls (optional)
  if (byId("simDraws")) byId("simDraws").value = state.sim.draws;
  if (byId("simSeed")) byId("simSeed").value = state.sim.seed;
  if (byId("simPricePct")) byId("simPricePct").value = Math.round(state.sim.pricePct * 100);
  if (byId("simYieldMultPct")) byId("simYieldMultPct").value = Math.round(state.sim.yieldMultPct * 100);
  if (byId("simCostMultPct")) byId("simCostMultPct").value = Math.round(state.sim.costMultPct * 100);
  if (byId("simDiscountPts")) byId("simDiscountPts").value = Math.round(state.sim.discountPts * 1000) / 10; // in pp
  if (byId("simDurationYears")) byId("simDurationYears").value = state.sim.durationYears;
  if (byId("simDistribution")) byId("simDistribution").value = state.sim.distribution;
}

function renderHalfLifeVisibility(){
  const decay = state.assumptions.decay;
  const row = byId("halfLifeRow");
  if (!row) return;
  row.hidden = decay !== "exp";
}

function renderSanity(){
  const box = byId("sanityBox");
  if (!box) return;
  if (!state.raw.length || !state.map.yield || !state.map.cost){
    box.innerHTML = `<div class="muted">Load data and set yield/cost mapping to see diagnostics.</div>`;
    return;
  }

  const yScale = Number(state.assumptions.yieldScale) || 1;
  const cScale = Number(state.assumptions.costScale) || 1;

  const ys = state.raw.map(r => toNum(r[state.map.yield]) * yScale).filter(Number.isFinite);
  const cs = state.raw.map(r => toNum(r[state.map.cost]) * cScale).filter(Number.isFinite);

  const yP50 = percentile(ys, 50), yP5 = percentile(ys, 5), yP95 = percentile(ys, 95);
  const cP50 = percentile(cs, 50), cP5 = percentile(cs, 5), cP95 = percentile(cs, 95);

  const warnings = [];
  if (!Number.isFinite(yP50) || yP50 <= 0) warnings.push("Yield median is non-positive; check yield mapping/scaling.");
  if (!Number.isFinite(cP50)) warnings.push("Cost median invalid; check cost mapping/scaling.");
  if (Number.isFinite(yP95) && yP95 > 50) warnings.push("Yield P95 looks very high for t/ha; consider yield scale.");
  if (Number.isFinite(cP95) && cP95 > 100000) warnings.push("Cost P95 looks extremely high; confirm the cost column is correct.");

  const warnHtml = warnings.length ? `
    <div class="sanity-item">
      <div class="sanity-item__title">Warnings</div>
      <div class="sanity-item__text">${warnings.map(w => `• ${escapeHtml(w)}`).join("<br>")}</div>
    </div>` : "";

  box.innerHTML = `
    <div class="sanity-item">
      <div class="sanity-item__title">Detected mapping</div>
      <div class="sanity-item__text">
        Treatment: <code>${escapeHtml(state.map.treatment || "—")}</code><br>
        Baseline: <code>${escapeHtml(state.map.baseline || "—")}</code><br>
        Yield: <code>${escapeHtml(state.map.yield || "—")}</code><br>
        Cost: <code>${escapeHtml(state.map.cost || "—")}</code>
      </div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Yield distribution (scaled)</div>
      <div class="sanity-item__text">P5=${escapeHtml(fmtNumber(yP5,3))}, P50=${escapeHtml(fmtNumber(yP50,3))}, P95=${escapeHtml(fmtNumber(yP95,3))}</div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Cost distribution (scaled)</div>
      <div class="sanity-item__text">P5=${escapeHtml(fmtMoney(cP5,0))}, P50=${escapeHtml(fmtMoney(cP50,0))}, P95=${escapeHtml(fmtMoney(cP95,0))}</div>
    </div>

    ${warnHtml}
  `;
}

function renderTrialSummary(){
  const wrap = byId("trialSummaryWrap");
  if (!wrap) return;
  const sum = buildTrialSummary();
  if (!sum.length){
    wrap.innerHTML = `<div class="muted">Load data and confirm mappings.</div>`;
    return;
  }

  let html = `<table>
    <thead><tr>
      <th>Treatment</th>
      <th>N</th>
      <th>Yield mean (t/ha)</th>
      <th>Cost mean (/ha)</th>
      <th>Δ Yield vs baseline</th>
      <th>Δ Cost vs baseline</th>
    </tr></thead><tbody>`;

  for (const r of sum){
    const isBase = r.name === state.map.baseline;
    const badge = isBase ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(r.name)}${badge}</td>
      <td class="mono">${fmtInt(r.n)}</td>
      <td class="mono">${fmtNumber(r.yield_mean,3)}</td>
      <td class="mono">${fmtMoney(r.cost_mean,0)}</td>
      <td class="mono">${fmtNumber(r.delta_yield,3)}</td>
      <td class="mono">${fmtMoney(r.delta_cost,0)}</td>
    </tr>`;
  }
  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

function verticalCbaTable(cbaList){
  const names = cbaList.map(x => x.name);

  const rows = [
    {label:"PV benefits", key:"pvBenefits"},
    {label:"PV costs", key:"pvCosts"},
    {label:"NPV", key:"npv"},
    {label:"BCR", key:"bcr"},
    {label:"ROI", key:"roi"},
    {label:"Rank (by current metric)", key:"rank"}
  ];

  const metric = state.ui.rankBy;
  const scored = cbaList.map(x => {
    const v = (metric === "npv") ? x.npv : (metric === "bcr" ? x.bcr : x.roi);
    return {name:x.name, v};
  });

  scored.sort((a,b) => {
    const av = a.v, bv = b.v;
    const aBad = !Number.isFinite(av);
    const bBad = !Number.isFinite(bv);
    if (aBad && bBad) return a.name.localeCompare(b.name);
    if (aBad) return 1;
    if (bBad) return -1;
    return (bv - av);
  });

  const rankMap = {};
  for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

  let html = `<table><thead><tr><th>Indicator</th>`;
  for (const n of names){
    const isBase = n === state.map.baseline;
    html += `<th>${escapeHtml(n)}${isBase ? "<br><span class='muted'>(baseline)</span>":""}</th>`;
  }
  html += `</tr></thead><tbody>`;

  for (const r of rows){
    html += `<tr><td>${escapeHtml(r.label)}</td>`;
    for (const t of cbaList){
      let v = "";
      if (r.key === "rank"){
        v = rankMap[t.name] ? String(rankMap[t.name]) : "—";
      } else if (r.key === "bcr" || r.key === "roi"){
        v = fmtNumber(t[r.key], 3);
      } else {
        v = fmtMoney(t[r.key], 0);
      }
      html += `<td class="mono">${escapeHtml(v)}</td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody></table>`;
  return html;
}

function renderResults(){
  const cbaWrap = byId("cbaVerticalWrap");
  const chart = byId("npvChart");
  if (!cbaWrap) return;

  if (!state.raw.length || !state.map.treatment || !state.map.baseline || !state.map.yield || !state.map.cost){
    cbaWrap.innerHTML = `<div class="muted">Load data and confirm mappings (treatment, baseline, yield, cost).</div>`;
    if (byId("trialSummaryWrap")) byId("trialSummaryWrap").innerHTML = "";
    if (chart) drawEmpty(chart, "Load data to see NPV chart");
    return;
  }

  const cba = buildCbaAllTreatments();
  if (!cba.length){
    cbaWrap.innerHTML = `<div class="muted">Could not compute results. Check mappings and assumptions.</div>`;
    if (chart) drawEmpty(chart, "No results yet");
    return;
  }

  cbaWrap.innerHTML = verticalCbaTable(cba);
  renderTrialSummary();
  if (chart) drawNpvChart(chart, cba);

  // Optional: display note about adders if any enabled
  const enabledAdders = state.adders.items.map(normAdder).filter(a=>a.enabled);
  const note = byId("resultsNote");
  if (note && enabledAdders.length){
    note.textContent = `Definitions: PV benefits and PV costs are discounted sums. NPV = PV benefits − PV costs. BCR = PV benefits ÷ PV costs. ROI = NPV ÷ PV costs. Additional benefits/costs are included (${enabledAdders.length} adders enabled).`;
  }
}

function renderCashflows(){
  const tableWrap = byId("cashflowsTableWrap");
  const chart = byId("cashflowChart");
  if (!tableWrap) return;

  if (!state.raw.length || !state.map.baseline){
    tableWrap.innerHTML = `<div class="muted">Load data first.</div>`;
    if (chart) drawEmpty(chart, "Load data to see cashflows");
    return;
  }

  const tnames = getTreatmentNames();
  if (!tnames.length) return;

  const sel = byId("cashflowTreatment");
  const current = sel ? (sel.value || state.ui.cashflowTreatment || state.map.baseline) : (state.ui.cashflowTreatment || state.map.baseline);
  state.ui.cashflowTreatment = current;

  const cf = buildCashflowsForTreatment(current);
  if (!cf){
    tableWrap.innerHTML = `<div class="muted">Select a treatment.</div>`;
    if (chart) drawEmpty(chart, "No cashflows");
    return;
  }

  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));
  let html = `<table>
    <thead><tr>
      <th>Year</th>
      <th>Price ($/t)</th>
      <th>Benefit factor</th>
      <th>Benefits</th>
      <th>Costs</th>
      <th>Net</th>
      <th>Discount factor</th>
      <th>PV net</th>
    </tr></thead><tbody>`;

  for (let y=0; y<=T; y++){
    const price = (y === 0) ? "—" : fmtNumber(priceForYear(y), 2);
    const bf = (y === 0) ? "—" : fmtNumber(benefitFactorByYear(y), 3);
    const df = fmtNumber(discountFactor(y), 4);
    const pvnet = cf.net[y] * discountFactor(y);
    html += `<tr>
      <td class="mono">${y}</td>
      <td class="mono">${escapeHtml(price)}</td>
      <td class="mono">${escapeHtml(bf)}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.benefits[y],0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.costs[y],0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(cf.net[y],0))}</td>
      <td class="mono">${escapeHtml(df)}</td>
      <td class="mono">${escapeHtml(fmtMoney(pvnet,0))}</td>
    </tr>`;
  }
  html += `</tbody></table>`;
  tableWrap.innerHTML = html;

  if (byId("cfPvBenefits")) byId("cfPvBenefits").textContent = fmtMoney(cf.pvBenefits, 0);
  if (byId("cfPvCosts")) byId("cfPvCosts").textContent = fmtMoney(cf.pvCosts, 0);
  if (byId("cfNpv")) byId("cfNpv").textContent = fmtMoney(cf.npv, 0);
  if (byId("cfBcr")) byId("cfBcr").textContent = Number.isFinite(cf.bcr) ? fmtNumber(cf.bcr, 3) : "—";

  if (chart) drawCashflowChart(chart, cf);
}

function renderSensitivity(){
  const wrap = byId("sensitivityWrap");
  if (!wrap) return;

  if (!state.cache.lastSensitivity){
    wrap.innerHTML = `<div class="muted">Run sensitivity to see results.</div>`;
    return;
  }
  wrap.innerHTML = sensitivityTableHtml(state.cache.lastSensitivity);
}

/* =========================
   New tab renderers
   ========================= */
function renderIntro(){
  const el = byId("introWrap");
  if (!el) return;
  const base = state.map.baseline || "Control";
  const ds = state.source?.name || "your dataset";

  const enabledAdders = state.adders.items.map(normAdder).filter(a=>a.enabled);
  const addersLine = enabledAdders.length
    ? `This analysis also includes ${enabledAdders.length} additional benefit/cost items (for example, monitoring costs or co-benefits) that you can turn on/off and tailor.`
    : `You can add extra benefits and extra costs (beyond yield and amendment costs) in the Additional Benefits & Costs tab if needed.`;

  const intro = `
    <div class="prose">
      <p><strong>What this tool does</strong>: It compares each treatment against the baseline (<code>${escapeHtml(base)}</code>) using your imported data (${escapeHtml(ds)}). It converts trial differences into dollar outcomes over time using price, discounting, and an effect duration/decay model.</p>
      <p><strong>Who it is for</strong>: Farmers can see which drivers move dollars (yield, costs, timing, adoption). Policy makers can see assumptions, net benefits, and uncertainty. Researchers can audit mappings, cashflows, and exports.</p>
      <p><strong>What drives the results</strong>: Incremental yield (t/ha) and incremental costs (/ha) relative to baseline. Benefits are valued using grain price assumptions. Costs can be one-off or annual. Adoption scales results to whole-farm totals.</p>
      <p>${escapeHtml(addersLine)}</p>
      <p><strong>Decision support</strong>: Results show trade-offs and sensitivities. The tool does not tell you what to choose.</p>
    </div>
  `;
  el.innerHTML = intro;
}

function renderAdders(){
  const wrap = byId("addersTableWrap");
  if (!wrap) return;

  const items = state.adders.items.map(normAdder);
  if (!items.length){
    wrap.innerHTML = `<div class="muted">No additional benefit/cost items yet. Use “Add benefit” or “Add cost” to create one.</div>`;
    return;
  }

  const tnames = getTreatmentNames();

  let html = `<table>
    <thead><tr>
      <th data-tooltip="Enable/disable this extra benefit or cost.">On</th>
      <th data-tooltip="Benefit adds to benefits stream; Cost adds to costs stream.">Kind</th>
      <th data-tooltip="Short label used in reports and exports.">Label</th>
      <th data-tooltip="Amount value in dollars. Interpretation depends on basis.">Amount</th>
      <th data-tooltip="per_ha scales by area and adoption; total is already whole-farm; per_t multiplies incremental yield.">Basis</th>
      <th data-tooltip="When the item applies (setup vs annual).">Timing</th>
      <th data-tooltip="Start year (0 allowed for setup costs).">Start</th>
      <th data-tooltip="End year (used for duration/custom).">End</th>
      <th data-tooltip="Annual growth in percent for recurring items.">Growth %</th>
      <th data-tooltip="Apply to all treatments or selected ones.">Applies to</th>
      <th data-tooltip="Optional note for audit trail and policy briefs.">Notes</th>
      <th></th>
    </tr></thead><tbody>`;

  const basisOptions = ["per_ha","total","per_t"];
  const timingOptions = ["y0_only","y1_only","annual_duration","annual_horizon","custom"];

  for (const a of items){
    const applies = (a.appliesTo === "all") ? "all" : "selected";
    html += `<tr>
      <td class="mono"><input type="checkbox" data-adder="${escapeHtml(a.id)}" data-field="enabled" ${a.enabled ? "checked":""}></td>
      <td>
        <select class="tselect" data-adder="${escapeHtml(a.id)}" data-field="kind">
          <option value="benefit" ${a.kind==="benefit"?"selected":""}>benefit</option>
          <option value="cost" ${a.kind==="cost"?"selected":""}>cost</option>
        </select>
      </td>
      <td><input class="tinput" style="width:220px" type="text" data-adder="${escapeHtml(a.id)}" data-field="label" value="${escapeHtml(a.label)}"></td>
      <td><input class="tinput" type="number" step="0.1" data-adder="${escapeHtml(a.id)}" data-field="amount" value="${escapeHtml(String(a.amount))}"></td>
      <td>
        <select class="tselect" data-adder="${escapeHtml(a.id)}" data-field="amountBasis">
          ${basisOptions.map(b=>`<option value="${b}" ${a.amountBasis===b?"selected":""}>${b}</option>`).join("")}
        </select>
      </td>
      <td>
        <select class="tselect" data-adder="${escapeHtml(a.id)}" data-field="timing">
          ${timingOptions.map(t=>`<option value="${t}" ${a.timing===t?"selected":""}>${t}</option>`).join("")}
        </select>
      </td>
      <td><input class="tinput" type="number" step="1" data-adder="${escapeHtml(a.id)}" data-field="startYear" value="${escapeHtml(String(a.startYear))}"></td>
      <td><input class="tinput" type="number" step="1" data-adder="${escapeHtml(a.id)}" data-field="endYear" value="${escapeHtml(String(a.endYear))}"></td>
      <td><input class="tinput" type="number" step="0.1" data-adder="${escapeHtml(a.id)}" data-field="growthPct" value="${escapeHtml(String(a.growthPct))}"></td>
      <td>
        <select class="tselect" data-adder="${escapeHtml(a.id)}" data-field="appliesToMode">
          <option value="all" ${applies==="all"?"selected":""}>all</option>
          <option value="selected" ${applies==="selected"?"selected":""}>selected</option>
        </select>
      </td>
      <td><input class="tinput" style="width:220px" type="text" data-adder="${escapeHtml(a.id)}" data-field="notes" value="${escapeHtml(a.notes)}"></td>
      <td class="mono"><button class="btn btn--ghost" type="button" data-adder-del="${escapeHtml(a.id)}">Remove</button></td>
    </tr>`;

    // If selected mode, show multi-select row
    if (applies === "selected"){
      const selected = Array.isArray(a.appliesTo) ? a.appliesTo : [];
      html += `<tr>
        <td></td>
        <td colspan="10">
          <div class="muted" style="margin-bottom:6px;">Select treatments:</div>
          <div style="display:flex; flex-wrap:wrap; gap:8px;">
            ${tnames.map(n=>{
              const chk = selected.includes(n);
              return `<label class="chip" style="display:inline-flex; align-items:center; gap:6px; font-family: var(--mono);">
                <input type="checkbox" data-adder="${escapeHtml(a.id)}" data-field="appliesToPick" data-tname="${escapeHtml(n)}" ${chk?"checked":""}>
                ${escapeHtml(n)}
              </label>`;
            }).join("")}
          </div>
        </td>
        <td></td>
      </tr>`;
    }
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;

  // Bind changes
  wrap.querySelectorAll("[data-adder][data-field]").forEach(el => {
    on(el, "change", (e) => {
      const id = e.target.getAttribute("data-adder");
      const field = e.target.getAttribute("data-field");
      let ad = state.adders.items.find(x => String(x.id) === id);
      if (!ad){
        ad = normAdder({id});
        state.adders.items.push(ad);
      }

      if (field === "enabled") ad.enabled = e.target.checked;
      else if (field === "kind") ad.kind = e.target.value === "cost" ? "cost" : "benefit";
      else if (field === "label") ad.label = e.target.value;
      else if (field === "amount") ad.amount = Number(e.target.value || 0);
      else if (field === "amountBasis") ad.amountBasis = e.target.value;
      else if (field === "timing") ad.timing = e.target.value;
      else if (field === "startYear") ad.startYear = Math.floor(Number(e.target.value || 0));
      else if (field === "endYear") ad.endYear = Math.floor(Number(e.target.value || 0));
      else if (field === "growthPct") ad.growthPct = Number(e.target.value || 0);
      else if (field === "notes") ad.notes = e.target.value;
      else if (field === "appliesToMode"){
        if (e.target.value === "all") ad.appliesTo = "all";
        else ad.appliesTo = Array.isArray(ad.appliesTo) ? ad.appliesTo : [];
      } else if (field === "appliesToPick"){
        const tname = e.target.getAttribute("data-tname");
        const arr = Array.isArray(ad.appliesTo) ? ad.appliesTo : [];
        const has = arr.includes(tname);
        if (e.target.checked && !has) arr.push(tname);
        if (!e.target.checked && has) ad.appliesTo = arr.filter(x=>x!==tname);
        else ad.appliesTo = arr;
      }

      invalidateCache();
      renderResults();
      renderCashflows();
      renderCopilot();
      // rerender to show selected list when toggling applies mode
      if (field === "appliesToMode") renderAdders();
    });
  });

  wrap.querySelectorAll("[data-adder-del]").forEach(btn => {
    on(btn, "click", (e) => {
      const id = e.target.getAttribute("data-adder-del");
      state.adders.items = state.adders.items.filter(a => String(a.id) !== id);
      invalidateCache();
      renderAdders();
      renderResults();
      renderCashflows();
      renderCopilot();
      toast("Removed item");
    });
  });
}

function renderSimulations(){
  const wrap = byId("simResultsWrap");
  if (!wrap) return;

  if (!state.cache.lastSim){
    wrap.innerHTML = `<div class="muted">Run simulation to see uncertainty results.</div>`;
    return;
  }

  const rows = state.cache.lastSim.summary;

  let html = `<table>
    <thead><tr>
      <th>Treatment</th>
      <th>NPV P10</th>
      <th>NPV P50</th>
      <th>NPV P90</th>
      <th>Pr(NPV&gt;0)</th>
      <th>BCR P50</th>
      <th>Pr(BCR&gt;1)</th>
      <th>Pr(best by NPV)</th>
    </tr></thead><tbody>`;

  for (const r of rows){
    const badge = r.isBaseline ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(r.treatment)}${badge}</td>
      <td class="mono">${escapeHtml(fmtMoney(r.npv_p10,0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(r.npv_p50,0))}</td>
      <td class="mono">${escapeHtml(fmtMoney(r.npv_p90,0))}</td>
      <td class="mono">${escapeHtml(fmtNumber(r.prob_npv_gt_0*100,1))}%</td>
      <td class="mono">${escapeHtml(fmtNumber(r.bcr_p50,3))}</td>
      <td class="mono">${escapeHtml(fmtNumber(r.prob_bcr_gt_1*100,1))}%</td>
      <td class="mono">${escapeHtml(fmtNumber(r.prob_best_by_npv*100,1))}%</td>
    </tr>`;
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;

  const note = byId("simNote");
  if (note){
    const s = state.cache.lastSim.settings;
    note.textContent = `Simulation ran ${s.draws} draws (${s.distribution}) around: price ±${Math.round(s.pricePct*100)}%, yield mult ±${Math.round(s.yieldMultPct*100)}%, cost mult ±${Math.round(s.costMultPct*100)}%, discount ±${(s.discountPts*100).toFixed(1)} pp, duration ±${s.durationYears} years.`;
  }
}

function renderCopilot(){
  const ta = byId("copilotPrompt");
  const meta = byId("copilotMeta");
  if (!ta && !meta) return;

  // If results aren't ready, still produce a scaffold prompt
  const payload = buildCopilotPayload();
  const json = JSON.stringify(payload, null, 2);

  if (ta) ta.value = json;

  if (meta){
    meta.innerHTML = `
      <div class="muted">
        Audience: <code>${escapeHtml(state.copilot.audience)}</code> · Tone: <code>${escapeHtml(state.copilot.tone)}</code> · Length: <code>${escapeHtml(state.copilot.length)}</code>
        · Baseline: <code>${escapeHtml(state.map.baseline || "—")}</code>
      </div>
    `;
  }
}

/* =========================
   Charts (simple canvas)
   ========================= */
function drawEmpty(canvas, msg){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);
  ctx.fillStyle = "rgba(255,255,255,0.72)";
  ctx.font = "16px system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial";
  ctx.fillText(msg, 16, 28);
}

function drawNpvChart(canvas, cbaList){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);

  const metric = state.ui.rankBy;
  const items = cbaList.slice().map(x => {
    const v = (metric === "npv") ? x.npv : (metric === "bcr" ? x.bcr : x.roi);
    return {name:x.name, v, npv:x.npv};
  });

  items.sort((a,b) => {
    const av = a.v, bv = b.v;
    const aBad = !Number.isFinite(av);
    const bBad = !Number.isFinite(bv);
    if (aBad && bBad) return a.name.localeCompare(b.name);
    if (aBad) return 1;
    if (bBad) return -1;
    return (bv - av);
  });

  const values = items.map(x => x.npv).filter(Number.isFinite);
  const minV = Math.min(0, ...values);
  const maxV = Math.max(0, ...values);

  const padL = 50, padR = 10, padT = 18, padB = 90;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;

  ctx.strokeStyle = "rgba(255,255,255,0.18)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  const n = items.length;
  const gap = 6;
  const barW = Math.max(6, (plotW - gap*(n-1)) / n);

  function yScale(v){
    if (maxV === minV) return padT + plotH/2;
    return padT + (maxV - v) * (plotH / (maxV - minV));
  }
  const y0 = yScale(0);

  for (let i=0; i<n; i++){
    const x = padL + i*(barW + gap);
    const v = items[i].npv;
    const y = yScale(v);
    const top = Math.min(y, y0);
    const bh = Math.abs(y0 - y);

    ctx.fillStyle = (items[i].name === state.map.baseline) ? "rgba(167,243,208,0.55)" : "rgba(125,211,252,0.55)";
    ctx.fillRect(x, top, barW, bh);

    ctx.save();
    ctx.translate(x + barW/2, padT + plotH + 10);
    ctx.rotate(-Math.PI/3);
    ctx.fillStyle = "rgba(255,255,255,0.75)";
    ctx.font = "12px ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas";
    ctx.textAlign = "right";
    ctx.fillText(items[i].name.length > 18 ? (items[i].name.slice(0,18)+"…") : items[i].name, 0, 0);
    ctx.restore();
  }

  ctx.fillStyle = "rgba(255,255,255,0.86)";
  ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto";
  ctx.fillText(`NPV chart (rank by ${metric.toUpperCase()})`, 16, 18);
}

function drawCashflowChart(canvas, cf){
  const ctx = canvas.getContext("2d");
  const w = canvas.width, h = canvas.height;
  ctx.clearRect(0,0,w,h);

  const T = cf.net.length - 1;
  const values = cf.net.slice(0).filter(Number.isFinite);
  const minV = Math.min(0, ...values);
  const maxV = Math.max(0, ...values);

  const padL = 50, padR = 10, padT = 18, padB = 50;
  const plotW = w - padL - padR;
  const plotH = h - padT - padB;

  ctx.strokeStyle = "rgba(255,255,255,0.18)";
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, padT);
  ctx.lineTo(padL, padT + plotH);
  ctx.lineTo(padL + plotW, padT + plotH);
  ctx.stroke();

  function yScale(v){
    if (maxV === minV) return padT + plotH/2;
    return padT + (maxV - v) * (plotH / (maxV - minV));
  }
  const y0 = yScale(0);

  const n = T + 1;
  const gap = 6;
  const barW = Math.max(8, (plotW - gap*(n-1)) / n);

  for (let i=0; i<n; i++){
    const x = padL + i*(barW + gap);
    const v = cf.net[i];
    const y = yScale(v);
    const top = Math.min(y, y0);
    const bh = Math.abs(y0 - y);

    ctx.fillStyle = (v >= 0) ? "rgba(167,243,208,0.55)" : "rgba(252,165,165,0.45)";
    ctx.fillRect(x, top, barW, bh);

    ctx.fillStyle = "rgba(255,255,255,0.70)";
    ctx.font = "11px ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas";
    ctx.textAlign = "center";
    ctx.fillText(String(i), x + barW/2, padT + plotH + 16);
  }

  ctx.fillStyle = "rgba(255,255,255,0.86)";
  ctx.font = "14px system-ui, -apple-system, Segoe UI, Roboto";
  ctx.fillText(`Net cashflow by year: ${cf.name}`, 16, 18);
}

/* =========================
   Sensitivity
   ========================= */
function computeMetric(cf, metric){
  if (metric === "npv") return cf.npv;
  if (metric === "bcr") return cf.bcr;
  if (metric === "roi") return cf.roi;
  return cf.npv;
}

function cloneAssumptions(){
  return JSON.parse(JSON.stringify(state.assumptions));
}

function runSensitivity(){
  const treatment = byId("sensTreatment")?.value;
  const metric = byId("sensMetric")?.value || "npv";
  if (!treatment) return toast("Select a treatment first");

  const pricePct = Number(byId("sensPricePct")?.value || 0) / 100;
  const costPct = Number(byId("sensCostPct")?.value || 0) / 100;
  const yieldPct = Number(byId("sensYieldPct")?.value || 0) / 100;
  const discPts = Number(byId("sensDiscountPts")?.value || 0);
  const durDelta = Math.floor(Number(byId("sensDurationDelta")?.value || 0));

  const baseAssump = cloneAssumptions();
  const baseOv = JSON.parse(JSON.stringify(state.treatmentOverrides));

  function evalWith(modFn){
    const savedA = state.assumptions;
    const savedO = state.treatmentOverrides;

    state.assumptions = cloneAssumptions();
    state.treatmentOverrides = JSON.parse(JSON.stringify(baseOv));

    modFn();

    invalidateCache();
    const cf = buildCashflowsForTreatment(treatment);

    state.assumptions = savedA;
    state.treatmentOverrides = savedO;
    invalidateCache();

    return cf ? computeMetric(cf, metric) : NaN;
  }

  const baseVal = evalWith(() => {});

  const res = [
    {
      driver: "Price",
      low: evalWith(() => { state.assumptions.priceYear1 = baseAssump.priceYear1 * (1 - pricePct); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.priceYear1 = baseAssump.priceYear1 * (1 + pricePct); })
    },
    {
      driver: "Incremental cost multiplier",
      low: evalWith(() => { state.treatmentOverrides[treatment].costMult = (baseOv[treatment].costMult || 1) * (1 - costPct); }),
      base: baseVal,
      high: evalWith(() => { state.treatmentOverrides[treatment].costMult = (baseOv[treatment].costMult || 1) * (1 + costPct); })
    },
    {
      driver: "Incremental yield multiplier",
      low: evalWith(() => { state.treatmentOverrides[treatment].yieldMult = (baseOv[treatment].yieldMult || 1) * (1 - yieldPct); }),
      base: baseVal,
      high: evalWith(() => { state.treatmentOverrides[treatment].yieldMult = (baseOv[treatment].yieldMult || 1) * (1 + yieldPct); })
    },
    {
      driver: "Discount rate",
      low: evalWith(() => { state.assumptions.discountRatePct = Math.max(0, baseAssump.discountRatePct - discPts); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.discountRatePct = Math.max(0, baseAssump.discountRatePct + discPts); })
    },
    {
      driver: "Effect duration",
      low: evalWith(() => { state.assumptions.effectDurationYears = Math.max(1, baseAssump.effectDurationYears - durDelta); }),
      base: baseVal,
      high: evalWith(() => { state.assumptions.effectDurationYears = Math.max(1, baseAssump.effectDurationYears + durDelta); })
    }
  ];

  const payload = {
    treatment,
    metric,
    inputs: {pricePct, costPct, yieldPct, discPts, durDelta},
    rows: res
  };

  state.cache.lastSensitivity = payload;
  renderSensitivity();
  renderCopilot();
  toast("Sensitivity updated");
}

function sensitivityTableHtml(payload){
  const metric = payload.metric;
  const isRatio = (metric === "bcr" || metric === "roi");

  let html = `<table>
    <thead><tr>
      <th>Driver</th>
      <th>Low</th>
      <th>Base</th>
      <th>High</th>
    </tr></thead><tbody>`;

  for (const r of payload.rows){
    const f = (x) => isRatio ? fmtNumber(x, 3) : fmtMoney(x, 0);
    html += `<tr>
      <td>${escapeHtml(r.driver)}</td>
      <td class="mono">${escapeHtml(f(r.low))}</td>
      <td class="mono">${escapeHtml(f(r.base))}</td>
      <td class="mono">${escapeHtml(f(r.high))}</td>
    </tr>`;
  }

  html += `</tbody></table>`;
  return html;
}

/* =========================
   Exports
   ========================= */
function toCsv(rows, columns){
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : []);
  const esc = (v) => {
    const s = (v === null || v === undefined) ? "" : String(v);
    if (/[",\n]/.test(s)) return `"${s.replaceAll('"','""')}"`;
    return s;
  };
  const lines = [];
  lines.push(cols.map(esc).join(","));
  for (const r of rows){
    lines.push(cols.map(c => esc(r[c])).join(","));
  }
  return lines.join("\n");
}

function downloadBlob(filename, blob){
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function downloadText(filename, text, mime="text/plain"){
  downloadBlob(filename, new Blob([text], {type: mime}));
}

function exportVerticalCsv(){
  const cba = buildCbaAllTreatments();
  if (!cba.length) return toast("Nothing to export");
  const metric = state.ui.rankBy;
  const scored = cba.map(x => ({name:x.name, v: (metric==="npv"?x.npv:(metric==="bcr"?x.bcr:x.roi))}))
    .sort((a,b)=> {
      const aBad = !Number.isFinite(a.v);
      const bBad = !Number.isFinite(b.v);
      if (aBad && bBad) return a.name.localeCompare(b.name);
      if (aBad) return 1;
      if (bBad) return -1;
      return (b.v - a.v);
    });
  const rankMap = {};
  for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

  const indicators = [
    {Indicator:"PV benefits", key:"pvBenefits"},
    {Indicator:"PV costs", key:"pvCosts"},
    {Indicator:"NPV", key:"npv"},
    {Indicator:"BCR", key:"bcr"},
    {Indicator:"ROI", key:"roi"},
    {Indicator:`Rank (by ${metric.toUpperCase()})`, key:"rank"}
  ];

  const cols = ["Indicator", ...cba.map(x => x.name)];
  const rows = indicators.map(ind => {
    const obj = {Indicator: ind.Indicator};
    for (const t of cba){
      let v = "";
      if (ind.key === "rank") v = rankMap[t.name] || "";
      else v = t[ind.key];
      obj[t.name] = v;
    }
    return obj;
  });

  downloadText("cba_vertical.csv", toCsv(rows, cols), "text/csv");
}

function exportTrialSummaryCsv(){
  const sum = buildTrialSummary();
  if (!sum.length) return toast("Nothing to export");
  const cols = ["name","n","yield_mean","cost_mean","delta_yield","delta_cost"];
  downloadText("trial_summary.csv", toCsv(sum, cols), "text/csv");
}

function exportStateJson(){
  const payload = {
    source: state.source,
    columns: state.columns,
    map: state.map,
    assumptions: state.assumptions,
    treatmentOverrides: state.treatmentOverrides,
    adders: state.adders,
    sim: state.sim,
    copilot: state.copilot,
    raw: state.raw
  };
  downloadText("tool_state.json", JSON.stringify(payload, null, 2), "application/json");
}

function exportSensitivityXlsx(){
  if (!state.cache.lastSensitivity) return toast("Run sensitivity first");
  if (!window.XLSX) return toast("XLSX library missing");

  const wb = XLSX.utils.book_new();
  const s = state.cache.lastSensitivity;

  const meta = [
    {key:"Treatment", value:s.treatment},
    {key:"Metric", value:s.metric},
    {key:"Price change +/-", value:(s.inputs.pricePct*100)+"%"},
    {key:"Cost mult change +/-", value:(s.inputs.costPct*100)+"%"},
    {key:"Yield mult change +/-", value:(s.inputs.yieldPct*100)+"%"},
    {key:"Discount rate change +/- (pp)", value:s.inputs.discPts},
    {key:"Duration change +/- (years)", value:s.inputs.durDelta}
  ];

  const wsMeta = XLSX.utils.json_to_sheet(meta);
  XLSX.utils.book_append_sheet(wb, wsMeta, "Sensitivity_Meta");

  const ws = XLSX.utils.json_to_sheet(s.rows);
  XLSX.utils.book_append_sheet(wb, ws, "Sensitivity");

  const out = XLSX.write(wb, {bookType:"xlsx", type:"array"});
  downloadBlob("sensitivity.xlsx", new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
}

function exportFullWorkbookXlsx(){
  if (!window.XLSX) return toast("XLSX library missing");
  if (!state.raw.length) return toast("Load data first");

  const wb = XLSX.utils.book_new();

  const assumpRows = Object.keys(state.assumptions).map(k => ({key:k, value: state.assumptions[k]}));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(assumpRows), "Assumptions");

  const trial = buildTrialSummary();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(trial), "TrialMeans");

  const cba = buildCbaAllTreatments();
  if (cba.length){
    const metric = state.ui.rankBy;
    const scored = cba.map(x => ({name:x.name, v: (metric==="npv"?x.npv:(metric==="bcr"?x.bcr:x.roi))}))
      .sort((a,b)=> {
        const aBad = !Number.isFinite(a.v);
        const bBad = !Number.isFinite(b.v);
        if (aBad && bBad) return a.name.localeCompare(b.name);
        if (aBad) return 1;
        if (bBad) return -1;
        return (b.v - a.v);
      });
    const rankMap = {};
    for (let i=0; i<scored.length; i++) rankMap[scored[i].name] = i+1;

    const indicators = [
      {Indicator:"PV benefits", key:"pvBenefits"},
      {Indicator:"PV costs", key:"pvCosts"},
      {Indicator:"NPV", key:"npv"},
      {Indicator:"BCR", key:"bcr"},
      {Indicator:"ROI", key:"roi"},
      {Indicator:`Rank (by ${metric.toUpperCase()})`, key:"rank"}
    ];
    const cols = ["Indicator", ...cba.map(x => x.name)];
    const rows = indicators.map(ind => {
      const obj = {Indicator: ind.Indicator};
      for (const t of cba){
        obj[t.name] = (ind.key === "rank") ? (rankMap[t.name] || "") : t[ind.key];
      }
      return obj;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows, {header: cols}), "CBA_Vertical");
  }

  // Cashflows (all treatments)
  const cashRows = [];
  const T = Math.max(1, Math.floor(Number(state.assumptions.horizonYears) || 1));
  for (const t of cba){
    for (let y=0; y<=T; y++){
      cashRows.push({
        treatment: t.name,
        year: y,
        price: (y===0 ? "" : priceForYear(y)),
        benefit_factor: (y===0 ? "" : benefitFactorByYear(y)),
        benefits: t.benefits[y],
        costs: t.costs[y],
        net: t.net[y],
        discount_factor: discountFactor(y),
        pv_net: t.net[y] * discountFactor(y)
      });
    }
  }
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cashRows), "Cashflows_All");

  // Additional adders
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.adders.items.map(normAdder)), "Adders");

  // Sensitivity (last run)
  if (state.cache.lastSensitivity){
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.cache.lastSensitivity.rows), "Sensitivity_Last");
  }

  // Simulation (last run)
  if (state.cache.lastSim){
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.cache.lastSim.summary), "Simulation_Last");
  }

  // Copilot prompt
  const cp = buildCopilotPayload();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{json: JSON.stringify(cp)}]), "CopilotPrompt");

  // Raw data
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.raw), "RawData");

  const out = XLSX.write(wb, {bookType:"xlsx", type:"array"});
  downloadBlob("farming_cba_tool_output.xlsx", new Blob([out], {type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
}

/* =========================
   Events
   ========================= */
function bindEvents(){
  // Tabs (existing + new)
  $$(".tab").forEach(btn => on(btn, "click", () => setActiveTab(btn.dataset.tab)));
  bindTabKeyboardNav();

  // Import
  on(byId("fileInput"), "change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try{
      const buf = await f.arrayBuffer();
      await importExcelArrayBuffer(buf, f.name);
      setActiveTab("mapping");
    } catch(err){
      console.error(err);
      toast(`Import failed: ${err.message || err}`);
    }
  });

  on(byId("btnLoadBundledXlsx"), "click", async () => {
    try{
      const ok = await tryLoadBundled();
      if (!ok) toast("Bundled XLSX not found. Use Import or embedded sample.");
      else setActiveTab("mapping");
    } catch(err){
      console.error(err);
      toast(`Bundled load failed: ${err.message || err}`);
    }
  });

  on(byId("btnLoadEmbedded"), "click", () => {
    loadEmbedded();
    setActiveTab("mapping");
  });

  on(byId("btnReset"), "click", resetApp);

  on(byId("btnGoMapping"), "click", () => setActiveTab("mapping"));
  on(byId("btnGoResults"), "click", () => setActiveTab("results"));
  on(byId("btnGoAssumptions"), "click", () => setActiveTab("assumptions"));
  on(byId("btnGoImport"), "click", () => setActiveTab("import"));
  on(byId("btnGoResults2"), "click", () => setActiveTab("results"));

  // Mapping selectors
  on(byId("mapTreatment"), "change", (e) => {
    state.map.treatment = e.target.value;
    const d = detectDefaultMappings(state.raw, state.columns);
    initTreatmentOverrides();
    state.map.baseline = d.baseline || state.map.baseline;
    invalidateCache();
    renderAll();
  });

  on(byId("mapBaseline"), "change", (e) => {
    state.map.baseline = e.target.value;
    if (state.treatmentOverrides[state.map.baseline]){
      state.treatmentOverrides[state.map.baseline].enabled = true;
      state.treatmentOverrides[state.map.baseline].adoption = 1;
    }
    invalidateCache();
    renderAll();
  });

  on(byId("mapYield"), "change", (e) => {
    state.map.yield = e.target.value || null;
    invalidateCache();
    renderAll();
  });

  on(byId("mapCost"), "change", (e) => {
    state.map.cost = e.target.value || null;
    invalidateCache();
    renderAll();
  });

  on(byId("mapOptional1"), "change", (e) => {
    state.map.optional1 = e.target.value || null;
    renderAll();
  });

  // Assumptions inputs (guarded)
  const bindAssump = (id, key, parser=(x)=>x) => {
    on(byId(id), "input", (e) => {
      state.assumptions[key] = parser(e.target.value);
      invalidateCache();
      renderAll();
    });
  };

  bindAssump("assumpArea", "areaHa", (v)=>Number(v));
  bindAssump("assumpHorizon", "horizonYears", (v)=>Math.max(1, Math.floor(Number(v||1))));
  bindAssump("assumpDiscount", "discountRatePct", (v)=>Math.max(0, Number(v)));
  bindAssump("assumpPrice", "priceYear1", (v)=>Math.max(0, Number(v)));
  bindAssump("assumpPriceGrowth", "priceGrowthPct", (v)=>Number(v));

  on(byId("assumpYieldScale"), "change", (e)=>{ state.assumptions.yieldScale = Number(e.target.value||1); invalidateCache(); renderAll(); });
  on(byId("assumpCostScale"), "change", (e)=>{ state.assumptions.costScale = Number(e.target.value||1); invalidateCache(); renderAll(); });

  bindAssump("assumpBenefitStart", "benefitStartYear", (v)=>Math.max(1, Math.floor(Number(v||1))));
  bindAssump("assumpDuration", "effectDurationYears", (v)=>Math.max(1, Math.floor(Number(v||1))));
  on(byId("assumpDecay"), "change", (e)=>{ state.assumptions.decay = e.target.value; invalidateCache(); renderAll(); });
  bindAssump("assumpHalfLife", "halfLifeYears", (v)=>Math.max(0.1, Number(v||2)));
  on(byId("assumpCostTimingDefault"), "change", (e)=>{ state.assumptions.costTimingDefault = e.target.value; invalidateCache(); renderAll(); });

  on(byId("btnResetAssumptions"), "click", () => {
    state.assumptions = {
      areaHa: 100,
      horizonYears: 10,
      discountRatePct: 7,
      priceYear1: 450,
      priceGrowthPct: 0,
      yieldScale: 1,
      costScale: 1,
      benefitStartYear: 1,
      effectDurationYears: 5,
      decay: "linear",
      halfLifeYears: 2,
      costTimingDefault: "y1_only"
    };
    invalidateCache();
    renderAll();
    toast("Assumptions reset");
  });

  // Results controls
  on(byId("resultMetric"), "change", (e)=>{ state.ui.rankBy = e.target.value; renderResults(); renderCopilot(); });
  on(byId("resultView"), "change", (e)=>{ state.ui.view = e.target.value; invalidateCache(); renderAll(); });

  on(byId("btnGoCashflows"), "click", ()=> setActiveTab("cashflows"));
  on(byId("cashflowTreatment"), "change", ()=>{ renderCashflows(); renderCopilot(); });
  on(byId("btnBackToResults"), "click", ()=> setActiveTab("results"));

  // Sensitivity
  on(byId("btnRunSensitivity"), "click", runSensitivity);
  on(byId("btnResetSensitivity"), "click", () => {
    if (byId("sensPricePct")) byId("sensPricePct").value = 10;
    if (byId("sensCostPct")) byId("sensCostPct").value = 10;
    if (byId("sensYieldPct")) byId("sensYieldPct").value = 10;
    if (byId("sensDiscountPts")) byId("sensDiscountPts").value = 2;
    if (byId("sensDurationDelta")) byId("sensDurationDelta").value = 2;
    toast("Sensitivity inputs reset");
  });
  on(byId("btnSensitivityExportXlsx"), "click", exportSensitivityXlsx);

  // Export
  on(byId("btnExportXlsx"), "click", exportFullWorkbookXlsx);
  on(byId("btnExportVerticalCsv"), "click", exportVerticalCsv);
  on(byId("btnExportTrialSummaryCsv"), "click", exportTrialSummaryCsv);
  on(byId("btnExportStateJson"), "click", exportStateJson);

  // Adders controls (optional)
  on(byId("btnAddBenefitAdder"), "click", () => {
    state.adders.items.push(normAdder({kind:"benefit", label:"New benefit", amount:0, amountBasis:"per_ha", timing:"annual_duration", startYear:1, endYear: state.assumptions.effectDurationYears, enabled:true, appliesTo:"all"}));
    invalidateCache();
    renderAdders();
    renderResults();
    renderCashflows();
    renderCopilot();
  });
  on(byId("btnAddCostAdder"), "click", () => {
    state.adders.items.push(normAdder({kind:"cost", label:"New cost", amount:0, amountBasis:"per_ha", timing:"annual_duration", startYear:1, endYear: state.assumptions.effectDurationYears, enabled:true, appliesTo:"all"}));
    invalidateCache();
    renderAdders();
    renderResults();
    renderCashflows();
    renderCopilot();
  });
  on(byId("btnClearAdders"), "click", () => {
    state.adders.items = [];
    invalidateCache();
    renderAdders();
    renderResults();
    renderCashflows();
    renderCopilot();
    toast("Cleared additional items");
  });

  // Simulation controls (optional)
  on(byId("btnRunSim"), "click", () => {
    // pull settings from UI if present
    if (byId("simDraws")) state.sim.draws = Math.max(50, Math.floor(Number(byId("simDraws").value || state.sim.draws)));
    if (byId("simSeed")) state.sim.seed = Math.floor(Number(byId("simSeed").value || state.sim.seed));
    if (byId("simPricePct")) state.sim.pricePct = Math.max(0, Number(byId("simPricePct").value || 0) / 100);
    if (byId("simYieldMultPct")) state.sim.yieldMultPct = Math.max(0, Number(byId("simYieldMultPct").value || 0) / 100);
    if (byId("simCostMultPct")) state.sim.costMultPct = Math.max(0, Number(byId("simCostMultPct").value || 0) / 100);
    if (byId("simDiscountPts")) state.sim.discountPts = Math.max(0, Number(byId("simDiscountPts").value || 0) / 100);
    if (byId("simDurationYears")) state.sim.durationYears = Math.max(0, Math.floor(Number(byId("simDurationYears").value || 0)));
    if (byId("simDistribution")) state.sim.distribution = byId("simDistribution").value || state.sim.distribution;

    const out = runMonteCarlo();
    if (!out) return toast("Simulation not available yet. Load data first.");
    renderSimulations();
    renderCopilot();
    toast("Simulation complete");
  });

  // Copilot controls (optional)
  on(byId("copilotAudience"), "change", (e)=>{ state.copilot.audience = e.target.value; renderCopilot(); });
  on(byId("copilotTone"), "change", (e)=>{ state.copilot.tone = e.target.value; renderCopilot(); });
  on(byId("copilotLength"), "change", (e)=>{ state.copilot.length = e.target.value; renderCopilot(); });
  on(byId("copilotIncludeTables"), "change", (e)=>{ state.copilot.includeTables = e.target.checked; renderCopilot(); });
  on(byId("copilotIncludeAssumptions"), "change", (e)=>{ state.copilot.includeAssumptions = e.target.checked; renderCopilot(); });
  on(byId("copilotIncludeCashflows"), "change", (e)=>{ state.copilot.includeCashflows = e.target.checked; renderCopilot(); });
  on(byId("copilotIncludeSensitivity"), "change", (e)=>{ state.copilot.includeSensitivity = e.target.checked; renderCopilot(); });
  on(byId("copilotIncludeSimulations"), "change", (e)=>{ state.copilot.includeSimulations = e.target.checked; renderCopilot(); });

  on(byId("btnCopilotCopy"), "click", async () => {
    const ta = byId("copilotPrompt");
    if (!ta) return;
    await copyTextToClipboard(ta.value || "");
  });

  on(byId("btnCopilotDownload"), "click", () => {
    const ta = byId("copilotPrompt");
    if (!ta) return;
    downloadText("copilot_prompt.json", ta.value || "{}", "application/json");
  });

  on(byId("btnCopilotRefresh"), "click", () => {
    renderCopilot();
    toast("Prompt refreshed");
  });
}

/* =========================
   Reset + boot
   ========================= */
function resetApp(){
  state.source = null;
  state.raw = [];
  state.columns = [];
  state.map = {treatment:null, baseline:null, yield:null, cost:null, optional1:null};
  state.assumptions = {
    areaHa: 100,
    horizonYears: 10,
    discountRatePct: 7,
    priceYear1: 450,
    priceGrowthPct: 0,
    yieldScale: 1,
    costScale: 1,
    benefitStartYear: 1,
    effectDurationYears: 5,
    decay: "linear",
    halfLifeYears: 2,
    costTimingDefault: "y1_only"
  };
  state.treatmentOverrides = {};
  state.adders.items = [];
  state.ui = {activeTab:"import", rankBy:"npv", view:"whole_farm", cashflowTreatment:null};
  state.cache = {
    trialSummary:null,
    cbaPerTreatment:null,
    lastSensitivity: state.cache.lastSensitivity || null,
    lastSim: null
  };

  const fi = byId("fileInput");
  if (fi) fi.value = "";
  toast("Reset complete");
  renderAll();
  setActiveTab("import");
}

(function boot(){
  bindEvents();

  // Sensitivity defaults if controls exist
  if (byId("sensPricePct")) byId("sensPricePct").value = 10;
  if (byId("sensCostPct")) byId("sensCostPct").value = 10;
  if (byId("sensYieldPct")) byId("sensYieldPct").value = 10;
  if (byId("sensDiscountPts")) byId("sensDiscountPts").value = 2;
  if (byId("sensDurationDelta")) byId("sensDurationDelta").value = 2;

  renderAll();
  setActiveTab("import");

  // Auto-load: try bundled; otherwise embedded
  (async () => {
    try{
      const ok = await tryLoadBundled();
      if (!ok) loadEmbedded();
    } catch(e){
      loadEmbedded();
    }
  })();
})();
