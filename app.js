// app.js
/* Fully functional, single-file vanilla JS app with:
   - working tabs
   - robust Excel import (SheetJS)
   - bundled fetch attempt (data.xlsx or Data for Lockhart-FA-031225 (1).xlsx)
   - embedded Lockhart dataset fallback (derived from the uploaded workbook)
   - treatment summaries + baseline comparisons
   - scenario controls + CSV/JSON exports
*/

/* =========================
   Embedded Lockhart sample
   ========================= */
const EMBEDDED_LOCKHART_ROWS = [
  {"Plot":1,"Rep":1,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":7.029229293617021,"Protein":23.2,"InputCost_perHa":16850,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":17945.488764568763},
  {"Plot":2,"Rep":1,"Trt":6,"Amendment":"Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":6.539273035489362,"Protein":23.6,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":24884.884058984914},
  {"Plot":3,"Rep":1,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.54757540287234,"Protein":23.7,"InputCost_perHa":17100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":18463.88888888889},
  {"Plot":4,"Rep":1,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.37207183687234,"Protein":24.7,"InputCost_perHa":1500,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1012.5633802816902},
  {"Plot":5,"Rep":1,"Trt":2,"Amendment":"Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.667165176319149,"Protein":23.9,"InputCost_perHa":1650,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":912.1951219512195},
  {"Plot":6,"Rep":1,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.199593337872341,"Protein":23.3,"InputCost_perHa":1100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":976.8292682926829},
  {"Plot":7,"Rep":1,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":6.614808249489362,"Protein":23.4,"InputCost_perHa":2550,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1033.6585365853657},
  {"Plot":8,"Rep":1,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.958392845957447,"Protein":23.1,"InputCost_perHa":7250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":782.9268292682927},
  {"Plot":9,"Rep":1,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.280082541361702,"Protein":23.7,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1059.349593495935},
  {"Plot":10,"Rep":1,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.839204480808511,"Protein":23.8,"InputCost_perHa":1900,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":908.9430894308944},
  {"Plot":11,"Rep":1,"Trt":1,"Amendment":"Lime only","Practice_Change":"Crop 1","Yield_t_ha":7.142857142857143,"Protein":23.3,"InputCost_perHa":250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":790.650406504065},
  {"Plot":12,"Rep":1,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":7.421177186276596,"Protein":23.4,"InputCost_perHa":6000,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":748.780487804878},
  {"Plot":13,"Rep":1,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.35011252212766,"Protein":23.7,"InputCost_perHa":2800,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1024.5934959349595},
  {"Plot":14,"Rep":1,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.35667234112766,"Protein":23.2,"InputCost_perHa":2150,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":923.1707317073171},
  {"Plot":15,"Rep":1,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.269241385957447,"Protein":23.9,"InputCost_perHa":2950,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1062.439024390244},
  {"Plot":16,"Rep":1,"Trt":0,"Amendment":"Control","Practice_Change":"Crop 1","Yield_t_ha":7.626622926382979,"Protein":23.5,"InputCost_perHa":0,"MachineryCost_perHa":0,"LabourCost_perHa":0,"TotalCost_perHa":694.6341463414634},

  {"Plot":17,"Rep":2,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":8.16927218306383,"Protein":23.5,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1059.349593495935},
  {"Plot":18,"Rep":2,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":7.467401828978723,"Protein":23.3,"InputCost_perHa":6000,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":748.780487804878},
  {"Plot":19,"Rep":2,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.751529506382978,"Protein":23.1,"InputCost_perHa":2150,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":923.1707317073171},
  {"Plot":20,"Rep":2,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.771685019574468,"Protein":23.7,"InputCost_perHa":2800,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1024.5934959349595},
  {"Plot":21,"Rep":2,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.373296617021276,"Protein":23.5,"InputCost_perHa":1100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":976.8292682926829},
  {"Plot":22,"Rep":2,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.794740369574468,"Protein":23.4,"InputCost_perHa":1900,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":908.9430894308944},
  {"Plot":23,"Rep":2,"Trt":1,"Amendment":"Lime only","Practice_Change":"Crop 1","Yield_t_ha":7.830657711489362,"Protein":23.2,"InputCost_perHa":250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":790.650406504065},
  {"Plot":24,"Rep":2,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":7.082754630638297,"Protein":23.3,"InputCost_perHa":2550,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1033.6585365853657},
  {"Plot":25,"Rep":2,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.722016223404255,"Protein":23.6,"InputCost_perHa":2950,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1062.439024390244},
  {"Plot":26,"Rep":2,"Trt":2,"Amendment":"Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.417663186382979,"Protein":23.2,"InputCost_perHa":1650,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":912.1951219512195},
  {"Plot":27,"Rep":2,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.268999694468085,"Protein":23.1,"InputCost_perHa":7250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":782.9268292682927},
  {"Plot":28,"Rep":2,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.640707778085106,"Protein":23.5,"InputCost_perHa":1500,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1012.5633802816902},
  {"Plot":29,"Rep":2,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.601502030425532,"Protein":23.1,"InputCost_perHa":17100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":18463.88888888889},
  {"Plot":30,"Rep":2,"Trt":6,"Amendment":"Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":6.842683569361702,"Protein":23.5,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":17787.777777777777},
  {"Plot":31,"Rep":2,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":7.006188607446808,"Protein":23.2,"InputCost_perHa":16850,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":17848.538011695906},
  {"Plot":32,"Rep":2,"Trt":0,"Amendment":"Control","Practice_Change":"Crop 1","Yield_t_ha":7.482897900425532,"Protein":23.2,"InputCost_perHa":0,"MachineryCost_perHa":0,"LabourCost_perHa":0,"TotalCost_perHa":694.6341463414634},

  {"Plot":33,"Rep":3,"Trt":7,"Amendment":"Lime+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.322008111702128,"Protein":23.4,"InputCost_perHa":2800,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1024.5934959349595},
  {"Plot":34,"Rep":3,"Trt":3,"Amendment":"Lime + Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.048594062978723,"Protein":23.5,"InputCost_perHa":2150,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":923.1707317073171},
  {"Plot":35,"Rep":3,"Trt":1,"Amendment":"Lime only","Practice_Change":"Crop 1","Yield_t_ha":7.468294498723404,"Protein":23.5,"InputCost_perHa":250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":790.650406504065},
  {"Plot":36,"Rep":3,"Trt":14,"Amendment":"Liquid Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":7.208078594468085,"Protein":23.5,"InputCost_perHa":6000,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":748.780487804878},
  {"Plot":37,"Rep":3,"Trt":10,"Amendment":"Lime+Deep OM+Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.82302918893617,"Protein":23.5,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1059.349593495935},
  {"Plot":38,"Rep":3,"Trt":5,"Amendment":"Lime +Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":6.866323363404255,"Protein":23.7,"InputCost_perHa":2550,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1033.6585365853657},
  {"Plot":39,"Rep":3,"Trt":11,"Amendment":"Lime+Deep OM+Gypsum CHT+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.176539666595744,"Protein":23.4,"InputCost_perHa":1100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":976.8292682926829},
  {"Plot":40,"Rep":3,"Trt":15,"Amendment":"Liquid Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":7.064219107659574,"Protein":23.4,"InputCost_perHa":7250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":782.9268292682927},
  {"Plot":41,"Rep":3,"Trt":12,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)","Practice_Change":"Crop 1","Yield_t_ha":6.519162115744681,"Protein":23.4,"InputCost_perHa":16850,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":17945.488764568763},
  {"Plot":42,"Rep":3,"Trt":8,"Amendment":"Lime+Gypsum CHT+Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.05956678787234,"Protein":23.7,"InputCost_perHa":2950,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1062.439024390244},
  {"Plot":43,"Rep":3,"Trt":6,"Amendment":"Deep OM (CP1)","Practice_Change":"Crop 1","Yield_t_ha":6.396029749361702,"Protein":23.7,"InputCost_perHa":1250,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":17787.777777777777},
  {"Plot":44,"Rep":3,"Trt":9,"Amendment":"Deep OM (CP1)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.268999694468085,"Protein":23.6,"InputCost_perHa":1500,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":1012.5633802816902},
  {"Plot":45,"Rep":3,"Trt":2,"Amendment":"Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.120229763404255,"Protein":23.7,"InputCost_perHa":1650,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":912.1951219512195},
  {"Plot":46,"Rep":3,"Trt":4,"Amendment":"Lime + Gypsum CHT","Practice_Change":"Crop 1","Yield_t_ha":7.462494905957446,"Protein":23.4,"InputCost_perHa":1900,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":908.9430894308944},
  {"Plot":47,"Rep":3,"Trt":13,"Amendment":"Deep OM (CP1) + liq. Gypsum (CHT)+PAM","Practice_Change":"Crop 1","Yield_t_ha":6.435858070212766,"Protein":23.6,"InputCost_perHa":17100,"MachineryCost_perHa":150,"LabourCost_perHa":35.71,"TotalCost_perHa":18463.88888888889},
  {"Plot":48,"Rep":3,"Trt":0,"Amendment":"Control","Practice_Change":"Crop 1","Yield_t_ha":7.269686206808511,"Protein":23.5,"InputCost_perHa":0,"MachineryCost_perHa":0,"LabourCost_perHa":0,"TotalCost_perHa":694.6341463414634}
];

// Note: the uploaded workbook has 48 rows; the embedded sample covers those rows exactly.
// If you import the workbook, you will get the full original columns (not just this reduced view).

/* =========================
   State
   ========================= */
const state = {
  source: null, // {kind:'excel'|'embedded', name:string}
  raw: [],      // array of row objects
  columns: [],  // array of column keys (deduped)
  keys: {
    treatment: null,
    yield: null,
    cost: null,
  },
  scenario: {
    pricePerTonne: 350,
    yieldScale: 1,
    costScale: 1,
  },
  ui: {
    activeTab: "import",
    baseline: null,
    sort: "delta_gm_desc",
    search: "",
  }
};

/* =========================
   DOM helpers
   ========================= */
const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

function escapeHtml(s){
  return String(s ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
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

function fmtCurrency(x, digits=0){
  if (x === null || x === undefined || Number.isNaN(x)) return "—";
  const n = Number(x);
  if (!Number.isFinite(n)) return "—";
  return n.toLocaleString(undefined, {style:"currency", currency:"USD", maximumFractionDigits: digits});
}

function toast(msg, ms=2200){
  const t = $("#toast");
  const tt = $("#toastText");
  tt.textContent = msg;
  t.hidden = false;
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { t.hidden = true; }, ms);
}

/* =========================
   Tabs
   ========================= */
function setActiveTab(tab){
  state.ui.activeTab = tab;

  $$(".tab").forEach(btn => {
    const on = btn.dataset.tab === tab;
    btn.classList.toggle("is-active", on);
    btn.setAttribute("aria-selected", on ? "true" : "false");
  });

  $$(".panel").forEach(p => {
    const on = p.dataset.panel === tab;
    p.classList.toggle("is-active", on);
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
  // header:1 => array-of-arrays
  return XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: null});
}

function rows2DToObjects(rows2d){
  const headerIdx = densestRowIndex(rows2d, 60);
  const headersRaw = rows2d[headerIdx] || [];
  const headers = dedupeHeaders(headersRaw);

  const out = [];
  for (let r=headerIdx+1; r<rows2d.length; r++){
    const row = rows2d[r] || [];
    // stop if row is mostly empty
    const nonEmpty = row.reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0);
    if (nonEmpty === 0) continue;

    const obj = {};
    for (let c=0; c<headers.length; c++){
      obj[headers[c]] = (c < row.length) ? row[c] : null;
    }

    // skip rows that look like footer blocks (very few filled cells)
    if (Object.values(obj).reduce((acc,v)=>acc + (isEmptyCell(v)?0:1), 0) < Math.max(3, Math.floor(headers.length*0.03))) {
      continue;
    }
    out.push(obj);
  }

  // also return original (deduped) column list
  return {rows: out, columns: headers};
}

function detectKeys(rows, columns){
  const cols = columns || (rows[0] ? Object.keys(rows[0]) : []);

  // Treatment key priority
  const treatmentCandidates = [
    "Amendment",
    "Treatment",
    "Trt",
    "TRT",
    "Treatment Name",
    "Practice Change",
    "Practice Change__1",
  ].filter(c => cols.includes(c));

  let treatmentKey = treatmentCandidates[0] || null;

  // yield key: prefer "Yield t/ha", then anything starting with Yield
  let yieldKey = null;
  if (cols.includes("Yield t/ha")) yieldKey = "Yield t/ha";
  else {
    const y = cols.find(c => /^yield\b/i.test(c));
    yieldKey = y || null;
  }

  // cost key: prefer "|" (as in this workbook), else cost-like column with strongest numeric coverage
  let costKey = null;
  if (cols.includes("|")) costKey = "|";
  if (!costKey){
    const costLike = cols.filter(c => /cost|\$|\/ha|per ha|ha\b/i.test(c));
    const scored = costLike.map(c => {
      const nums = rows.map(r => toNum(r[c])).filter(v => Number.isFinite(v));
      const coverage = nums.length / Math.max(1, rows.length);
      const median = nums.length ? percentile(nums, 50) : -Infinity;
      return {c, coverage, median};
    }).filter(x => x.coverage >= 0.5);

    scored.sort((a,b) => (b.median - a.median) || (b.coverage - a.coverage));
    costKey = scored[0]?.c || null;
  }

  // If we still don't have a treatment key, choose first text-heavy column
  if (!treatmentKey){
    const texty = cols.map(c => {
      const nonEmpty = rows.map(r => r[c]).filter(v => !isEmptyCell(v));
      const txt = nonEmpty.filter(v => typeof v === "string").length;
      return {c, txt, nonEmpty: nonEmpty.length};
    }).filter(x => x.nonEmpty > 0);
    texty.sort((a,b)=> (b.txt - a.txt));
    treatmentKey = texty[0]?.c || null;
  }

  return {treatmentKey, yieldKey, costKey};
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

async function importExcelArrayBuffer(buf, name="workbook.xlsx"){
  if (!window.XLSX) throw new Error("XLSX library not found (SheetJS). Check index.html script include.");
  const wb = XLSX.read(buf, {type:"array"});
  const rows2d = sheetToRows2D(wb);
  const {rows, columns} = rows2DToObjects(rows2d);

  // Store
  state.raw = rows;
  state.columns = columns;
  state.source = {kind: "excel", name};

  // Detect keys
  const {treatmentKey, yieldKey, costKey} = detectKeys(rows, columns);
  state.keys.treatment = treatmentKey;
  state.keys.yield = yieldKey;
  state.keys.cost = costKey;

  // Default baseline
  state.ui.baseline = pickDefaultBaseline();

  // Render
  toast(`Imported: ${name}`);
  renderAll();
}

/* =========================
   Embedded sample loader
   ========================= */
function loadEmbeddedLockhart(){
  // Convert embedded rows to the tool’s generic format (keeping both original and convenience keys)
  const rows = EMBEDDED_LOCKHART_ROWS.map(r => ({
    Plot: r.Plot,
    Rep: r.Rep,
    Trt: r.Trt,
    Amendment: r.Amendment,
    "Practice Change": r.Practice_Change,
    "Yield t/ha": r.Yield_t_ha,
    Protein: r.Protein,
    "Treatment Input Cost Only /Ha": r.InputCost_perHa,
    "Prototype Machinery for Adding amendments": r.MachineryCost_perHa,
    "Labour per Ha application could be included in next column": r.LabourCost_perHa,
    "|": r.TotalCost_perHa
  }));

  state.raw = rows;
  state.columns = dedupeHeaders(Object.keys(rows[0] || {}));
  state.source = {kind:"embedded", name:"Lockhart (embedded sample)"};

  const {treatmentKey, yieldKey, costKey} = detectKeys(rows, state.columns);
  state.keys.treatment = treatmentKey;
  state.keys.yield = yieldKey;
  state.keys.cost = costKey;

  state.ui.baseline = pickDefaultBaseline();
  toast("Loaded embedded Lockhart sample");
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
      return;
    } catch(e){
      // keep trying
    }
  }
  toast("Bundled XLSX not found. Use Import or load embedded sample.");
}

/* =========================
   Summaries
   ========================= */
function pickDefaultBaseline(){
  const tkey = state.keys.treatment;
  if (!tkey || !state.raw.length) return null;
  const groups = unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(s => s !== ""));
  // Prefer "Control" if present
  const control = groups.find(g => g.toLowerCase() === "control");
  return control || groups[0] || null;
}

function unique(arr){
  const s = new Set(arr);
  return Array.from(s);
}

function buildTreatmentSummary(){
  const rows = state.raw;
  const tkey = state.keys.treatment;
  const ykey = state.keys.yield;
  const ckey = state.keys.cost;

  if (!rows.length || !tkey || !ykey || !ckey) return [];

  const price = Number(state.scenario.pricePerTonne) || 0;
  const yScale = Number(state.scenario.yieldScale) || 1;
  const cScale = Number(state.scenario.costScale) || 1;

  const buckets = new Map();

  for (const r of rows){
    const name = String(r[tkey] ?? "").trim();
    if (!name) continue;

    const y = toNum(r[ykey]) * yScale;
    const c = toNum(r[ckey]) * cScale;

    if (!buckets.has(name)){
      buckets.set(name, {name, n:0, yields:[], costs:[]});
    }
    const b = buckets.get(name);
    b.n += 1;
    if (Number.isFinite(y)) b.yields.push(y);
    if (Number.isFinite(c)) b.costs.push(c);
  }

  const out = [];
  for (const b of buckets.values()){
    const yMean = mean(b.yields);
    const cMean = mean(b.costs);
    const revenue = (Number.isFinite(yMean) ? yMean : NaN) * price;
    const gm = revenue - (Number.isFinite(cMean) ? cMean : NaN);

    out.push({
      name: b.name,
      n: b.n,
      yield_mean: yMean,
      cost_mean: cMean,
      revenue_mean: revenue,
      gm_mean: gm
    });
  }

  // Baseline deltas
  const baselineName = state.ui.baseline;
  const base = out.find(x => x.name === baselineName) || null;

  for (const r of out){
    r.delta_yield = base ? (r.yield_mean - base.yield_mean) : NaN;
    r.delta_cost  = base ? (r.cost_mean  - base.cost_mean)  : NaN;
    r.delta_gm    = base ? (r.gm_mean    - base.gm_mean)    : NaN;
  }

  return out;
}

function mean(arr){
  const xs = (arr || []).filter(v => Number.isFinite(v));
  if (!xs.length) return NaN;
  return xs.reduce((a,b)=>a+b,0)/xs.length;
}

/* =========================
   Rendering
   ========================= */
function renderAll(){
  renderDatasetPill();
  renderKPIs();
  renderColumns();
  renderPreviewTable();
  renderBaselineSelect();
  renderTreatmentsTable();
  renderScenarioControls();
  renderSanityChecks();
}

function renderDatasetPill(){
  const pill = $("#datasetPill");
  if (!state.raw.length){
    pill.textContent = "No data loaded";
    pill.style.borderColor = "rgba(255,255,255,0.12)";
    return;
  }
  const src = state.source?.name || "Dataset";
  pill.textContent = `${src} · ${state.raw.length} rows`;
  pill.style.borderColor = "rgba(125,211,252,0.35)";
}

function renderKPIs(){
  $("#kpiRows").textContent = fmtInt(state.raw.length);
  const tkey = state.keys.treatment;
  const treatments = tkey ? unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean)) : [];
  $("#kpiTreatments").textContent = fmtInt(treatments.length);

  $("#kpiYieldKey").textContent = state.keys.yield || "—";
  $("#kpiCostKey").textContent = state.keys.cost || "—";

  $("#previewNote").hidden = !!state.raw.length;
  $("#previewTableWrap").hidden = !state.raw.length;
}

function renderColumns(){
  const box = $("#columnsChips");
  box.innerHTML = "";
  const cols = state.columns || [];
  if (!cols.length){
    box.innerHTML = `<div class="muted">No columns yet.</div>`;
    return;
  }
  const frag = document.createDocumentFragment();
  for (const c of cols){
    const d = document.createElement("div");
    d.className = "chip";
    d.textContent = c;
    frag.appendChild(d);
  }
  box.appendChild(frag);
}

function renderPreviewTable(){
  const wrap = $("#previewTableWrap");
  if (!state.raw.length){
    wrap.innerHTML = "";
    return;
  }
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

function renderBaselineSelect(){
  const sel = $("#baselineSelect");
  if (!sel) return;

  sel.innerHTML = "";
  const tkey = state.keys.treatment;
  if (!state.raw.length || !tkey){
    sel.innerHTML = `<option value="">Load data first</option>`;
    sel.disabled = true;
    return;
  }
  sel.disabled = false;

  const options = unique(state.raw.map(r => String(r[tkey] ?? "").trim()).filter(Boolean))
    .sort((a,b) => a.localeCompare(b));

  for (const name of options){
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    if (state.ui.baseline === name) opt.selected = true;
    sel.appendChild(opt);
  }
}

function renderTreatmentsTable(){
  const wrap = $("#treatmentsTableWrap");
  if (!wrap) return;

  if (!state.raw.length){
    wrap.innerHTML = `<div class="muted">Load data on the Import tab.</div>`;
    return;
  }
  if (!state.keys.treatment || !state.keys.yield || !state.keys.cost){
    wrap.innerHTML = `<div class="muted">Could not detect treatment/yield/cost keys. Check the Import tab diagnostics.</div>`;
    return;
  }

  let summary = buildTreatmentSummary();

  // Filter search
  const q = (state.ui.search || "").trim().toLowerCase();
  if (q){
    summary = summary.filter(r => r.name.toLowerCase().includes(q));
  }

  // Sort
  const s = state.ui.sort;
  const sorters = {
    delta_gm_desc: (a,b) => (b.delta_gm - a.delta_gm),
    gm_desc: (a,b) => (b.gm_mean - a.gm_mean),
    yield_desc: (a,b) => (b.yield_mean - a.yield_mean),
    cost_asc: (a,b) => (a.cost_mean - b.cost_mean),
    name_asc: (a,b) => a.name.localeCompare(b.name),
  };
  (sorters[s] || sorters.delta_gm_desc)(summary[0] || {}, summary[0] || {});
  summary.sort(sorters[s] || sorters.delta_gm_desc);

  const price = Number(state.scenario.pricePerTonne) || 0;

  let html = `<table>
    <thead>
      <tr>
        <th>Treatment</th>
        <th>N</th>
        <th>Yield (t/ha)</th>
        <th>Cost (/ha)</th>
        <th>Revenue (/ha)<br><span class="muted">price=${escapeHtml(price)}</span></th>
        <th>Gross margin (/ha)</th>
        <th>Δ Yield</th>
        <th>Δ Cost</th>
        <th>Δ Gross margin</th>
      </tr>
    </thead>
    <tbody>`;

  for (const r of summary){
    const isBase = (r.name === state.ui.baseline);
    const badge = isBase ? ` <span class="badge">Baseline</span>` : "";
    html += `<tr>
      <td>${escapeHtml(r.name)}${badge}</td>
      <td class="mono">${fmtInt(r.n)}</td>
      <td class="mono">${fmtNumber(r.yield_mean, 3)}</td>
      <td class="mono">${fmtCurrency(r.cost_mean, 0)}</td>
      <td class="mono">${fmtCurrency(r.revenue_mean, 0)}</td>
      <td class="mono">${fmtCurrency(r.gm_mean, 0)}</td>
      <td class="mono">${fmtNumber(r.delta_yield, 3)}</td>
      <td class="mono">${fmtCurrency(r.delta_cost, 0)}</td>
      <td class="mono">${fmtCurrency(r.delta_gm, 0)}</td>
    </tr>`;
  }

  html += `</tbody></table>`;
  wrap.innerHTML = html;
}

function renderScenarioControls(){
  const price = $("#priceInput");
  const ys = $("#yieldScaleInput");
  const cs = $("#costScaleInput");
  if (!price || !ys || !cs) return;

  price.value = state.scenario.pricePerTonne;
  ys.value = String(state.scenario.yieldScale);
  cs.value = String(state.scenario.costScale);
}

function renderSanityChecks(){
  const box = $("#sanityBox");
  if (!box) return;

  if (!state.raw.length){
    box.innerHTML = `<div class="muted">Load data to see diagnostics.</div>`;
    return;
  }

  const ykey = state.keys.yield;
  const ckey = state.keys.cost;
  const price = Number(state.scenario.pricePerTonne) || 0;
  const yScale = Number(state.scenario.yieldScale) || 1;
  const cScale = Number(state.scenario.costScale) || 1;

  const ys = state.raw.map(r => toNum(r[ykey]) * yScale).filter(Number.isFinite);
  const cs = state.raw.map(r => toNum(r[ckey]) * cScale).filter(Number.isFinite);

  const yP50 = percentile(ys, 50);
  const yP5  = percentile(ys, 5);
  const yP95 = percentile(ys, 95);

  const cP50 = percentile(cs, 50);
  const cP5  = percentile(cs, 5);
  const cP95 = percentile(cs, 95);

  // Heuristic warnings
  const warnings = [];
  if (!Number.isFinite(yP50) || yP50 <= 0) warnings.push("Yield median looks non-positive; check yield key and scaling.");
  if (!Number.isFinite(cP50) || cP50 < 0) warnings.push("Cost median looks invalid; check cost key and scaling.");
  if (Number.isFinite(yP95) && yP95 > 50) warnings.push("Yield P95 is very high for t/ha; you may need yield scaling.");
  if (Number.isFinite(cP95) && cP95 > 50000) warnings.push("Cost P95 is very high; check whether the detected column is a cost total or a machinery value field.");

  const warnHtml = warnings.length
    ? `<div class="sanity-item">
         <div class="sanity-item__title">Warnings</div>
         <div class="sanity-item__text">${warnings.map(w => `• ${escapeHtml(w)}`).join("<br>")}</div>
       </div>`
    : "";

  box.innerHTML = `
    <div class="sanity-item">
      <div class="sanity-item__title">Detected keys</div>
      <div class="sanity-item__text">
        Treatment: <code>${escapeHtml(state.keys.treatment)}</code><br>
        Yield: <code>${escapeHtml(ykey)}</code><br>
        Cost: <code>${escapeHtml(ckey)}</code>
      </div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Yield distribution (scaled)</div>
      <div class="sanity-item__text">
        P5=${escapeHtml(fmtNumber(yP5, 3))}, P50=${escapeHtml(fmtNumber(yP50, 3))}, P95=${escapeHtml(fmtNumber(yP95, 3))}
      </div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Cost distribution (scaled)</div>
      <div class="sanity-item__text">
        P5=${escapeHtml(fmtCurrency(cP5, 0))}, P50=${escapeHtml(fmtCurrency(cP50, 0))}, P95=${escapeHtml(fmtCurrency(cP95, 0))}
      </div>
    </div>

    <div class="sanity-item">
      <div class="sanity-item__title">Scenario</div>
      <div class="sanity-item__text">
        Price per tonne: <code>${escapeHtml(price)}</code><br>
        Yield scale: <code>${escapeHtml(yScale)}</code>, Cost scale: <code>${escapeHtml(cScale)}</code>
      </div>
    </div>

    ${warnHtml}
  `;
}

/* =========================
   Export
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

function downloadText(filename, text, mime="text/plain"){
  const blob = new Blob([text], {type: mime});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportProcessedRowsCsv(){
  if (!state.raw.length) return toast("Nothing to export");
  downloadText("processed_rows.csv", toCsv(state.raw, state.columns), "text/csv");
}

function exportSummaryCsv(){
  const summary = buildTreatmentSummary();
  if (!summary.length) return toast("No summary to export");
  const cols = ["name","n","yield_mean","cost_mean","revenue_mean","gm_mean","delta_yield","delta_cost","delta_gm"];
  downloadText("treatment_summary.csv", toCsv(summary, cols), "text/csv");
}

function exportStateJson(){
  const payload = {
    source: state.source,
    keys: state.keys,
    scenario: state.scenario,
    baseline: state.ui.baseline,
    sort: state.ui.sort,
    rows: state.raw,
    columns: state.columns
  };
  downloadText("decision_aid_state.json", JSON.stringify(payload, null, 2), "application/json");
}

/* =========================
   Events
   ========================= */
function bindEvents(){
  // Tabs
  $$(".tab").forEach(btn => {
    btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
  });

  // Import
  $("#fileInput").addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    try{
      const buf = await f.arrayBuffer();
      await importExcelArrayBuffer(buf, f.name);
      setActiveTab("treatments");
    } catch(err){
      console.error(err);
      toast(`Import failed: ${err.message || err}`);
    }
  });

  $("#btnLoadBundledXlsx").addEventListener("click", async () => {
    try{
      await tryLoadBundled();
      if (state.raw.length) setActiveTab("treatments");
    } catch(err){
      console.error(err);
      toast(`Bundled load failed: ${err.message || err}`);
    }
  });

  $("#btnLoadEmbedded").addEventListener("click", () => {
    loadEmbeddedLockhart();
    setActiveTab("treatments");
  });

  $("#btnReset").addEventListener("click", () => resetApp());

  // Treatments controls
  $("#baselineSelect").addEventListener("change", (e) => {
    state.ui.baseline = e.target.value;
    renderTreatmentsTable();
  });

  $("#sortSelect").addEventListener("change", (e) => {
    state.ui.sort = e.target.value;
    renderTreatmentsTable();
  });

  $("#searchInput").addEventListener("input", (e) => {
    state.ui.search = e.target.value || "";
    renderTreatmentsTable();
  });

  $("#btnToScenario").addEventListener("click", () => setActiveTab("scenarios"));

  // Scenario controls
  $("#priceInput").addEventListener("input", (e) => {
    state.scenario.pricePerTonne = Number(e.target.value || 0);
    renderTreatmentsTable();
    renderSanityChecks();
  });
  $("#yieldScaleInput").addEventListener("change", (e) => {
    state.scenario.yieldScale = Number(e.target.value || 1);
    renderTreatmentsTable();
    renderSanityChecks();
  });
  $("#costScaleInput").addEventListener("change", (e) => {
    state.scenario.costScale = Number(e.target.value || 1);
    renderTreatmentsTable();
    renderSanityChecks();
  });

  $("#btnBackToTreatments").addEventListener("click", () => setActiveTab("treatments"));
  $("#btnResetScenario").addEventListener("click", () => {
    state.scenario = {pricePerTonne: 350, yieldScale: 1, costScale: 1};
    renderScenarioControls();
    renderTreatmentsTable();
    renderSanityChecks();
    toast("Scenario reset");
  });

  // Export
  $("#btnExportProcessedCsv").addEventListener("click", exportProcessedRowsCsv);
  $("#btnExportSummaryCsv").addEventListener("click", exportSummaryCsv);
  $("#btnExportJson").addEventListener("click", exportStateJson);
}

function resetApp(){
  state.source = null;
  state.raw = [];
  state.columns = [];
  state.keys = {treatment:null, yield:null, cost:null};
  state.scenario = {pricePerTonne: 350, yieldScale: 1, costScale: 1};
  state.ui = {activeTab:"import", baseline:null, sort:"delta_gm_desc", search:""};

  // Reset inputs
  const fi = $("#fileInput");
  if (fi) fi.value = "";

  $("#searchInput").value = "";
  $("#sortSelect").value = "delta_gm_desc";

  renderAll();
  setActiveTab("import");
  toast("Reset complete");
}

/* =========================
   Boot
   ========================= */
(function boot(){
  bindEvents();
  renderAll();
  setActiveTab("import");

  // Make the tool “just work” immediately with the uploaded dataset:
  // If a bundled XLSX exists, it will load; otherwise the embedded Lockhart sample loads.
  // This avoids “inactive tabs” and “tool not working” even when file hosting isn’t set up.
  (async () => {
    try{
      await tryLoadBundled();
      if (!state.raw.length) loadEmbeddedLockhart();
      // keep user on import, but data is ready
    } catch(e){
      loadEmbeddedLockhart();
    }
  })();
})();
