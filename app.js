/* Farming CBA Decision Tool 2
   - Fully functional tabs (click + keyboard)
   - Reads uploaded Excel data (Lockhart template auto-detected) and a generic fallback
   - Computes PV benefits/costs, NPV, BCR, ROI, ranking with control alongside all treatments
   - Additional benefits/costs
   - Simulations for uncertainty
   - Excel export + copy-to-Word table
*/

/* global XLSX */

(function () {
  'use strict';

  // ---------- DOM helpers ----------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));
  const el = (tag, attrs = {}, children = []) => {
    const n = document.createElement(tag);
    Object.entries(attrs).forEach(([k, v]) => {
      if (k === 'class') n.className = v;
      else if (k === 'text') n.textContent = v;
      else if (k.startsWith('on') && typeof v === 'function') n.addEventListener(k.slice(2), v);
      else n.setAttribute(k, v);
    });
    children.forEach((c) => n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
    return n;
  };

  const toast = (msg) => {
    const t = $('#toast');
    if (!t) return;
    t.textContent = msg;
    t.classList.add('show');
    window.clearTimeout(toast._timer);
    toast._timer = window.setTimeout(() => t.classList.remove('show'), 2200);
  };

  const fmtMoney = (x) => (Number.isFinite(x) ? x.toLocaleString(undefined, { maximumFractionDigits: 0 }) : '—');
  const fmtMoney2 = (x) => (Number.isFinite(x) ? x.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—');
  const fmtNum2 = (x) => (Number.isFinite(x) ? x.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—');
  const clamp = (v, a, b) => Math.min(Math.max(v, a), b);

  // ---------- State ----------
  const state = {
    wb: null,
    fileName: '',
    detected: {
      format: '—',
      sheet: '',
      columns: {
        treatment: null,
        yield: null,
        cost: null,
      },
      rawColumns: [],
      numericColumns: [],
    },
    rows: [], // normalised rows
    treatments: [], // unique treatment labels
    control: 'Control',

    // derived stats by treatment from uploaded data
    stats: new Map(), // treatment -> {n, yieldMean, yieldSD, costMean, costSD}

    // user scenario inputs
    inputs: {
      areaHa: 100,
      horizonY: 10,
      discountRate: 0.07,
      grainPrice: 450,
      benefitPersistence: 'flat',
      decayParam: 1.5,
      baseCostTiming: 'upfront',
      costInflation: 0.0, // 0.00 = none
      overrideYieldCol: '',
      overrideCostCol: '',
    },

    // extras
    addBenefits: [], // {id, appliesTo, name, amtPerHaPerYr, startY, endY}
    addCosts: [], // {id, appliesTo, name, amtPerHa, type, years} type: upfront|annual|custom, years: "1" or "2-5" etc

    // results
    results: null, // {cols:[treatments...], rows:[{metric, values:{t:val}}], byTreatment: Map(t->{...})}
    cashflows: new Map(), // treatment -> array years: {y, ben, cost, net, df, pvnet}
    lastCalcISO: '',
  };

  // ---------- Tabs (fix “inactive tabs”) ----------
  function initTabs() {
    const tabs = $$('.tab');
    const panes = $$('.tabpane');

    const activate = (id) => {
      tabs.forEach((t) => {
        const on = t.dataset.tab === id;
        t.classList.toggle('is-active', on);
        t.setAttribute('aria-selected', on ? 'true' : 'false');
        t.tabIndex = on ? 0 : -1;
      });
      panes.forEach((p) => p.classList.toggle('is-active', p.id === id));
      const pane = $('#' + id);
      if (pane) pane.focus({ preventScroll: true });
    };

    tabs.forEach((t, idx) => {
      t.addEventListener('click', () => activate(t.dataset.tab));
      t.addEventListener('keydown', (e) => {
        const key = e.key;
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(key)) return;
        e.preventDefault();
        let j = idx;
        if (key === 'ArrowLeft') j = (idx - 1 + tabs.length) % tabs.length;
        if (key === 'ArrowRight') j = (idx + 1) % tabs.length;
        if (key === 'Home') j = 0;
        if (key === 'End') j = tabs.length - 1;
        tabs[j].focus();
        activate(tabs[j].dataset.tab);
      });
    });

    // default active already set in HTML; ensure consistent ARIA
    const active = tabs.find((t) => t.classList.contains('is-active')) || tabs[0];
    if (active) activate(active.dataset.tab);
  }

  // ---------- Excel parsing ----------
  function sheetToAOA(ws) {
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  function makeUniqueHeaders(headers) {
    const seen = new Map();
    return headers.map((h) => {
      const base = (h == null ? '' : String(h)).trim() || '(blank)';
      const n = seen.get(base) || 0;
      seen.set(base, n + 1);
      return n === 0 ? base : `${base}__${n}`;
    });
  }

  function toNumber(v) {
    if (v == null) return NaN;
    if (typeof v === 'number') return v;
    const s = String(v).trim();
    if (s === '' || s === '?' || s === 'NA' || s === 'N/A') return NaN;
    const cleaned = s.replace(/[$,%]/g, '').replace(/,/g, '');
    const x = Number(cleaned);
    return Number.isFinite(x) ? x : NaN;
  }

  function detectLockhartTemplate(wb) {
    const targetSheet = wb.SheetNames.find((n) => String(n).trim().toLowerCase() === 'input-output analysis');
    if (!targetSheet) return null;

    const ws = wb.Sheets[targetSheet];
    const aoa = sheetToAOA(ws);
    if (!aoa || aoa.length < 6) return null;

    // In the provided file, row index 3 (Excel row 4) holds headers
    const headerRowIndex = 3;
    const headersRaw = aoa[headerRowIndex] || [];
    const headers = makeUniqueHeaders(headersRaw);

    const idx = (name) => headers.findIndex((h) => String(h).toLowerCase() === name.toLowerCase());

    const colYield = headers.findIndex((h) => String(h).toLowerCase().includes('yield t/ha'));
    const colCost = headers.findIndex((h) => String(h).toLowerCase().includes('treatment input cost only /ha'));
    const colTreatment = idx('Amendment') >= 0 ? idx('Amendment') : headers.findIndex((h) => String(h).toLowerCase().includes('treatment'));

    if (colTreatment < 0 || colYield < 0) return null;

    // data starts at aoa row index 4 (Excel row 5)
    const rows = [];
    for (let r = headerRowIndex + 1; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const treat = row[colTreatment];
      const y = toNumber(row[colYield]);
      const c = colCost >= 0 ? toNumber(row[colCost]) : NaN;

      if (treat == null && !Number.isFinite(y) && !Number.isFinite(c)) continue;

      // stop if we reach an obvious blank tail (many templates have trailing blanks)
      const anyNonNull = row.some((v) => v != null && String(v).trim() !== '');
      if (!anyNonNull) continue;

      rows.push({
        treatment: treat == null ? '' : String(treat).trim(),
        yield: y,
        cost: c,
        _raw: row,
      });
    }

    const rawColumns = headers.slice();
    const numericColumns = headers
      .map((h, j) => {
        // quick numeric test: look at first 20 data rows
        let count = 0;
        let num = 0;
        for (let r = headerRowIndex + 1; r < Math.min(aoa.length, headerRowIndex + 1 + 20); r++) {
          const v = aoa[r]?.[j];
          if (v == null || String(v).trim() === '') continue;
          count++;
          if (Number.isFinite(toNumber(v))) num++;
        }
        return { h, j, count, num };
      })
      .filter((x) => x.count >= 5 && x.num / x.count >= 0.7)
      .map((x) => x.h);

    return {
      format: 'Lockhart Input-Output Analysis',
      sheet: targetSheet,
      columns: { treatment: 'Amendment', yield: headers[colYield], cost: colCost >= 0 ? headers[colCost] : null },
      rawColumns,
      numericColumns,
      rows,
      headerRowIndex,
      headers,
      aoa,
    };
  }

  function detectGenericTemplate(wb) {
    // Generic: any sheet with headers in first row and columns like treatment + yield/benefit + cost
    const sname = wb.SheetNames[0];
    const ws = wb.Sheets[sname];
    const json = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });
    if (!json || json.length === 0) return null;

    const keys = Object.keys(json[0] || {});
    const lower = keys.map((k) => String(k).toLowerCase());

    const findKey = (re) => {
      const i = lower.findIndex((k) => re.test(k));
      return i >= 0 ? keys[i] : null;
    };

    const kTreatment = findKey(/treat|amend|practice|option|arm/);
    const kYield = findKey(/yield|t\/ha|ton\/ha|t_per_ha/);
    const kBenefit = findKey(/benefit|revenue|income|gross/);
    const kCost = findKey(/cost|expense/);

    // If yield exists, we can still compute revenue via price; else use benefit column
    if (!kTreatment || (!kYield && !kBenefit) || !kCost) return null;

    const rows = json.map((r) => ({
      treatment: String(r[kTreatment] ?? '').trim(),
      yield: kYield ? toNumber(r[kYield]) : NaN,
      benefit: kBenefit ? toNumber(r[kBenefit]) : NaN,
      cost: toNumber(r[kCost]),
      _raw: r,
    })).filter((r) => r.treatment);

    return {
      format: 'Generic (treatment + yield/benefit + cost)',
      sheet: sname,
      columns: { treatment: kTreatment, yield: kYield, cost: kCost },
      rawColumns: keys.slice(),
      numericColumns: keys.filter((k) => rows.some((r) => Number.isFinite(toNumber(r._raw[k])))),
      rows,
    };
  }

  function detectAndLoadWorkbook(wb, fileName) {
    state.wb = wb;
    state.fileName = fileName || '';
    state.rows = [];
    state.treatments = [];
    state.stats = new Map();

    const lock = detectLockhartTemplate(wb);
    const det = lock || detectGenericTemplate(wb);

    if (!det) {
      state.detected = { format: 'Unknown', sheet: '', columns: { treatment: null, yield: null, cost: null }, rawColumns: [], numericColumns: [] };
      updateStatus();
      renderDetectedColumns();
      toast('Could not detect required columns. Please check your sheet format.');
      return false;
    }

    state.detected = {
      format: det.format,
      sheet: det.sheet,
      columns: det.columns,
      rawColumns: det.rawColumns,
      numericColumns: det.numericColumns,
    };

    // Normalise rows for core calculation: treatment + yield + cost (per ha)
    if (det.format.startsWith('Generic')) {
      state.rows = det.rows.map((r) => ({
        treatment: r.treatment,
        yield: Number.isFinite(r.yield) ? r.yield : NaN,
        cost: Number.isFinite(r.cost) ? r.cost : NaN,
        benefit: Number.isFinite(r.benefit) ? r.benefit : NaN,
      }));
    } else {
      state.rows = det.rows.map((r) => ({
        treatment: r.treatment,
        yield: r.yield,
        cost: r.cost,
      }));
    }

    // Treatments
    const set = new Set();
    state.rows.forEach((r) => { if (r.treatment) set.add(r.treatment); });
    state.treatments = Array.from(set).sort((a, b) => a.localeCompare(b));

    // Default control if present
    const guess = state.treatments.find((t) => t.toLowerCase() === 'control') ||
                  state.treatments.find((t) => t.toLowerCase().includes('control')) ||
                  state.treatments[0] || 'Control';
    state.control = guess;

    // Stats by treatment
    computeTreatmentStats();

    updateStatus();
    renderControlSelect();
    renderDetectedColumns();
    renderTreatmentSummary();
    populateOverrides();
    populateTreatmentSelectors();
    return true;
  }

  function computeTreatmentStats() {
    const groups = new Map();
    state.rows.forEach((r) => {
      if (!r.treatment) return;
      if (!groups.has(r.treatment)) groups.set(r.treatment, []);
      groups.get(r.treatment).push(r);
    });

    const mean = (arr) => arr.reduce((a, b) => a + b, 0) / (arr.length || 1);
    const sd = (arr) => {
      if (arr.length < 2) return 0;
      const m = mean(arr);
      const v = arr.reduce((s, x) => s + (x - m) * (x - m), 0) / (arr.length - 1);
      return Math.sqrt(v);
    };

    state.stats = new Map();
    for (const [t, rows] of groups.entries()) {
      const y = rows.map((r) => r.yield).filter(Number.isFinite);
      const c = rows.map((r) => r.cost).filter(Number.isFinite);
      state.stats.set(t, {
        n: rows.length,
        yieldMean: y.length ? mean(y) : NaN,
        yieldSD: y.length ? sd(y) : NaN,
        costMean: c.length ? mean(c) : NaN,
        costSD: c.length ? sd(c) : NaN,
      });
    }
  }

  // ---------- UI rendering ----------
  function updateStatus() {
    $('#statusData').textContent = state.rows.length ? 'Loaded' : 'No file loaded';
    $('#statusFormat').textContent = state.detected.format || '—';
    $('#statusTreatments').textContent = state.treatments.length ? `${state.treatments.length} detected` : '—';
    $('#fileMeta').textContent = state.fileName ? `File: ${state.fileName} · Sheet: ${state.detected.sheet || '—'}` : '';
    $('#treatmentCount').value = state.treatments.length ? String(state.treatments.length) : '';
  }

  function renderControlSelect() {
    const sel = $('#controlSelect');
    sel.innerHTML = '';
    state.treatments.forEach((t) => sel.appendChild(el('option', { value: t, text: t })));
    sel.value = state.control;

    sel.addEventListener('change', () => {
      state.control = sel.value;
      recalcAll();
      toast('Control updated');
    });
  }

  function renderDetectedColumns() {
    const tbody = $('#detectedTable tbody');
    tbody.innerHTML = '';

    const add = (role, col, notes) => {
      tbody.appendChild(el('tr', {}, [
        el('td', { text: role }),
        el('td', { text: col || '—' }),
        el('td', { text: notes }),
      ]));
    };

    add('Treatment label', state.detected.columns.treatment, 'Used to group rows into treatments (including Control).');
    add('Yield', state.detected.columns.yield, 'Used to estimate income changes via grain price (unless benefit provided).');
    add('Cost (per ha)', state.detected.columns.cost, 'Used as treatment input cost; timing controlled under Assumptions.');

    // overrides selected
    const oy = state.inputs.overrideYieldCol || '—';
    const oc = state.inputs.overrideCostCol || '—';
    add('Yield override', oy, 'Optional override chosen in Assumptions tab.');
    add('Cost override', oc, 'Optional override chosen in Assumptions tab.');
  }

  function renderTreatmentSummary() {
    const tbody = $('#treatmentSummaryTable tbody');
    tbody.innerHTML = '';

    const treatments = state.treatments.slice();
    treatments.forEach((t) => {
      const s = state.stats.get(t) || {};
      tbody.appendChild(el('tr', {}, [
        el('td', { text: t }),
        el('td', { text: String(s.n ?? '—') }),
        el('td', { text: fmtNum2(s.yieldMean) }),
        el('td', { text: fmtNum2(s.yieldSD) }),
        el('td', { text: Number.isFinite(s.costMean) ? fmtMoney2(s.costMean) : '—' }),
        el('td', { text: Number.isFinite(s.costSD) ? fmtMoney2(s.costSD) : '—' }),
      ]));
    });
  }

  function populateOverrides() {
    const ySel = $('#overrideYieldCol');
    const cSel = $('#overrideCostCol');
    ySel.innerHTML = '';
    cSel.innerHTML = '';

    const opts = [''].concat(state.detected.numericColumns || []);
    opts.forEach((o) => {
      ySel.appendChild(el('option', { value: o, text: o === '' ? '(use detected)' : o }));
      cSel.appendChild(el('option', { value: o, text: o === '' ? '(use detected)' : o }));
    });

    ySel.value = state.inputs.overrideYieldCol || '';
    cSel.value = state.inputs.overrideCostCol || '';

    ySel.addEventListener('change', () => {
      state.inputs.overrideYieldCol = ySel.value;
      renderDetectedColumns();
      recalcAll();
    });
    cSel.addEventListener('change', () => {
      state.inputs.overrideCostCol = cSel.value;
      renderDetectedColumns();
      recalcAll();
    });
  }

  function populateTreatmentSelectors() {
    const cashSel = $('#cashflowTreatment');
    const simSel = $('#simTreatment');
    [cashSel, simSel].forEach((s) => { s.innerHTML = ''; });

    state.treatments.forEach((t) => {
      cashSel.appendChild(el('option', { value: t, text: t }));
      simSel.appendChild(el('option', { value: t, text: t }));
    });

    cashSel.value = state.treatments.includes(state.control) ? state.control : (state.treatments[0] || '');
    simSel.value = state.treatments.find((t) => t !== state.control) || state.control || (state.treatments[0] || '');

    cashSel.addEventListener('change', () => renderCashflowTable());
    $('#cashflowMode').addEventListener('change', () => renderCashflowTable());
  }

  // ---------- Scenario inputs ----------
  function readInputsFromUI() {
    state.inputs.areaHa = clamp(toNumber($('#areaHa').value), 0, 1e9);
    state.inputs.horizonY = Math.max(1, Math.round(clamp(toNumber($('#horizonY').value), 1, 200)));
    state.inputs.discountRate = clamp(toNumber($('#discountRate').value) / 100, 0, 1);
    state.inputs.grainPrice = clamp(toNumber($('#grainPrice').value), 0, 1e9);
    state.inputs.benefitPersistence = $('#benefitPersistence').value;
    state.inputs.decayParam = clamp(toNumber($('#decayParam').value), 0, 50);
    state.inputs.baseCostTiming = $('#baseCostTiming').value;
    state.inputs.costInflation = clamp(toNumber($('#costInflation').value) / 100, -0.5, 1);
  }

  function fillInputsToUI() {
    $('#areaHa').value = state.inputs.areaHa;
    $('#horizonY').value = state.inputs.horizonY;
    $('#discountRate').value = (state.inputs.discountRate * 100).toFixed(1).replace(/\.0$/, '');
    $('#grainPrice').value = state.inputs.grainPrice;
    $('#benefitPersistence').value = state.inputs.benefitPersistence;
    $('#decayParam').value = state.inputs.decayParam;
    $('#baseCostTiming').value = state.inputs.baseCostTiming;
    $('#costInflation').value = (state.inputs.costInflation * 100).toFixed(1).replace(/\.0$/, '');
  }

  // ---------- Core CBA math ----------
  function discountFactor(year, r) {
    return 1 / Math.pow(1 + r, year - 1);
  }

  function persistenceMultiplier(year, horizon, mode, param) {
    if (horizon <= 1) return 1;
    const t = (year - 1) / (horizon - 1); // 0..1
    if (mode === 'flat') return 1;
    if (mode === 'linearDecay') return 1 - t;
    if (mode === 'expDecay') return Math.exp(-param * t);
    if (mode === 'rampUp') {
      // ramp from 0 to 1 quickly then stay near 1
      const k = Math.max(0.0001, param);
      return 1 - Math.exp(-k * t);
    }
    return 1;
  }

  function parseYearSpec(spec, horizon) {
    // "1" or "2-5" or "all"
    const s = String(spec || '').trim().toLowerCase();
    if (!s || s === 'all') return { start: 1, end: horizon };
    if (/^\d+$/.test(s)) {
      const y = Math.max(1, Math.min(horizon, Number(s)));
      return { start: y, end: y };
    }
    const m = s.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) {
      let a = Number(m[1]), b = Number(m[2]);
      a = Math.max(1, Math.min(horizon, a));
      b = Math.max(1, Math.min(horizon, b));
      return { start: Math.min(a, b), end: Math.max(a, b) };
    }
    return { start: 1, end: horizon };
  }

  function additionalBenefitsByYear(treatment) {
    const H = state.inputs.horizonY;
    const out = new Array(H).fill(0);
    state.addBenefits.forEach((b) => {
      if (!(b.appliesTo === 'ALL' || b.appliesTo === treatment)) return;
      const a = clamp(toNumber(b.amtPerHaPerYr), -1e9, 1e9);
      const s = Math.max(1, Math.min(H, Math.round(toNumber(b.startY) || 1)));
      const e = Math.max(1, Math.min(H, Math.round(toNumber(b.endY) || H)));
      for (let y = s; y <= e; y++) out[y - 1] += a;
    });
    return out;
  }

  function additionalCostsByYear(treatment) {
    const H = state.inputs.horizonY;
    const out = new Array(H).fill(0);
    state.addCosts.forEach((c) => {
      if (!(c.appliesTo === 'ALL' || c.appliesTo === treatment)) return;
      const amt = clamp(toNumber(c.amtPerHa), -1e9, 1e9);
      const type = c.type || 'upfront';
      if (type === 'upfront') {
        out[0] += amt;
      } else if (type === 'annual') {
        for (let y = 1; y <= H; y++) out[y - 1] += amt;
      } else {
        const yr = parseYearSpec(c.years, H);
        for (let y = yr.start; y <= yr.end; y++) out[y - 1] += amt;
      }
    });
    return out;
  }

  function baseCostStream(costPerHa) {
    const H = state.inputs.horizonY;
    const inf = state.inputs.costInflation;
    const timing = state.inputs.baseCostTiming;

    const out = new Array(H).fill(0);
    const c = Number.isFinite(costPerHa) ? costPerHa : 0;

    if (timing === 'upfront') out[0] = c;
    else if (timing === 'annual') for (let y = 1; y <= H; y++) out[y - 1] = c;
    else {
      // amortise evenly across years
      const per = c / H;
      for (let y = 1; y <= H; y++) out[y - 1] = per;
    }

    // apply inflation
    if (inf !== 0) {
      for (let y = 1; y <= H; y++) {
        out[y - 1] = out[y - 1] * Math.pow(1 + inf, y - 1);
      }
    }
    return out;
  }

  function buildCashflowForTreatment(treatment) {
    const H = state.inputs.horizonY;
    const r = state.inputs.discountRate;
    const area = state.inputs.areaHa;
    const price = state.inputs.grainPrice;

    const ctrl = state.control;

    const sT = state.stats.get(treatment) || {};
    const sC = state.stats.get(ctrl) || {};

    const yT = Number.isFinite(sT.yieldMean) ? sT.yieldMean : NaN;
    const yC = Number.isFinite(sC.yieldMean) ? sC.yieldMean : NaN;

    // Yield benefit difference (t/ha) vs control
    const deltaYield = (Number.isFinite(yT) && Number.isFinite(yC)) ? (yT - yC) : 0;

    // Base cost difference vs control
    const cT = Number.isFinite(sT.costMean) ? sT.costMean : 0;
    const cC = Number.isFinite(sC.costMean) ? sC.costMean : 0;
    const deltaCost = cT - cC;

    const benExtra = additionalBenefitsByYear(treatment);
    const costExtra = additionalCostsByYear(treatment);

    const baseCost = baseCostStream(deltaCost);

    const rows = [];
    let pvBen = 0, pvCost = 0, pvNet = 0;

    for (let y = 1; y <= H; y++) {
      const mult = persistenceMultiplier(y, H, state.inputs.benefitPersistence, state.inputs.decayParam);

      // benefit per ha from yield difference
      const benYieldPerHa = deltaYield * price * mult;

      // add extra benefits
      const benPerHa = benYieldPerHa + (benExtra[y - 1] || 0);

      // costs per ha: base + extra
      const costPerHa = (baseCost[y - 1] || 0) + (costExtra[y - 1] || 0);

      const netPerHa = benPerHa - costPerHa;

      const df = discountFactor(y, r);
      const pvBenY = benPerHa * df;
      const pvCostY = costPerHa * df;
      const pvNetY = netPerHa * df;

      pvBen += pvBenY;
      pvCost += pvCostY;
      pvNet += pvNetY;

      rows.push({
        y,
        benPerHa,
        costPerHa,
        netPerHa,
        df,
        pvNetPerHa: pvNetY,
        // for whole farm preview
        benFarm: benPerHa * area,
        costFarm: costPerHa * area,
        netFarm: netPerHa * area,
        pvNetFarm: pvNetY * area,
      });
    }

    const NPV = pvNet;
    const PVBenefits = pvBen;
    const PVCosts = pvCost;
    const BCR = (PVBenefits >= 0 && PVCosts > 0) ? (PVBenefits / PVCosts) : (PVCosts === 0 ? (PVBenefits > 0 ? Infinity : 0) : PVBenefits / PVCosts);
    const ROI = (PVCosts !== 0) ? (NPV / PVCosts) : (NPV > 0 ? Infinity : 0);

    return {
      treatment,
      deltaYield,
      deltaCost,
      PVBenefits,
      PVCosts,
      NPV,
      BCR,
      ROI,
      rows,
    };
  }

  function recalcAll() {
    if (!state.rows.length || !state.treatments.length) return;

    readInputsFromUI();
    renderDetectedColumns();

    // Build cashflows for each treatment (including control)
    state.cashflows = new Map();
    const byT = new Map();

    state.treatments.forEach((t) => {
      const res = buildCashflowForTreatment(t);
      byT.set(t, res);
      state.cashflows.set(t, res.rows);
    });

    // Results table layout: indicators as rows, treatments as columns
    const cols = state.treatments.slice(); // includes control
    const metrics = [
      { key: 'PVBenefits', label: 'Present value of benefits (PV benefits)' },
      { key: 'PVCosts', label: 'Present value of costs (PV costs)' },
      { key: 'NPV', label: 'Net present value (NPV)' },
      { key: 'BCR', label: 'Benefit–cost ratio (BCR)' },
      { key: 'ROI', label: 'Return on investment (ROI)' },
      { key: 'Rank', label: 'Ranking (by NPV, highest = 1)' },
    ];

    // rank by NPV (descending)
    const ranked = cols
      .map((t) => ({ t, npv: byT.get(t)?.NPV ?? -Infinity }))
      .sort((a, b) => (b.npv - a.npv));
    const rankMap = new Map();
    ranked.forEach((x, i) => rankMap.set(x.t, i + 1));

    const tableRows = metrics.map((m) => {
      const values = {};
      cols.forEach((t) => {
        if (m.key === 'Rank') values[t] = rankMap.get(t);
        else values[t] = byT.get(t)?.[m.key];
      });
      return { metric: m.label, key: m.key, values };
    });

    state.results = { cols, rows: tableRows, byTreatment: byT, rankMap };
    state.lastCalcISO = new Date().toISOString();
    $('#lastCalc').textContent = `Last calculated: ${new Date().toLocaleString()}`;

    renderResultsTable();
    renderDriverNotes();
    renderCashflowTable();
    updateStatus();
  }

  // ---------- Results rendering ----------
  function renderResultsTable() {
    const tbl = $('#resultsTable');
    const thead = tbl.querySelector('thead');
    const tbody = tbl.querySelector('tbody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!state.results) return;

    const cols = state.results.cols;

    const trh = el('tr');
    trh.appendChild(el('th', { text: 'Indicator' }));
    cols.forEach((t) => trh.appendChild(el('th', { text: t })));
    thead.appendChild(trh);

    state.results.rows.forEach((r) => {
      const tr = el('tr');
      tr.appendChild(el('td', { text: r.metric }));

      cols.forEach((t) => {
        let v = r.values[t];
        let text = '—';
        if (r.key === 'BCR' || r.key === 'ROI') {
          text = Number.isFinite(v) ? fmtNum2(v) : (v === Infinity ? '∞' : '—');
        } else if (r.key === 'Rank') {
          text = Number.isFinite(v) ? String(v) : '—';
        } else {
          text = Number.isFinite(v) ? fmtMoney(v) : '—';
        }
        tr.appendChild(el('td', { text }));
      });

      tbody.appendChild(tr);
    });
  }

  function renderDriverNotes() {
    const box = $('#driverNotes');
    box.innerHTML = '';
    if (!state.results) return;

    const ctrl = state.control;
    const byT = state.results.byTreatment;

    const ctrlRes = byT.get(ctrl);
    const price = state.inputs.grainPrice;

    const items = state.treatments
      .filter((t) => t !== ctrl)
      .map((t) => {
        const r = byT.get(t);
        if (!r) return null;

        // contribution intuition: benefit from yield difference vs cost difference (PV-level, roughly)
        const dy = r.deltaYield;
        const dc = r.deltaCost;

        const yieldSign = dy > 0 ? 'higher' : (dy < 0 ? 'lower' : 'similar');
        const costSign = dc > 0 ? 'higher' : (dc < 0 ? 'lower' : 'similar');

        const lines = [];
        lines.push(`Compared with the control, this option has ${yieldSign} yield (Δ ${fmtNum2(dy)} t/ha).`);
        lines.push(`Using grain price $${fmtMoney(price)}/t, that shifts income in the same direction as the yield difference.`);
        if (Number.isFinite(dc)) lines.push(`Its treatment input cost is ${costSign} than the control (Δ ${fmtMoney2(dc)} per ha, applied as ${state.inputs.baseCostTiming}).`);
        lines.push(`Overall, the NPV is ${Number.isFinite(r.NPV) ? (r.NPV >= 0 ? 'positive' : 'negative') : 'not available'} and it ranks ${state.results.rankMap.get(t) ?? '—'} by NPV in this scenario.`);

        // improvement levers if weak
        if (Number.isFinite(r.NPV) && r.NPV < 0) {
          lines.push('If you wanted to improve performance, options include lowering treatment costs, targeting paddocks where yield response is more likely, or using a more conservative benefit persistence setting.');
        }

        return { t, lines };
      })
      .filter(Boolean);

    if (!items.length) {
      box.appendChild(el('div', { class: 'noteitem' }, [
        el('div', { class: 'noteitem__title', text: 'No driver notes yet' }),
        el('div', { class: 'noteitem__text', text: 'Load data and recalculate to generate interpretation notes.' }),
      ]));
      return;
    }

    items.forEach((it) => {
      box.appendChild(el('div', { class: 'noteitem' }, [
        el('div', { class: 'noteitem__title', text: it.t }),
        el('div', { class: 'noteitem__text', text: it.lines.join(' ') }),
      ]));
    });

    // also mention control
    if (ctrlRes) {
      box.appendChild(el('div', { class: 'noteitem' }, [
        el('div', { class: 'noteitem__title', text: 'Control (baseline)' }),
        el('div', { class: 'noteitem__text', text: 'The control is included as a column so all figures line up for side-by-side comparison. Differences versus control drive the treatment cashflows.' }),
      ]));
    }
  }

  function renderCashflowTable() {
    const tbody = $('#cashflowTable tbody');
    tbody.innerHTML = '';
    if (!state.cashflows || !state.cashflows.size) return;

    const t = $('#cashflowTreatment').value;
    const mode = $('#cashflowMode').value;
    const rows = state.cashflows.get(t) || [];

    rows.forEach((r) => {
      const ben = mode === 'wholeFarm' ? r.benFarm : r.benPerHa;
      const cost = mode === 'wholeFarm' ? r.costFarm : r.costPerHa;
      const net = mode === 'wholeFarm' ? r.netFarm : r.netPerHa;
      const pvnet = mode === 'wholeFarm' ? r.pvNetFarm : r.pvNetPerHa;

      tbody.appendChild(el('tr', {}, [
        el('td', { text: String(r.y) }),
        el('td', { text: fmtMoney2(ben) }),
        el('td', { text: fmtMoney2(cost) }),
        el('td', { text: fmtMoney2(net) }),
        el('td', { text: fmtNum2(r.df) }),
        el('td', { text: fmtMoney2(pvnet) }),
      ]));
    });
  }

  // ---------- Additional benefits/costs tables ----------
  function renderBenefitsTable() {
    const tbody = $('#benefitsTable tbody');
    tbody.innerHTML = '';

    state.addBenefits.forEach((b) => {
      const tr = el('tr', {}, [
        tdSelectApplies(b.appliesTo, (v) => { b.appliesTo = v; recalcAll(); }),
        el('td', {}, [inputText(b.name, (v) => { b.name = v; })]),
        el('td', {}, [inputNum(b.amtPerHaPerYr, (v) => { b.amtPerHaPerYr = v; recalcAll(); })]),
        el('td', {}, [inputNum(b.startY, (v) => { b.startY = v; recalcAll(); })]),
        el('td', {}, [inputNum(b.endY, (v) => { b.endY = v; recalcAll(); })]),
        el('td', {}, [btnRemove(() => { state.addBenefits = state.addBenefits.filter(x => x.id !== b.id); renderBenefitsTable(); recalcAll(); })]),
      ]);
      tbody.appendChild(tr);
    });
  }

  function renderCostsTable() {
    const tbody = $('#costsTable tbody');
    tbody.innerHTML = '';

    state.addCosts.forEach((c) => {
      const tr = el('tr', {}, [
        tdSelectApplies(c.appliesTo, (v) => { c.appliesTo = v; recalcAll(); }),
        el('td', {}, [inputText(c.name, (v) => { c.name = v; })]),
        el('td', {}, [inputNum(c.amtPerHa, (v) => { c.amtPerHa = v; recalcAll(); })]),
        el('td', {}, [selectType(c.type, (v) => { c.type = v; renderCostsTable(); recalcAll(); })]),
        el('td', {}, [typeYearsCell(c)]),
        el('td', {}, [btnRemove(() => { state.addCosts = state.addCosts.filter(x => x.id !== c.id); renderCostsTable(); recalcAll(); })]),
      ]);
      tbody.appendChild(tr);
    });
  }

  function tdSelectApplies(current, onChange) {
    const s = el('select');
    s.appendChild(el('option', { value: 'ALL', text: 'ALL' }));
    state.treatments.forEach((t) => s.appendChild(el('option', { value: t, text: t })));
    s.value = current || 'ALL';
    s.addEventListener('change', () => onChange(s.value));
    return el('td', {}, [s]);
  }

  function inputText(val, onChange) {
    const i = el('input', { type: 'text', value: val || '' });
    i.addEventListener('input', () => onChange(i.value));
    return i;
  }

  function inputNum(val, onChange) {
    const i = el('input', { type: 'number', step: '0.01', value: val == null ? '' : val });
    i.addEventListener('input', () => onChange(i.value));
    return i;
  }

  function selectType(val, onChange) {
    const s = el('select');
    s.appendChild(el('option', { value: 'upfront', text: 'Upfront (Year 1)' }));
    s.appendChild(el('option', { value: 'annual', text: 'Annual (every year)' }));
    s.appendChild(el('option', { value: 'custom', text: 'Custom years' }));
    s.value = val || 'upfront';
    s.addEventListener('change', () => onChange(s.value));
    return s;
  }

  function typeYearsCell(costItem) {
    if ((costItem.type || 'upfront') !== 'custom') {
      return el('div', { class: 'muted', text: costItem.type === 'annual' ? 'All years' : 'Year 1' });
    }
    const i = el('input', { type: 'text', value: costItem.years || '1', placeholder: 'e.g. 2-5 or all' });
    i.addEventListener('input', () => { costItem.years = i.value; recalcAll(); });
    return i;
  }

  function btnRemove(onClick) {
    return el('button', { class: 'btn btn--ghost', type: 'button', text: 'Remove', onclick: onClick });
  }

  // ---------- Simulations ----------
  function pooledYieldSD() {
    const vals = [];
    state.stats.forEach((s) => {
      if (Number.isFinite(s.yieldSD) && s.yieldSD > 0) vals.push(s.yieldSD);
    });
    if (!vals.length) return 0;
    const m = vals.reduce((a, b) => a + b, 0) / vals.length;
    return m;
  }

  function normalRand() {
    // Box–Muller
    let u = 0, v = 0;
    while (u === 0) u = Math.random();
    while (v === 0) v = Math.random();
    return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
  }

  function runSim() {
    if (!state.results) recalcAll();
    if (!state.results) return;

    const N = Math.max(200, Math.round(toNumber($('#simN').value) || 2000));
    const priceSDpct = Math.max(0, toNumber($('#simPriceSD').value) || 0) / 100;
    const costSDpct = Math.max(0, toNumber($('#simCostSDPct').value) || 0) / 100;
    const yieldMode = $('#simYieldSDMode').value;
    const fixedPct = Math.max(0, toNumber($('#simFixedYieldPct').value) || 0) / 100;

    const t = $('#simTreatment').value;
    const ctrl = state.control;

    const sT = state.stats.get(t) || {};
    const sC = state.stats.get(ctrl) || {};

    const meanYieldT = Number.isFinite(sT.yieldMean) ? sT.yieldMean : 0;
    const meanYieldC = Number.isFinite(sC.yieldMean) ? sC.yieldMean : 0;

    const meanCostT = Number.isFinite(sT.costMean) ? sT.costMean : 0;
    const meanCostC = Number.isFinite(sC.costMean) ? sC.costMean : 0;

    let sdYieldT = 0, sdYieldC = 0;
    if (yieldMode === 'fromData') {
      sdYieldT = Number.isFinite(sT.yieldSD) ? sT.yieldSD : 0;
      sdYieldC = Number.isFinite(sC.yieldSD) ? sC.yieldSD : 0;
    } else if (yieldMode === 'pooled') {
      sdYieldT = pooledYieldSD();
      sdYieldC = pooledYieldSD();
    } else {
      sdYieldT = meanYieldT * fixedPct;
      sdYieldC = meanYieldC * fixedPct;
    }

    const price0 = state.inputs.grainPrice;
    const H = state.inputs.horizonY;
    const r = state.inputs.discountRate;

    const baseCostTiming = state.inputs.baseCostTiming;
    const benefitPersistence = state.inputs.benefitPersistence;
    const decayParam = state.inputs.decayParam;

    const benExtra = additionalBenefitsByYear(t);
    const costExtra = additionalCostsByYear(t);

    const npvs = [];
    for (let i = 0; i < N; i++) {
      const price = Math.max(0, price0 * (1 + normalRand() * priceSDpct));

      const yT = Math.max(0, meanYieldT + normalRand() * sdYieldT);
      const yC = Math.max(0, meanYieldC + normalRand() * sdYieldC);
      const deltaYield = yT - yC;

      const cT = meanCostT * (1 + normalRand() * costSDpct);
      const cC = meanCostC * (1 + normalRand() * costSDpct);
      const deltaCost = cT - cC;

      const baseCost = (() => {
        const out = new Array(H).fill(0);
        if (baseCostTiming === 'upfront') out[0] = deltaCost;
        else if (baseCostTiming === 'annual') for (let y = 1; y <= H; y++) out[y - 1] = deltaCost;
        else {
          const per = deltaCost / H;
          for (let y = 1; y <= H; y++) out[y - 1] = per;
        }
        // inflation ignored in sim to keep it transparent; already available deterministically
        return out;
      })();

      let pvNet = 0;
      for (let y = 1; y <= H; y++) {
        const mult = persistenceMultiplier(y, H, benefitPersistence, decayParam);
        const benPerHa = deltaYield * price * mult + (benExtra[y - 1] || 0);
        const costPerHa = (baseCost[y - 1] || 0) + (costExtra[y - 1] || 0);
        const net = benPerHa - costPerHa;
        pvNet += net * discountFactor(y, r);
      }
      npvs.push(pvNet);
    }

    npvs.sort((a, b) => a - b);
    const q = (p) => npvs[Math.min(npvs.length - 1, Math.max(0, Math.floor(p * (npvs.length - 1))))];

    const mean = npvs.reduce((a, b) => a + b, 0) / npvs.length;
    const pPos = npvs.filter((x) => x > 0).length / npvs.length;

    const tbody = $('#simTable tbody');
    tbody.innerHTML = '';
    const addRow = (k, v) => tbody.appendChild(el('tr', {}, [el('td', { text: k }), el('td', { text: v })]));

    addRow('Treatment', t);
    addRow('Control', ctrl);
    addRow('Iterations', String(N));
    addRow('Mean NPV (PV net, per ha)', fmtMoney2(mean));
    addRow('Median NPV (per ha)', fmtMoney2(q(0.5)));
    addRow('5th percentile NPV (per ha)', fmtMoney2(q(0.05)));
    addRow('95th percentile NPV (per ha)', fmtMoney2(q(0.95)));
    addRow('Probability NPV > 0', `${(pPos * 100).toFixed(1)}%`);

    toast('Simulations complete');
  }

  // ---------- Copy + Export ----------
  function buildResultsTSV() {
    if (!state.results) return '';
    const cols = ['Indicator'].concat(state.results.cols);
    const lines = [cols.join('\t')];

    state.results.rows.forEach((r) => {
      const row = [r.metric];
      state.results.cols.forEach((t) => {
        const v = r.values[t];
        if (r.key === 'BCR' || r.key === 'ROI') row.push(Number.isFinite(v) ? String(v.toFixed(2)) : (v === Infinity ? 'Infinity' : ''));
        else if (r.key === 'Rank') row.push(Number.isFinite(v) ? String(v) : '');
        else row.push(Number.isFinite(v) ? String(Math.round(v)) : '');
      });
      lines.push(row.join('\t'));
    });

    return lines.join('\n');
  }

  async function copyResultsTable() {
    const tsv = buildResultsTSV();
    if (!tsv) return;

    try {
      await navigator.clipboard.writeText(tsv);
      toast('Results copied (TSV for Word/Excel)');
    } catch {
      // fallback
      const ta = el('textarea', { style: 'position:fixed;left:-9999px;top:-9999px' });
      ta.value = tsv;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
      toast('Results copied');
    }
  }

  function downloadResultsExcel() {
    if (!state.results) recalcAll();
    if (!state.results) return;

    const wb = XLSX.utils.book_new();

    // Results sheet (vertical table)
    const aoa = [];
    aoa.push(['Indicator'].concat(state.results.cols));
    state.results.rows.forEach((r) => {
      const row = [r.metric];
      state.results.cols.forEach((t) => row.push(r.values[t]));
      aoa.push(row);
    });
    const wsResults = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, wsResults, 'Results');

    // Assumptions
    const A = state.inputs;
    const wsAss = XLSX.utils.aoa_to_sheet([
      ['Assumption', 'Value'],
      ['Farm area (ha)', A.areaHa],
      ['Time horizon (years)', A.horizonY],
      ['Discount rate', A.discountRate],
      ['Grain price ($/t)', A.grainPrice],
      ['Benefit persistence', A.benefitPersistence],
      ['Persistence parameter', A.decayParam],
      ['Base cost timing', A.baseCostTiming],
      ['Cost inflation', A.costInflation],
      ['Control', state.control],
    ]);
    XLSX.utils.book_append_sheet(wb, wsAss, 'Assumptions');

    // Treatment summary
    const sum = [['Treatment', 'n', 'Yield mean', 'Yield SD', 'Cost mean', 'Cost SD']];
    state.treatments.forEach((t) => {
      const s = state.stats.get(t) || {};
      sum.push([t, s.n ?? null, s.yieldMean ?? null, s.yieldSD ?? null, s.costMean ?? null, s.costSD ?? null]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sum), 'TreatmentSummary');

    // Additional benefits/costs
    const bAOA = [['Applies to', 'Name', 'Amount ($/ha/year)', 'Start year', 'End year']];
    state.addBenefits.forEach((b) => bAOA.push([b.appliesTo, b.name, toNumber(b.amtPerHaPerYr), toNumber(b.startY), toNumber(b.endY)]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bAOA), 'AddBenefits');

    const cAOA = [['Applies to', 'Name', 'Amount ($/ha)', 'Type', 'Years']];
    state.addCosts.forEach((c) => cAOA.push([c.appliesTo, c.name, toNumber(c.amtPerHa), c.type, c.years]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cAOA), 'AddCosts');

    const fname = `Farming_CBA_Results_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, fname);
    toast('Downloaded results workbook');
  }

  // ---------- Copilot prompt builder ----------
  function buildBriefPrompt() {
    if (!state.results) recalcAll();
    if (!state.results) return '';

    const out = {
      tool: 'Farming CBA Decision Tool 2',
      purpose: 'Draft a policy or farm decision brief comparing treatments against a control using discounted cashflow results.',
      scenario: {
        control: state.control,
        areaHa: state.inputs.areaHa,
        horizonYears: state.inputs.horizonY,
        discountRate: state.inputs.discountRate,
        grainPrice: state.inputs.grainPrice,
        benefitPersistence: state.inputs.benefitPersistence,
        persistenceParameter: state.inputs.decayParam,
        baseCostTiming: state.inputs.baseCostTiming,
        costInflation: state.inputs.costInflation,
      },
      data_detection: {
        format: state.detected.format,
        sheet: state.detected.sheet,
        columns: state.detected.columns,
        overrides: { yield: state.inputs.overrideYieldCol || null, cost: state.inputs.overrideCostCol || null },
      },
      treatment_summary: state.treatments.map((t) => {
        const s = state.stats.get(t) || {};
        return { treatment: t, n: s.n, yieldMean: s.yieldMean, yieldSD: s.yieldSD, costMean: s.costMean, costSD: s.costSD };
      }),
      results_table: (() => {
        const rows = [];
        state.results.rows.forEach((r) => {
          const obj = { indicator: r.metric };
          state.results.cols.forEach((t) => { obj[t] = r.values[t]; });
          rows.push(obj);
        });
        return rows;
      })(),
      ranking: state.treatments.map((t) => ({ treatment: t, rankByNPV: state.results.rankMap.get(t) })),
      interpretation_notes: (() => {
        const notes = [];
        const ctrl = state.control;
        const byT = state.results.byTreatment;
        state.treatments.forEach((t) => {
          const r = byT.get(t);
          if (!r) return;
          notes.push({
            treatment: t,
            deltaYield_vs_control_t_per_ha: r.deltaYield,
            deltaCost_vs_control_per_ha: r.deltaCost,
            NPV: r.NPV,
            PVBenefits: r.PVBenefits,
            PVCosts: r.PVCosts,
            BCR: r.BCR,
            ROI: r.ROI,
            isControl: t === ctrl,
          });
        });
        return notes;
      })(),
      requested_output_format: {
        audience_versions: ['farmer/plain language', 'policy maker', 'research/technical appendix'],
        include_tables: ['Results table (vertical indicators x treatments)', 'Assumptions table', 'Top drivers summary'],
        include_sections: ['Executive summary', 'Scenario assumptions', 'Results and interpretation vs control', 'Uncertainty/sensitivity discussion', 'Implementation considerations', 'Limitations'],
      },
    };

    const prompt =
`You are writing a decision brief using the structured JSON below.
Write in clear prose with headings. Provide:
1) A farmer-friendly version (plain language, what drives outcomes, options to improve).
2) A policy version (decision-relevant framing, distributional and implementation considerations).
3) A technical appendix (definitions, equations for PV/NPV/BCR/ROI, assumptions, and replication notes).
Include a clean table that matches the results table.

JSON:
${JSON.stringify(out, null, 2)}`;

    return prompt;
  }

  async function copyBrief() {
    const text = $('#briefBox').value || '';
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      toast('Briefing prompt copied');
    } catch {
      toast('Copy failed in this browser');
    }
  }

  // ---------- Control auto-detect ----------
  function autoDetectControl() {
    const cand =
      state.treatments.find((t) => t.toLowerCase() === 'control') ||
      state.treatments.find((t) => t.toLowerCase().includes('control')) ||
      state.treatments.find((t) => /baseline|business as usual|bau/.test(t.toLowerCase())) ||
      state.treatments[0];

    if (!cand) return;
    state.control = cand;
    $('#controlSelect').value = cand;
    recalcAll();
    toast('Control auto-detected');
  }

  // ---------- File load ----------
  function loadFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ok = detectAndLoadWorkbook(wb, file.name);
        if (ok) {
          fillInputsToUI();
          recalcAll();
          toast('File loaded and analysed');
        }
      } catch (err) {
        console.error(err);
        toast('Error reading file');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- Wire up ----------
  function init() {
    initTabs();

    // Buttons
    $('#btnRecalcResults').addEventListener('click', () => { recalcAll(); toast('Recalculated'); });
    $('#btnApplyAssumptions').addEventListener('click', () => { recalcAll(); toast('Assumptions applied'); });
    $('#btnUseDemo').addEventListener('click', () => { recalcAll(); toast('Recalculated'); });

    $('#btnDetectControl').addEventListener('click', autoDetectControl);

    $('#btnAddBenefit').addEventListener('click', () => {
      const id = `b_${Math.random().toString(16).slice(2)}`;
      state.addBenefits.push({ id, appliesTo: 'ALL', name: 'New benefit', amtPerHaPerYr: 0, startY: 1, endY: state.inputs.horizonY });
      renderBenefitsTable();
      recalcAll();
    });

    $('#btnAddCost').addEventListener('click', () => {
      const id = `c_${Math.random().toString(16).slice(2)}`;
      state.addCosts.push({ id, appliesTo: 'ALL', name: 'New cost', amtPerHa: 0, type: 'upfront', years: '1' });
      renderCostsTable();
      recalcAll();
    });

    $('#btnRunSim').addEventListener('click', runSim);

    $('#btnCopyResults').addEventListener('click', copyResultsTable);
    $('#btnDownloadResults').addEventListener('click', downloadResultsExcel);

    $('#btnBuildBrief').addEventListener('click', () => {
      const p = buildBriefPrompt();
      $('#briefBox').value = p;
      toast('Briefing prompt built');
    });
    $('#btnCopyBrief').addEventListener('click', copyBrief);

    // File input
    $('#fileInput').addEventListener('change', (e) => {
      const f = e.target.files && e.target.files[0];
      if (f) loadFile(f);
    });

    // Initial tables
    renderBenefitsTable();
    renderCostsTable();
    fillInputsToUI();
    updateStatus();
    renderDetectedColumns();
  }

  // ---------- Start ----------
  document.addEventListener('DOMContentLoaded', init);
})();
