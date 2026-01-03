/* Farming CBA Decision Tool 2 — Fully functional UI + analysis + simulations + Copilot pack
   - Reads uploaded .xlsx (via SheetJS)
   - Auto-detects key fields (treatment, yield, costs)
   - Computes PV Benefits, PV Costs, NPV, BCR, ROI over user horizon & discount rate
   - Compares all treatments against a control
   - Supports Additional Benefits/Costs (annual or one-off, per ha or total)
   - Monte Carlo simulations (price/yield/cost uncertainty)
   - Exports results to Excel
   - Generates a Copilot-ready policy brief pack (prompt + JSON + tables)
*/

(() => {
  'use strict';

  // -----------------------------
  // Utilities
  // -----------------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const clamp = (x, a, b) => Math.max(a, Math.min(b, x));
  const fmt = new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 });
  const fmt2 = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
  const money0 = (x) => (isFinite(x) ? `$${fmt.format(x)}` : '—');
  const num2 = (x) => (isFinite(x) ? fmt2.format(x) : '—');
  const pct1 = (x) => (isFinite(x) ? `${(x * 100).toFixed(1)}%` : '—');

  function uid(prefix = 'id') {
    return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
  }

  function safeText(s) {
    return String(s ?? '').replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));
  }

  function asNumber(v) {
    if (v === null || v === undefined) return NaN;
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    const s = String(v).trim();
    if (!s) return NaN;
    // Strip currency/commas/units
    const cleaned = s
      .replace(/\$/g, '')
      .replace(/,/g, '')
      .replace(/[^\d.\-+eE]/g, '');
    const n = Number(cleaned);
    return isFinite(n) ? n : NaN;
  }

  function mean(arr) {
    const xs = arr.filter((x) => isFinite(x));
    if (!xs.length) return NaN;
    return xs.reduce((a, b) => a + b, 0) / xs.length;
  }

  function sum(arr) {
    const xs = arr.filter((x) => isFinite(x));
    return xs.reduce((a, b) => a + b, 0);
  }

  function stdev(arr) {
    const xs = arr.filter((x) => isFinite(x));
    if (xs.length < 2) return NaN;
    const m = mean(xs);
    const v = xs.reduce((acc, x) => acc + (x - m) ** 2, 0) / (xs.length - 1);
    return Math.sqrt(v);
  }

  function quantile(arr, q) {
    const xs = arr.filter((x) => isFinite(x)).sort((a, b) => a - b);
    if (!xs.length) return NaN;
    const pos = (xs.length - 1) * clamp(q, 0, 1);
    const base = Math.floor(pos);
    const rest = pos - base;
    if (xs[base + 1] === undefined) return xs[base];
    return xs[base] + rest * (xs[base + 1] - xs[base]);
  }

  function discountFactor(r, t) {
    return 1 / Math.pow(1 + r, t);
  }

  function downloadBlob(filename, blob) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 250);
  }

  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch {
      // Fallback
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      let ok = false;
      try {
        ok = document.execCommand('copy');
      } catch {
        ok = false;
      }
      ta.remove();
      return ok;
    }
  }

  function toCSV(rows) {
    return rows
      .map((r) =>
        r
          .map((cell) => {
            const s = String(cell ?? '');
            if (/[,"\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
            return s;
          })
          .join(',')
      )
      .join('\n');
  }

  // -----------------------------
  // State
  // -----------------------------
  const state = {
    workbook: null,
    sheetName: null,
    rows: [], // array of objects
    columns: [], // array of column names (unique)
    meta: {
      treatmentKey: null, // column name
      treatmentLabelKey: null, // column name
      controlPredicate: null, // function(row) => bool
      yieldKey: null, // column name (t/ha)
      yieldAltKey: null, // alt yield key (kg) sometimes
      costKeys: [], // array of column names to sum (per ha)
      hasTimeKey: false,
      timeKey: null
    },
    assumptions: {
      areaHa: 100,
      horizonYears: 10,
      discountRate: 0.07,
      grainPrice: 450, // $/t
      priceGrowth: 0.0, // annual growth rate
      yieldGrowth: 0.0, // annual growth rate
      costGrowth: 0.0, // annual growth rate
      extrapolation: 'steady' // 'steady' | 'repeat_last' | 'average'
    },
    addBenefits: [], // [{id,name,value,unit,mode,startYear,endYear}]
    addCosts: [],    // same structure
    results: null,   // computed results object
    simulations: null
  };

  // -----------------------------
  // Tab system (accessible, never "inactive")
  // -----------------------------
  function initTabs() {
    const tablist = $('[data-tablist]');
    if (!tablist) return;

    const tabs = $$('[role="tab"]', tablist);
    const panels = $$('[role="tabpanel"]');

    function activate(tab) {
      tabs.forEach((t) => {
        const selected = t === tab;
        t.setAttribute('aria-selected', selected ? 'true' : 'false');
        t.classList.toggle('is-active', selected);
        t.tabIndex = selected ? 0 : -1;
      });
      panels.forEach((p) => {
        const match = p.id === tab.getAttribute('aria-controls');
        p.hidden = !match;
        p.classList.toggle('is-active', match);
      });

      // Scroll to top of content on tab switch for clarity on mobile
      const main = $('#mainContent');
      if (main) main.scrollTop = 0;
    }

    tabs.forEach((t) => {
      t.addEventListener('click', (e) => {
        e.preventDefault();
        activate(t);
      });
      t.addEventListener('keydown', (e) => {
        const idx = tabs.indexOf(t);
        if (e.key === 'ArrowRight') {
          e.preventDefault();
          const next = tabs[(idx + 1) % tabs.length];
          next.focus();
          activate(next);
        } else if (e.key === 'ArrowLeft') {
          e.preventDefault();
          const prev = tabs[(idx - 1 + tabs.length) % tabs.length];
          prev.focus();
          activate(prev);
        } else if (e.key === 'Home') {
          e.preventDefault();
          tabs[0].focus();
          activate(tabs[0]);
        } else if (e.key === 'End') {
          e.preventDefault();
          tabs[tabs.length - 1].focus();
          activate(tabs[tabs.length - 1]);
        }
      });
    });

    // Activate first
    const initially = tabs.find((t) => t.dataset.default === 'true') || tabs[0];
    if (initially) activate(initially);
  }

  // -----------------------------
  // Data ingestion
  // -----------------------------
  function uniquifyHeaders(headers) {
    const seen = new Map();
    return headers.map((h, i) => {
      const base = (h ?? '').toString().trim();
      const key = base || `col_${i}`;
      const count = seen.get(key) ?? 0;
      seen.set(key, count + 1);
      return count === 0 ? key : `${key}__${count}`;
    });
  }

  function sheetToRows(sheet) {
    // Read as arrays to preserve header row as it exists
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
    // Find best header row: choose row with most non-null cells
    let bestIdx = 0;
    let bestCount = -1;
    for (let i = 0; i < Math.min(aoa.length, 60); i++) {
      const row = aoa[i] || [];
      const count = row.filter((v) => v !== null && v !== undefined && String(v).trim() !== '').length;
      if (count > bestCount) {
        bestCount = count;
        bestIdx = i;
      }
    }

    const header = uniquifyHeaders((aoa[bestIdx] || []).map((v) => (v === null ? '' : String(v))));
    const rows = [];
    for (let r = bestIdx + 1; r < aoa.length; r++) {
      const line = aoa[r] || [];
      // Ignore empty lines
      const nonEmpty = line.some((v) => v !== null && v !== undefined && String(v).trim() !== '');
      if (!nonEmpty) continue;

      const obj = {};
      for (let c = 0; c < header.length; c++) obj[header[c]] = line[c] ?? null;
      rows.push(obj);
    }

    // Remove rows that look like repeated headers (common in templates)
    const pruned = rows.filter((row) => {
      const vals = Object.values(row).slice(0, 6).map((v) => String(v ?? '').toLowerCase());
      const looksHeader = vals.join(' ').includes('plot') && vals.join(' ').includes('trt');
      return !looksHeader;
    });

    return { header, rows: pruned, headerRowIndex: bestIdx };
  }

  function detectFields(columns, rows) {
    const lc = (s) => String(s).toLowerCase();

    const pick = (patterns) => {
      for (const p of patterns) {
        const found = columns.find((c) => p.test(lc(c)));
        if (found) return found;
      }
      return null;
    };

    const treatmentKey =
      pick([/\btrt\b/, /\btreat(ment)?\b/, /\boption\b/]) ||
      columns[0];

    const treatmentLabelKey =
      pick([/\bamendment\b/, /\bpractice\b/, /\bname\b/, /\bscenario\b/, /\bdescription\b/]) ||
      treatmentKey;

    // Yield candidates
    const yieldKey = pick([/yield.*t\/?ha/, /grain.*t\/?ha/, /t\/?ha.*yield/, /yield.*\(t\/?ha\)/]);
    const yieldAltKey = pick([/grain.*yield.*\bkg\b/, /yield.*\bkg\b/, /grain.*yield.*\bt\)/]);

    // Identify a time/period column (optional). If too many unique values, we treat as non-time.
    const timeCandidate = pick([/\byear\b/, /\bperiod\b/, /\bseason\b/, /\bcrop\b/, /\bwave\b/]);
    let hasTimeKey = false;
    let timeKey = null;
    if (timeCandidate) {
      const uniq = new Set(rows.map((r) => String(r[timeCandidate] ?? '').trim()).filter(Boolean));
      // If it looks like a short sequence (<= horizon or <= 15), treat as time
      if (uniq.size > 1 && uniq.size <= 15) {
        hasTimeKey = true;
        timeKey = timeCandidate;
      }
    }

    // Cost columns: prefer an explicit "total cost" if present, else sum plausible cost items
    const totalCostKey = pick([
      /total.*(variable|input|operat|treatment).*cost.*\/?ha/,
      /total.*cost.*\/?ha/,
      /cost.*total.*\/?ha/
    ]);

    let costKeys = [];
    if (totalCostKey) {
      costKeys = [totalCostKey];
    } else {
      const costRegex = /(cost|labou?r|machinery|seed|fert|urea|ammonia|lime|gypsum|herb|fung|insect|chemical|spray|application|planting|cultivation|harvest|trucking|fuel|diesel|cartage|contract)/i;
      const excludeRegex = /(protein|moisture|test weight|biomass|weed score|satellite|soil|plot length|plot width|area|rep\b|block|ndvi)/i;

      costKeys = columns.filter((c) => costRegex.test(c) && !excludeRegex.test(c));
      // As a fallback, if nothing found, include numeric columns beyond the first few identifiers
      if (!costKeys.length) {
        const sample = rows.slice(0, 30);
        const numericCols = columns.filter((c) => sample.some((r) => isFinite(asNumber(r[c]))));
        costKeys = numericCols.slice(Math.min(6, numericCols.length));
      }
    }

    // Detect control: Amendment contains "control" OR treatment id equals 1 and label indicates control.
    const controlPredicate = (row) => {
      const a = String(row[treatmentLabelKey] ?? '').toLowerCase();
      const t = String(row[treatmentKey] ?? '').toLowerCase();
      return a.includes('control') || t === '1' || t === '0';
    };

    return {
      treatmentKey,
      treatmentLabelKey,
      yieldKey,
      yieldAltKey,
      costKeys,
      hasTimeKey,
      timeKey,
      controlPredicate
    };
  }

  function summariseData(rows, meta) {
    const treatKey = meta.treatmentKey;
    const labelKey = meta.treatmentLabelKey;

    // Group rows by treatment id (string)
    const groups = new Map();
    for (const r of rows) {
      const id = String(r[treatKey] ?? '').trim() || 'Unknown';
      const label = String(r[labelKey] ?? '').trim() || id;
      if (!groups.has(id)) groups.set(id, { id, label, rows: [] });
      groups.get(id).rows.push(r);
      // Keep a more informative label if we find it
      if (label && label !== id && label.length > (groups.get(id).label || '').length) {
        groups.get(id).label = label;
      }
    }
    return Array.from(groups.values()).sort((a, b) => {
      // numeric sort if possible
      const na = Number(a.id);
      const nb = Number(b.id);
      if (isFinite(na) && isFinite(nb)) return na - nb;
      return a.id.localeCompare(b.id);
    });
  }

  // -----------------------------
  // Core CBA calculations
  // -----------------------------
  function computePerHaForGroup(group, meta) {
    const rows = group.rows;
    const yKey = meta.yieldKey;
    const yAlt = meta.yieldAltKey;
    const costKeys = meta.costKeys;

    // Yield: prefer t/ha; else derive from kg and area if possible; else NaN
    let yieldTHa = NaN;
    if (yKey) {
      yieldTHa = mean(rows.map((r) => asNumber(r[yKey])));
    }
    if (!isFinite(yieldTHa) && yAlt) {
      // If alt yield appears to be kg from a plot, try to infer t/ha using area if present
      // Common columns: "Area (m2)" or "Area" (m2). If we can find an area column, convert.
      const areaKey = Object.keys(rows[0] || {}).find((c) => /area/i.test(c) && !/ha/i.test(c)) || null;
      if (areaKey) {
        const derived = rows.map((r) => {
          const kg = asNumber(r[yAlt]);
          const m2 = asNumber(r[areaKey]);
          if (!isFinite(kg) || !isFinite(m2) || m2 <= 0) return NaN;
          const ha = m2 / 10000;
          return (kg / 1000) / ha; // (t) / ha
        });
        yieldTHa = mean(derived);
      }
    }

    // Costs per ha: sum cost columns (per ha)
    const perHaCosts = rows.map((r) => {
      const vals = costKeys.map((k) => asNumber(r[k]));
      return sum(vals);
    });
    const costPerHa = mean(perHaCosts);

    return { yieldTHa, costPerHa };
  }

  function expandAddItems(items, assumptions, year) {
    // returns {benefitsExtra, costsExtra} for that year (farm total, not per ha)
    let total = 0;
    for (const it of items) {
      const start = Number(it.startYear ?? 1);
      const end = Number(it.endYear ?? assumptions.horizonYears);
      if (year < start || year > end) continue;

      if (it.mode === 'oneoff' && year !== start) continue;

      const value = asNumber(it.value);
      if (!isFinite(value)) continue;

      if (it.unit === 'per_ha') total += value * assumptions.areaHa;
      else total += value;
    }
    return total;
  }

  function computeTreatmentPV(perHa, assumptions, addBenefits, addCosts) {
    const r = assumptions.discountRate;
    const T = Math.max(1, Math.floor(assumptions.horizonYears));

    const area = assumptions.areaHa;
    const p0 = assumptions.grainPrice;

    let pvBenefits = 0;
    let pvCosts = 0;

    for (let t = 1; t <= T; t++) {
      const df = discountFactor(r, t);

      // Growth assumptions applied to steady annual flows
      const price = p0 * Math.pow(1 + assumptions.priceGrowth, t - 1);
      const y = (perHa.yieldTHa || 0) * Math.pow(1 + assumptions.yieldGrowth, t - 1);
      const c = (perHa.costPerHa || 0) * Math.pow(1 + assumptions.costGrowth, t - 1);

      const benefitsBase = y * price * area; // $ per year
      const costsBase = c * area;

      const benefitsExtra = expandAddItems(addBenefits, assumptions, t);
      const costsExtra = expandAddItems(addCosts, assumptions, t);

      pvBenefits += (benefitsBase + benefitsExtra) * df;
      pvCosts += (costsBase + costsExtra) * df;
    }

    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts !== 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts !== 0 ? npv / pvCosts : NaN;

    return { pvBenefits, pvCosts, npv, bcr, roi };
  }

  function computeAllResults() {
    if (!state.rows.length) return null;

    const groups = summariseData(state.rows, state.meta);

    // Identify control group
    let control = groups.find((g) => g.rows.some(state.meta.controlPredicate));
    if (!control) control = groups[0] || null;

    const perHaById = new Map();
    const metricsById = new Map();

    for (const g of groups) {
      const perHa = computePerHaForGroup(g, state.meta);
      perHaById.set(g.id, perHa);
      const pv = computeTreatmentPV(perHa, state.assumptions, state.addBenefits, state.addCosts);
      metricsById.set(g.id, { ...pv, perHa, label: g.label, id: g.id });
    }

    // Rank by NPV (descending)
    const list = Array.from(metricsById.values()).sort((a, b) => (b.npv ?? -Infinity) - (a.npv ?? -Infinity));
    const rank = new Map();
    list.forEach((m, i) => rank.set(m.id, i + 1));

    const controlM = control ? metricsById.get(control.id) : null;

    // Build results table structure
    const columns = groups.map((g) => ({ id: g.id, label: g.label, isControl: control ? g.id === control.id : false }));
    const rows = [
      { key: 'pvBenefits', label: 'Present value of benefits (PV Benefits)', format: money0 },
      { key: 'pvCosts', label: 'Present value of costs (PV Costs)', format: money0 },
      { key: 'npv', label: 'Net present value (NPV)', format: money0 },
      { key: 'bcr', label: 'Benefit–cost ratio (BCR)', format: num2 },
      { key: 'roi', label: 'Return on investment (ROI)', format: num2 },
      { key: 'rank', label: 'Ranking (by NPV)', format: (x) => (isFinite(x) ? String(x) : '—') }
    ];

    const deltas = [
      { key: 'dNPV', label: 'Δ NPV vs control', format: money0 },
      { key: 'dPVBenefits', label: 'Δ PV benefits vs control', format: money0 },
      { key: 'dPVCosts', label: 'Δ PV costs vs control', format: money0 }
    ];

    // Produce cell matrix
    const table = {
      controlId: control ? control.id : null,
      columns,
      rows: rows.map((r) => {
        const cells = {};
        for (const c of columns) {
          const m = metricsById.get(c.id);
          cells[c.id] = r.key === 'rank' ? rank.get(c.id) : (m ? m[r.key] : NaN);
        }
        return { ...r, cells };
      }),
      deltas: deltas.map((r) => {
        const cells = {};
        for (const c of columns) {
          const m = metricsById.get(c.id);
          if (!m || !controlM) {
            cells[c.id] = NaN;
          } else {
            if (r.key === 'dNPV') cells[c.id] = m.npv - controlM.npv;
            if (r.key === 'dPVBenefits') cells[c.id] = m.pvBenefits - controlM.pvBenefits;
            if (r.key === 'dPVCosts') cells[c.id] = m.pvCosts - controlM.pvCosts;
          }
        }
        return { ...r, cells };
      }),
      raw: {
        groups,
        perHaById,
        metricsById,
        rank
      }
    };

    return table;
  }

  // -----------------------------
  // Simulations
  // -----------------------------
  function randn() {
    // Box–Muller
    let u = 0, v = 0;
    while (u === 0) u = Math.random();
    while (v === 0) v = Math.random();
    return Math.sqrt(-2.0 * Math.log(u)) * Math.cos(2.0 * Math.PI * v);
  }

  function runSimulations() {
    if (!state.results) return null;

    const draws = Math.floor(asNumber($('#simDraws')?.value) || 2000);
    const priceSD = Math.max(0, asNumber($('#simPriceSD')?.value) || 0.2); // proportional
    const yieldSD = Math.max(0, asNumber($('#simYieldSD')?.value) || 0.15);
    const costSD = Math.max(0, asNumber($('#simCostSD')?.value) || 0.15);

    const { columns, raw } = state.results;
    const out = {};

    for (const col of columns) {
      const base = raw.metricsById.get(col.id);
      if (!base) continue;

      const npvs = [];
      const bcrs = [];
      const rois = [];

      for (let i = 0; i < draws; i++) {
        const priceMult = Math.max(0, 1 + priceSD * randn());
        const yieldMult = Math.max(0, 1 + yieldSD * randn());
        const costMult = Math.max(0, 1 + costSD * randn());

        // Temporarily adjust assumptions + perHa
        const perHa = {
          yieldTHa: (base.perHa.yieldTHa || 0) * yieldMult,
          costPerHa: (base.perHa.costPerHa || 0) * costMult
        };
        const assumptions = { ...state.assumptions, grainPrice: state.assumptions.grainPrice * priceMult };
        const pv = computeTreatmentPV(perHa, assumptions, state.addBenefits, state.addCosts);

        npvs.push(pv.npv);
        bcrs.push(pv.bcr);
        rois.push(pv.roi);
      }

      out[col.id] = {
        id: col.id,
        label: col.label,
        draws,
        npv: {
          mean: mean(npvs),
          sd: stdev(npvs),
          p10: quantile(npvs, 0.10),
          p50: quantile(npvs, 0.50),
          p90: quantile(npvs, 0.90),
          probPositive: npvs.filter((x) => x > 0).length / npvs.length
        },
        bcr: {
          mean: mean(bcrs),
          p10: quantile(bcrs, 0.10),
          p50: quantile(bcrs, 0.50),
          p90: quantile(bcrs, 0.90)
        },
        roi: {
          mean: mean(rois),
          p10: quantile(rois, 0.10),
          p50: quantile(rois, 0.50),
          p90: quantile(rois, 0.90)
        }
      };
    }

    return { meta: { draws, priceSD, yieldSD, costSD }, byTreatment: out };
  }

  function renderHistogram(canvas, values, bins = 24) {
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    const w = canvas.width = canvas.clientWidth * devicePixelRatio;
    const h = canvas.height = canvas.clientHeight * devicePixelRatio;
    ctx.clearRect(0, 0, w, h);

    const xs = values.filter((v) => isFinite(v));
    if (!xs.length) return;

    const min = Math.min(...xs);
    const max = Math.max(...xs);
    const span = max - min || 1;

    const counts = Array.from({ length: bins }, () => 0);
    for (const x of xs) {
      const idx = clamp(Math.floor(((x - min) / span) * bins), 0, bins - 1);
      counts[idx]++;
    }
    const maxC = Math.max(...counts);

    const padL = 32 * devicePixelRatio;
    const padR = 12 * devicePixelRatio;
    const padT = 12 * devicePixelRatio;
    const padB = 22 * devicePixelRatio;

    const plotW = w - padL - padR;
    const plotH = h - padT - padB;

    // axes
    ctx.globalAlpha = 1;
    ctx.lineWidth = 1 * devicePixelRatio;
    ctx.strokeStyle = getComputedStyle(document.documentElement).getPropertyValue('--wb-border').trim() || '#cfd8dc';
    ctx.beginPath();
    ctx.moveTo(padL, padT);
    ctx.lineTo(padL, padT + plotH);
    ctx.lineTo(padL + plotW, padT + plotH);
    ctx.stroke();

    // bars
    const barW = plotW / bins;
    for (let i = 0; i < bins; i++) {
      const c = counts[i];
      const bh = (c / maxC) * plotH;
      const x0 = padL + i * barW;
      const y0 = padT + plotH - bh;
      ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--wb-accent').trim() || '#0b7285';
      ctx.globalAlpha = 0.85;
      ctx.fillRect(x0 + 1, y0, Math.max(1, barW - 2), bh);
    }
    ctx.globalAlpha = 1;

    // simple labels
    ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--wb-text').trim() || '#102a43';
    ctx.font = `${12 * devicePixelRatio}px ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Arial`;
    ctx.textBaseline = 'middle';
    ctx.fillText('Count', 4 * devicePixelRatio, padT + plotH / 2);

    ctx.textBaseline = 'top';
    ctx.fillText(money0(min), padL, padT + plotH + 4 * devicePixelRatio);
    ctx.textAlign = 'right';
    ctx.fillText(money0(max), padL + plotW, padT + plotH + 4 * devicePixelRatio);
    ctx.textAlign = 'left';
  }

  // -----------------------------
  // Rendering
  // -----------------------------
  function setStatus(msg, tone = 'info') {
    const el = $('#statusBox');
    if (!el) return;
    el.dataset.tone = tone;
    el.innerHTML = safeText(msg);
    el.hidden = !msg;
  }

  function renderDataSummary() {
    const box = $('#dataSummary');
    if (!box) return;

    if (!state.rows.length) {
      box.innerHTML = `<div class="muted">No data loaded yet.</div>`;
      return;
    }

    const groups = summariseData(state.rows, state.meta);
    const n = state.rows.length;
    const c = state.columns.length;

    const yieldKey = state.meta.yieldKey || state.meta.yieldAltKey || 'Not detected';
    const costKeyText = state.meta.costKeys.length ? `${state.meta.costKeys.length} cost columns detected` : 'No cost columns detected';

    box.innerHTML = `
      <div class="grid2">
        <div class="card soft">
          <div class="kicker">Dataset</div>
          <div class="metric">${fmt.format(n)} rows</div>
          <div class="sub">${fmt.format(c)} columns · ${fmt.format(groups.length)} treatments detected</div>
        </div>
        <div class="card soft">
          <div class="kicker">Auto-detected fields</div>
          <div class="sub">
            <div><span class="pill">Treatment</span> ${safeText(state.meta.treatmentKey || '—')}</div>
            <div><span class="pill">Label</span> ${safeText(state.meta.treatmentLabelKey || '—')}</div>
            <div><span class="pill">Yield</span> ${safeText(yieldKey)}</div>
            <div><span class="pill">Costs</span> ${safeText(costKeyText)}</div>
          </div>
        </div>
      </div>
    `;
  }

  function renderPreviewTable() {
    const wrap = $('#dataPreview');
    if (!wrap) return;

    if (!state.rows.length) {
      wrap.innerHTML = '';
      return;
    }

    const cols = state.columns.slice(0, 10);
    const rows = state.rows.slice(0, 8);

    const thead = cols.map((c) => `<th title="${safeText(c)}">${safeText(c)}</th>`).join('');
    const tbody = rows
      .map((r) => `<tr>${cols.map((c) => `<td>${safeText(r[c] ?? '')}</td>`).join('')}</tr>`)
      .join('');

    wrap.innerHTML = `
      <div class="tableWrap">
        <table class="table">
          <thead><tr>${thead}</tr></thead>
          <tbody>${tbody}</tbody>
        </table>
      </div>
      <div class="muted mt8">Preview shows first 10 columns and first 8 rows for a quick check.</div>
    `;
  }

  function renderAssumptions() {
    // Inputs already exist; just reflect state if needed
    $('#areaHa').value = String(state.assumptions.areaHa);
    $('#horizonYears').value = String(state.assumptions.horizonYears);
    $('#discountRate').value = String(Math.round(state.assumptions.discountRate * 1000) / 10);
    $('#grainPrice').value = String(state.assumptions.grainPrice);
    $('#priceGrowth').value = String(Math.round(state.assumptions.priceGrowth * 1000) / 10);
    $('#yieldGrowth').value = String(Math.round(state.assumptions.yieldGrowth * 1000) / 10);
    $('#costGrowth').value = String(Math.round(state.assumptions.costGrowth * 1000) / 10);
  }

  function renderAddList(kind) {
    const listEl = kind === 'benefit' ? $('#addBenefitsList') : $('#addCostsList');
    if (!listEl) return;

    const items = kind === 'benefit' ? state.addBenefits : state.addCosts;
    if (!items.length) {
      listEl.innerHTML = `<div class="muted">None added.</div>`;
      return;
    }

    listEl.innerHTML = items
      .map((it) => {
        const unitText = it.unit === 'per_ha' ? 'per ha' : 'total';
        const modeText = it.mode === 'annual' ? 'annual' : 'one-off';
        const years = it.mode === 'oneoff'
          ? `Year ${it.startYear}`
          : `Years ${it.startYear}–${it.endYear}`;
        return `
          <div class="rowItem">
            <div class="rowItem__main">
              <div class="rowItem__title">${safeText(it.name || '(Untitled)')}</div>
              <div class="rowItem__meta">
                <span class="badge">${safeText(modeText)}</span>
                <span class="badge">${safeText(unitText)}</span>
                <span class="badge">${safeText(years)}</span>
              </div>
            </div>
            <div class="rowItem__value">${money0(asNumber(it.value))}</div>
            <button class="btn btn--ghost" data-remove="${safeText(it.id)}" aria-label="Remove">Remove</button>
          </div>
        `;
      })
      .join('');

    $$('button[data-remove]', listEl).forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.getAttribute('data-remove');
        if (kind === 'benefit') state.addBenefits = state.addBenefits.filter((x) => x.id !== id);
        else state.addCosts = state.addCosts.filter((x) => x.id !== id);
        renderAddList(kind);
        recomputeAndRender();
      });
    });
  }

  function renderResultsTable() {
    const wrap = $('#resultsTable');
    const wrapDelta = $('#resultsDeltaTable');
    const callout = $('#resultsCallout');
    if (!wrap || !wrapDelta || !callout) return;

    if (!state.results) {
      wrap.innerHTML = '';
      wrapDelta.innerHTML = '';
      callout.innerHTML = `<div class="muted">Load data and run analysis to see results.</div>`;
      return;
    }

    const { columns, rows, deltas, controlId, raw } = state.results;
    const control = columns.find((c) => c.id === controlId);

    // Callout summary
    const ranked = Array.from(raw.metricsById.values()).sort((a, b) => b.npv - a.npv);
    const best = ranked[0];
    const worst = ranked[ranked.length - 1];
    callout.innerHTML = `
      <div class="callout">
        <div class="callout__title">Quick read</div>
        <div class="callout__body">
          <div><span class="pill">Control</span> ${safeText(control?.label || '—')}</div>
          <div class="mt6"><span class="pill">Highest NPV</span> ${safeText(best?.label || '—')} · ${money0(best?.npv)}</div>
          <div class="mt6"><span class="pill">Lowest NPV</span> ${safeText(worst?.label || '—')} · ${money0(worst?.npv)}</div>
        </div>
      </div>
    `;

    const head = `
      <tr>
        <th class="stickyLeft">Economic indicator</th>
        ${columns
          .map(
            (c) => `<th class="${c.isControl ? 'isControl' : ''}">
              <div class="thCol">
                <div class="thTitle">${safeText(c.label)}</div>
                ${c.isControl ? `<div class="thSub">Control</div>` : `<div class="thSub">Treatment</div>`}
              </div>
            </th>`
          )
          .join('')}
      </tr>
    `;

    const body = rows
      .map((r) => {
        const cells = columns
          .map((c) => {
            const v = r.cells[c.id];
            const isHighlight = r.key === 'npv';
            return `<td class="${isHighlight ? 'hi' : ''}">${safeText(r.format(v))}</td>`;
          })
          .join('');
        return `<tr><td class="stickyLeft">${safeText(r.label)}</td>${cells}</tr>`;
      })
      .join('');

    wrap.innerHTML = `
      <div class="tableWrap">
        <table class="table table--wide">
          <thead>${head}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    `;

    const bodyDelta = deltas
      .map((r) => {
        const cells = columns
          .map((c) => {
            const v = r.cells[c.id];
            const cls = c.isControl ? 'mutedCell' : '';
            return `<td class="${cls}">${safeText(r.format(v))}</td>`;
          })
          .join('');
        return `<tr><td class="stickyLeft">${safeText(r.label)}</td>${cells}</tr>`;
      })
      .join('');

    wrapDelta.innerHTML = `
      <div class="tableWrap">
        <table class="table table--wide">
          <thead>
            <tr>
              <th class="stickyLeft">Compared with control</th>
              ${columns.map((c) => `<th class="${c.isControl ? 'isControl' : ''}">${safeText(c.label)}</th>`).join('')}
            </tr>
          </thead>
          <tbody>${bodyDelta}</tbody>
        </table>
      </div>
      <div class="muted mt8">Δ rows show each treatment relative to the control. The control column is shown for alignment.</div>
    `;
  }

  function renderSimulations() {
    const box = $('#simSummary');
    const table = $('#simTable');
    const canvas = $('#simChart');
    if (!box || !table || !canvas) return;

    if (!state.simulations) {
      box.innerHTML = `<div class="muted">Run simulations to see uncertainty ranges.</div>`;
      table.innerHTML = '';
      const ctx = canvas.getContext('2d');
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      return;
    }

    const { meta, byTreatment } = state.simulations;

    box.innerHTML = `
      <div class="card soft">
        <div class="kicker">Simulation settings</div>
        <div class="sub">
          <div><span class="pill">Draws</span> ${fmt.format(meta.draws)}</div>
          <div class="mt6"><span class="pill">Price uncertainty</span> ${pct1(meta.priceSD)}</div>
          <div class="mt6"><span class="pill">Yield uncertainty</span> ${pct1(meta.yieldSD)}</div>
          <div class="mt6"><span class="pill">Cost uncertainty</span> ${pct1(meta.costSD)}</div>
        </div>
      </div>
    `;

    const rows = Object.values(byTreatment);
    const head = `
      <tr>
        <th>Treatment</th>
        <th>NPV (P10)</th>
        <th>NPV (Median)</th>
        <th>NPV (P90)</th>
        <th>Pr(NPV &gt; 0)</th>
        <th>BCR (Median)</th>
      </tr>
    `;
    const body = rows
      .map((r) => `
        <tr>
          <td>${safeText(r.label)}</td>
          <td>${money0(r.npv.p10)}</td>
          <td class="hi">${money0(r.npv.p50)}</td>
          <td>${money0(r.npv.p90)}</td>
          <td>${pct1(r.npv.probPositive)}</td>
          <td>${num2(r.bcr.p50)}</td>
        </tr>
      `)
      .join('');

    table.innerHTML = `
      <div class="tableWrap">
        <table class="table">
          <thead>${head}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    `;

    // Chart: histogram for the currently selected treatment (default control)
    const sel = $('#simTreatmentSelect');
    if (sel) {
      const selectedId = sel.value;
      const t = byTreatment[selectedId] ? selectedId : Object.keys(byTreatment)[0];
      // We didn't store draws arrays for memory reasons; recreate a representative histogram from summary is not possible.
      // Instead, re-run small on-demand draws for the chart only.
      const small = 1200;
      const base = state.results.raw.metricsById.get(t);
      if (base) {
        const npvs = [];
        for (let i = 0; i < small; i++) {
          const priceMult = Math.max(0, 1 + meta.priceSD * randn());
          const yieldMult = Math.max(0, 1 + meta.yieldSD * randn());
          const costMult = Math.max(0, 1 + meta.costSD * randn());
          const perHa = {
            yieldTHa: (base.perHa.yieldTHa || 0) * yieldMult,
            costPerHa: (base.perHa.costPerHa || 0) * costMult
          };
          const assumptions = { ...state.assumptions, grainPrice: state.assumptions.grainPrice * priceMult };
          const pv = computeTreatmentPV(perHa, assumptions, state.addBenefits, state.addCosts);
          npvs.push(pv.npv);
        }
        renderHistogram(canvas, npvs, 28);
      }
    }
  }

  function renderSimTreatmentSelect() {
    const sel = $('#simTreatmentSelect');
    if (!sel || !state.results) return;
    const { columns, controlId } = state.results;
    sel.innerHTML = columns
      .map((c) => `<option value="${safeText(c.id)}" ${c.id === controlId ? 'selected' : ''}>${safeText(c.label)}</option>`)
      .join('');
    sel.addEventListener('change', () => renderSimulations());
  }

  function buildCopilotPack() {
    if (!state.results) return null;

    const { columns, rows, deltas, controlId } = state.results;
    const control = columns.find((c) => c.id === controlId);

    const tableRows = rows.map((r) => {
      const obj = { indicator: r.label };
      for (const c of columns) obj[c.label] = r.key === 'rank' ? r.cells[c.id] : r.cells[c.id];
      return obj;
    });

    const deltaRows = deltas.map((r) => {
      const obj = { indicator: r.label };
      for (const c of columns) obj[c.label] = r.cells[c.id];
      return obj;
    });

    const pack = {
      tool: 'Farming CBA Decision Tool 2',
      assumptions: { ...state.assumptions },
      control: { id: controlId, label: control?.label || null },
      indicators: rows.map((r) => r.label),
      treatments: columns.map((c) => ({ id: c.id, label: c.label, isControl: c.isControl })),
      results: tableRows,
      deltasVsControl: deltaRows,
      additionalBenefits: state.addBenefits,
      additionalCosts: state.addCosts,
      notes: {
        interpretation: [
          'PV Benefits and PV Costs are discounted over the chosen horizon using the chosen discount rate.',
          'NPV = PV Benefits − PV Costs.',
          'BCR = PV Benefits ÷ PV Costs.',
          'ROI = NPV ÷ PV Costs.',
          'Δ rows compare each option to the control.'
        ]
      }
    };

    const prompt = [
      'Write a detailed policy brief using the JSON provided.',
      'Audience: farmers, policy makers, and researchers. Use plain language but keep technical accuracy.',
      'Requirements:',
      '1) Start with a short executive summary and then a results narrative.',
      '2) Explain what PV Benefits, PV Costs, NPV, BCR, and ROI mean in practical terms.',
      '3) Compare each treatment to the control explicitly using the Δ rows (what changes and why).',
      '4) Identify the main drivers of PV Benefits and PV Costs (yield and costs) and how assumptions (price, discount rate, horizon) influence results.',
      '5) Provide a clean table (Word-ready) that mirrors the indicators-as-rows, treatments-as-columns layout.',
      '6) Add a short “How to improve performance” section for options with low BCR or negative NPV (e.g., reduce costs, lift yield, improve price, adjust practices), framed as options.',
      '7) If simulations are provided, summarise uncertainty and probability NPV>0.',
      'Use headings and short paragraphs. Avoid jargon where possible.'
    ].join('\n');

    // Word-ready markdown tables
    const colLabels = columns.map((c) => c.label);
    const mdTable = (() => {
      const header = `| Economic indicator | ${colLabels.join(' | ')} |`;
      const sep = `|---|${colLabels.map(() => '---:').join('|')}|`;
      const lines = rows.map((r) => {
        const vals = columns.map((c) => {
          const v = r.cells[c.id];
          return r.key === 'bcr' || r.key === 'roi' ? num2(v) : r.key === 'rank' ? (isFinite(v) ? String(v) : '—') : money0(v);
        });
        return `| ${r.label} | ${vals.join(' | ')} |`;
      });
      return [header, sep, ...lines].join('\n');
    })();

    const mdDelta = (() => {
      const header = `| Δ vs control | ${colLabels.join(' | ')} |`;
      const sep = `|---|${colLabels.map(() => '---:').join('|')}|`;
      const lines = deltas.map((r) => {
        const vals = columns.map((c) => money0(r.cells[c.id]));
        return `| ${r.label} | ${vals.join(' | ')} |`;
      });
      return [header, sep, ...lines].join('\n');
    })();

    return { pack, prompt, mdTable, mdDelta };
  }

  function renderCopilot() {
    const box = $('#copilotBox');
    const promptEl = $('#copilotPrompt');
    const jsonEl = $('#copilotJSON');
    const mdEl = $('#copilotMD');
    if (!box || !promptEl || !jsonEl || !mdEl) return;

    if (!state.results) {
      box.innerHTML = `<div class="muted">Run analysis first to generate the Copilot pack.</div>`;
      promptEl.value = '';
      jsonEl.value = '';
      mdEl.value = '';
      return;
    }

    const built = buildCopilotPack();
    const json = JSON.stringify(built.pack, null, 2);

    promptEl.value = built.prompt;
    jsonEl.value = json;
    mdEl.value = `${built.mdTable}\n\n${built.mdDelta}`;
    box.innerHTML = `
      <div class="callout">
        <div class="callout__title">Copilot / AI briefing pack</div>
        <div class="callout__body">
          Copy the Prompt, then copy the JSON, paste both into your AI tool (e.g., Microsoft Copilot), and request a full policy brief and tables.
        </div>
      </div>
    `;
  }

  // -----------------------------
  // Export
  // -----------------------------
  function exportResultsExcel() {
    if (!state.results) {
      setStatus('Run analysis first, then export.', 'warn');
      return;
    }

    const { columns, rows, deltas, controlId } = state.results;
    const control = columns.find((c) => c.id === controlId);

    const wb = XLSX.utils.book_new();

    // Summary sheet
    const summary = [
      ['Farming CBA Decision Tool 2 — Results export'],
      ['Generated', new Date().toISOString()],
      [],
      ['Assumptions'],
      ['Farm area (ha)', state.assumptions.areaHa],
      ['Time horizon (years)', state.assumptions.horizonYears],
      ['Discount rate', state.assumptions.discountRate],
      ['Grain price ($/t)', state.assumptions.grainPrice],
      ['Price growth (per year)', state.assumptions.priceGrowth],
      ['Yield growth (per year)', state.assumptions.yieldGrowth],
      ['Cost growth (per year)', state.assumptions.costGrowth],
      [],
      ['Control', control?.label || '']
    ];
    const ws0 = XLSX.utils.aoa_to_sheet(summary);
    XLSX.utils.book_append_sheet(wb, ws0, 'Summary');

    // Results table (indicator rows x treatments columns)
    const header = ['Economic indicator', ...columns.map((c) => c.label)];
    const aoa = [header];
    for (const r of rows) {
      const line = [r.label];
      for (const c of columns) {
        const v = r.key === 'rank' ? r.cells[c.id] : r.cells[c.id];
        line.push(v);
      }
      aoa.push(line);
    }
    const ws1 = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws1, 'Results');

    // Delta table
    const header2 = ['Δ vs control', ...columns.map((c) => c.label)];
    const aoa2 = [header2];
    for (const r of deltas) {
      const line = [r.label];
      for (const c of columns) line.push(r.cells[c.id]);
      aoa2.push(line);
    }
    const ws2 = XLSX.utils.aoa_to_sheet(aoa2);
    XLSX.utils.book_append_sheet(wb, ws2, 'DeltasVsControl');

    // Additional items
    const addAoa = [
      ['Type', 'Name', 'Value', 'Unit', 'Mode', 'StartYear', 'EndYear'],
      ...state.addBenefits.map((x) => ['Benefit', x.name, x.value, x.unit, x.mode, x.startYear, x.endYear]),
      ...state.addCosts.map((x) => ['Cost', x.name, x.value, x.unit, x.mode, x.startYear, x.endYear])
    ];
    const ws3 = XLSX.utils.aoa_to_sheet(addAoa);
    XLSX.utils.book_append_sheet(wb, ws3, 'AdditionalItems');

    // Copilot pack (JSON)
    const built = buildCopilotPack();
    const ws4 = XLSX.utils.aoa_to_sheet([['Copilot JSON'], [JSON.stringify(built.pack, null, 2)]]);
    XLSX.utils.book_append_sheet(wb, ws4, 'CopilotPack');

    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    downloadBlob(`Farming_CBA_Results_${new Date().toISOString().slice(0, 10)}.xlsx`, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }

  // -----------------------------
  // Recompute pipeline
  // -----------------------------
  function recomputeAndRender() {
    if (!state.rows.length) {
      state.results = null;
      state.simulations = null;
      renderResultsTable();
      renderSimulations();
      renderCopilot();
      return;
    }

    state.results = computeAllResults();
    renderResultsTable();
    renderSimTreatmentSelect();
    renderCopilot();

    // Clear simulations when anything changes
    state.simulations = null;
    renderSimulations();
  }

  // -----------------------------
  // Event wiring
  // -----------------------------
  function initAssumptionInputs() {
    const bindNum = (id, setter) => {
      const el = $(id);
      if (!el) return;
      el.addEventListener('input', () => {
        setter(el.value);
        recomputeAndRender();
      });
    };

    bindNum('#areaHa', (v) => (state.assumptions.areaHa = Math.max(0, asNumber(v) || 0)));
    bindNum('#horizonYears', (v) => (state.assumptions.horizonYears = clamp(Math.floor(asNumber(v) || 1), 1, 60)));
    bindNum('#discountRate', (v) => (state.assumptions.discountRate = clamp((asNumber(v) || 0) / 100, 0, 0.5)));
    bindNum('#grainPrice', (v) => (state.assumptions.grainPrice = Math.max(0, asNumber(v) || 0)));
    bindNum('#priceGrowth', (v) => (state.assumptions.priceGrowth = clamp((asNumber(v) || 0) / 100, -0.5, 0.5)));
    bindNum('#yieldGrowth', (v) => (state.assumptions.yieldGrowth = clamp((asNumber(v) || 0) / 100, -0.5, 0.5)));
    bindNum('#costGrowth', (v) => (state.assumptions.costGrowth = clamp((asNumber(v) || 0) / 100, -0.5, 0.5)));

    const btn = $('#runAnalysisBtn');
    if (btn) btn.addEventListener('click', () => recomputeAndRender());
  }

  function initAddItemForms() {
    function addItem(kind) {
      const prefix = kind === 'benefit' ? 'ben' : 'cost';
      const name = $(`#${prefix}Name`)?.value?.trim();
      const value = asNumber($(`#${prefix}Value`)?.value);
      const unit = $(`#${prefix}Unit`)?.value || 'per_ha';
      const mode = $(`#${prefix}Mode`)?.value || 'annual';
      const startYear = Math.max(1, Math.floor(asNumber($(`#${prefix}Start`)?.value) || 1));
      const endYear = Math.max(startYear, Math.floor(asNumber($(`#${prefix}End`)?.value) || state.assumptions.horizonYears));

      if (!name || !isFinite(value)) {
        setStatus(`Please provide a name and a valid numeric value for the ${kind}.`, 'warn');
        return;
      }

      const item = { id: uid(kind), name, value, unit, mode, startYear, endYear };
      if (kind === 'benefit') state.addBenefits.push(item);
      else state.addCosts.push(item);

      // Clear
      $(`#${prefix}Name`).value = '';
      $(`#${prefix}Value`).value = '';
      $(`#${prefix}Start`).value = '1';
      $(`#${prefix}End`).value = String(state.assumptions.horizonYears);

      renderAddList(kind);
      recomputeAndRender();
      setStatus(`${kind === 'benefit' ? 'Benefit' : 'Cost'} added.`, 'ok');
    }

    $('#addBenefitBtn')?.addEventListener('click', () => addItem('benefit'));
    $('#addCostBtn')?.addEventListener('click', () => addItem('cost'));
  }

  function initSimulationControls() {
    $('#runSimBtn')?.addEventListener('click', () => {
      if (!state.results) {
        setStatus('Run the base analysis first, then run simulations.', 'warn');
        return;
      }
      state.simulations = runSimulations();
      renderSimulations();
      setStatus('Simulations completed.', 'ok');
    });
  }

  function initCopilotButtons() {
    $('#copyPromptBtn')?.addEventListener('click', async () => {
      const ok = await copyToClipboard($('#copilotPrompt')?.value || '');
      setStatus(ok ? 'Prompt copied.' : 'Could not copy prompt. Select and copy manually.', ok ? 'ok' : 'warn');
    });
    $('#copyJSONBtn')?.addEventListener('click', async () => {
      const ok = await copyToClipboard($('#copilotJSON')?.value || '');
      setStatus(ok ? 'JSON copied.' : 'Could not copy JSON. Select and copy manually.', ok ? 'ok' : 'warn');
    });
    $('#copyMDBtn')?.addEventListener('click', async () => {
      const ok = await copyToClipboard($('#copilotMD')?.value || '');
      setStatus(ok ? 'Tables copied.' : 'Could not copy tables. Select and copy manually.', ok ? 'ok' : 'warn');
    });
  }

  function initExportButtons() {
    $('#exportExcelBtn')?.addEventListener('click', exportResultsExcel);
    $('#exportCSVBtn')?.addEventListener('click', () => {
      if (!state.results) {
        setStatus('Run analysis first, then export.', 'warn');
        return;
      }
      const { columns, rows, deltas } = state.results;
      const header = ['Economic indicator', ...columns.map((c) => c.label)];
      const out1 = [header, ...rows.map((r) => [r.label, ...columns.map((c) => r.key === 'rank' ? r.cells[c.id] : r.cells[c.id])])];
      const out2 = [['Δ vs control', ...columns.map((c) => c.label)], ...deltas.map((r) => [r.label, ...columns.map((c) => r.cells[c.id])])];

      const csv = `RESULTS\n${toCSV(out1)}\n\nDELTAS_VS_CONTROL\n${toCSV(out2)}\n`;
      downloadBlob(`Farming_CBA_Results_${new Date().toISOString().slice(0, 10)}.csv`, new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    });
  }

  function initFileUpload() {
    const input = $('#fileInput');
    const sheetSelect = $('#sheetSelect');

    if (!input) return;

    input.addEventListener('change', async (e) => {
      const file = e.target.files?.[0];
      if (!file) return;

      setStatus('Reading Excel file…', 'info');

      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: 'array' });
        state.workbook = wb;
        const names = wb.SheetNames || [];
        state.sheetName = names[0] || null;

        // Populate sheets
        if (sheetSelect) {
          sheetSelect.innerHTML = names.map((n) => `<option value="${safeText(n)}">${safeText(n)}</option>`).join('');
          sheetSelect.value = state.sheetName || '';
          sheetSelect.disabled = names.length <= 1;
          sheetSelect.addEventListener('change', () => {
            state.sheetName = sheetSelect.value;
            loadSheetAndCompute();
          });
        }

        await loadSheetAndCompute();
        setStatus('Data loaded. Review the Data tab, then run analysis.', 'ok');
      } catch (err) {
        console.error(err);
        setStatus('Could not read the Excel file. Please confirm it is a valid .xlsx and try again.', 'warn');
      }
    });

    async function loadSheetAndCompute() {
      if (!state.workbook || !state.sheetName) return;

      const sheet = state.workbook.Sheets[state.sheetName];
      const { header, rows } = sheetToRows(sheet);

      // Store
      state.columns = header;
      state.rows = rows;

      // Detect fields
      state.meta = detectFields(state.columns, state.rows);

      // Render
      renderDataSummary();
      renderPreviewTable();
      renderAssumptions();

      // Pre-set add item end years to horizon
      const benEnd = $('#benEnd');
      const costEnd = $('#costEnd');
      if (benEnd) benEnd.value = String(state.assumptions.horizonYears);
      if (costEnd) costEnd.value = String(state.assumptions.horizonYears);

      // Recompute
      recomputeAndRender();

      // Populate sim select once results exist
      if (state.results) renderSimTreatmentSelect();

      // Refresh copilot pack
      renderCopilot();
    }
  }

  // -----------------------------
  // Init
  // -----------------------------
  document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initFileUpload();
    initAssumptionInputs();
    initAddItemForms();
    initSimulationControls();
    initCopilotButtons();
    initExportButtons();

    // Initial renders
    renderDataSummary();
    renderPreviewTable();
    renderAssumptions();
    renderAddList('benefit');
    renderAddList('cost');
    renderResultsTable();
    renderCopilot();
    renderSimulations();

    setStatus('Ready. Upload your Excel file to begin.', 'info');
  });

})();
