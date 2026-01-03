// Farming CBA Decision Tool 2
// Robust, fully working: tabs, dynamic inputs, correct PV/NPV/BCR/ROI, vertical comparison table,
// Excel-first workflow (template + sample + import with validation), clean exports, simulation,
// AI prompt generator, and full inclusion of the provided default dataset (raw rows).

(() => {
  "use strict";

  // -------------------- SMALL HELPERS --------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  function clamp(v, a, b) {
    return Math.max(a, Math.min(b, v));
  }

  function isFiniteNumber(x) {
    return typeof x === "number" && Number.isFinite(x);
  }

  function toNum(x, fallback = 0) {
    if (x === null || x === undefined) return fallback;
    if (typeof x === "number") return Number.isFinite(x) ? x : fallback;
    const s = String(x).trim();
    if (!s) return fallback;
    const cleaned = s.replace(/[$,]/g, "").replace(/\s+/g, "");
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : fallback;
  }

  function money(n) {
    if (!Number.isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    const fmt = abs >= 1000
      ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
      : n.toLocaleString(undefined, { maximumFractionDigits: 2 });
    return "$" + fmt;
  }

  function num(n) {
    if (!Number.isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    return abs >= 1000
      ? n.toLocaleString(undefined, { maximumFractionDigits: 0 })
      : n.toLocaleString(undefined, { maximumFractionDigits: 3 });
  }

  function pct(n) {
    if (!Number.isFinite(n)) return "n/a";
    return num(n) + "%";
  }

  function showToast(message) {
    const root = $("#toast-root") || document.body;
    const t = document.createElement("div");
    t.className = "toast";
    t.textContent = message;
    root.appendChild(t);
    void t.offsetWidth;
    t.classList.add("show");
    setTimeout(() => {
      t.classList.remove("show");
      setTimeout(() => t.remove(), 220);
    }, 3000);
  }

  function downloadText(filename, content) {
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 500);
  }

  function downloadBlob(filename, blob) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 500);
  }

  function slug(s) {
    return (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  }

  // RNG for simulation
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

  function triangular(u, a, c, b) {
    const F = (c - a) / (b - a);
    if (u < F) return a + Math.sqrt(u * (b - a) * (c - a));
    return b - Math.sqrt((1 - u) * (b - a) * (b - c));
  }

  // -------------------- TOOLTIP (NEAR INDICATOR) --------------------
  const tooltipEl = $("#tooltip");
  let tooltipTarget = null;

  function showTooltip(el) {
    const text = el.getAttribute("data-tooltip");
    if (!text || !tooltipEl) return;
    tooltipTarget = el;

    tooltipEl.textContent = text;
    tooltipEl.setAttribute("aria-hidden", "false");
    tooltipEl.classList.add("show");

    positionTooltip(el);
  }

  function hideTooltip() {
    if (!tooltipEl) return;
    tooltipTarget = null;
    tooltipEl.classList.remove("show");
    tooltipEl.setAttribute("aria-hidden", "true");
  }

  function positionTooltip(el) {
    if (!tooltipEl || !el) return;
    const r = el.getBoundingClientRect();
    const pad = 10;
    const w = tooltipEl.offsetWidth || 320;
    const h = tooltipEl.offsetHeight || 60;

    let left = r.left + Math.min(r.width / 2, 18);
    let top = r.bottom + 8;

    if (left + w + pad > window.innerWidth) left = window.innerWidth - w - pad;
    if (left < pad) left = pad;

    if (top + h + pad > window.innerHeight) {
      top = r.top - h - 8;
    }
    if (top < pad) top = pad;

    tooltipEl.style.left = left + "px";
    tooltipEl.style.top = top + "px";
  }

  document.addEventListener("mouseover", (e) => {
    const t = e.target.closest("[data-tooltip]");
    if (t) showTooltip(t);
  });

  document.addEventListener("focusin", (e) => {
    const t = e.target.closest("[data-tooltip]");
    if (t) showTooltip(t);
  });

  document.addEventListener("mouseout", (e) => {
    if (!tooltipTarget) return;
    const rel = e.relatedTarget;
    if (rel && tooltipTarget.contains(rel)) return;
    hideTooltip();
  });

  document.addEventListener("focusout", () => hideTooltip());
  window.addEventListener("scroll", () => tooltipTarget && positionTooltip(tooltipTarget), { passive: true });
  window.addEventListener("resize", () => tooltipTarget && positionTooltip(tooltipTarget));

  // -------------------- DEFAULT DATA (FULL RAW ROWS INCLUDED) --------------------
  // Raw rows are included in full as provided (for transparency and export).
  // A structured key-column list is also created for calibration (yield, moisture, protein, costs).
  const FABA_2022_RAW_HEADER =
`Faba Beans  Production Costs per hectare  2022 (raw rows exactly as provided)
Columns in each row include: Plot, Trt, Rep, Amendment, Practice Change, plot dimensions, agronomy outputs, application rate, treatment input cost, many cost components, capital purchase prices, and the final total cost figure shown at the end of each row.`;

  const FABA_2022_RAW_ROWS = [
`1\t12\t1\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 1\t20\t2.5\t50\t34\t7.03\t11.8\t23.2\t8.40\t15.51\tCrop 1\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t$210.00\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$4.24\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,945`,
`2\t3\t1\tDeep OM (CP1)\tCrop 2\t20\t2.5\t50\t27\t5.18\t10.6\t23.6\t14.83\t16.46\tCrop 2\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385`,
`3\t11\t1\tDeep Ripping\tCrop 3\t20\t2.5\t50\t33\t7.26\t10.7\t23.4\t17.89\t16.41\tCrop 3\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`4\t1\t1\tControl\tCrop 4\t20\t2.5\t50\t29\t6.20\t10\t22.7\t12.28\t15.19\tCrop 4\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695`,
`5\t5\t1\tDeep Carbon-coated mineral (CCM)\tCrop 5\t20\t2.5\t50\t28\t6.13\t10.2\t22.8\t12.69\t13.28\tCrop 5\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110`,
`6\t10\t1\tDeep OM (CP1) + PAM\tCrop 6\t20\t2.5\t50\t28\t7.27\t11.6\t23.4\t16.13\t15.20\tCrop 6\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`7\t9\t1\tSurface Silicon\tCrop 7\t20\t2.5\t50\t29\t6.78\t10.5\t23.4\t12.23\t15.29\tCrop 7\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835`,
`8\t4\t1\tDeep OM + Gypsum (CP2)\tCrop 8\t20\t2.5\t50\t31\t7.60\t10.3\t25.2\t13.87\t14.46\tCrop 8\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885`,
`9\t6\t1\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 9\t20\t2.5\t50\t31\t5.88\t10.3\t24.4\t14.19\t17.95\tCrop 9\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110`,
`10\t7\t1\tDeep liq. NPKS\tCrop 10\t20\t2.5\t50\t25\t7.23\t11.5\t23.2\t12.12\t15.57\tCrop 10\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`11\t2\t1\tDeep Gypsum\tCrop 11\t20\t2.5\t50\t22\t6.29\t9.9\t22.8\t9.85\t15.45\tCrop 11\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385`,
`12\t8\t1\tDeep liq. Gypsum (CHT)\tCrop 12\t20\t2.5\t50\t26\t5.88\t9.9\t23.5\t10.48\t11.69\tCrop 12\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235`,
`13\t6\t2\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 13\t20\t2.5\t50\t33\t4.79\t9.8\t24.4\t14.49\t13.62\tCrop 13\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110`,
`14\t7\t2\tDeep liq. NPKS\tCrop 14\t20\t2.5\t50\t29\t4.88\t10.4\t23.7\t12.81\t13.49\tCrop 14\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`15\t5\t2\tDeep Carbon-coated mineral (CCM)\tCrop 15\t20\t2.5\t50\t26\t5.39\t10.5\t23.7\t11.97\t12.77\tCrop 15\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110`,
`16\t3\t2\tDeep OM (CP1)\tCrop 16\t20\t2.5\t50\t24\t4.96\t10.2\t23.2\t13.85\t14.44\tCrop 16\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385`,
`17\t1\t2\tControl\tCrop 17\t20\t2.5\t50\t24\t4.99\t10.3\t23.3\t15.61\t10.63\tCrop 17\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695`,
`18\t9\t2\tSurface Silicon\tCrop 18\t20\t2.5\t50\t27\t5.79\t10.6\t21.1\t8.59\t10.63\tCrop 18\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835`,
`19\t4\t2\tDeep OM + Gypsum (CP2)\tCrop 19\t20\t2.5\t50\t22\t5.45\t11.2\t23\t12.34\t15.59\tCrop 19\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885`,
`20\t11\t2\tDeep Ripping\tCrop 20\t20\t2.5\t50\t27\t6.30\t10.4\t22.9\t12.34\t15.28\tCrop 20\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`21\t8\t2\tDeep liq. Gypsum (CHT)\tCrop 21\t20\t2.5\t50\t24\t6.57\t9.8\t23.2\t16.16\t11.35\tCrop 21\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235`,
`22\t12\t2\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 22\t20\t2.5\t50\t26\t6.10\t10.3\t23.6\t14.16\t12.21\tCrop 22\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735`,
`23\t10\t2\tDeep OM (CP1) + PAM\tCrop 23\t20\t2.5\t50\t24\t6.34\t10\t22.8\t15.68\t12.70\tCrop 23\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`24\t2\t2\tDeep Gypsum\tCrop 24\t20\t2.5\t50\t25\t5.44\t9.8\t23.1\t12.70\t13.24\tCrop 24\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385`,
`25\t6\t3\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 25\t20\t2.5\t50\t19\t5.04\t11.2\t24.5\t15.45\t10.97\tCrop 25\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110`,
`26\t11\t3\tDeep Ripping\tCrop 26\t20\t2.5\t50\t21\t6.35\t11.2\t27.3\t19.73\t20.65\tCrop 26\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`27\t2\t3\tDeep Gypsum\tCrop 27\t20\t2.5\t50\t21\t6.94\t10.2\t24.6\t16.39\t14.52\tCrop 27\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385`,
`28\t5\t3\tDeep Carbon-coated mineral (CCM)\tCrop 28\t20\t2.5\t50\t19\t6.31\t10.2\t23\t11.23\t15.58\tCrop 28\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110`,
`29\t8\t3\tDeep liq. Gypsum (CHT)\tCrop 29\t20\t2.5\t50\t26\t6.64\t11.2\t23.5\t13.36\t14.23\tCrop 29\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235`,
`30\t12\t3\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 30\t20\t2.5\t50\t22\t5.96\t10.4\t23.8\t12.01\t13.71\tCrop 30\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735`,
`31\t10\t3\tDeep OM (CP1) + PAM\tCrop 31\t20\t2.5\t50\t22\t7.58\t10.2\t24.2\t12.73\t11.98\tCrop 31\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`32\t4\t3\tDeep OM + Gypsum (CP2)\tCrop 32\t20\t2.5\t50\t25\t6.68\t10.3\t24.6\t13.34\t13.12\tCrop 32\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885`,
`33\t7\t3\tDeep liq. NPKS\tCrop 33\t20\t2.5\t50\t23\t7.33\t10.1\t23.3\t13.06\t12.18\tCrop 33\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`34\t1\t3\tControl\tCrop 34\t20\t2.5\t50\t25\t7.37\t10.3\t23.3\t15.30\t9.52\tCrop 34\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695`,
`35\t3\t3\tDeep OM (CP1)\tCrop 35\t20\t2.5\t50\t23\t5.29\t10.5\t23.7\t12.61\t11.73\tCrop 35\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385`,
`36\t9\t3\tSurface Silicon\tCrop 36\t20\t2.5\t50\t18\t6.81\t10\t23.8\t14.04\t17.68\tCrop 36\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835`,
`37\t5\t4\tDeep Carbon-coated mineral (CCM)\tCrop 37\t20\t2.5\t50\t20\t6.42\t11.1\t23.4\t13.51\t13.34\tCrop 37\t5 t/ha\t$3,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$4,110`,
`38\t6\t4\tDeep OM (CP1) + Carbon-coated mineral (CCM)\tCrop 38\t20\t2.5\t50\t20\t6.18\t10.6\t24.9\t14.50\t13.16\tCrop 38\t15 t/ha ; 5 t/ha\t$21,225.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$22,110`,
`39\t9\t4\tSurface Silicon\tCrop 39\t20\t2.5\t50\t21\t6.69\t10.8\t24.6\t13.72\t15.00\tCrop 39\t2 t/ha\t?\t$35.71\t$100.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$835`,
`40\t10\t4\tDeep OM (CP1) + PAM\tCrop 40\t20\t2.5\t50\t21\t7.72\t10.2\t23.3\t16.55\t18.02\tCrop 40\t15 t/ha ; 5 t/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`41\t11\t4\tDeep Ripping\tCrop 41\t20\t2.5\t50\t23\t6.28\t10.6\t23.4\t10.25\t14.71\tCrop 41\tn/a\t$0.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`42\t2\t4\tDeep Gypsum\tCrop 42\t20\t2.5\t50\t19\t5.85\t9.8\t23.1\t10.66\t11.19\tCrop 42\t5 t/ha\t$500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,385`,
`43\t7\t4\tDeep liq. NPKS\tCrop 43\t20\t2.5\t50\t23\t6.40\t10.1\t23.6\t13.28\t10.18\tCrop 43\t750 L/ha\t?\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$885`,
`44\t4\t4\tDeep OM + Gypsum (CP2)\tCrop 44\t20\t2.5\t50\t33\t5.30\t9.7\t25.5\t16.80\t13.87\tCrop 44\t15 t/ha ; 5 t/ha\t$24,000.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$24,885`,
`45\t1\t4\tControl\tCrop 45\t20\t2.5\t50\t24\t6.21\t9.9\t22.1\t10.02\t14.31\tCrop 45\tn/a\t$0.00\t$0.00\t$0.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$0.00\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$695`,
`46\t3\t4\tDeep OM (CP1)\tCrop 46\t20\t2.5\t50\t28\t5.85\t10.9\t23.9\t13.05\t13.28\tCrop 46\t15 t/ha\t$16,500.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,385`,
`47\t8\t4\tDeep liq. Gypsum (CHT)\tCrop 47\t20\t2.5\t50\t27\t5.85\t9.6\t24.2\t20.66\t12.83\tCrop 47\t0.5 t/ha\t$350.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$1,235`,
`48\t12\t4\tDeep OM (CP1) + liq. Gypsum (CHT)\tCrop 48\t20\t2.5\t50\t23\t6.06\t10\t25.1\t15.65\t11.32\tCrop 48\t15 t/ha ; 0.5 t/ha\t$16,850.00\t$35.71\t$150.00\t$45.00\t$50.00\t$3.33\t$105.00\t\t$193.00\t$30.40\t$1.14\t$1.14\t$0.86\t$18.45\t$13.00\t$14.25\t$8.39\t$4.88\t$12.00\t$10.67\t$16.50\t$16.95\t$16.50\t$16.50\t$2.20\t$2.20\t$1.31\t\t\t$1.11\t$5.56\t$3.67\t$6.25\t$3.53\t$4.55\t$3.64\t$7.58\t$4.93\t$3.48\t$13.64\t$20.00\t$2.12\t$21.21\t$12.12\t$2.12\t\t$125,000.00\t$259,000.00\t$162,800.00\t$792,000.00\t\t$17,735`
  ];

  // Key-column parse from raw TSV row:
  // Plot, Trt, Rep, Amendment, ..., Plants/m2, Yield t/ha, Moisture, Protein, ... TreatmentInputCostOnly/Ha, ... TotalCost (final token)
  function parseKeyColumnsFromRawRow(rawRow) {
    const parts = rawRow.split("\t");
    const plot = toNum(parts[0], NaN);
    const trt = toNum(parts[1], NaN);
    const rep = toNum(parts[2], NaN);
    const amendment = (parts[3] || "").trim();
    // yield, moisture, protein are in fixed positions early in the row in the provided data
    const plantsPerM2 = toNum(parts[8], NaN);
    const yieldT = toNum(parts[9], NaN);
    const moisture = toNum(parts[10], NaN);
    const protein = toNum(parts[11], NaN);

    // Treatment input cost is at parts[16] in the rows that have it; but some rows have "?"
    const inputCostRaw = (parts[16] || "").trim();
    const inputCost = inputCostRaw === "?" ? NaN : toNum(inputCostRaw, NaN);

    // Total cost is the last non-empty token
    let last = "";
    for (let i = parts.length - 1; i >= 0; i--) {
      if (String(parts[i] || "").trim() !== "") { last = String(parts[i]).trim(); break; }
    }
    const totalCost = toNum(last, NaN);

    return { plot, trt, rep, amendment, plantsPerM2, yieldT, moisture, protein, inputCost, totalCost, rawRow };
  }

  const FABA_2022_KEY = FABA_2022_RAW_ROWS.map(parseKeyColumnsFromRawRow);

  const TRT_LABELS = {
    1: "Control",
    2: "Deep Gypsum",
    3: "Deep OM (CP1)",
    4: "Deep OM + Gypsum (CP2)",
    5: "Deep Carbon-coated mineral (CCM)",
    6: "Deep OM (CP1) + Carbon-coated mineral (CCM)",
    7: "Deep liq. NPKS",
    8: "Deep liq. Gypsum (CHT)",
    9: "Surface Silicon",
    10: "Deep OM (CP1) + PAM",
    11: "Deep Ripping",
    12: "Deep OM (CP1) + liq. Gypsum (CHT)"
  };

  // -------------------- MODEL --------------------
  const model = {
    meta: {
      toolName: "Farming CBA Decision Tool 2",
      version: "2.0.0"
    },
    project: {
      name: "Faba beans soil amendment trial (2022)",
      lead: "",
      analysts: "",
      team: "",
      organisation: "Newcastle Business School",
      contactEmail: "",
      contactPhone: "",
      lastUpdated: new Date().toISOString().slice(0, 10),
      summary: "Cost–benefit comparison of soil amendment treatments in faba beans (2022 dataset).",
      goal: "Compare treatments against a control using PV benefits, PV costs, NPV, BCR, ROI, and ranking.",
      withProject: "",
      withoutProject: "",
      objectives: "",
      activities: "",
      stakeholders: ""
    },
    settings: {
      startYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      systemType: "single",
      adoptLow: 0.6,
      adoptBase: 0.9,
      adoptHigh: 1.0,
      riskLow: 0.05,
      riskBase: 0.15,
      riskHigh: 0.30,
      assumeExtraCostsOneOff: true,
      assumptions: "Default calibration uses the provided 2022 faba bean dataset. Treatment extra costs are assumed one-off (year 0) unless changed."
    },
    outputs: [],
    treatments: [],
    benefits: [],
    otherCosts: [],
    runtime: {
      lastResults: null,
      lastComparisonTableHtml: "",
      lastAiPrompt: ""
    },
    rawData: {
      header: FABA_2022_RAW_HEADER,
      rawRows: FABA_2022_RAW_ROWS.slice(),
      keyRows: FABA_2022_KEY.slice()
    }
  };

  // -------------------- CALIBRATION: DEFAULT OUTPUTS + TREATMENTS --------------------
  function groupBy(arr, keyFn) {
    const m = new Map();
    for (const x of arr) {
      const k = keyFn(x);
      if (!m.has(k)) m.set(k, []);
      m.get(k).push(x);
    }
    return m;
  }

  function mean(arr) {
    const xs = arr.filter(Number.isFinite);
    if (xs.length === 0) return NaN;
    return xs.reduce((a, b) => a + b, 0) / xs.length;
  }

  function seedFromFabaBeans2022() {
    // Outputs (farm-friendly, minimal)
    model.outputs = [
      { id: uid(), name: "Grain yield", unit: "t/ha", unitValue: 450, notes: "Default grain price per tonne." },
      { id: uid(), name: "Protein", unit: "%", unitValue: 0, notes: "Set a premium per percentage point if relevant." },
      { id: uid(), name: "Moisture", unit: "%", unitValue: 0, notes: "Set a penalty per percentage point if relevant." }
    ];

    const byTrt = groupBy(model.rawData.keyRows, r => r.trt);
    const controlRows = byTrt.get(1) || [];
    const controlCostPerHa = mean(controlRows.map(r => r.totalCost));
    const baseAnnualCostPerHa = Number.isFinite(controlCostPerHa) ? controlCostPerHa : 0;

    const treatments = [];
    for (const [trt, rows] of byTrt.entries()) {
      const trtNum = Number(trt);
      if (!Number.isFinite(trtNum)) continue;

      const label = TRT_LABELS[trtNum] || (rows[0]?.amendment || `Treatment ${trtNum}`);
      const avgYield = mean(rows.map(r => r.yieldT));
      const avgProtein = mean(rows.map(r => r.protein));
      const avgMoisture = mean(rows.map(r => r.moisture));
      const avgTotalCost = mean(rows.map(r => r.totalCost));
      const avgInputCost = mean(rows.map(r => r.inputCost));

      const areaHa = 100; // as per "worked out on a 100 ha paddock" note in the provided sheet
      const extraCostPerHa = Number.isFinite(avgTotalCost) && Number.isFinite(baseAnnualCostPerHa)
        ? Math.max(0, avgTotalCost - baseAnnualCostPerHa)
        : 0;

      const capitalYear0 = model.settings.assumeExtraCostsOneOff ? (extraCostPerHa * areaHa) : 0;
      const annualExtraPerHa = model.settings.assumeExtraCostsOneOff ? 0 : extraCostPerHa;

      // Decompose annual cost into baseline + any recurring extra
      const annualCostPerHa = baseAnnualCostPerHa + annualExtraPerHa;

      const outputLevels = {};
      outputLevels[model.outputs[0].id] = Number.isFinite(avgYield) ? avgYield : 0;
      outputLevels[model.outputs[1].id] = Number.isFinite(avgProtein) ? avgProtein : 0;
      outputLevels[model.outputs[2].id] = Number.isFinite(avgMoisture) ? avgMoisture : 0;

      // Optional info only (not required for the CBA math)
      const noteBits = [];
      if (!Number.isFinite(avgInputCost)) noteBits.push("Input cost is missing ('?') in raw data for at least one replicate.");
      const notes = noteBits.join(" ");

      treatments.push({
        id: uid(),
        name: label,
        isControl: trtNum === 1,
        areaHa,
        adoption: 1,
        outputLevels,
        // Cost components (annual per ha)
        labourPerHa: 0,
        materialsPerHa: 0,
        servicesPerHa: annualCostPerHa, // store the annual cost here as default, editable by user
        // Capital (year 0)
        capitalYear0,
        // Display-only references
        refAvgTotalCostPerHa: avgTotalCost,
        refAvgInputCostPerHa: avgInputCost,
        notes
      });
    }

    // Ensure exactly one control
    const controls = treatments.filter(t => t.isControl);
    if (controls.length === 0 && treatments.length > 0) treatments[0].isControl = true;
    if (controls.length > 1) {
      treatments.forEach((t, i) => (t.isControl = i === 0));
    }

    // Sort control first, then by trt label
    treatments.sort((a, b) => (b.isControl - a.isControl) || a.name.localeCompare(b.name));
    model.treatments = treatments;

    // Default: no extra benefits/costs
    model.benefits = [];
    model.otherCosts = [];

    // Store the raw dataset box content
    const raw = [
      model.rawData.header.trim(),
      "",
      "RAW ROWS (tab-separated):",
      "Plot\tTrt\tRep\tAmendment\t...\tTotalCost",
      ...model.rawData.rawRows
    ].join("\n");
    $("#rawDataBox").value = raw;

    renderCalibrationSummary();
  }

  // -------------------- UI: TABS --------------------
  function setActiveTab(tabId) {
    $$(".tab-link").forEach(btn => {
      const active = btn.dataset.tab === tabId;
      btn.classList.toggle("active", active);
      btn.setAttribute("aria-selected", active ? "true" : "false");
    });
    $$(".tab-panel").forEach(p => {
      const active = p.dataset.tabPanel === tabId;
      p.classList.toggle("show", active);
      p.classList.toggle("active", active);
      p.setAttribute("aria-hidden", active ? "false" : "true");
    });
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function initTabs() {
    $$(".tab-link").forEach(btn => {
      btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
    });
    $$("[data-tab-jump]").forEach(btn => {
      btn.addEventListener("click", () => setActiveTab(btn.dataset.tabJump));
    });
    $("#startBtn").addEventListener("click", () => setActiveTab("project"));
    $("#startBtn-duplicate").addEventListener("click", () => setActiveTab("project"));
  }

  // -------------------- UI: BIND PROJECT + SETTINGS --------------------
  function bindProjectAndSettings() {
    const P = model.project;
    $("#projectName").value = P.name || "";
    $("#projectLead").value = P.lead || "";
    $("#analystNames").value = P.analysts || "";
    $("#projectTeam").value = P.team || "";
    $("#organisation").value = P.organisation || "";
    $("#contactEmail").value = P.contactEmail || "";
    $("#contactPhone").value = P.contactPhone || "";
    $("#lastUpdated").value = P.lastUpdated || "";
    $("#projectSummary").value = P.summary || "";
    $("#projectGoal").value = P.goal || "";
    $("#withProject").value = P.withProject || "";
    $("#withoutProject").value = P.withoutProject || "";
    $("#projectObjectives").value = P.objectives || "";
    $("#projectActivities").value = P.activities || "";
    $("#stakeholderGroups").value = P.stakeholders || "";

    const S = model.settings;
    $("#startYear").value = S.startYear;
    $("#years").value = S.years;
    $("#discBase").value = S.discBase;
    $("#systemType").value = S.systemType;
    $("#adoptLow").value = S.adoptLow;
    $("#adoptBase").value = S.adoptBase;
    $("#adoptHigh").value = S.adoptHigh;
    $("#riskLow").value = S.riskLow;
    $("#riskBase").value = S.riskBase;
    $("#riskHigh").value = S.riskHigh;
    $("#assumeExtraCostsOneOff").value = String(S.assumeExtraCostsOneOff);
    $("#outputAssumptions").value = S.assumptions || "";

    const wire = (id, fn) => $(id).addEventListener("input", fn);
    wire("#projectName", e => (P.name = e.target.value));
    wire("#projectLead", e => (P.lead = e.target.value));
    wire("#analystNames", e => (P.analysts = e.target.value));
    wire("#projectTeam", e => (P.team = e.target.value));
    wire("#organisation", e => (P.organisation = e.target.value));
    wire("#contactEmail", e => (P.contactEmail = e.target.value));
    wire("#contactPhone", e => (P.contactPhone = e.target.value));
    wire("#lastUpdated", e => (P.lastUpdated = e.target.value));
    wire("#projectSummary", e => (P.summary = e.target.value));
    wire("#projectGoal", e => (P.goal = e.target.value));
    wire("#withProject", e => (P.withProject = e.target.value));
    wire("#withoutProject", e => (P.withoutProject = e.target.value));
    wire("#projectObjectives", e => (P.objectives = e.target.value));
    wire("#projectActivities", e => (P.activities = e.target.value));
    wire("#stakeholderGroups", e => (P.stakeholders = e.target.value));

    wire("#startYear", e => (S.startYear = toNum(e.target.value, S.startYear)));
    wire("#years", e => (S.years = Math.max(1, toNum(e.target.value, S.years))));
    wire("#discBase", e => (S.discBase = toNum(e.target.value, S.discBase)));
    $("#systemType").addEventListener("change", e => (S.systemType = e.target.value));

    wire("#adoptLow", e => (S.adoptLow = clamp(toNum(e.target.value, S.adoptLow), 0, 1)));
    wire("#adoptBase", e => (S.adoptBase = clamp(toNum(e.target.value, S.adoptBase), 0, 1)));
    wire("#adoptHigh", e => (S.adoptHigh = clamp(toNum(e.target.value, S.adoptHigh), 0, 1)));

    wire("#riskLow", e => (S.riskLow = clamp(toNum(e.target.value, S.riskLow), 0, 1)));
    wire("#riskBase", e => (S.riskBase = clamp(toNum(e.target.value, S.riskBase), 0, 1)));
    wire("#riskHigh", e => (S.riskHigh = clamp(toNum(e.target.value, S.riskHigh), 0, 1)));

    $("#assumeExtraCostsOneOff").addEventListener("change", e => {
      S.assumeExtraCostsOneOff = (e.target.value === "true");
      // Re-seed costs from references so behaviour matches the assumption
      reseedCostsUsingAssumption();
      renderTreatments();
      computeAndRenderAll();
    });

    $("#outputAssumptions").addEventListener("input", e => (S.assumptions = e.target.value));
  }

  function reseedCostsUsingAssumption() {
    // recompute baseline from control references
    const control = getControlTreatment();
    const controlAnnual = control ? getAnnualCostPerHa(control) : 0;

    model.treatments.forEach(t => {
      const refTotal = toNum(t.refAvgTotalCostPerHa, NaN);
      const extraPerHa = Number.isFinite(refTotal) ? Math.max(0, refTotal - controlAnnual) : 0;
      if (model.settings.assumeExtraCostsOneOff) {
        t.capitalYear0 = extraPerHa * (toNum(t.areaHa, 0));
        // keep annual at control default
        t.servicesPerHa = controlAnnual;
        t.labourPerHa = 0;
        t.materialsPerHa = 0;
      } else {
        t.capitalYear0 = 0;
        t.servicesPerHa = controlAnnual + extraPerHa;
        t.labourPerHa = 0;
        t.materialsPerHa = 0;
      }
    });
  }

  // -------------------- OUTPUTS UI --------------------
  function renderOutputs() {
    const root = $("#outputsList");
    root.innerHTML = "";

    model.outputs.forEach((o, idx) => {
      const card = document.createElement("div");
      card.className = "card subtle";

      card.innerHTML = `
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Name shown throughout the tool and exports.">Output name</label>
            <input type="text" data-out="name" value="${escapeHtml(o.name)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Units per hectare (e.g., t/ha).">Unit</label>
            <input type="text" data-out="unit" value="${escapeHtml(o.unit)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Dollar value per unit (e.g., grain price per tonne).">Unit value ($ per unit)</label>
            <input type="number" step="0.01" data-out="unitValue" value="${Number.isFinite(o.unitValue) ? o.unitValue : 0}"/>
          </div>
        </div>
        <div class="row-2">
          <div class="field">
            <label data-tooltip="Optional notes to preserve assumptions.">Notes</label>
            <input type="text" data-out="notes" value="${escapeHtml(o.notes || "")}"/>
          </div>
          <div class="field right">
            <button class="btn small ghost" type="button" data-action="removeOutput">Remove output</button>
          </div>
        </div>
      `;

      const inputs = $$("input", card);
      inputs.forEach(inp => {
        inp.addEventListener("input", () => {
          const k = inp.dataset.out;
          if (k === "unitValue") o.unitValue = toNum(inp.value, 0);
          else o[k] = inp.value;
          // ensure treatments have this output id
          model.treatments.forEach(t => {
            if (!(o.id in t.outputLevels)) t.outputLevels[o.id] = 0;
          });
          renderTreatments(); // output columns change
          computeAndRenderAll();
        });
      });

      $("[data-action='removeOutput']", card).addEventListener("click", () => {
        if (model.outputs.length <= 1) {
          showToast("Keep at least one output.");
          return;
        }
        model.outputs.splice(idx, 1);
        // remove levels from treatments
        model.treatments.forEach(t => {
          for (const oid of Object.keys(t.outputLevels)) {
            if (!model.outputs.some(o2 => o2.id === oid)) delete t.outputLevels[oid];
          }
        });
        renderOutputs();
        renderTreatments();
        computeAndRenderAll();
      });

      root.appendChild(card);
    });
  }

  $("#addOutput").addEventListener("click", () => {
    model.outputs.push({ id: uid(), name: "New output", unit: "unit/ha", unitValue: 0, notes: "" });
    model.treatments.forEach(t => (t.outputLevels[model.outputs[model.outputs.length - 1].id] = 0));
    renderOutputs();
    renderTreatments();
  });

  // -------------------- TREATMENTS UI (CAPITAL BEFORE TOTAL COST) --------------------
  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, c => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[c]));
  }

  function getControlTreatment() {
    const c = model.treatments.find(t => t.isControl);
    return c || model.treatments[0] || null;
  }

  function getAnnualCostPerHa(t) {
    // annual per ha cost comes from cost components (labour + materials + services)
    const labour = toNum(t.labourPerHa, 0);
    const materials = toNum(t.materialsPerHa, 0);
    const services = toNum(t.servicesPerHa, 0);
    return labour + materials + services;
  }

  function getFirstYearCostPerHa(t) {
    const annual = getAnnualCostPerHa(t);
    const cap = toNum(t.capitalYear0, 0);
    const area = Math.max(1e-9, toNum(t.areaHa, 0));
    return annual + (cap / area);
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    root.innerHTML = "";

    // ensure outputLevels keys exist
    model.treatments.forEach(t => {
      model.outputs.forEach(o => {
        if (!(o.id in t.outputLevels)) t.outputLevels[o.id] = 0;
      });
    });

    model.treatments.forEach((t, idx) => {
      const card = document.createElement("div");
      card.className = "card subtle";

      const annualCostPerHa = getAnnualCostPerHa(t);
      const totalFirstYearPerHa = getFirstYearCostPerHa(t);

      const outputsGrid = model.outputs.map(o => {
        const v = toNum(t.outputLevels[o.id], 0);
        return `
          <div class="field">
            <label data-tooltip="Enter the output level for this treatment (per hectare). The control provides the baseline for comparisons.">
              ${escapeHtml(o.name)} (${escapeHtml(o.unit)})
            </label>
            <input type="number" step="0.01" data-oid="${o.id}" value="${v}"/>
          </div>
        `;
      }).join("");

      card.innerHTML = `
        <div class="row-2">
          <div class="field">
            <label data-tooltip="Name shown in the Results table and exports.">Treatment name</label>
            <input type="text" data-t="name" value="${escapeHtml(t.name)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Flag exactly one control treatment.">Control</label>
            <select data-t="isControl">
              <option value="false"${t.isControl ? "" : " selected"}>No</option>
              <option value="true"${t.isControl ? " selected" : ""}>Yes (control)</option>
            </select>
          </div>
        </div>

        <div class="row-3">
          <div class="field">
            <label data-tooltip="Area in hectares that this scenario applies to. Used to scale costs and benefits.">Area (ha)</label>
            <input type="number" step="1" min="0" data-t="areaHa" value="${toNum(t.areaHa, 0)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Treatment-specific adoption share (0–1). A base adoption multiplier is also applied in Settings.">Treatment adoption (0–1)</label>
            <input type="number" step="0.05" min="0" max="1" data-t="adoption" value="${clamp(toNum(t.adoption, 1), 0, 1)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Optional notes for data provenance or interpretation.">Notes</label>
            <input type="text" data-t="notes" value="${escapeHtml(t.notes || "")}"/>
          </div>
        </div>

        <hr />

        <h4>Outputs (levels per hectare)</h4>
        <div class="row-3">${outputsGrid}</div>

        <hr />

        <h4>Costs (dynamic)</h4>
        <div class="row-4">
          <div class="field">
            <label data-tooltip="Annual labour cost per hectare.">Labour ($/ha/yr)</label>
            <input type="number" step="0.01" min="0" data-t="labourPerHa" value="${toNum(t.labourPerHa, 0)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Annual materials cost per hectare (e.g., consumables).">Materials ($/ha/yr)</label>
            <input type="number" step="0.01" min="0" data-t="materialsPerHa" value="${toNum(t.materialsPerHa, 0)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Annual services/other cost per hectare (default holds the baseline operating cost).">Services/other ($/ha/yr)</label>
            <input type="number" step="0.01" min="0" data-t="servicesPerHa" value="${toNum(t.servicesPerHa, 0)}"/>
          </div>
          <div class="field">
            <label data-tooltip="Upfront cost in year 0 for this treatment (e.g., one-off amendment, equipment hire, setup).">Capital cost ($, year 0)</label>
            <input type="number" step="1" min="0" data-t="capitalYear0" value="${toNum(t.capitalYear0, 0)}"/>
          </div>
        </div>

        <div class="row-3">
          <div class="metricBox">
            <div class="metricLabel" data-tooltip="Sum of annual labour + materials + services costs per hectare.">Annual operating cost ($/ha/yr)</div>
            <div class="metricValue small" data-live="annualCostPerHa">${money(annualCostPerHa)}</div>
          </div>
          <div class="metricBox">
            <div class="metricLabel" data-tooltip="Annual cost plus capital cost spread over area for a first-year view.">Total first-year cost ($/ha)</div>
            <div class="metricValue small" data-live="firstYearCostPerHa">${money(totalFirstYearPerHa)}</div>
          </div>
          <div class="field right">
            <button class="btn small ghost" type="button" data-action="removeTreatment">Remove treatment</button>
          </div>
        </div>

        ${t.refAvgTotalCostPerHa !== undefined ? `
        <p class="small muted">
          Default dataset reference: mean total cost/ha=${money(toNum(t.refAvgTotalCostPerHa, NaN))}${Number.isFinite(toNum(t.refAvgInputCostPerHa, NaN)) ? `, mean input cost/ha=${money(toNum(t.refAvgInputCostPerHa, NaN))}` : ""}.
        </p>` : ""}
      `;

      // Wire basic fields
      const nameInput = $("[data-t='name']", card);
      const ctrlSelect = $("[data-t='isControl']", card);
      const areaInput = $("[data-t='areaHa']", card);
      const adoptInput = $("[data-t='adoption']", card);
      const notesInput = $("[data-t='notes']", card);

      nameInput.addEventListener("input", () => {
        t.name = nameInput.value;
        computeAndRenderAll();
      });

      ctrlSelect.addEventListener("change", () => {
        const v = ctrlSelect.value === "true";
        model.treatments.forEach(x => (x.isControl = false));
        t.isControl = v;
        // ensure one control
        if (!model.treatments.some(x => x.isControl) && model.treatments.length) model.treatments[0].isControl = true;
        renderTreatments();
        computeAndRenderAll();
      });

      areaInput.addEventListener("input", () => {
        t.areaHa = Math.max(0, toNum(areaInput.value, 0));
        updateTreatmentLiveCosts(card, t);
        computeAndRenderAll();
      });

      adoptInput.addEventListener("input", () => {
        t.adoption = clamp(toNum(adoptInput.value, 1), 0, 1);
        computeAndRenderAll();
      });

      notesInput.addEventListener("input", () => {
        t.notes = notesInput.value;
      });

      // Wire outputs
      $$("input[data-oid]", card).forEach(inp => {
        inp.addEventListener("input", () => {
          const oid = inp.dataset.oid;
          t.outputLevels[oid] = toNum(inp.value, 0);
          computeAndRenderAll();
        });
      });

      // Wire costs
      ["labourPerHa", "materialsPerHa", "servicesPerHa", "capitalYear0"].forEach(k => {
        const inp = $(`input[data-t='${k}']`, card);
        inp.addEventListener("input", () => {
          t[k] = Math.max(0, toNum(inp.value, 0));
          updateTreatmentLiveCosts(card, t);
          computeAndRenderAll();
        });
      });

      // Remove
      $("[data-action='removeTreatment']", card).addEventListener("click", () => {
        if (model.treatments.length <= 2) {
          showToast("Keep at least two scenarios (control plus at least one treatment).");
          return;
        }
        const wasControl = t.isControl;
        model.treatments.splice(idx, 1);
        if (wasControl) {
          model.treatments.forEach((x, i) => (x.isControl = i === 0));
        }
        renderTreatments();
        computeAndRenderAll();
      });

      root.appendChild(card);
    });

    refreshTreatmentSelectors();
  }

  function updateTreatmentLiveCosts(card, t) {
    const annual = getAnnualCostPerHa(t);
    const first = getFirstYearCostPerHa(t);
    const a = $("[data-live='annualCostPerHa']", card);
    const f = $("[data-live='firstYearCostPerHa']", card);
    if (a) a.textContent = money(annual);
    if (f) f.textContent = money(first);
  }

  $("#addTreatment").addEventListener("click", () => {
    const ctrl = getControlTreatment();
    const baseOutputs = {};
    model.outputs.forEach(o => (baseOutputs[o.id] = ctrl ? toNum(ctrl.outputLevels[o.id], 0) : 0));
    model.treatments.push({
      id: uid(),
      name: "New treatment",
      isControl: false,
      areaHa: ctrl ? toNum(ctrl.areaHa, 100) : 100,
      adoption: 1,
      outputLevels: baseOutputs,
      labourPerHa: 0,
      materialsPerHa: 0,
      servicesPerHa: ctrl ? getAnnualCostPerHa(ctrl) : 0,
      capitalYear0: 0,
      notes: ""
    });
    renderTreatments();
    computeAndRenderAll();
  });

  // -------------------- BENEFITS + COSTS (SIMPLE, WORKING) --------------------
  function renderBenefits() {
    const root = $("#benefitsList");
    root.innerHTML = "";
    model.benefits.forEach((b, idx) => {
      const card = document.createElement("div");
      card.className = "card subtle";
      card.innerHTML = `
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Short label used in exports and AI prompt.">Benefit label</label>
            <input type="text" value="${escapeHtml(b.label)}" data-b="label"/>
          </div>
          <div class="field">
            <label data-tooltip="Annual amount in dollars (project-wide).">Annual amount ($/yr)</label>
            <input type="number" step="1" value="${toNum(b.annual, 0)}" data-b="annual"/>
          </div>
          <div class="field">
            <label data-tooltip="Start year index (1 means year 1 of the analysis).">Start year (1..N)</label>
            <input type="number" step="1" min="1" value="${toNum(b.start, 1)}" data-b="start"/>
          </div>
        </div>
        <div class="row-3">
          <div class="field">
            <label data-tooltip="End year index.">End year (1..N)</label>
            <input type="number" step="1" min="1" value="${toNum(b.end, 1)}" data-b="end"/>
          </div>
          <div class="field">
            <label data-tooltip="If on, applies adoption and risk multipliers.">Link to adoption and risk</label>
            <select data-b="link">
              <option value="true"${b.link ? " selected" : ""}>Yes</option>
              <option value="false"${b.link ? "" : " selected"}>No</option>
            </select>
          </div>
          <div class="field right">
            <button class="btn small ghost" type="button" data-action="removeBenefit">Remove</button>
          </div>
        </div>
      `;
      $$("input,select", card).forEach(inp => {
        inp.addEventListener("input", () => {
          const k = inp.dataset.b;
          if (k === "label") b.label = inp.value;
          if (k === "annual") b.annual = Math.max(0, toNum(inp.value, 0));
          if (k === "start") b.start = Math.max(1, toNum(inp.value, 1));
          if (k === "end") b.end = Math.max(1, toNum(inp.value, 1));
          if (k === "link") b.link = (inp.value === "true");
          computeAndRenderAll();
        });
      });
      $("[data-action='removeBenefit']", card).addEventListener("click", () => {
        model.benefits.splice(idx, 1);
        renderBenefits();
        computeAndRenderAll();
      });
      root.appendChild(card);
    });
  }

  $("#addBenefit").addEventListener("click", () => {
    model.benefits.push({ id: uid(), label: "New benefit", annual: 0, start: 1, end: 1, link: true });
    renderBenefits();
  });

  function renderOtherCosts() {
    const root = $("#costsList");
    root.innerHTML = "";
    model.otherCosts.forEach((c, idx) => {
      const card = document.createElement("div");
      card.className = "card subtle";
      card.innerHTML = `
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Short label used in exports and AI prompt.">Cost label</label>
            <input type="text" value="${escapeHtml(c.label)}" data-c="label"/>
          </div>
          <div class="field">
            <label data-tooltip="Annual amount in dollars (project-wide).">Annual amount ($/yr)</label>
            <input type="number" step="1" value="${toNum(c.annual, 0)}" data-c="annual"/>
          </div>
          <div class="field">
            <label data-tooltip="Start year index (1 means year 1 of the analysis).">Start year (1..N)</label>
            <input type="number" step="1" min="1" value="${toNum(c.start, 1)}" data-c="start"/>
          </div>
        </div>
        <div class="row-3">
          <div class="field">
            <label data-tooltip="End year index.">End year (1..N)</label>
            <input type="number" step="1" min="1" value="${toNum(c.end, 1)}" data-c="end"/>
          </div>
          <div class="field">
            <label data-tooltip="One-off cost in year 0 (project-wide).">Capital (year 0, $)</label>
            <input type="number" step="1" value="${toNum(c.capitalYear0, 0)}" data-c="capitalYear0"/>
          </div>
          <div class="field right">
            <button class="btn small ghost" type="button" data-action="removeCost">Remove</button>
          </div>
        </div>
      `;
      $$("input", card).forEach(inp => {
        inp.addEventListener("input", () => {
          const k = inp.dataset.c;
          if (k === "label") c.label = inp.value;
          if (k === "annual") c.annual = Math.max(0, toNum(inp.value, 0));
          if (k === "start") c.start = Math.max(1, toNum(inp.value, 1));
          if (k === "end") c.end = Math.max(1, toNum(inp.value, 1));
          if (k === "capitalYear0") c.capitalYear0 = Math.max(0, toNum(inp.value, 0));
          computeAndRenderAll();
        });
      });
      $("[data-action='removeCost']", card).addEventListener("click", () => {
        model.otherCosts.splice(idx, 1);
        renderOtherCosts();
        computeAndRenderAll();
      });
      root.appendChild(card);
    });
  }

  $("#addCost").addEventListener("click", () => {
    model.otherCosts.push({ id: uid(), label: "New cost", annual: 0, start: 1, end: 1, capitalYear0: 0 });
    renderOtherCosts();
  });

  // -------------------- CALCULATION CORE --------------------
  function annualRevenuePerHa(t, outputs = model.outputs) {
    let sum = 0;
    for (const o of outputs) {
      const level = toNum(t.outputLevels[o.id], 0);
      const v = toNum(o.unitValue, 0);
      sum += level * v;
    }
    return sum;
  }

  function buildCashflowsForTreatment(t, opts = {}) {
    const years = Math.max(1, toNum(opts.years ?? model.settings.years, model.settings.years));
    const disc = toNum(opts.discBase ?? model.settings.discBase, model.settings.discBase) / 100;
    const adoptBase = clamp(toNum(opts.adoptBase ?? model.settings.adoptBase, model.settings.adoptBase), 0, 1);
    const riskBase = clamp(toNum(opts.riskBase ?? model.settings.riskBase, model.settings.riskBase), 0, 1);

    const area = Math.max(0, toNum(t.areaHa, 0));
    const adopt = clamp(toNum(t.adoption, 1), 0, 1);
    const effArea = area * adopt * adoptBase;

    const revPerHa = toNum(opts.revenuePerHa ?? annualRevenuePerHa(t, opts.outputs ?? model.outputs), 0);
    const annualRev = revPerHa * effArea * (1 - riskBase);

    const annualCostPerHa = toNum(opts.annualCostPerHa ?? getAnnualCostPerHa(t), 0);
    const annualCost = annualCostPerHa * effArea;

    const cap0 = toNum(opts.capitalYear0 ?? t.capitalYear0, 0);

    // Project-wide benefits/costs
    const extraAnnualBenefits = new Array(years + 1).fill(0);
    const extraAnnualCosts = new Array(years + 1).fill(0);

    // Benefits (year 1..years)
    for (const b of (opts.benefits ?? model.benefits)) {
      const start = clamp(toNum(b.start, 1), 1, years);
      const end = clamp(toNum(b.end, 1), 1, years);
      for (let y = start; y <= end; y++) {
        const amt = toNum(b.annual, 0);
        const scaled = b.link ? amt * adoptBase * (1 - riskBase) : amt;
        extraAnnualBenefits[y] += scaled;
      }
    }

    // Costs (year 1..years) + capital year0
    let extraCap0 = 0;
    for (const c of (opts.otherCosts ?? model.otherCosts)) {
      const start = clamp(toNum(c.start, 1), 1, years);
      const end = clamp(toNum(c.end, 1), 1, years);
      for (let y = start; y <= end; y++) {
        extraAnnualCosts[y] += toNum(c.annual, 0);
      }
      extraCap0 += toNum(c.capitalYear0, 0);
    }

    // Build cashflows array for IRR: CF[0] is year0 net (benefits - costs)
    const cash = new Array(years + 1).fill(0);
    // year0: negative capital costs (treatment + project wide)
    cash[0] = -(cap0 + extraCap0);

    for (let y = 1; y <= years; y++) {
      const ben = annualRev + extraAnnualBenefits[y];
      const cost = annualCost + extraAnnualCosts[y];
      cash[y] = ben - cost;
    }

    // PV calculations
    let pvBenefits = 0;
    let pvCosts = 0;
    for (let y = 0; y <= years; y++) {
      const df = (disc === 0) ? 1 : Math.pow(1 + disc, y);
      if (y === 0) {
        pvCosts += (cap0 + extraCap0); // year0 cost
      } else {
        pvBenefits += (annualRev + extraAnnualBenefits[y]) / df;
        pvCosts += (annualCost + extraAnnualCosts[y]) / df;
      }
    }

    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts === 0 ? NaN : (pvBenefits / pvCosts);
    const roi = pvCosts === 0 ? NaN : (npv / pvCosts);

    const irr = computeIRR(cash);
    const payback = computePayback(cash, disc);

    return {
      years, disc,
      effArea,
      annualRev, annualCost,
      cap0, extraCap0,
      cashflows: cash,
      pvBenefits, pvCosts, npv, bcr, roi, irr, payback
    };
  }

  function computeIRR(cashflows) {
    // robust bisection in [-0.99, 3.0] if sign change exists
    const f = (r) => {
      let npv = 0;
      for (let t = 0; t < cashflows.length; t++) {
        npv += cashflows[t] / Math.pow(1 + r, t);
      }
      return npv;
    };

    let lo = -0.99;
    let hi = 3.0;
    let flo = f(lo);
    let fhi = f(hi);
    if (!Number.isFinite(flo) || !Number.isFinite(fhi)) return NaN;
    if (flo === 0) return lo * 100;
    if (fhi === 0) return hi * 100;

    // Try to bracket by scanning
    if (flo * fhi > 0) {
      let prevR = lo, prevF = flo;
      for (let r = -0.9; r <= 3.0; r += 0.1) {
        const fr = f(r);
        if (Number.isFinite(fr) && prevF * fr <= 0) {
          lo = prevR; hi = r; flo = prevF; fhi = fr;
          break;
        }
        prevR = r; prevF = fr;
      }
      if (flo * fhi > 0) return NaN;
    }

    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2;
      const fmid = f(mid);
      if (!Number.isFinite(fmid)) return NaN;
      if (Math.abs(fmid) < 1e-7) return mid * 100;
      if (flo * fmid <= 0) {
        hi = mid; fhi = fmid;
      } else {
        lo = mid; flo = fmid;
      }
    }
    return ((lo + hi) / 2) * 100;
  }

  function computePayback(cashflows, disc) {
    // earliest year when discounted cumulative net >= 0
    let cum = 0;
    for (let t = 0; t < cashflows.length; t++) {
      const df = (disc === 0) ? 1 : Math.pow(1 + disc, t);
      cum += cashflows[t] / df;
      if (cum >= 0) return t; // year index
    }
    return NaN;
  }

  // -------------------- RESULTS TABLE (VERTICAL, CONTROL INCLUDED) --------------------
  function computeResultsBaseCase() {
    const control = getControlTreatment();
    const sortMode = $("#resultsSort").value || "npv_desc";
    const showDelta = ($("#showDeltaLine").value || "true") === "true";
    const indicatorsMode = $("#resultsIndicators").value || "core";

    // compute all
    const results = model.treatments.map(t => ({
      t,
      res: buildCashflowsForTreatment(t)
    }));

    // control baseline
    const controlRes = results.find(x => x.t.isControl)?.res || results[0]?.res || null;

    // sorting (keep control first)
    const treatmentsOnly = results.filter(x => !x.t.isControl);

    const getKey = (x) => {
      if (sortMode === "name_asc") return x.t.name.toLowerCase();
      if (sortMode === "pvcosts_asc") return x.res.pvCosts;
      if (sortMode === "pvbenefits_desc") return -x.res.pvBenefits;
      if (sortMode === "bcr_desc") return -x.res.bcr;
      if (sortMode === "roi_desc") return -x.res.roi;
      // default NPV desc
      return -x.res.npv;
    };

    treatmentsOnly.sort((a, b) => {
      const ka = getKey(a);
      const kb = getKey(b);
      if (typeof ka === "string" && typeof kb === "string") return ka.localeCompare(kb);
      return (ka - kb);
    });

    const ordered = [];
    const ctrlItem = results.find(x => x.t.isControl) || results[0];
    if (ctrlItem) ordered.push(ctrlItem);
    ordered.push(...treatmentsOnly);

    // ranking by NPV (desc) among non-control
    const rankSorted = treatmentsOnly
      .slice()
      .sort((a, b) => (b.res.npv - a.res.npv));
    const rankMap = new Map(rankSorted.map((x, i) => [x.t.id, i + 1]));

    // Headline cards (best by NPV, BCR, ROI)
    const bestNPV = treatmentsOnly.slice().sort((a, b) => (b.res.npv - a.res.npv))[0];
    const bestBCR = treatmentsOnly.slice().sort((a, b) => (b.res.bcr - a.res.bcr))[0];
    const bestROI = treatmentsOnly.slice().sort((a, b) => (b.res.roi - a.res.roi))[0];

    $("#headlineBestNPV").textContent = bestNPV ? money(bestNPV.res.npv) : "n/a";
    $("#headlineBestNPVName").textContent = bestNPV ? bestNPV.t.name : "n/a";

    $("#headlineBestBCR").textContent = bestBCR ? num(bestBCR.res.bcr) : "n/a";
    $("#headlineBestBCRName").textContent = bestBCR ? bestBCR.t.name : "n/a";

    $("#headlineBestROI").textContent = bestROI ? num(bestROI.res.roi) : "n/a";
    $("#headlineBestROIName").textContent = bestROI ? bestROI.t.name : "n/a";

    $("#headlineControlName").textContent = ctrlItem ? ctrlItem.t.name : "n/a";

    // Build table
    const indicators = (indicatorsMode === "full")
      ? [
          { key: "pvBenefits", label: "Present value of benefits", fmt: money, deltaType: "money" },
          { key: "pvCosts", label: "Present value of costs", fmt: money, deltaType: "money" },
          { key: "npv", label: "Net present value", fmt: money, deltaType: "money" },
          { key: "bcr", label: "Benefit–cost ratio", fmt: num, deltaType: "ratio" },
          { key: "roi", label: "Return on investment", fmt: num, deltaType: "ratio" },
          { key: "irr", label: "Internal rate of return", fmt: (x) => Number.isFinite(x) ? pct(x) : "n/a", deltaType: "ratio" },
          { key: "payback", label: "Payback (years)", fmt: (x) => Number.isFinite(x) ? String(x) : "n/a", deltaType: "diff" },
          { key: "rank", label: "Ranking (by NPV)", fmt: (x) => (x === 0 ? "Control" : (Number.isFinite(x) ? String(x) : "n/a")), deltaType: "none" }
        ]
      : [
          { key: "pvBenefits", label: "Present value of benefits", fmt: money, deltaType: "money" },
          { key: "pvCosts", label: "Present value of costs", fmt: money, deltaType: "money" },
          { key: "npv", label: "Net present value", fmt: money, deltaType: "money" },
          { key: "bcr", label: "Benefit–cost ratio", fmt: num, deltaType: "ratio" },
          { key: "roi", label: "Return on investment", fmt: num, deltaType: "ratio" },
          { key: "rank", label: "Ranking (by NPV)", fmt: (x) => (x === 0 ? "Control" : (Number.isFinite(x) ? String(x) : "n/a")), deltaType: "none" }
        ];

    // Table HTML
    const table = document.createElement("table");

    const thead = document.createElement("thead");
    const trh = document.createElement("tr");

    const th0 = document.createElement("th");
    th0.className = "sticky-left";
    th0.textContent = "Indicator";
    trh.appendChild(th0);

    ordered.forEach((x, j) => {
      const th = document.createElement("th");
      th.textContent = x.t.name + (x.t.isControl ? " (Control)" : "");
      if (j === 0) th.classList.add("sticky-left"); // header aligns with sticky indicator column style
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    for (const ind of indicators) {
      const tr = document.createElement("tr");

      const tdLabel = document.createElement("td");
      tdLabel.className = "sticky-left";
      tdLabel.innerHTML = `<div class="cellMain">${escapeHtml(ind.label)}</div>`;
      tr.appendChild(tdLabel);

      ordered.forEach(x => {
        const td = document.createElement("td");

        let val;
        if (ind.key === "rank") val = x.t.isControl ? 0 : (rankMap.get(x.t.id) ?? NaN);
        else val = x.res[ind.key];

        const main = ind.fmt(val);

        let sub = "";
        if (showDelta && controlRes && !x.t.isControl && ind.deltaType !== "none") {
          const cval = controlRes[ind.key];
          if (ind.deltaType === "money") {
            const d = (Number.isFinite(val) && Number.isFinite(cval)) ? (val - cval) : NaN;
            sub = fmtDeltaMoney(d);
          } else if (ind.deltaType === "ratio") {
            const d = (Number.isFinite(val) && Number.isFinite(cval)) ? (val - cval) : NaN;
            const r = (Number.isFinite(val) && Number.isFinite(cval) && cval !== 0) ? (val / cval) : NaN;
            sub = fmtDeltaRatio(d, r);
          } else if (ind.deltaType === "diff") {
            const d = (Number.isFinite(val) && Number.isFinite(cval)) ? (val - cval) : NaN;
            sub = fmtDeltaPlain(d);
          }
        } else if (showDelta && x.t.isControl && ind.deltaType !== "none") {
          // control: delta line shown as baseline
          if (ind.deltaType === "money") sub = `<span class="muted">Δ $0</span>`;
          if (ind.deltaType === "ratio") sub = `<span class="muted">Δ 0 (×1.00)</span>`;
          if (ind.deltaType === "diff") sub = `<span class="muted">Δ 0</span>`;
        }

        td.innerHTML = `
          <div class="cellMain">${escapeHtml(main)}</div>
          ${showDelta && ind.deltaType !== "none" ? `<div class="cellSub">${sub}</div>` : ""}
        `;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    }

    table.appendChild(tbody);

    const wrap = $("#comparisonTableWrap");
    wrap.innerHTML = "";
    const sc = document.createElement("div");
    sc.className = "table-scroll";
    sc.appendChild(table);
    wrap.appendChild(sc);

    // Save for exports (HTML)
    model.runtime.lastComparisonTableHtml = wrap.innerHTML;

    // Update detail selector + breakdown
    populateDetailBreakdown(ordered);

    model.runtime.lastResults = { ordered, results, controlRes, rankMap };
    return model.runtime.lastResults;
  }

  function fmtDeltaMoney(d) {
    if (!Number.isFinite(d)) return "Δ n/a";
    const cls = d >= 0 ? "pos" : "neg";
    const sign = d >= 0 ? "+" : "−";
    return `Δ <span class="${cls}">${sign}${money(Math.abs(d))}</span>`;
  }

  function fmtDeltaPlain(d) {
    if (!Number.isFinite(d)) return "Δ n/a";
    const cls = d >= 0 ? "pos" : "neg";
    const sign = d >= 0 ? "+" : "−";
    return `Δ <span class="${cls}">${sign}${num(Math.abs(d))}</span>`;
  }

  function fmtDeltaRatio(d, r) {
    const dTxt = Number.isFinite(d) ? (d >= 0 ? `+${num(d)}` : `−${num(Math.abs(d))}`) : "n/a";
    const rTxt = Number.isFinite(r) ? `×${num(r)}` : "×n/a";
    const cls = Number.isFinite(d) ? (d >= 0 ? "pos" : "neg") : "";
    return `Δ <span class="${cls}">${dTxt}</span> (${rTxt})`;
  }

  // -------------------- BREAKDOWN PANEL --------------------
  function refreshTreatmentSelectors() {
    const sel1 = $("#detailTreatment");
    const sel2 = $("#simTarget");
    if (!sel1 || !sel2) return;

    const prev1 = sel1.value;
    const prev2 = sel2.value;

    sel1.innerHTML = "";
    sel2.innerHTML = "";

    model.treatments.forEach(t => {
      const o1 = document.createElement("option");
      o1.value = t.id;
      o1.textContent = t.name + (t.isControl ? " (Control)" : "");
      sel1.appendChild(o1);

      const o2 = document.createElement("option");
      o2.value = t.id;
      o2.textContent = t.name + (t.isControl ? " (Control)" : "");
      sel2.appendChild(o2);
    });

    if (prev1 && model.treatments.some(t => t.id === prev1)) sel1.value = prev1;
    if (prev2 && model.treatments.some(t => t.id === prev2)) sel2.value = prev2;

    sel1.addEventListener("change", () => computeAndRenderAll());
    sel2.addEventListener("change", () => { /* no auto-run */ });
  }

  function populateDetailBreakdown(ordered) {
    const sel = $("#detail
