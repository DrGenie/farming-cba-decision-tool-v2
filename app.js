// app.js
// Farming CBA Decision Tool 2 (commercial-grade single-file JS)
// Robust tabs + working buttons + Excel-first workflow + results vs control table + simulation + AI prompt generator
// No external frameworks required (optional: SheetJS XLSX via CDN already included in HTML).

(() => {
  "use strict";

  // -----------------------------
  // 0) DOM + small utilities
  // -----------------------------
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const nowISO = () => new Date().toISOString().slice(0, 10);

  function clamp(v, lo, hi) {
    v = Number(v);
    if (!isFinite(v)) return lo;
    return Math.max(lo, Math.min(hi, v));
  }

  function toNum(v, fallback = 0) {
    const n = Number(v);
    return isFinite(n) ? n : fallback;
  }

  function parseMoney(x) {
    if (x === null || x === undefined) return NaN;
    const s = String(x).trim();
    if (!s) return NaN;
    if (s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    // keep minus, digits, dot
    const cleaned = s.replace(/\$/g, "").replace(/,/g, "").replace(/[^\d.\-]/g, "");
    const n = Number(cleaned);
    return isFinite(n) ? n : NaN;
  }

  function parseMaybeNumber(x) {
    if (x === null || x === undefined) return NaN;
    const s = String(x).trim();
    if (!s) return NaN;
    if (s === "?" || s.toLowerCase() === "na" || s.toLowerCase() === "n/a") return NaN;
    const cleaned = s.replace(/,/g, "").replace(/[^\d.\-]/g, "");
    const n = Number(cleaned);
    return isFinite(n) ? n : NaN;
  }

  function fmt(n, maxDP = 2) {
    if (!isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    const opts =
      abs >= 1000
        ? { maximumFractionDigits: 0 }
        : { minimumFractionDigits: 0, maximumFractionDigits: maxDP };
    return n.toLocaleString(undefined, opts);
  }

  function money(n) {
    return isFinite(n) ? "$" + fmt(n, 2) : "n/a";
  }

  function ratio(n) {
    return isFinite(n) ? fmt(n, 3) : "n/a";
  }

  function percent(n) {
    return isFinite(n) ? fmt(n, 2) + "%" : "n/a";
  }

  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  function slug(s) {
    return (s || "project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_|_$/g, "");
  }

  function showToast(msg) {
    const root = $("#toast-root") || document.body;
    const t = document.createElement("div");
    t.className = "toast";
    t.textContent = msg;
    root.appendChild(t);
    // trigger
    void t.offsetWidth;
    t.classList.add("show");
    setTimeout(() => {
      t.classList.remove("show");
      setTimeout(() => t.remove(), 200);
    }, 3200);
  }

  // deterministic RNG (mulberry32)
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

  // -----------------------------
  // 1) Tooltip (near the indicator)
  // -----------------------------
  function installTooltipSystem() {
    // minimal CSS injected (keeps tooltips close to hovered label)
    const css = `
      #_tool2_tip {
        position: fixed;
        z-index: 99999;
        max-width: 360px;
        padding: 10px 12px;
        border-radius: 10px;
        background: rgba(17,24,39,0.95);
        color: #fff;
        font-size: 12.5px;
        line-height: 1.35;
        box-shadow: 0 10px 30px rgba(0,0,0,0.25);
        pointer-events: none;
        opacity: 0;
        transform: translateY(4px);
        transition: opacity 120ms ease, transform 120ms ease;
      }
      #_tool2_tip.show { opacity: 1; transform: translateY(0); }
      [data-tooltip] { cursor: help; }
    `;
    const style = document.createElement("style");
    style.textContent = css;
    document.head.appendChild(style);

    const tip = document.createElement("div");
    tip.id = "_tool2_tip";
    document.body.appendChild(tip);

    let active = null;

    function positionTip(target) {
      const r = target.getBoundingClientRect();
      const pad = 10;
      let x = r.left + Math.min(24, r.width);
      let y = r.bottom + 10;
      // keep on screen
      const tw = tip.offsetWidth || 260;
      const th = tip.offsetHeight || 60;
      x = Math.min(x, window.innerWidth - tw - pad);
      y = Math.min(y, window.innerHeight - th - pad);
      if (y < pad) y = pad;
      tip.style.left = x + "px";
      tip.style.top = y + "px";
    }

    function show(target) {
      const txt = target.getAttribute("data-tooltip");
      if (!txt) return;
      tip.textContent = txt;
      tip.classList.add("show");
      positionTip(target);
      active = target;
    }

    function hide() {
      tip.classList.remove("show");
      active = null;
    }

    document.addEventListener("mouseover", (e) => {
      const el = e.target.closest("[data-tooltip]");
      if (!el) return;
      show(el);
    });

    document.addEventListener("mousemove", () => {
      if (!active) return;
      positionTip(active);
    });

    document.addEventListener("mouseout", (e) => {
      const el = e.target.closest("[data-tooltip]");
      if (!el) return;
      // if moving within the same element, ignore
      if (active && (e.relatedTarget === active || active.contains(e.relatedTarget))) return;
      hide();
    });

    window.addEventListener("scroll", () => {
      if (active) positionTip(active);
    });
    window.addEventListener("resize", () => {
      if (active) positionTip(active);
    });
  }

  // -----------------------------
  // 2) Default dataset (FULL rows preserved as raw lines)
  // -----------------------------
  // IMPORTANT: These are the exact data rows supplied by the user (do not omit).
  // The tool uses Yield t/ha and the last $ value on each row as Total cost ($/ha),
  // and keeps the full raw row text for export and traceability.
  const FABA_2022_RAW_LINES = [
    `1	12	1	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 1	20	2.5	50	34	7.03	11.8	23.2	8.40	15.51	Crop 1	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00	$210.00	$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$4.24	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,945`,
    `2	3	1	Deep OM (CP1)	Crop 2	20	2.5	50	27	5.18	10.6	23.6	14.83	16.46	Crop 2	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385`,
    `3	11	1	Deep Ripping	Crop 3	20	2.5	50	33	7.26	10.7	23.4	17.89	16.41	Crop 3	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `4	1	1	Control	Crop 4	20	2.5	50	29	6.20	10	22.7	12.28	15.19	Crop 4	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695`,
    `5	5	1	Deep Carbon-coated mineral (CCM)	Crop 5	20	2.5	50	28	6.13	10.2	22.8	12.69	13.28	Crop 5	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110`,
    `6	10	1	Deep OM (CP1) + PAM	Crop 6	20	2.5	50	28	7.27	11.6	23.4	16.13	15.20	Crop 6	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `7	9	1	Surface Silicon	Crop 7	20	2.5	50	29	6.78	10.5	23.4	12.23	15.29	Crop 7	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835`,
    `8	4	1	Deep OM + Gypsum (CP2)	Crop 8	20	2.5	50	31	7.60	10.3	25.2	13.87	14.46	Crop 8	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885`,
    `9	6	1	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 9	20	2.5	50	31	5.88	10.3	24.4	14.19	17.95	Crop 9	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110`,
    `10	7	1	Deep liq. NPKS	Crop 10	20	2.5	50	25	7.23	11.5	23.2	12.12	15.57	Crop 10	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `11	2	1	Deep Gypsum	Crop 11	20	2.5	50	22	6.29	9.9	22.8	9.85	15.45	Crop 11	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385`,
    `12	8	1	Deep liq. Gypsum (CHT)	Crop 12	20	2.5	50	26	5.88	9.9	23.5	10.48	11.69	Crop 12	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235`,
    `13	6	2	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 13	20	2.5	50	33	4.79	9.8	24.4	14.49	13.62	Crop 13	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110`,
    `14	7	2	Deep liq. NPKS	Crop 14	20	2.5	50	29	4.88	10.4	23.7	12.81	13.49	Crop 14	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `15	5	2	Deep Carbon-coated mineral (CCM)	Crop 15	20	2.5	50	26	5.39	10.5	23.7	11.97	12.77	Crop 15	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110`,
    `16	3	2	Deep OM (CP1)	Crop 16	20	2.5	50	24	4.96	10.2	23.2	13.85	14.44	Crop 16	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385`,
    `17	1	2	Control	Crop 17	20	2.5	50	24	4.99	10.3	23.3	15.61	10.63	Crop 17	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695`,
    `18	9	2	Surface Silicon	Crop 18	20	2.5	50	27	5.79	10.6	21.1	8.59	10.63	Crop 18	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835`,
    `19	4	2	Deep OM + Gypsum (CP2)	Crop 19	20	2.5	50	22	5.45	11.2	23	12.34	15.59	Crop 19	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885`,
    `20	11	2	Deep Ripping	Crop 20	20	2.5	50	27	6.30	10.4	22.9	12.34	15.28	Crop 20	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `21	8	2	Deep liq. Gypsum (CHT)	Crop 21	20	2.5	50	24	6.57	9.8	23.2	16.16	11.35	Crop 21	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235`,
    `22	12	2	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 22	20	2.5	50	26	6.10	10.3	23.6	14.16	12.21	Crop 22	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735`,
    `23	10	2	Deep OM (CP1) + PAM	Crop 23	20	2.5	50	24	6.34	10	22.8	15.68	12.70	Crop 23	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `24	2	2	Deep Gypsum	Crop 24	20	2.5	50	25	5.44	9.8	23.1	12.70	13.24	Crop 24	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385`,
    `25	6	3	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 25	20	2.5	50	19	5.04	11.2	24.5	15.45	10.97	Crop 25	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110`,
    `26	11	3	Deep Ripping	Crop 26	20	2.5	50	21	6.35	11.2	27.3	19.73	20.65	Crop 26	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `27	2	3	Deep Gypsum	Crop 27	20	2.5	50	21	6.94	10.2	24.6	16.39	14.52	Crop 27	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385`,
    `28	5	3	Deep Carbon-coated mineral (CCM)	Crop 28	20	2.5	50	19	6.31	10.2	23	11.23	15.58	Crop 28	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110`,
    `29	8	3	Deep liq. Gypsum (CHT)	Crop 29	20	2.5	50	26	6.64	11.2	23.5	13.36	14.23	Crop 29	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235`,
    `30	12	3	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 30	20	2.5	50	22	5.96	10.4	23.8	12.01	13.71	Crop 30	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735`,
    `31	10	3	Deep OM (CP1) + PAM	Crop 31	20	2.5	50	22	7.58	10.2	24.2	12.73	11.98	Crop 31	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `32	4	3	Deep OM + Gypsum (CP2)	Crop 32	20	2.5	50	25	6.68	10.3	24.6	13.34	13.12	Crop 32	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885`,
    `33	7	3	Deep liq. NPKS	Crop 33	20	2.5	50	23	7.33	10.1	23.3	13.06	12.18	Crop 33	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `34	1	3	Control	Crop 34	20	2.5	50	25	7.37	10.3	23.3	15.30	9.52	Crop 34	n/a	$0.00	$0.00	$0.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695`,
    `35	3	3	Deep OM (CP1)	Crop 35	20	2.5	50	23	5.29	10.5	23.7	12.61	11.73	Crop 35	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385`,
    `36	9	3	Surface Silicon	Crop 36	20	2.5	50	18	6.81	10	23.8	14.04	17.68	Crop 36	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835`,
    `37	5	4	Deep Carbon-coated mineral (CCM)	Crop 37	20	2.5	50	20	6.42	11.1	23.4	13.51	13.34	Crop 37	5 t/ha	$3,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$4,110`,
    `38	6	4	Deep OM (CP1) + Carbon-coated mineral (CCM)	Crop 38	20	2.5	50	20	6.18	10.6	24.9	14.50	13.16	Crop 38	15 t/ha ; 5 t/ha	$21,225.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$22,110`,
    `39	9	4	Surface Silicon	Crop 39	20	2.5	50	21	6.69	10.8	24.6	13.72	15.00	Crop 39	2 t/ha	?	$35.71	$100.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$835`,
    `40	10	4	Deep OM (CP1) + PAM	Crop 40	20	2.5	50	21	7.72	10.2	23.3	16.55	18.02	Crop 40	15 t/ha ; 5 t/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `41	11	4	Deep Ripping	Crop 41	20	2.5	50	23	6.28	10.6	23.4	10.25	14.71	Crop 41	n/a	$0.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `42	2	4	Deep Gypsum	Crop 42	20	2.5	50	19	5.85	9.8	23.1	10.66	11.19	Crop 42	5 t/ha	$500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,385`,
    `43	7	4	Deep liq. NPKS	Crop 43	20	2.5	50	23	6.40	10.1	23.6	13.28	10.18	Crop 43	750 L/ha	?	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$885`,
    `44	4	4	Deep OM + Gypsum (CP2)	Crop 44	20	2.5	50	33	5.30	9.7	25.5	16.80	13.87	Crop 44	15 t/ha ; 5 t/ha	$24,000.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$24,885`,
    `45	1	4	Control	Crop 45	20	2.5	50	24	6.21	9.9	22.1	10.02	14.31	Crop 45	n/a	$0.00	$0.00	$0.0	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$0.00	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$695`,
    `46	3	4	Deep OM (CP1)	Crop 46	20	2.5	50	28	5.85	10.9	23.9	13.05	13.28	Crop 46	15 t/ha	$16,500.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,385`,
    `47	8	4	Deep liq. Gypsum (CHT)	Crop 47	20	2.5	50	27	5.85	9.6	24.2	20.66	12.83	Crop 47	0.5 t/ha	$350.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$1,235`,
    `48	12	4	Deep OM (CP1) + liq. Gypsum (CHT)	Crop 48	20	2.5	50	23	6.06	10	25.1	15.65	11.32	Crop 48	15 t/ha ; 0.5 t/ha	$16,850.00	$35.71	$150.00	$45.00	$50.00	$3.33	$105.00		$193.00	$30.40	$1.14	$1.14	$0.86	$18.45	$13.00	$14.25	$8.39	$4.88	$12.00	$10.67	$16.50	$16.95	$16.50	$16.50	$2.20	$2.20	$1.31			$1.11	$5.56	$3.67	$6.25	$3.53	$4.55	$3.64	$7.58	$4.93	$3.48	$13.64	$20.00	$2.12	$21.21	$12.12	$2.12		$125,000.00	$259,000.00	$162,800.00	$792,000.00		$17,735`
  ];

  function parseFabaRawLines(lines) {
    const rows = [];
    const rx = /^(\d+)\s+(\d+)\s+(\d+)\s+(.*?)\s+(Crop\s+\d+)\s+(\d+)\s+([\d.]+)\s+(\d+)\s+(\d+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+(Crop\s+\d+)\s+(.*?)\s+(\$[0-9,?.]+|\?)\s+/i;

    for (const raw of lines) {
      const line = String(raw || "").trim();
      if (!line) continue;

      // grab all money tokens to identify last $... (TotalCostPerHa)
      const moneyTokens = line.match(/\$[0-9,]+(?:\.[0-9]+)?/g) || [];
      const totalCostPerHa = moneyTokens.length ? parseMoney(moneyTokens[moneyTokens.length - 1]) : NaN;

      // capital assets (kept as trace info; not forced into model unless user uses them)
      const capitalCandidates = moneyTokens.map(parseMoney).filter((v) => isFinite(v) && v >= 100000);
      const capital = capitalCandidates.slice(0, 4);

      const m = line.match(rx);
      if (!m) {
        rows.push({
          ok: false,
          rawLine: line,
          Plot: NaN,
          Trt: NaN,
          Rep: NaN,
          Amendment: "",
          Crop: "",
          Yield_t_ha: NaN,
          TreatmentInputCostOnly_perHa: NaN,
          TotalCost_perHa: totalCostPerHa,
          capital
        });
        continue;
      }

      const Plot = parseInt(m[1], 10);
      const Trt = parseInt(m[2], 10);
      const Rep = parseInt(m[3], 10);
      const Amendment = m[4].trim();
      const Crop = m[5].trim();
      const PlotLength_m = parseMaybeNumber(m[6]);
      const PlotWidth_m = parseMaybeNumber(m[7]);
      const PlotArea_m2 = parseMaybeNumber(m[8]);
      const Plants_per_m2 = parseMaybeNumber(m[9]);
      const Yield_t_ha = parseMaybeNumber(m[10]);
      const Moisture = parseMaybeNumber(m[11]);
      const Protein = parseMaybeNumber(m[12]);
      const AnthesisBiomass_t_ha = parseMaybeNumber(m[13]);
      const HarvestBiomass_t_ha = parseMaybeNumber(m[14]);
      const Crop2 = m[15].trim();
      const ApplicationRate = m[16].trim();
      const TreatmentInputCostOnly_perHa = parseMoney(m[17]);

      rows.push({
        ok: true,
        rawLine: line,
        Plot,
        Trt,
        Rep,
        Amendment,
        Crop,
        PlotLength_m,
        PlotWidth_m,
        PlotArea_m2,
        Plants_per_m2,
        Yield_t_ha,
        Moisture,
        Protein,
        AnthesisBiomass_t_ha,
        HarvestBiomass_t_ha,
        Crop2,
        ApplicationRate,
        TreatmentInputCostOnly_perHa,
        TotalCost_perHa: totalCostPerHa,
        capital
      });
    }
    return rows;
  }

  // -----------------------------
  // 3) Model (single source of truth)
  // -----------------------------
  const model = {
    meta: {
      toolName: "Farming CBA Decision Tool 2",
      version: "2.0.0",
      lastLoaded: nowISO()
    },
    project: {
      name: "Faba Beans: 2022 soil amendment trial",
      lead: "",
      analysts: "",
      team: "",
      organisation: "Newcastle Business School",
      contactEmail: "",
      contactPhone: "",
      summary:
        "Cost-benefit analysis comparing faba bean treatments (soil amendments and practice changes) against a control, using 2022 production cost data.",
      goal: "Compare economic performance of treatments vs control using NPV, PV benefits, PV costs, BCR, ROI and ranking.",
      withProject: "Treatments are applied and outcomes are tracked over the analysis horizon.",
      withoutProject: "Control practice continues without the additional treatment inputs.",
      objectives: "",
      activities: "",
      stakeholders: "",
      lastUpdated: nowISO()
    },
    time: {
      startYear: new Date().getFullYear(),
      projectStartYear: new Date().getFullYear(),
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4
    },
    discountSchedule: [
      { label: "2025 to 2034", low: 2, base: 4, high: 6 },
      { label: "2035 to 2044", low: 4, base: 7, high: 10 },
      { label: "2045 to 2054", low: 4, base: 7, high: 10 },
      { label: "2055 to 2064", low: 3, base: 6, high: 9 },
      { label: "2065 to 2074", low: 2, base: 5, high: 8 }
    ],
    outputsMeta: {
      systemType: "single",
      assumptions: ""
    },
    outputs: [
      {
        id: uid(),
        name: "Grain yield",
        unit: "t/ha",
        value: 450, // $/t
        source: "Default",
        baselineQuantity: 0 // set from control yield after calibration
      }
    ],
    // treatments populated from raw data calibration
    treatments: [],
    benefits: [],
    otherCosts: [],
    adoption: { low: 0.6, base: 0.9, high: 1.0 },
    risk: { low: 0.05, base: 0.15, high: 0.3, tech: 0.05, nonCoop: 0.04, socio: 0.02, fin: 0.03, man: 0.02 },
    sim: {
      n: 1000,
      targetBCR: 2,
      bcrMode: "all",
      seed: null,
      variationPct: 20,
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false,
      results: { npv: [], bcr: [] }
    },
    raw: {
      faba2022: {
        rows: []
      }
    }
  };

  function ensureSingleControl() {
    const controls = model.treatments.filter((t) => t.isControl);
    if (controls.length === 0 && model.treatments.length) model.treatments[0].isControl = true;
    if (controls.length > 1) {
      // keep first
      let kept = false;
      model.treatments.forEach((t) => {
        if (t.isControl && !kept) kept = true;
        else if (t.isControl && kept) t.isControl = false;
      });
    }
  }

  function recomputeTreatmentTotalCost(t) {
    const labour = toNum(t.labourCost_perHa, 0);
    const materials = toNum(t.materialsCost_perHa, 0);
    const services = toNum(t.servicesCost_perHa, 0);
    const otherVar = toNum(t.otherVarCost_perHa, 0);
    t.totalCost_perHa = labour + materials + services + otherVar;
  }

  function initDeltasForAllTreatments() {
    model.treatments.forEach((t) => {
      t.deltas = t.deltas || {};
      model.outputs.forEach((o) => {
        if (!(o.id in t.deltas)) t.deltas[o.id] = 0;
      });
    });
  }

  // -----------------------------
  // 4) Calibration from default raw dataset
  // -----------------------------
  function calibrateTreatmentsFromFabaRaw() {
    const rows = parseFabaRawLines(FABA_2022_RAW_LINES);
    model.raw.faba2022.rows = rows;

    // group by Trt number + Amendment string (robust)
    const groups = new Map();
    for (const r of rows) {
      if (!r.ok) continue;
      const key = `${r.Trt}||${r.Amendment}`;
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    }

    // identify control group
    let controlKey = null;
    for (const [k, arr] of groups.entries()) {
      const any = arr[0];
      const a = (any.Amendment || "").toLowerCase();
      if (any.Trt === 1 || a.includes("control")) {
        controlKey = k;
        break;
      }
    }
    if (!controlKey && groups.size) controlKey = Array.from(groups.keys())[0];

    function mean(vals) {
      const v = vals.filter((x) => isFinite(x));
      if (!v.length) return NaN;
      return v.reduce((a, b) => a + b, 0) / v.length;
    }

    const controlRows = groups.get(controlKey) || [];
    const controlYield = mean(controlRows.map((r) => r.Yield_t_ha));
    const controlTotalCost = mean(controlRows.map((r) => r.TotalCost_perHa));

    // update output baseline for grain yield
    if (model.outputs[0]) model.outputs[0].baselineQuantity = isFinite(controlYield) ? controlYield : 0;

    const treatments = [];
    for (const [k, arr] of groups.entries()) {
      const first = arr[0];
      const trt = first.Trt;
      const name = first.Amendment || `Treatment ${trt}`;
      const avgYield = mean(arr.map((r) => r.Yield_t_ha));
      const avgTotal = mean(arr.map((r) => r.TotalCost_perHa));
      const avgInput = mean(arr.map((r) => r.TreatmentInputCostOnly_perHa));
      const isControl = k === controlKey;

      // map to cost components while keeping dynamic totals
      const materials = isFinite(avgInput) ? avgInput : 0;
      const total = isFinite(avgTotal) ? avgTotal : 0;
      const otherVar = Math.max(0, total - materials);

      const t = {
        id: uid(),
        trtCode: trt,
        name: name,
        area_ha: 100, // default from trial header
        adoptionMultiplier: 1,
        isControl: !!isControl,
        // cost components (per ha)
        capitalCost_year0: 0,
        labourCost_perHa: 0,
        materialsCost_perHa: materials,
        servicesCost_perHa: 0,
        otherVarCost_perHa: otherVar,
        totalCost_perHa: total,
        // output deltas vs control baseline
        deltas: {},
        notes: "",
        source: "Faba beans 2022 raw data",
        _calibration: {
          meanYield_t_ha: avgYield,
          meanTotalCost_perHa: avgTotal,
          meanInputCostOnly_perHa: avgInput
        }
      };

      // deltas (only grain yield by default)
      const o0 = model.outputs[0];
      if (o0) {
        const dy = isFinite(avgYield) && isFinite(controlYield) ? avgYield - controlYield : 0;
        t.deltas[o0.id] = isControl ? 0 : dy;
      }
      treatments.push(t);
    }

    // sort: control first, then by trtCode ascending
    treatments.sort((a, b) => {
      if (a.isControl && !b.isControl) return -1;
      if (!a.isControl && b.isControl) return 1;
      return (a.trtCode || 0) - (b.trtCode || 0);
    });

    model.treatments = treatments;
    ensureSingleControl();
    initDeltasForAllTreatments();

    // Set a useful default assumption string
    model.outputsMeta.assumptions =
      "Default values use the supplied 2022 faba bean trial dataset. Grain price is configurable. Costs are interpreted as total production cost per hectare per year. Benefits are revenue from grain yield using the configured price. Results compare each treatment against the control column shown alongside.";

    // If controlTotalCost is known, store it in notes for traceability
    const ctrl = model.treatments.find((t) => t.isControl);
    if (ctrl && isFinite(controlTotalCost)) {
      ctrl.notes = `Calibrated from raw data. Mean control yield=${fmt(controlYield)} t/ha; mean control total cost=${money(controlTotalCost)}/ha.`;
    }
  }

  // -----------------------------
  // 5) Core economics
  // -----------------------------
  function discountFactor(tYearIndex, rPct) {
    const r = rPct / 100;
    return 1 / Math.pow(1 + r, tYearIndex);
  }

  function computeTreatmentAnnualPerHa(t, scenario) {
    // scenario: { adoption, risk, priceMultiplier, costMultiplier, years, discBase, ...}
    // Benefits: sum outputs value * (baseline + delta)
    let annualBenefits = 0;
    for (const o of model.outputs) {
      const baseline = toNum(o.baselineQuantity, 0);
      const delta = toNum((t.deltas || {})[o.id], 0);
      const qty = baseline + delta;
      const value = toNum(o.value, 0) * toNum(scenario.priceMultiplier, 1);
      annualBenefits += qty * value;
    }

    // Costs: treatment total cost per ha (constructed from components)
    const totalCost = toNum(t.totalCost_perHa, 0) * toNum(scenario.costMultiplier, 1);

    // Risk: interpreted as reducing realised benefits
    const risk = clamp(toNum(scenario.risk, 0), 0, 1);
    const realisedBenefits = annualBenefits * (1 - risk);

    // Adoption: multiplier on area elsewhere; here keep per-ha
    return {
      annualBenefits_perHa: realisedBenefits,
      annualCosts_perHa: totalCost,
      annualNet_perHa: realisedBenefits - totalCost
    };
  }

  function computePVForTreatmentPerHa(t, years, discBasePct, scenario) {
    // Year 0: capital cost (per ha). If area-based capital is used, user can convert themselves; here per-ha.
    const cap0 = toNum(t.capitalCost_year0, 0);

    let pvBenefits = 0;
    let pvCosts = 0;

    // Costs include year0 capital (undiscounted)
    pvCosts += cap0;

    for (let y = 1; y <= years; y++) {
      const df = discountFactor(y, discBasePct);
      const a = computeTreatmentAnnualPerHa(t, scenario);
      pvBenefits += a.annualBenefits_perHa * df;
      pvCosts += a.annualCosts_perHa * df;
    }

    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? npv / pvCosts : NaN;

    return { pvBenefits, pvCosts, npv, bcr, roi };
  }

  function computeProjectTotalsBase() {
    const years = toNum(model.time.years, 10);
    const r = toNum(model.time.discBase, 7);

    const adoption = clamp(toNum(model.adoption.base, 1), 0, 1);
    const risk = clamp(toNum(model.risk.base, 0), 0, 1);

    const scenario = {
      adoption,
      risk,
      priceMultiplier: 1,
      costMultiplier: 1
    };

    let pvBenefits = 0;
    let pvCosts = 0;

    for (const t of model.treatments) {
      const perHa = computePVForTreatmentPerHa(t, years, r, scenario);
      const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1);
      const adoptFactor = t.isControl ? 1 : adoption;
      pvBenefits += perHa.pvBenefits * area * adoptFactor;
      pvCosts += perHa.pvCosts * area * adoptFactor;
    }

    const npv = pvBenefits - pvCosts;
    const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
    const roi = pvCosts > 0 ? npv / pvCosts : NaN;

    return { pvBenefits, pvCosts, npv, bcr, roi };
  }

  // IRR (simple numeric solve on net cashflows per ha)
  function irrFromCashflows(cashflows) {
    // cashflows indexed t=0..N
    // Solve NPV(r)=0 using bisection on r in [-0.99, 10] (i.e., -99% to 1000%)
    function npv(r) {
      let s = 0;
      for (let t = 0; t < cashflows.length; t++) s += cashflows[t] / Math.pow(1 + r, t);
      return s;
    }
    let lo = -0.99;
    let hi = 10;
    let fLo = npv(lo);
    let fHi = npv(hi);
    if (!isFinite(fLo) || !isFinite(fHi) || fLo === 0) return lo;
    if (fHi === 0) return hi;
    // if no sign change, return NaN
    if (fLo * fHi > 0) return NaN;

    for (let i = 0; i < 80; i++) {
      const mid = (lo + hi) / 2;
      const fMid = npv(mid);
      if (!isFinite(fMid)) return NaN;
      if (Math.abs(fMid) < 1e-8) return mid;
      if (fLo * fMid <= 0) {
        hi = mid;
        fHi = fMid;
      } else {
        lo = mid;
        fLo = fMid;
      }
    }
    return (lo + hi) / 2;
  }

  function paybackPeriod(cashflows, discPct) {
    let cum = 0;
    for (let t = 0; t < cashflows.length; t++) {
      const df = t === 0 ? 1 : discountFactor(t, discPct);
      cum += cashflows[t] * df;
      if (cum >= 0) return t;
    }
    return NaN;
  }

  // -----------------------------
  // 6) Rendering (lists + settings)
  // -----------------------------
  function setBrandingToTool2() {
    document.title = "Farming CBA Decision Tool 2 - Newcastle Business School";
    const brandTitle = $(".brand-title");
    if (brandTitle) brandTitle.textContent = "Farming CBA Decision Tool 2";
    const footer = $(".app-footer .footer-left .small");
    if (footer) footer.textContent = "Farming CBA Decision Tool 2, Newcastle Business School";
  }

  function bindProjectFields() {
    const map = [
      ["projectName", ["project", "name"]],
      ["projectLead", ["project", "lead"]],
      ["analystNames", ["project", "analysts"]],
      ["projectTeam", ["project", "team"]],
      ["organisation", ["project", "organisation"]],
      ["lastUpdated", ["project", "lastUpdated"]],
      ["contactEmail", ["project", "contactEmail"]],
      ["contactPhone", ["project", "contactPhone"]],
      ["projectSummary", ["project", "summary"]],
      ["projectGoal", ["project", "goal"]],
      ["withProject", ["project", "withProject"]],
      ["withoutProject", ["project", "withoutProject"]],
      ["projectObjectives", ["project", "objectives"]],
      ["projectActivities", ["project", "activities"]],
      ["stakeholderGroups", ["project", "stakeholders"]]
    ];

    for (const [id, path] of map) {
      const el = $("#" + id);
      if (!el) continue;
      // set
      el.value = model[path[0]][path[1]] ?? "";
      el.addEventListener("input", () => {
        model[path[0]][path[1]] = el.value;
        if (id === "projectName") setBrandingToTool2();
      });
    }
  }

  function bindSettingsFields() {
    const startYear = $("#startYear");
    const projectStartYear = $("#projectStartYear");
    const years = $("#years");
    const systemType = $("#systemType");
    const discBase = $("#discBase");
    const discLow = $("#discLow");
    const discHigh = $("#discHigh");
    const outputAssumptions = $("#outputAssumptions");
    const mirrFinance = $("#mirrFinance");
    const mirrReinvest = $("#mirrReinvest");

    if (startYear) startYear.value = model.time.startYear;
    if (projectStartYear) projectStartYear.value = model.time.projectStartYear;
    if (years) years.value = model.time.years;
    if (systemType) systemType.value = model.outputsMeta.systemType || "single";
    if (discBase) discBase.value = model.time.discBase;
    if (discLow) discLow.value = model.time.discLow;
    if (discHigh) discHigh.value = model.time.discHigh;
    if (outputAssumptions) outputAssumptions.value = model.outputsMeta.assumptions || "";
    if (mirrFinance) mirrFinance.value = model.time.mirrFinance;
    if (mirrReinvest) mirrReinvest.value = model.time.mirrReinvest;

    const onNum = (el, setter, min = -1e9, max = 1e9) => {
      if (!el) return;
      el.addEventListener("input", () => {
        setter(clamp(el.value, min, max));
      });
    };

    onNum(startYear, (v) => (model.time.startYear = v), 1900, 2500);
    onNum(projectStartYear, (v) => (model.time.projectStartYear = v), 1900, 2500);
    onNum(years, (v) => (model.time.years = Math.max(1, Math.round(v))), 1, 100);
    onNum(discBase, (v) => (model.time.discBase = v), -50, 200);
    onNum(discLow, (v) => (model.time.discLow = v), -50, 200);
    onNum(discHigh, (v) => (model.time.discHigh = v), -50, 200);
    onNum(mirrFinance, (v) => (model.time.mirrFinance = v), -50, 200);
    onNum(mirrReinvest, (v) => (model.time.mirrReinvest = v), -50, 200);

    if (systemType) {
      systemType.addEventListener("change", () => {
        model.outputsMeta.systemType = systemType.value;
      });
    }

    if (outputAssumptions) {
      outputAssumptions.addEventListener("input", () => {
        model.outputsMeta.assumptions = outputAssumptions.value;
      });
    }

    // adoption
    const adoptLow = $("#adoptLow");
    const adoptBase = $("#adoptBase");
    const adoptHigh = $("#adoptHigh");
    if (adoptLow) adoptLow.value = model.adoption.low;
    if (adoptBase) adoptBase.value = model.adoption.base;
    if (adoptHigh) adoptHigh.value = model.adoption.high;
    onNum(adoptLow, (v) => (model.adoption.low = clamp(v, 0, 1)), 0, 1);
    onNum(adoptBase, (v) => (model.adoption.base = clamp(v, 0, 1)), 0, 1);
    onNum(adoptHigh, (v) => (model.adoption.high = clamp(v, 0, 1)), 0, 1);

    // risk
    const riskLow = $("#riskLow");
    const riskBase = $("#riskBase");
    const riskHigh = $("#riskHigh");
    if (riskLow) riskLow.value = model.risk.low;
    if (riskBase) riskBase.value = model.risk.base;
    if (riskHigh) riskHigh.value = model.risk.high;
    onNum(riskLow, (v) => (model.risk.low = clamp(v, 0, 1)), 0, 1);
    onNum(riskBase, (v) => (model.risk.base = clamp(v, 0, 1)), 0, 1);
    onNum(riskHigh, (v) => (model.risk.high = clamp(v, 0, 1)), 0, 1);

    const rTech = $("#rTech"),
      rNonCoop = $("#rNonCoop"),
      rSocio = $("#rSocio"),
      rFin = $("#rFin"),
      rMan = $("#rMan");
    if (rTech) rTech.value = model.risk.tech;
    if (rNonCoop) rNonCoop.value = model.risk.nonCoop;
    if (rSocio) rSocio.value = model.risk.socio;
    if (rFin) rFin.value = model.risk.fin;
    if (rMan) rMan.value = model.risk.man;
    onNum(rTech, (v) => (model.risk.tech = clamp(v, 0, 1)), 0, 1);
    onNum(rNonCoop, (v) => (model.risk.nonCoop = clamp(v, 0, 1)), 0, 1);
    onNum(rSocio, (v) => (model.risk.socio = clamp(v, 0, 1)), 0, 1);
    onNum(rFin, (v) => (model.risk.fin = clamp(v, 0, 1)), 0, 1);
    onNum(rMan, (v) => (model.risk.man = clamp(v, 0, 1)), 0, 1);

    const combinedRiskOut = $("#combinedRiskOut .value");
    const calcCombinedRisk = $("#calcCombinedRisk");
    if (calcCombinedRisk) {
      calcCombinedRisk.addEventListener("click", () => {
        const parts = [model.risk.tech, model.risk.nonCoop, model.risk.socio, model.risk.fin, model.risk.man].map((v) =>
          clamp(toNum(v, 0), 0, 1)
        );
        // combined risk: 1 - product(1 - ri)
        const comb = 1 - parts.reduce((p, ri) => p * (1 - ri), 1);
        model.risk.base = clamp(comb, 0, 1);
        if (riskBase) riskBase.value = model.risk.base;
        if (combinedRiskOut) combinedRiskOut.textContent = percent(comb * 100);
        showToast("Combined base risk updated.");
      });
    }

    // discount schedule table (inputs already in HTML)
    $$("[data-disc-period]").forEach((el) => {
      const p = parseInt(el.getAttribute("data-disc-period"), 10);
      const sc = el.getAttribute("data-scenario");
      const row = model.discountSchedule[p];
      if (!row) return;
      el.value = row[sc];
      el.addEventListener("input", () => {
        row[sc] = clamp(el.value, -50, 200);
      });
    });
  }

  function renderOutputs() {
    const root = $("#outputsList");
    if (!root) return;
    root.innerHTML = "";

    model.outputs.forEach((o, idx) => {
      const wrap = document.createElement("div");
      wrap.className = "card subtle";
      wrap.style.marginBottom = "12px";
      wrap.innerHTML = `
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Name of the output (e.g., Grain yield).">Output name</label>
            <input type="text" value="${esc(o.name)}" data-out="name" data-id="${o.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Unit for the output (e.g., t/ha).">Unit</label>
            <input type="text" value="${esc(o.unit)}" data-out="unit" data-id="${o.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Dollar value per unit (e.g., grain price $/t).">Value per unit ($)</label>
            <input type="number" step="0.01" value="${esc(o.value)}" data-out="value" data-id="${o.id}" />
          </div>
        </div>
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Baseline quantity in the control (used so treatment deltas become absolute quantities).">Baseline quantity (control)</label>
            <input type="number" step="0.0001" value="${esc(o.baselineQuantity ?? 0)}" data-out="baselineQuantity" data-id="${o.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Optional note on where this output value came from.">Source</label>
            <input type="text" value="${esc(o.source || "")}" data-out="source" data-id="${o.id}" />
          </div>
          <div class="field" style="display:flex; align-items:end; justify-content:flex-end;">
            <button class="btn small ghost" data-action="remove-output" data-id="${o.id}" ${
        idx === 0 ? "disabled" : ""
      }>Remove</button>
          </div>
        </div>
      `;
      root.appendChild(wrap);
    });
  }

  function renderTreatments() {
    const root = $("#treatmentsList");
    if (!root) return;
    root.innerHTML = "";

    model.treatments.forEach((t, i) => {
      recomputeTreatmentTotalCost(t);

      const wrap = document.createElement("div");
      wrap.className = "card subtle";
      wrap.style.marginBottom = "12px";

      // delta inputs for each output
      const deltaHtml = model.outputs
        .map((o) => {
          const dv = toNum((t.deltas || {})[o.id], 0);
          const tip =
            "Change in this output relative to control. The tool uses: (control baseline quantity + delta) x value per unit.";
          return `
            <div class="field">
              <label data-tooltip="${esc(tip)}">${esc(o.name)} delta (${esc(o.unit)})</label>
              <input type="number" step="0.0001" value="${esc(dv)}" data-trt="delta" data-oid="${o.id}" data-id="${t.id}" />
            </div>
          `;
        })
        .join("");

      wrap.innerHTML = `
        <div class="row-3">
          <div class="field">
            <label data-tooltip="Treatment name. For the default dataset this matches the amendment label.">Treatment name</label>
            <input type="text" value="${esc(t.name)}" data-trt="name" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Area to which this treatment applies (hectares).">Area (ha)</label>
            <input type="number" step="1" min="0" value="${esc(t.area_ha)}" data-trt="area_ha" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Mark exactly one option as the control so comparisons are shown clearly.">Control</label>
            <div style="display:flex; gap:10px; align-items:center; padding-top:8px;">
              <input type="radio" name="controlRadio" ${t.isControl ? "checked" : ""} data-trt="isControl" data-id="${t.id}" />
              <span class="small muted">${t.isControl ? "Control" : "Treatment"}</span>
            </div>
          </div>
        </div>

        <div class="row-3">
          <div class="field">
            <label data-tooltip="Optional multiplier on area for this specific treatment (kept separate from the scenario adoption setting).">Treatment multiplier</label>
            <input type="number" step="0.05" min="0" value="${esc(t.adoptionMultiplier ?? 1)}" data-trt="adoptionMultiplier" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Capital cost in year 0 (entered as $ per hectare, unless you convert a farm-level cost yourself). IMPORTANT: This field is placed before Total cost as requested.">Capital cost ($, year 0)</label>
            <input type="number" step="0.01" value="${esc(t.capitalCost_year0 ?? 0)}" data-trt="capitalCost_year0" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Total annual production cost per hectare. This is computed from the cost components below and updates dynamically.">Total cost ($/ha)</label>
            <input type="number" step="0.01" value="${esc(t.totalCost_perHa ?? 0)}" data-trt="totalCost_perHa" data-id="${t.id}" />
          </div>
        </div>

        <div class="row-4">
          <div class="field">
            <label data-tooltip="Labour cost per hectare per year for this treatment.">Labour ($/ha)</label>
            <input type="number" step="0.01" value="${esc(t.labourCost_perHa ?? 0)}" data-trt="labourCost_perHa" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Materials cost per hectare per year (inputs). For the default dataset this starts from 'Treatment Input Cost Only /Ha' where available.">Materials ($/ha)</label>
            <input type="number" step="0.01" value="${esc(t.materialsCost_perHa ?? 0)}" data-trt="materialsCost_perHa" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Contracting or services cost per hectare per year.">Services ($/ha)</label>
            <input type="number" step="0.01" value="${esc(t.servicesCost_perHa ?? 0)}" data-trt="servicesCost_perHa" data-id="${t.id}" />
          </div>
          <div class="field">
            <label data-tooltip="Other variable costs per hectare per year. In the default dataset this holds the remainder so Total cost matches the provided Total cost ($/ha).">Other variable ($/ha)</label>
            <input type="number" step="0.01" value="${esc(t.otherVarCost_perHa ?? 0)}" data-trt="otherVarCost_perHa" data-id="${t.id}" />
          </div>
        </div>

        <div class="row-3">
          ${deltaHtml}
        </div>

        <div class="row-2">
          <div class="field">
            <label data-tooltip="Notes that explain what the treatment is and why costs or yields differ.">Notes</label>
            <textarea rows="2" data-trt="notes" data-id="${t.id}">${esc(t.notes || "")}</textarea>
          </div>
          <div class="field" style="display:flex; align-items:end; justify-content:flex-end; gap:10px;">
            <button class="btn small" data-action="jump-results">See results</button>
            <button class="btn small ghost" data-action="remove-treatment" data-id="${t.id}" ${
        model.treatments.length <= 2 ? "disabled" : ""
      }>Remove</button>
          </div>
        </div>
      `;
      root.appendChild(wrap);
    });
  }

  function renderDatabase() {
    const dbO = $("#dbOutputs");
    const dbT = $("#dbTreatments");
    if (dbO) {
      dbO.innerHTML =
        model.outputs
          .map(
            (o) => `
            <div class="card subtle" style="margin-bottom:10px;">
              <div class="row-3">
                <div class="field"><div class="small muted">Output</div><div><strong>${esc(o.name)}</strong></div></div>
                <div class="field"><div class="small muted">Unit</div><div>${esc(o.unit)}</div></div>
                <div class="field"><div class="small muted">Source</div><div>${esc(o.source || "")}</div></div>
              </div>
            </div>
          `
          )
          .join("") || `<div class="small muted">No outputs yet.</div>`;
    }
    if (dbT) {
      dbT.innerHTML =
        model.treatments
          .map(
            (t) => `
            <div class="card subtle" style="margin-bottom:10px;">
              <div class="row-3">
                <div class="field"><div class="small muted">Treatment</div><div><strong>${esc(t.name)}</strong></div></div>
                <div class="field"><div class="small muted">Control?</div><div>${t.isControl ? "Yes" : "No"}</div></div>
                <div class="field"><div class="small muted">Source</div><div>${esc(t.source || "")}</div></div>
              </div>
            </div>
          `
          )
          .join("") || `<div class="small muted">No treatments yet.</div>`;
    }
  }

  // -----------------------------
  // 7) Tabs (fully working)
  // -----------------------------
  function openTab(tabName) {
    const links = $$(".tab-link");
    const panels = $$(".tab-panel");
    links.forEach((b) => {
      const active = b.getAttribute("data-tab") === tabName;
      b.classList.toggle("active", active);
      b.setAttribute("aria-selected", active ? "true" : "false");
    });
    panels.forEach((p) => {
      const active = p.getAttribute("data-tab-panel") === tabName;
      p.classList.toggle("active", active);
      p.classList.toggle("show", active);
      p.setAttribute("aria-hidden", active ? "false" : "true");
    });
    // scroll within main
    const main = $(".app-main");
    if (main) main.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  function bindTabs() {
    $$(".tab-link").forEach((b) => {
      b.addEventListener("click", () => openTab(b.getAttribute("data-tab")));
    });
    $$("[data-tab-jump]").forEach((b) => {
      b.addEventListener("click", () => openTab(b.getAttribute("data-tab-jump")));
    });

    const startBtn = $("#startBtn");
    const startBtnDup = $("#startBtn-duplicate");
    [startBtn, startBtnDup].forEach((btn) => {
      if (!btn) return;
      btn.addEventListener("click", () => openTab("project"));
    });
  }

  // -----------------------------
  // 8) Results rendering (snapshot + comparison table)
  // -----------------------------
  function ensureComparisonTableContainer() {
    const resultsPanel = $("#tab-results .card");
    if (!resultsPanel) return null;

    let box = $("#comparisonBox");
    if (box) return box;

    // Insert before existing "Ranking of treatments"
    const anchor = resultsPanel.querySelector("#treatmentSummary")?.parentElement || $("#treatmentSummary");
    box = document.createElement("div");
    box.id = "comparisonBox";
    box.innerHTML = `
      <h3>Control and treatment comparison table</h3>
      <p class="small muted">
        Indicators are rows and options are columns. The control is shown alongside all treatments for direct comparison.
      </p>
      <div class="table-scroll">
        <table id="comparisonTable" class="summary-table"></table>
      </div>
      <hr />
    `;

    if (anchor && anchor.parentElement) anchor.parentElement.insertBefore(box, anchor);
    else resultsPanel.appendChild(box);

    return box;
  }

  function buildComparisonTable(perHaMetricsById) {
    ensureComparisonTableContainer();
    const tbl = $("#comparisonTable");
    if (!tbl) return;

    const treatments = model.treatments.slice();
    ensureSingleControl();
    treatments.sort((a, b) => {
      if (a.isControl && !b.isControl) return -1;
      if (!a.isControl && b.isControl) return 1;
      return (a.trtCode || 0) - (b.trtCode || 0);
    });

    const control = treatments.find((t) => t.isControl) || treatments[0];

    // ranking by BCR (desc), tie-break NPV (desc)
    const rankList = treatments
      .filter((t) => t.id !== (control?.id || ""))
      .map((t) => ({ id: t.id, bcr: perHaMetricsById[t.id]?.bcr, npv: perHaMetricsById[t.id]?.npv }))
      .sort((x, y) => {
        const xb = isFinite(x.bcr) ? x.bcr : -1e18;
        const yb = isFinite(y.bcr) ? y.bcr : -1e18;
        if (yb !== xb) return yb - xb;
        const xn = isFinite(x.npv) ? x.npv : -1e18;
        const yn = isFinite(y.npv) ? y.npv : -1e18;
        return yn - xn;
      });

    const rankMap = new Map();
    rankList.forEach((x, i) => rankMap.set(x.id, i + 1));
    if (control) rankMap.set(control.id, 0);

    const header = `
      <thead>
        <tr>
          <th style="min-width:220px;" data-tooltip="Economic indicator (row).">Indicator</th>
          ${treatments
            .map((t) => {
              const label = t.isControl ? `${t.name} (Control)` : t.name;
              return `<th style="min-width:220px;" data-tooltip="Option column for comparison.">${esc(label)}</th>`;
            })
            .join("")}
        </tr>
      </thead>
    `;

    const rows = [];

    function row(label, formatter, getter, tip) {
      rows.push(`
        <tr>
          <td data-tooltip="${esc(tip || "")}"><strong>${esc(label)}</strong></td>
          ${treatments
            .map((t) => {
              const v = getter(t);
              return `<td>${formatter(v)}</td>`;
            })
            .join("")}
        </tr>
      `);
    }

    row(
      "PV benefits ($/ha)",
      money,
      (t) => perHaMetricsById[t.id]?.pvBenefits,
      "Present value of benefits per hectare using the base discount rate."
    );
    row(
      "PV costs ($/ha)",
      money,
      (t) => perHaMetricsById[t.id]?.pvCosts,
      "Present value of costs per hectare using the base discount rate (includes any year 0 capital cost entered)."
    );
    row(
      "NPV ($/ha)",
      money,
      (t) => perHaMetricsById[t.id]?.npv,
      "Net present value per hectare (PV benefits minus PV costs)."
    );
    row(
      "BCR (PV benefits / PV costs)",
      ratio,
      (t) => perHaMetricsById[t.id]?.bcr,
      "Benefit-cost ratio. Values above 1 mean benefits exceed costs in present value terms."
    );
    row(
      "ROI (NPV / PV costs)",
      ratio,
      (t) => perHaMetricsById[t.id]?.roi,
      "Return on investment as net gain per dollar of PV cost."
    );
    row(
      "Rank (by BCR)",
      (n) => (n === 0 ? "Control" : isFinite(n) ? String(n) : "n/a"),
      (t) => rankMap.get(t.id),
      "Treatments ranked from highest to lowest BCR (control shown as 'Control')."
    );

    // Differences vs control (helps "compared against control" directly)
    if (control) {
      row(
        "Delta NPV vs control ($/ha)",
        money,
        (t) => perHaMetricsById[t.id]?.npv - perHaMetricsById[control.id]?.npv,
        "Difference in NPV per hectare compared with the control."
      );
      row(
        "Delta PV benefits vs control ($/ha)",
        money,
        (t) => perHaMetricsById[t.id]?.pvBenefits - perHaMetricsById[control.id]?.pvBenefits,
        "Difference in PV benefits per hectare compared with the control."
      );
      row(
        "Delta PV costs vs control ($/ha)",
        money,
        (t) => perHaMetricsById[t.id]?.pvCosts - perHaMetricsById[control.id]?.pvCosts,
        "Difference in PV costs per hectare compared with the control."
      );
    }

    tbl.innerHTML = header + `<tbody>${rows.join("")}</tbody>`;
  }

  function renderRankingCards(perHaMetricsById) {
    const root = $("#treatmentSummary");
    if (!root) return;

    ensureSingleControl();
    const control = model.treatments.find((t) => t.isControl) || model.treatments[0];

    const list = model.treatments
      .filter((t) => !t.isControl)
      .map((t) => ({ t, m: perHaMetricsById[t.id] }))
      .sort((a, b) => {
        const ab = isFinite(a.m?.bcr) ? a.m.bcr : -1e18;
        const bb = isFinite(b.m?.bcr) ? b.m.bcr : -1e18;
        if (bb !== ab) return bb - ab;
        const an = isFinite(a.m?.npv) ? a.m.npv : -1e18;
        const bn = isFinite(b.m?.npv) ? b.m.npv : -1e18;
        return bn - an;
      });

    root.innerHTML = list
      .map(({ t, m }, idx) => {
        const dNpv = control ? m.npv - perHaMetricsById[control.id]?.npv : NaN;
        return `
          <div class="card subtle" style="margin-bottom:10px;">
            <div style="display:flex; align-items:center; justify-content:space-between; gap:12px;">
              <div>
                <div class="small muted">Rank ${idx + 1} (by BCR)</div>
                <div style="font-weight:700;">${esc(t.name)}</div>
              </div>
              <button class="btn small ghost" data-action="focus-treatment" data-id="${t.id}">Edit</button>
            </div>
            <div class="row-4 metrics-grid" style="margin-top:10px;">
              <div class="field"><label>NPV ($/ha)</label><div class="metric"><div class="value">${money(m.npv)}</div></div></div>
              <div class="field"><label>BCR</label><div class="metric"><div class="value">${ratio(m.bcr)}</div></div></div>
              <div class="field"><label>ROI</label><div class="metric"><div class="value">${ratio(m.roi)}</div></div></div>
              <div class="field"><label>Delta NPV vs control ($/ha)</label><div class="metric"><div class="value">${money(dNpv)}</div></div></div>
            </div>
          </div>
        `;
      })
      .join("");
  }

  function renderTimeProjection() {
    const tblBody = $("#timeProjectionTable tbody");
    const canvas = $("#timeNpvChart");
    if (tblBody) tblBody.innerHTML = "";

    const yearsMax = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const horizons = [5, 10, 15, 20, 25].filter((h) => h <= yearsMax);

    const pts = [];

    for (const h of horizons) {
      const oldYears = model.time.years;
      model.time.years = h;
      const totals = computeProjectTotalsBase();
      model.time.years = oldYears;

      pts.push({ x: h, y: totals.npv });

      if (tblBody) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${h}</td>
          <td>${money(totals.pvBenefits)}</td>
          <td>${money(totals.pvCosts)}</td>
          <td>${money(totals.npv)}</td>
          <td>${ratio(totals.bcr)}</td>
        `;
        tblBody.appendChild(tr);
      }
    }

    if (canvas) drawSimpleLine(canvas, pts, "Years", "NPV");
  }

  function drawSimpleLine(canvas, pts, xLabel, yLabel) {
    const ctx = canvas.getContext("2d");
    const w = canvas.width,
      h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    if (!pts.length) return;

    const pad = 35;
    const xs = pts.map((p) => p.x);
    const ys = pts.map((p) => p.y);
    const xmin = Math.min(...xs),
      xmax = Math.max(...xs);
    const ymin = Math.min(...ys),
      ymax = Math.max(...ys);
    const ypad = (ymax - ymin) * 0.1 || 1;

    const xScale = (x) => pad + ((x - xmin) / (xmax - xmin || 1)) * (w - 2 * pad);
    const yScale = (y) => h - pad - ((y - (ymin - ypad)) / ((ymax + ypad) - (ymin - ypad) || 1)) * (h - 2 * pad);

    // axes
    ctx.globalAlpha = 0.9;
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, h - pad);
    ctx.lineTo(w - pad, h - pad);
    ctx.stroke();

    // line
    ctx.beginPath();
    pts.forEach((p, i) => {
      const x = xScale(p.x);
      const y = yScale(p.y);
      if (i === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    });
    ctx.stroke();

    // points
    pts.forEach((p) => {
      const x = xScale(p.x);
      const y = yScale(p.y);
      ctx.beginPath();
      ctx.arc(x, y, 3, 0, Math.PI * 2);
      ctx.fill();
    });

    // labels
    ctx.globalAlpha = 0.8;
    ctx.font = "12px system-ui, -apple-system, Segoe UI, Roboto, Arial";
    ctx.fillText(xLabel, w / 2 - 15, h - 10);
    ctx.save();
    ctx.translate(12, h / 2 + 15);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText(yLabel, 0, 0);
    ctx.restore();
  }

  function drawHistogram(canvas, values) {
    const ctx = canvas.getContext("2d");
    const w = canvas.width,
      h = canvas.height;
    ctx.clearRect(0, 0, w, h);

    const xs = values.filter((v) => isFinite(v));
    if (!xs.length) return;

    const bins = 20;
    const min = Math.min(...xs);
    const max = Math.max(...xs);
    const bw = (max - min) / bins || 1;

    const counts = new Array(bins).fill(0);
    xs.forEach((v) => {
      let i = Math.floor((v - min) / bw);
      if (i < 0) i = 0;
      if (i >= bins) i = bins - 1;
      counts[i] += 1;
    });

    const maxC = Math.max(...counts) || 1;
    const pad = 20;

    // axes
    ctx.beginPath();
    ctx.moveTo(pad, pad);
    ctx.lineTo(pad, h - pad);
    ctx.lineTo(w - pad, h - pad);
    ctx.stroke();

    const barW = (w - 2 * pad) / bins;

    for (let i = 0; i < bins; i++) {
      const x = pad + i * barW;
      const barH = ((h - 2 * pad) * counts[i]) / maxC;
      ctx.fillRect(x + 1, h - pad - barH, barW - 2, barH);
    }
  }

  function renderHeadlineMetrics(projectTotals) {
    // Whole project base case
    if ($("#pvBenefits")) $("#pvBenefits").textContent = money(projectTotals.pvBenefits);
    if ($("#pvCosts")) $("#pvCosts").textContent = money(projectTotals.pvCosts);
    if ($("#npv")) $("#npv").textContent = money(projectTotals.npv);
    if ($("#bcr")) $("#bcr").textContent = ratio(projectTotals.bcr);
    if ($("#roi")) $("#roi").textContent = ratio(projectTotals.roi);

    // IRR, MIRR, payback, margins (project-level, simplified)
    // Use net cashflows: year0 = -sum capital; years 1..N = annual net across treatments
    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const r = toNum(model.time.discBase, 7);
    const adoption = clamp(toNum(model.adoption.base, 1), 0, 1);
    const risk = clamp(toNum(model.risk.base, 0), 0, 1);

    const scenario = { adoption, risk, priceMultiplier: 1, costMultiplier: 1 };

    // cashflows total project
    const cash = new Array(years + 1).fill(0);
    for (const t of model.treatments) {
      const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1);
      const adoptFactor = t.isControl ? 1 : adoption;

      // year0: capital
      cash[0] -= toNum(t.capitalCost_year0, 0) * area * adoptFactor;

      for (let y = 1; y <= years; y++) {
        const a = computeTreatmentAnnualPerHa(t, scenario);
        cash[y] += a.annualNet_perHa * area * adoptFactor;
      }
    }

    const irr = irrFromCashflows(cash);
    if ($("#irr")) $("#irr").textContent = isFinite(irr) ? percent(irr * 100) : "n/a";

    // simplified MIRR
    const fin = toNum(model.time.mirrFinance, 6) / 100;
    const reinv = toNum(model.time.mirrReinvest, 4) / 100;
    let pvNeg = 0;
    let fvPos = 0;
    for (let t = 0; t < cash.length; t++) {
      const cf = cash[t];
      if (cf < 0) pvNeg += cf / Math.pow(1 + fin, t);
      if (cf > 0) fvPos += cf * Math.pow(1 + reinv, years - t);
    }
    const mirr = pvNeg !== 0 ? Math.pow(-fvPos / pvNeg, 1 / years) - 1 : NaN;
    if ($("#mirr")) $("#mirr").textContent = isFinite(mirr) ? percent(mirr * 100) : "n/a";

    const pb = paybackPeriod(cash, r);
    if ($("#payback")) $("#payback").textContent = isFinite(pb) ? `${pb} years` : "n/a";

    // annual gross margin and profit margin
    // (average annual net; profit margin as share of benefits)
    let annualBenefits = 0;
    let annualCosts = 0;
    for (const t of model.treatments) {
      const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1);
      const adoptFactor = t.isControl ? 1 : adoption;
      const a = computeTreatmentAnnualPerHa(t, scenario);
      annualBenefits += a.annualBenefits_perHa * area * adoptFactor;
      annualCosts += a.annualCosts_perHa * area * adoptFactor;
    }
    const gm = annualBenefits - annualCosts;
    const pm = annualBenefits > 0 ? (gm / annualBenefits) * 100 : NaN;
    if ($("#grossMargin")) $("#grossMargin").textContent = money(gm);
    if ($("#profitMargin")) $("#profitMargin").textContent = isFinite(pm) ? percent(pm) : "n/a";
  }

  function renderControlVsTreatments(perHaMetricsById) {
    const control = model.treatments.find((t) => t.isControl);
    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const r = toNum(model.time.discBase, 7);
    const scenario = { adoption: model.adoption.base, risk: model.risk.base, priceMultiplier: 1, costMultiplier: 1 };

    function aggregate(treatments) {
      let pvB = 0,
        pvC = 0;
      for (const t of treatments) {
        const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1);
        const adoptFactor = t.isControl ? 1 : clamp(toNum(model.adoption.base, 1), 0, 1);
        const perHa = computePVForTreatmentPerHa(t, years, r, scenario);
        pvB += perHa.pvBenefits * area * adoptFactor;
        pvC += perHa.pvCosts * area * adoptFactor;
      }
      const npv = pvB - pvC;
      const bcr = pvC > 0 ? pvB / pvC : NaN;
      const roi = pvC > 0 ? npv / pvC : NaN;
      return { pvB, pvC, npv, bcr, roi };
    }

    const ctrlAgg = control ? aggregate([control]) : null;
    const trtAgg = aggregate(model.treatments.filter((t) => !t.isControl));

    // control boxes
    const set = (id, val) => {
      const el = $("#" + id);
      if (el) el.textContent = val;
    };

    if (ctrlAgg) {
      set("pvBenefitsControl", money(ctrlAgg.pvB));
      set("pvCostsControl", money(ctrlAgg.pvC));
      set("npvControl", money(ctrlAgg.npv));
      set("bcrControl", ratio(ctrlAgg.bcr));
      set("roiControl", ratio(ctrlAgg.roi));

      // IRR and payback for control (per ha not perfect; project-level for that group)
      const cash = new Array(years + 1).fill(0);
      const area = toNum(control.area_ha, 0) * toNum(control.adoptionMultiplier, 1);
      cash[0] -= toNum(control.capitalCost_year0, 0) * area;
      for (let y = 1; y <= years; y++) {
        const a = computeTreatmentAnnualPerHa(control, scenario);
        cash[y] += a.annualNet_perHa * area;
      }
      const irr = irrFromCashflows(cash);
      set("irrControl", isFinite(irr) ? percent(irr * 100) : "n/a");
      const pb = paybackPeriod(cash, toNum(model.time.discBase, 7));
      set("paybackControl", isFinite(pb) ? `${pb} years` : "n/a");
      // gmControl as annual average net
      const a0 = computeTreatmentAnnualPerHa(control, scenario);
      set("gmControl", money(a0.annualNet_perHa * area));
    } else {
      ["pvBenefitsControl", "pvCostsControl", "npvControl", "bcrControl", "irrControl", "roiControl", "paybackControl", "gmControl"].forEach(
        (id) => set(id, "n/a")
      );
    }

    set("pvBenefitsTreat", money(trtAgg.pvB));
    set("pvCostsTreat", money(trtAgg.pvC));
    set("npvTreat", money(trtAgg.npv));
    set("bcrTreat", ratio(trtAgg.bcr));
    set("roiTreat", ratio(trtAgg.roi));

    // IRR and payback for combined treatments
    const cashT = new Array(years + 1).fill(0);
    for (const t of model.treatments.filter((x) => !x.isControl)) {
      const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1) * clamp(toNum(model.adoption.base, 1), 0, 1);
      cashT[0] -= toNum(t.capitalCost_year0, 0) * area;
      for (let y = 1; y <= years; y++) {
        const a = computeTreatmentAnnualPerHa(t, scenario);
        cashT[y] += a.annualNet_perHa * area;
      }
    }
    const irrT = irrFromCashflows(cashT);
    const pbT = paybackPeriod(cashT, toNum(model.time.discBase, 7));
    set("irrTreat", isFinite(irrT) ? percent(irrT * 100) : "n/a");
    set("paybackTreat", isFinite(pbT) ? `${pbT} years` : "n/a");
    // gmTreat: annual net
    let gmTreat = 0;
    for (const t of model.treatments.filter((x) => !x.isControl)) {
      const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1) * clamp(toNum(model.adoption.base, 1), 0, 1);
      const a = computeTreatmentAnnualPerHa(t, scenario);
      gmTreat += a.annualNet_perHa * area;
    }
    set("gmTreat", money(gmTreat));
  }

  function buildCopilotPayload(perHaMetricsById) {
    const toolName = model.meta.toolName;
    ensureSingleControl();
    const control = model.treatments.find((t) => t.isControl) || model.treatments[0];

    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const r = toNum(model.time.discBase, 7);
    const adoption = clamp(toNum(model.adoption.base, 1), 0, 1);
    const risk = clamp(toNum(model.risk.base, 0), 0, 1);

    const options = model.treatments
      .map((t) => {
        const m = perHaMetricsById[t.id] || {};
        const deltaVsCtrl =
          control && perHaMetricsById[control.id]
            ? {
                deltaNPV_perHa: m.npv - perHaMetricsById[control.id].npv,
                deltaPVBenefits_perHa: m.pvBenefits - perHaMetricsById[control.id].pvBenefits,
                deltaPVCosts_perHa: m.pvCosts - perHaMetricsById[control.id].pvCosts
              }
            : null;

        // drivers (simple: yield and cost vs control)
        const o0 = model.outputs[0];
        const ctrlYield = o0 ? toNum(o0.baselineQuantity, 0) : 0;
        const dy = o0 ? toNum((t.deltas || {})[o0.id], 0) : 0;
        const absYield = ctrlYield + dy;

        return {
          id: t.id,
          name: t.name,
          isControl: t.isControl,
          area_ha: toNum(t.area_ha, 0),
          capitalCost_year0_perHa: toNum(t.capitalCost_year0, 0),
          totalCost_perHa_perYear: toNum(t.totalCost_perHa, 0),
          outputs: model.outputs.map((o) => ({
            name: o.name,
            unit: o.unit,
            value_per_unit: toNum(o.value, 0),
            baseline_control_quantity: toNum(o.baselineQuantity, 0),
            delta_vs_control: toNum((t.deltas || {})[o.id], 0),
            implied_absolute_quantity: toNum(o.baselineQuantity, 0) + toNum((t.deltas || {})[o.id], 0)
          })),
          headline_perHa: {
            pvBenefits: m.pvBenefits,
            pvCosts: m.pvCosts,
            npv: m.npv,
            bcr: m.bcr,
            roi: m.roi
          },
          delta_vs_control_perHa: deltaVsCtrl,
          simple_drivers: {
            implied_yield_t_per_ha: absYield,
            cost_per_ha_per_year: toNum(t.totalCost_perHa, 0),
            yield_delta_t_per_ha_vs_control: dy,
            cost_delta_per_ha_vs_control: control ? toNum(t.totalCost_perHa, 0) - toNum(control.totalCost_perHa, 0) : null
          }
        };
      })
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;
        const ab = isFinite(a.headline_perHa?.bcr) ? a.headline_perHa.bcr : -1e18;
        const bb = isFinite(b.headline_perHa?.bcr) ? b.headline_perHa.bcr : -1e18;
        return bb - ab;
      });

    const instructions = {
      task:
        "Write a plain-English interpretation of the CBA results for a farmer or on-farm manager. Explain what each indicator means, compare each treatment against the control, and highlight trade-offs. Do not recommend a single choice or impose thresholds as rules.",
      required_sections: [
        "Brief context (what the tool is, what data was used)",
        "Explain the indicators (PV benefits, PV costs, NPV, BCR, ROI, ranking)",
        "Compare treatments vs the control (what drives differences: yield, price, costs, capital)",
        "What would need to change for underperforming treatments to improve (practical levers: reduce costs, raise yield, improve price, adjust practices) framed as reflection and options, not rules",
        "Sensitivities to check (grain price, yield stability, costs, discount rate, adoption, risk)",
        "Plain-language conclusion summarising the trade-offs without telling the user what to choose"
      ],
      guardrails: [
        "Do not say 'choose treatment X'.",
        "Do not hide low-performing treatments: discuss why they underperform.",
        "If BCR is low, suggest realistic improvement paths (cost reduction, yield improvement, price improvement, agronomic changes).",
        "If information is missing (e.g., unknown input cost in raw data), acknowledge uncertainty and point to fields that can be updated in the tool."
      ],
      tone: "Practical, farmer-friendly, non-prescriptive, decision-support."
    };

    return {
      tool: toolName,
      timestamp: new Date().toISOString(),
      scenario: {
        years,
        discount_rate_base_percent: r,
        adoption_base_multiplier: adoption,
        risk_base: risk,
        assumptions: model.outputsMeta.assumptions || ""
      },
      project: model.project,
      options,
      instructions
    };
  }

  function updateCopilotPreview(perHaMetricsById) {
    const preview = $("#copilotPreview");
    if (!preview) return;
    const payload = buildCopilotPayload(perHaMetricsById);
    preview.value = JSON.stringify(payload, null, 2);
  }

  function recalcAndRender() {
    // compute per-ha metrics for each treatment
    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const r = toNum(model.time.discBase, 7);
    const scenario = {
      adoption: clamp(toNum(model.adoption.base, 1), 0, 1),
      risk: clamp(toNum(model.risk.base, 0), 0, 1),
      priceMultiplier: 1,
      costMultiplier: 1
    };

    const perHaMetricsById = {};
    for (const t of model.treatments) {
      // ensure totals consistent
      recomputeTreatmentTotalCost(t);
      perHaMetricsById[t.id] = computePVForTreatmentPerHa(t, years, r, scenario);
    }

    // project totals
    const totals = computeProjectTotalsBase();
    renderHeadlineMetrics(totals);

    buildComparisonTable(perHaMetricsById);
    renderRankingCards(perHaMetricsById);
    renderControlVsTreatments(perHaMetricsById);
    renderTimeProjection();
    renderDatabase();
    updateCopilotPreview(perHaMetricsById);

    return { totals, perHaMetricsById };
  }

  // -----------------------------
  // 9) Simulation
  // -----------------------------
  function bindSimulationControls() {
    const simN = $("#simN");
    const targetBCR = $("#targetBCR");
    const bcrMode = $("#bcrMode");
    const randSeed = $("#randSeed");
    const simVarPct = $("#simVarPct");
    const simVaryOutputs = $("#simVaryOutputs");
    const simVaryTreatCosts = $("#simVaryTreatCosts");
    const simVaryInputCosts = $("#simVaryInputCosts");
    const simBcrTargetLabel = $("#simBcrTargetLabel");

    if (simN) simN.value = model.sim.n;
    if (targetBCR) targetBCR.value = model.sim.targetBCR;
    if (bcrMode) bcrMode.value = model.sim.bcrMode;
    if (randSeed) randSeed.value = model.sim.seed ?? "";
    if (simVarPct) simVarPct.value = model.sim.variationPct;
    if (simVaryOutputs) simVaryOutputs.value = String(model.sim.varyOutputs);
    if (simVaryTreatCosts) simVaryTreatCosts.value = String(model.sim.varyTreatCosts);
    if (simVaryInputCosts) simVaryInputCosts.value = String(model.sim.varyInputCosts);
    if (simBcrTargetLabel) simBcrTargetLabel.textContent = String(model.sim.targetBCR);

    const on = (el, fn) => el && el.addEventListener("input", fn);
    on(simN, () => (model.sim.n = Math.max(100, Math.round(toNum(simN.value, 1000)))));
    on(targetBCR, () => {
      model.sim.targetBCR = Math.max(0, toNum(targetBCR.value, 2));
      if (simBcrTargetLabel) simBcrTargetLabel.textContent = String(model.sim.targetBCR);
    });
    if (bcrMode) bcrMode.addEventListener("change", () => (model.sim.bcrMode = bcrMode.value));
    on(randSeed, () => (model.sim.seed = randSeed.value === "" ? null : Math.round(toNum(randSeed.value, 1))));
    on(simVarPct, () => (model.sim.variationPct = clamp(simVarPct.value, 0, 100)));
    if (simVaryOutputs) simVaryOutputs.addEventListener("change", () => (model.sim.varyOutputs = simVaryOutputs.value === "true"));
    if (simVaryTreatCosts) simVaryTreatCosts.addEventListener("change", () => (model.sim.varyTreatCosts = simVaryTreatCosts.value === "true"));
    if (simVaryInputCosts) simVaryInputCosts.addEventListener("change", () => (model.sim.varyInputCosts = simVaryInputCosts.value === "true"));

    const runSim = $("#runSim");
    if (runSim) {
      runSim.addEventListener("click", () => {
        const status = $("#simStatus");
        if (status) status.textContent = "Running simulation...";
        setTimeout(() => {
          const out = runSimulation();
          if (status) status.textContent = `Simulation complete (${out.n} runs).`;
        }, 50);
      });
    }
  }

  function runSimulation() {
    const n = Math.max(100, Math.round(toNum(model.sim.n, 1000)));
    const seed = model.sim.seed;
    const R = rng(seed || undefined);

    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));

    const discLow = toNum(model.time.discLow, 4);
    const discBase = toNum(model.time.discBase, 7);
    const discHigh = toNum(model.time.discHigh, 10);

    const adoptLow = clamp(toNum(model.adoption.low, 0.6), 0, 1);
    const adoptBase = clamp(toNum(model.adoption.base, 0.9), 0, 1);
    const adoptHigh = clamp(toNum(model.adoption.high, 1.0), 0, 1);

    const riskLow = clamp(toNum(model.risk.low, 0.05), 0, 1);
    const riskBase = clamp(toNum(model.risk.base, 0.15), 0, 1);
    const riskHigh = clamp(toNum(model.risk.high, 0.3), 0, 1);

    const varPct = clamp(toNum(model.sim.variationPct, 20), 0, 100) / 100;

    const npvArr = [];
    const bcrArr = [];

    // cache base outputs + costs
    const baseOutputValues = model.outputs.map((o) => toNum(o.value, 0));
    const baseTreatCosts = model.treatments.map((t) => toNum(t.totalCost_perHa, 0));

    for (let i = 0; i < n; i++) {
      const disc = triangular(R(), discLow, discBase, discHigh);
      const adoption = triangular(R(), adoptLow, adoptBase, adoptHigh);
      const risk = triangular(R(), riskLow, riskBase, riskHigh);

      const scenario = {
        adoption,
        risk,
        priceMultiplier: 1,
        costMultiplier: 1
      };

      // apply multiplicative noise where enabled
      if (model.sim.varyOutputs) {
        model.outputs.forEach((o, k) => {
          const m = 1 + (R() * 2 - 1) * varPct;
          o.value = baseOutputValues[k] * m;
        });
      }

      if (model.sim.varyTreatCosts) {
        model.treatments.forEach((t, k) => {
          const m = 1 + (R() * 2 - 1) * varPct;
          t.totalCost_perHa = baseTreatCosts[k] * m;
        });
      }

      // compute totals with drawn discount rate
      let pvB = 0,
        pvC = 0;
      for (const t of model.treatments) {
        const perHa = computePVForTreatmentPerHa(t, years, disc, scenario);
        const area = toNum(t.area_ha, 0) * toNum(t.adoptionMultiplier, 1);
        const adoptFactor = t.isControl ? 1 : adoption;
        pvB += perHa.pvBenefits * area * adoptFactor;
        pvC += perHa.pvCosts * area * adoptFactor;
      }
      const npv = pvB - pvC;
      const bcr = pvC > 0 ? pvB / pvC : NaN;

      npvArr.push(npv);
      bcrArr.push(bcr);
    }

    // restore base values
    model.outputs.forEach((o, k) => (o.value = baseOutputValues[k]));
    model.treatments.forEach((t, k) => (t.totalCost_perHa = baseTreatCosts[k]));

    model.sim.results = { npv: npvArr, bcr: bcrArr };

    // summarise
    const finite = (arr) => arr.filter((x) => isFinite(x));
    const aN = finite(npvArr);
    const aB = finite(bcrArr);

    const stat = (arr) => {
      if (!arr.length) return { min: NaN, max: NaN, mean: NaN, median: NaN };
      const s = arr.slice().sort((x, y) => x - y);
      const mean = s.reduce((p, c) => p + c, 0) / s.length;
      const median = s.length % 2 ? s[(s.length - 1) / 2] : 0.5 * (s[s.length / 2 - 1] + s[s.length / 2]);
      return { min: s[0], max: s[s.length - 1], mean, median };
    };

    const sn = stat(aN);
    const sb = stat(aB);

    const pNpv = aN.length ? aN.filter((x) => x > 0).length / aN.length : NaN;
    const pBcr1 = aB.length ? aB.filter((x) => x > 1).length / aB.length : NaN;
    const pBcrT = aB.length ? aB.filter((x) => x > toNum(model.sim.targetBCR, 2)).length / aB.length : NaN;

    const set = (id, v) => {
      const el = $("#" + id);
      if (!el) return;
      el.textContent = v;
    };

    set("simNpvMin", money(sn.min));
    set("simNpvMax", money(sn.max));
    set("simNpvMean", money(sn.mean));
    set("simNpvMedian", money(sn.median));
    set("simNpvProb", isFinite(pNpv) ? percent(pNpv * 100) : "n/a");

    set("simBcrMin", ratio(sb.min));
    set("simBcrMax", ratio(sb.max));
    set("simBcrMean", ratio(sb.mean));
    set("simBcrMedian", ratio(sb.median));
    set("simBcrProb1", isFinite(pBcr1) ? percent(pBcr1 * 100) : "n/a");
    set("simBcrProbTarget", isFinite(pBcrT) ? percent(pBcrT * 100) : "n/a");

    const histNpv = $("#histNpv");
    const histBcr = $("#histBcr");
    if (histNpv) drawHistogram(histNpv, aN);
    if (histBcr) drawHistogram(histBcr, aB);

    return { n };
  }

  // -----------------------------
  // 10) Excel-first workflow (export + import)
  // -----------------------------
  function hasXLSX() {
    return typeof window.XLSX !== "undefined";
  }

  function buildWorkbookFromModel(includeRaw = true) {
    if (!hasXLSX()) return null;

    const XLSX = window.XLSX;
    const wb = XLSX.utils.book_new();

    // Summary sheet (clean, Word-copy friendly)
    const years = Math.max(1, Math.round(toNum(model.time.years, 10)));
    const r = toNum(model.time.discBase, 7);
    const scenario = { adoption: model.adoption.base, risk: model.risk.base, priceMultiplier: 1, costMultiplier: 1 };

    const perHa = {};
    for (const t of model.treatments) perHa[t.id] = computePVForTreatmentPerHa(t, years, r, scenario);

    const control = model.treatments.find((t) => t.isControl) || model.treatments[0];

    const indicators = [
      ["PV benefits ($/ha)", (t) => perHa[t.id].pvBenefits],
      ["PV costs ($/ha)", (t) => perHa[t.id].pvCosts],
      ["NPV ($/ha)", (t) => perHa[t.id].npv],
      ["BCR", (t) => perHa[t.id].bcr],
      ["ROI", (t) => perHa[t.id].roi],
      ["Delta NPV vs control ($/ha)", (t) => perHa[t.id].npv - perHa[control.id].npv]
    ];

    const header = ["Indicator"].concat(
      model.treatments
        .slice()
        .sort((a, b) => {
          if (a.isControl && !b.isControl) return -1;
          if (!a.isControl && b.isControl) return 1;
          return (a.trtCode || 0) - (b.trtCode || 0);
        })
        .map((t) => (t.isControl ? `${t.name} (Control)` : t.name))
    );

    const aoa = [header];
    const ordered = model.treatments
      .slice()
      .sort((a, b) => {
        if (a.isControl && !b.isControl) return -1;
        if (!a.isControl && b.isControl) return 1;
        return (a.trtCode || 0) - (b.trtCode || 0);
      });

    for (const [lab, fn] of indicators) {
      const row = [lab].concat(
        ordered.map((t) => {
          const v = fn(t);
          return isFinite(v) ? v : null;
        })
      );
      aoa.push(row);
    }

    const wsSummary = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Results_vs_Control");

    // Treatments sheet (editable)
    const trtRows = model.treatments.map((t) => {
      const o0 = model.outputs[0];
      return {
        id: t.id,
        trtCode: t.trtCode,
        name: t.name,
        isControl: t.isControl ? 1 : 0,
        area_ha: toNum(t.area_ha, 0),
        adoptionMultiplier: toNum(t.adoptionMultiplier, 1),
        capitalCost_year0_perHa: toNum(t.capitalCost_year0, 0),
        labourCost_perHa: toNum(t.labourCost_perHa, 0),
        materialsCost_perHa: toNum(t.materialsCost_perHa, 0),
        servicesCost_perHa: toNum(t.servicesCost_perHa, 0),
        otherVarCost_perHa: toNum(t.otherVarCost_perHa, 0),
        totalCost_perHa: toNum(t.totalCost_perHa, 0),
        yield_delta_t_perHa: o0 ? toNum((t.deltas || {})[o0.id], 0) : 0,
        notes: t.notes || ""
      };
    });
    const wsTrt = XLSX.utils.json_to_sheet(trtRows);
    XLSX.utils.book_append_sheet(wb, wsTrt, "Treatments");

    // Outputs sheet
    const outRows = model.outputs.map((o) => ({
      id: o.id,
      name: o.name,
      unit: o.unit,
      value_per_unit: toNum(o.value, 0),
      baseline_control_quantity: toNum(o.baselineQuantity, 0),
      source: o.source || ""
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(outRows), "Outputs");

    // Settings sheet
    const setRows = [
      { key: "toolName", value: model.meta.toolName },
      { key: "years", value: toNum(model.time.years, 10) },
      { key: "discount_rate_base_percent", value: toNum(model.time.discBase, 7) },
      { key: "discount_rate_low_percent", value: toNum(model.time.discLow, 4) },
      { key: "discount_rate_high_percent", value: toNum(model.time.discHigh, 10) },
      { key: "adoption_low", value: toNum(model.adoption.low, 0.6) },
      { key: "adoption_base", value: toNum(model.adoption.base, 0.9) },
      { key: "adoption_high", value: toNum(model.adoption.high, 1.0) },
      { key: "risk_low", value: toNum(model.risk.low, 0.05) },
      { key: "risk_base", value: toNum(model.risk.base, 0.15) },
      { key: "risk_high", value: toNum(model.risk.high, 0.3) },
      { key: "assumptions", value: model.outputsMeta.assumptions || "" }
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(setRows), "Settings");

    // Raw sheet (full preservation via rawLine + parsed fields)
    if (includeRaw) {
      const rawRows = model.raw.faba2022.rows.map((r) => ({
        ok: r.ok ? 1 : 0,
        Plot: r.Plot,
        Trt: r.Trt,
        Rep: r.Rep,
        Amendment: r.Amendment,
        Crop: r.Crop,
        Yield_t_ha: r.Yield_t_ha,
        TreatmentInputCostOnly_perHa: r.TreatmentInputCostOnly_perHa,
        TotalCost_perHa: r.TotalCost_perHa,
        rawLine: r.rawLine
      }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rawRows), "FabaBeans_2022_Raw");
    }

    return wb;
  }

  function downloadWorkbook(wb, filename) {
    if (!hasXLSX() || !wb) return false;
    const XLSX = window.XLSX;
    XLSX.writeFile(wb, filename);
    return true;
  }

  function bindExcelControls() {
    const downloadTemplate = $("#downloadTemplate");
    const downloadSample = $("#downloadSample");
    const parseExcel = $("#parseExcel");
    const importExcel = $("#importExcel");

    // hidden input for excel files
    let excelInput = $("#_excelFileInput");
    if (!excelInput) {
      excelInput = document.createElement("input");
      excelInput.type = "file";
      excelInput.accept = ".xlsx,.xls";
      excelInput.id = "_excelFileInput";
      excelInput.style.display = "none";
      document.body.appendChild(excelInput);
    }

    let parsed = null;

    function parseWorkbook(file) {
      return new Promise((resolve, reject) => {
        if (!hasXLSX()) {
          reject(new Error("XLSX library not available."));
          return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const wb = window.XLSX.read(data, { type: "array" });
            resolve(wb);
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
      });
    }

    function safeSheetToJSON(wb, name) {
      const XLSX = window.XLSX;
      const ws = wb.Sheets[name];
      if (!ws) return null;
      return XLSX.utils.sheet_to_json(ws, { defval: null });
    }

    if (downloadTemplate) {
      downloadTemplate.addEventListener("click", () => {
        if (!hasXLSX()) {
          showToast("Excel export requires the XLSX library (SheetJS).");
          return;
        }
        // Blank template: same structure but minimal rows
        const XLSX = window.XLSX;
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(
          wb,
          XLSX.utils.json_to_sheet([
            { id: "", trtCode: "", name: "", isControl: 0, area_ha: "", adoptionMultiplier: "", capitalCost_year0_perHa: "", labourCost_perHa: "", materialsCost_perHa: "", servicesCost_perHa: "", otherVarCost_perHa: "", totalCost_perHa: "", yield_delta_t_perHa: "", notes: "" }
          ]),
          "Treatments"
        );
        XLSX.utils.book_append_sheet(
          wb,
          XLSX.utils.json_to_sheet([{ id: "", name: "", unit: "", value_per_unit: "", baseline_control_quantity: "", source: "" }]),
          "Outputs"
        );
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ key: "", value: "" }]), "Settings");
        XLSX.utils.book_append_sheet(
          wb,
          XLSX.utils.json_to_sheet([{ ok: 1, Plot: "", Trt: "", Rep: "", Amendment: "", Crop: "", Yield_t_ha: "", TreatmentInputCostOnly_perHa: "", TotalCost_perHa: "", rawLine: "" }]),
          "FabaBeans_2022_Raw"
        );

        const fn = `${slug(model.meta.toolName)}_template.xlsx`;
        XLSX.writeFile(wb, fn);
        showToast("Template downloaded.");
      });
    }

    if (downloadSample) {
      downloadSample.addEventListener("click", () => {
        const wb = buildWorkbookFromModel(true);
        if (!wb) {
          showToast("Excel export requires the XLSX library (SheetJS).");
          return;
        }
        const fn = `${slug(model.meta.toolName)}_sample_${slug(model.project.name)}.xlsx`;
        downloadWorkbook(wb, fn);
        showToast("Sample workbook downloaded.");
      });
    }

    if (parseExcel) {
      parseExcel.addEventListener("click", () => {
        if (!hasXLSX()) {
          showToast("Excel import requires the XLSX library (SheetJS).");
          return;
        }
        excelInput.value = "";
        excelInput.click();
      });
    }

    excelInput.addEventListener("change", async () => {
      const file = excelInput.files && excelInput.files[0];
      if (!file) return;
      try {
        const wb = await parseWorkbook(file);
        parsed = {
          wb,
          treatments: safeSheetToJSON(wb, "Treatments"),
          outputs: safeSheetToJSON(wb, "Outputs"),
          settings: safeSheetToJSON(wb, "Settings"),
          raw: safeSheetToJSON(wb, "FabaBeans_2022_Raw"),
          filename: file.name
        };
        showToast(`Excel parsed: ${file.name}. Click 'Apply parsed Excel data' to import.`);
      } catch (e) {
        console.error(e);
        showToast("Failed to parse Excel file.");
      }
    });

    if (importExcel) {
      importExcel.addEventListener("click", () => {
        if (!parsed) {
          showToast("No parsed Excel data yet. Click 'Parse Excel file' first.");
          return;
        }

        // Outputs
        if (Array.isArray(parsed.outputs) && parsed.outputs.length) {
          const outs = parsed.outputs
            .filter((r) => r && (r.name || r.id))
            .map((r) => ({
              id: r.id && String(r.id).trim() ? String(r.id).trim() : uid(),
              name: String(r.name || "Output").trim(),
              unit: String(r.unit || "").trim(),
              value: toNum(r.value_per_unit ?? r.value ?? 0, 0),
              baselineQuantity: toNum(r.baseline_control_quantity ?? r.baselineQuantity ?? 0, 0),
              source: String(r.source || "").trim()
            }));
          if (outs.length) model.outputs = outs;
        }

        // Treatments
        if (Array.isArray(parsed.treatments) && parsed.treatments.length) {
          const o0 = model.outputs[0];
          const trts = parsed.treatments
            .filter((r) => r && (r.name || r.id))
            .map((r) => {
              const id = r.id && String(r.id).trim() ? String(r.id).trim() : uid();
              const t = {
                id,
                trtCode: toNum(r.trtCode, NaN),
                name: String(r.name || "Treatment").trim(),
                isControl: toNum(r.isControl, 0) === 1,
                area_ha: toNum(r.area_ha, 0),
                adoptionMultiplier: toNum(r.adoptionMultiplier, 1),
                capitalCost_year0: toNum(r.capitalCost_year0_perHa ?? r.capitalCost_year0 ?? 0, 0),
                labourCost_perHa: toNum(r.labourCost_perHa ?? 0, 0),
                materialsCost_perHa: toNum(r.materialsCost_perHa ?? 0, 0),
                servicesCost_perHa: toNum(r.servicesCost_perHa ?? 0, 0),
                otherVarCost_perHa: toNum(r.otherVarCost_perHa ?? 0, 0),
                totalCost_perHa: toNum(r.totalCost_perHa ?? 0, 0),
                deltas: {},
                notes: String(r.notes || "").trim(),
                source: `Imported from ${parsed.filename}`
              };
              if (o0) {
                t.deltas[o0.id] = toNum(r.yield_delta_t_perHa ?? 0, 0);
              }
              recomputeTreatmentTotalCost(t);
              return t;
            });
          if (trts.length) model.treatments = trts;
        }

        // Settings
        if (Array.isArray(parsed.settings) && parsed.settings.length) {
          const map = new Map(parsed.settings.map((r) => [String(r.key || "").trim(), r.value]));
          if (map.has("years")) model.time.years = Math.max(1, Math.round(toNum(map.get("years"), model.time.years)));
          if (map.has("discount_rate_base_percent")) model.time.discBase = toNum(map.get("discount_rate_base_percent"), model.time.discBase);
          if (map.has("discount_rate_low_percent")) model.time.discLow = toNum(map.get("discount_rate_low_percent"), model.time.discLow);
          if (map.has("discount_rate_high_percent")) model.time.discHigh = toNum(map.get("discount_rate_high_percent"), model.time.discHigh);
          if (map.has("adoption_low")) model.adoption.low = clamp(toNum(map.get("adoption_low"), model.adoption.low), 0, 1);
          if (map.has("adoption_base")) model.adoption.base = clamp(toNum(map.get("adoption_base"), model.adoption.base), 0, 1);
          if (map.has("adoption_high")) model.adoption.high = clamp(toNum(map.get("adoption_high"), model.adoption.high), 0, 1);
          if (map.has("risk_low")) model.risk.low = clamp(toNum(map.get("risk_low"), model.risk.low), 0, 1);
          if (map.has("risk_base")) model.risk.base = clamp(toNum(map.get("risk_base"), model.risk.base), 0, 1);
          if (map.has("risk_high")) model.risk.high = clamp(toNum(map.get("risk_high"), model.risk.high), 0, 1);
          if (map.has("assumptions")) model.outputsMeta.assumptions = String(map.get("assumptions") ?? "");
        }

        ensureSingleControl();
        initDeltasForAllTreatments();

        // re-render UI + recalc
        bindProjectFields();
        bindSettingsFields();
        renderOutputs();
        renderTreatments();

        recalcAndRender();
        showToast("Excel data imported and results updated.");
      });
    }
  }

  // -----------------------------
  // 11) Export (CSV/Excel) + PDF print
  // -----------------------------
  function exportResultsToExcelOrCsv() {
    // Prefer XLSX workbook
    const wb = buildWorkbookFromModel(true);
    if (wb && hasXLSX()) {
      const fn = `${slug(model.meta.toolName)}_results_${slug(model.project.name)}_${nowISO()}.xlsx`;
      downloadWorkbook(wb, fn);
      return;
    }

    // CSV fallback: only the comparison table
    const tbl = $("#comparisonTable");
    if (!tbl) {
      showToast("Nothing to export yet. Recalculate results first.");
      return;
    }
    const rows = [];
    const trs = Array.from(tbl.querySelectorAll("tr"));
    trs.forEach((tr) => {
      const cells = Array.from(tr.querySelectorAll("th,td")).map((c) => `"${String(c.textContent || "").replace(/"/g, '""')}"`);
      rows.push(cells.join(","));
    });

    const blob = new Blob([rows.join("\n")], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${slug(model.meta.toolName)}_results_${nowISO()}.csv`;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function bindExports() {
    const exportCsv = $("#exportCsv");
    const exportCsvFoot = $("#exportCsvFoot");
    const exportPdf = $("#exportPdf");
    const exportPdfFoot = $("#exportPdfFoot");

    [exportCsv, exportCsvFoot].forEach((btn) => {
      if (!btn) return;
      btn.addEventListener("click", () => {
        exportResultsToExcelOrCsv();
        showToast("Results exported.");
      });
    });

    [exportPdf, exportPdfFoot].forEach((btn) => {
      if (!btn) return;
      btn.addEventListener("click", () => {
        openTab("results");
        // allow layout to settle then print
        setTimeout(() => window.print(), 150);
      });
    });
  }

  // -----------------------------
  // 12) Copilot helper button
  // -----------------------------
  function bindCopilot() {
    const openCopilot = $("#openCopilot");
    if (!openCopilot) return;
    openCopilot.addEventListener("click", async () => {
      // ensure preview is fresh
      const { perHaMetricsById } = recalcAndRender();
      updateCopilotPreview(perHaMetricsById);

      const txt = $("#copilotPreview")?.value || "";
      try {
        await navigator.clipboard.writeText(txt);
        showToast("Scenario summary copied to clipboard.");
      } catch (e) {
        // fallback
        const ta = $("#copilotPreview");
        if (ta) {
          ta.focus();
          ta.select();
          document.execCommand("copy");
          showToast("Scenario summary selected. Copy it using Ctrl/Cmd+C.");
        }
      }
    });
  }

  // -----------------------------
  // 13) Save / load project JSON
  // -----------------------------
  function bindSaveLoad() {
    const saveProject = $("#saveProject");
    const loadProject = $("#loadProject");
    const loadFile = $("#loadFile");

    if (saveProject) {
      saveProject.addEventListener("click", () => {
        const payload = JSON.stringify(model, null, 2);
        const blob = new Blob([payload], { type: "application/json;charset=utf-8" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = `${slug(model.meta.toolName)}_${slug(model.project.name)}_${nowISO()}.json`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        showToast("Project saved.");
      });
    }

    if (loadProject && loadFile) {
      loadProject.addEventListener("click", () => loadFile.click());
      loadFile.addEventListener("change", async () => {
        const file = loadFile.files && loadFile.files[0];
        if (!file) return;
        try {
          const text = await file.text();
          const obj = JSON.parse(text);
          // minimal validation then merge (replace)
          if (!obj || !obj.meta || !obj.project || !obj.time) throw new Error("Invalid project file.");
          // replace model in-place
          Object.keys(model).forEach((k) => delete model[k]);
          Object.assign(model, obj);

          // ensure required fallbacks
          model.meta = model.meta || { toolName: "Farming CBA Decision Tool 2" };
          model.meta.toolName = "Farming CBA Decision Tool 2";
          model.raw = model.raw || { faba2022: { rows: [] } };

          ensureSingleControl();
          initDeltasForAllTreatments();

          // rebind + render
          setBrandingToTool2();
          bindProjectFields();
          bindSettingsFields();
          renderOutputs();
          renderTreatments();
          renderDatabase();
          bindSimulationControls();
          recalcAndRender();

          showToast("Project loaded.");
        } catch (e) {
          console.error(e);
          showToast("Failed to load project JSON.");
        }
      });
    }
  }

  // -----------------------------
  // 14) List event delegation for dynamic inputs
  // -----------------------------
  function bindDynamicEditors() {
    // Outputs editor
    const outRoot = $("#outputsList");
    if (outRoot) {
      outRoot.addEventListener("input", (e) => {
        const el = e.target;
        if (!(el instanceof HTMLElement)) return;
        const id = el.getAttribute("data-id");
        const field = el.getAttribute("data-out");
        if (!id || !field) return;
        const o = model.outputs.find((x) => x.id === id);
        if (!o) return;
        if (field === "value") o.value = toNum(el.value, o.value);
        else if (field === "baselineQuantity") o.baselineQuantity = toNum(el.value, o.baselineQuantity);
        else o[field] = el.value;
        recalcAndRender();
      });

      outRoot.addEventListener("click", (e) => {
        const el = e.target.closest("[data-action]");
        if (!el) return;
        const act = el.getAttribute("data-action");
        if (act === "remove-output") {
          const id = el.getAttribute("data-id");
          if (!id) return;
          if (model.outputs.length <= 1) return;
          model.outputs = model.outputs.filter((o) => o.id !== id);
          // remove delta entries
          model.treatments.forEach((t) => {
            if (t.deltas) delete t.deltas[id];
          });
          renderOutputs();
          renderTreatments();
          recalcAndRender();
          showToast("Output removed.");
        }
      });
    }

    // Treatments editor
    const trtRoot = $("#treatmentsList");
    if (trtRoot) {
      trtRoot.addEventListener("input", (e) => {
        const el = e.target;
        if (!(el instanceof HTMLElement)) return;
        const id = el.getAttribute("data-id");
        const field = el.getAttribute("data-trt");
        if (!id || !field) return;
        const t = model.treatments.find((x) => x.id === id);
        if (!t) return;

        if (field === "delta") {
          const oid = el.getAttribute("data-oid");
          if (!oid) return;
          t.deltas = t.deltas || {};
          t.deltas[oid] = toNum(el.value, 0);
        } else if (field === "area_ha") t.area_ha = Math.max(0, toNum(el.value, 0));
        else if (field === "adoptionMultiplier") t.adoptionMultiplier = Math.max(0, toNum(el.value, 1));
        else if (field === "capitalCost_year0") t.capitalCost_year0 = toNum(el.value, 0);
        else if (field === "labourCost_perHa") t.labourCost_perHa = toNum(el.value, 0);
        else if (field === "materialsCost_perHa") t.materialsCost_perHa = toNum(el.value, 0);
        else if (field === "servicesCost_perHa") t.servicesCost_perHa = toNum(el.value, 0);
        else if (field === "otherVarCost_perHa") t.otherVarCost_perHa = toNum(el.value, 0);
        else if (field === "totalCost_perHa") {
          // If user edits total directly, store the difference in otherVar so totals remain consistent
          const newTotal = toNum(el.value, 0);
          const fixed = toNum(t.labourCost_perHa, 0) + toNum(t.materialsCost_perHa, 0) + toNum(t.servicesCost_perHa, 0);
          t.otherVarCost_perHa = Math.max(0, newTotal - fixed);
          t.totalCost_perHa = fixed + t.otherVarCost_perHa;
        } else if (field === "notes") t.notes = el.value;
        else if (field === "name") t.name = el.value;
        else t[field] = el.value;

        recomputeTreatmentTotalCost(t);
        recalcAndRender();
      });

      trtRoot.addEventListener("change", (e) => {
        const el = e.target;
        if (!(el instanceof HTMLElement)) return;
        const id = el.getAttribute("data-id");
        const field = el.getAttribute("data-trt");
        if (field !== "isControl" || !id) return;
        model.treatments.forEach((t) => (t.isControl = t.id === id));
        renderTreatments();
        recalcAndRender();
        showToast("Control updated.");
      });

      trtRoot.addEventListener("click", (e) => {
        const btn = e.target.closest("[data-action]");
        if (!btn) return;
        const act = btn.getAttribute("data-action");

        if (act === "remove-treatment") {
          const id = btn.getAttribute("data-id");
          if (!id) return;
          if (model.treatments.length <= 2) return;
          const wasControl = model.treatments.find((t) => t.id === id)?.isControl;
          model.treatments = model.treatments.filter((t) => t.id !== id);
          if (wasControl) ensureSingleControl();
          renderTreatments();
          recalcAndRender();
          showToast("Treatment removed.");
        }

        if (act === "focus-treatment") {
          const id = btn.getAttribute("data-id");
          openTab("treatments");
          // scroll card into view
          setTimeout(() => {
            const card = trtRoot.querySelector(`[data-id="${id}"]`);
            if (card) card.scrollIntoView({ behavior: "smooth", block: "center" });
          }, 100);
        }

        if (act === "jump-results") openTab("results");
      });
    }

    // add output / add treatment buttons
    const addOutput = $("#addOutput");
    if (addOutput) {
      addOutput.addEventListener("click", () => {
        model.outputs.push({ id: uid(), name: "New output", unit: "", value: 0, source: "", baselineQuantity: 0 });
        initDeltasForAllTreatments();
        renderOutputs();
        renderTreatments();
        recalcAndRender();
      });
    }

    const addTreatment = $("#addTreatment");
    if (addTreatment) {
      addTreatment.addEventListener("click", () => {
        const t = {
          id: uid(),
          trtCode: null,
          name: "New treatment",
          area_ha: 100,
          adoptionMultiplier: 1,
          isControl: false,
          capitalCost_year0: 0,
          labourCost_perHa: 0,
          materialsCost_perHa: 0,
          servicesCost_perHa: 0,
          otherVarCost_perHa: 0,
          totalCost_perHa: 0,
          deltas: {},
          notes: "",
          source: "User added"
        };
        model.outputs.forEach((o) => (t.deltas[o.id] = 0));
        model.treatments.push(t);
        renderTreatments();
        recalcAndRender();
        showToast("Treatment added.");
      });
    }

    // recalc button
    const recalc = $("#recalc");
    if (recalc) recalc.addEventListener("click", () => recalcAndRender());
  }

  // -----------------------------
  // 15) Init
  // -----------------------------
  function init() {
    installTooltipSystem();
    setBrandingToTool2();

    // default calibration uses the FULL supplied dataset
    calibrateTreatmentsFromFabaRaw();

    // bind + render
    bindTabs();
    bindProjectFields();
    bindSettingsFields();
    renderOutputs();
    renderTreatments();
    renderDatabase();
    bindDynamicEditors();

    bindExports();
    bindExcelControls();
    bindCopilot();
    bindSaveLoad();
    bindSimulationControls();

    // initial results
    recalcAndRender();

    // keep tool name consistent everywhere
    model.meta.toolName = "Farming CBA Decision Tool 2";
    showToast("Farming CBA Decision Tool 2 loaded.");
  }

  // wait for DOM
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
