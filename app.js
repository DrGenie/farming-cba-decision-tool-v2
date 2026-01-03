// Farming CBA Decision Tool 2 - Newcastle Business School
// Fully upgraded script with working tabs, CBA, Excel-first workflow, AI helper,
// treatment–control matrixsnew, and snapshot-first results.

(() => {
  "use strict";

  // ---------- HELPERS ----------

  function uid() {
    return Math.random().toString(36).slice(2, 10);
  }

  function clamp(v, a, b) {
    return Math.max(a, Math.min(b, v));
  }

  function fmtNumber(n, decimals = 2) {
    if (!isFinite(n)) return "n/a";
    const abs = Math.abs(n);
    if (abs >= 1000) {
      return n.toLocaleString(undefined, { maximumFractionDigits: 0 });
    }
    return n.toLocaleString(undefined, {
      minimumFractionDigits: 0,
      maximumFractionDigits: decimals
    });
  }

  function money(n) {
    if (!isFinite(n)) return "n/a";
    return "$" + fmtNumber(n, 0);
  }

  function ratio(n, decimals = 2) {
    if (!isFinite(n)) return "n/a";
    return fmtNumber(n, decimals);
  }

  function percent(n, decimals = 1) {
    if (!isFinite(n)) return "n/a";
    return fmtNumber(n, decimals) + "%";
  }

  function esc(s) {
    return (s ?? "")
      .toString()
      .replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function annuityFactor(years, rPct) {
    const r = rPct / 100;
    if (!isFinite(r) || r === 0) return years;
    return (1 - Math.pow(1 + r, -years)) / r;
  }

  function showToast(message) {
    const root = document.getElementById("toast-root") || document.body;
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

  // simple seeded RNG for simulation
  function makeRng(seed) {
    let t = (seed || Math.floor(Math.random() * 2 ** 31)) >>> 0;
    return () => {
      t += 0x6d2b79f5;
      let x = t;
      x = Math.imul(x ^ (x >>> 15), 1 | x);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }

  // ---------- CORE MODEL (DEFAULT SCENARIO) ----------

  const thisYear = new Date().getFullYear();

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
        "Applied faba bean trial comparing deep ripping, organic matter, gypsum, liquid fertiliser and silicon treatments against a control.",
      goal:
        "Identify soil amendment packages that deliver higher faba bean yields and acceptable economic returns after accounting for additional costs.",
      withProject:
        "Faba bean growers adopt high-performing amendment packages on similar soils, improving yields and soil function.",
      withoutProject:
        "Growers continue with baseline practice and have limited economic evidence on soil amendments.",
      objectives:
        "Quantify yield and gross margin impacts of candidate soil amendments and rank treatments on benefit–cost and risk.",
      activities:
        "Establish replicated field plots, collect plot-level yield and cost data, and summarise trial-wide economics.",
      stakeholders:
        "Producers, agronomists, farm consultants, government agencies, and research partners.",
      lastUpdated: new Date().toISOString().slice(0, 10)
    },
    time: {
      startYear: thisYear,
      projectStartYear: thisYear,
      years: 10,
      discBase: 7,
      discLow: 4,
      discHigh: 10,
      mirrFinance: 6,
      mirrReinvest: 4,
      discountSchedule: [
        { label: "2025-2034", low: 4, base: 7, high: 10 },
        { label: "2035-2044", low: 4, base: 7, high: 10 },
        { label: "2045-2054", low: 3, base: 6, high: 9 },
        { label: "2055-2064", low: 3, base: 6, high: 9 },
        { label: "2065-2074", low: 2, base: 5, high: 8 }
      ]
    },
    adoption: {
      low: 0.6,
      base: 0.9,
      high: 1.0
    },
    risk: {
      low: 0.05,
      base: 0.15,
      high: 0.3,
      tech: 0.05,
      nonCoop: 0.04,
      socio: 0.02,
      fin: 0.03,
      man: 0.02
    },
    outputs: [
      { id: uid(), name: "Grain yield", unit: "t/ha", unitValue: 450, source: "Market price" },
      { id: uid(), name: "Protein uplift", unit: "percentage point", unitValue: 10, source: "Quality premium" }
    ],
    treatments: [
      {
        id: uid(),
        name: "Control (no amendment)",
        areaHa: 100,
        isControl: true,
        capitalCost: 0,
        labourCostPerHa: 40,
        materialsCostPerHa: 0,
        servicesCostPerHa: 0,
        otherCostPerHa: 0,
        totalCostPerHa: 40,
        annualBenefitPerHa: 1100,
        notes: "Baseline faba bean practice without deep soil amendment."
      },
      {
        id: uid(),
        name: "Deep organic matter (CP1)",
        areaHa: 100,
        isControl: false,
        capitalCost: 125000,
        labourCostPerHa: 55,
        materialsCostPerHa: 16500,
        servicesCostPerHa: 0,
        otherCostPerHa: 0,
        totalCostPerHa: 16555,
        annualBenefitPerHa: 1850,
        notes: "Deep organic matter incorporation at 15 t/ha."
      },
      {
        id: uid(),
        name: "Deep OM + liquid gypsum (CP1 + CHT)",
        areaHa: 100,
        isControl: false,
        capitalCost: 125000,
        labourCostPerHa: 56,
        materialsCostPerHa: 16850,
        servicesCostPerHa: 0,
        otherCostPerHa: 0,
        totalCostPerHa: 16906,
        annualBenefitPerHa: 1900,
        notes: "Combination of deep OM and liquid gypsum."
      },
      {
        id: uid(),
        name: "Surface silicon",
        areaHa: 100,
        isControl: false,
        capitalCost: 0,
        labourCostPerHa: 45,
        materialsCostPerHa: 1000,
        servicesCostPerHa: 0,
        otherCostPerHa: 0,
        totalCostPerHa: 1045,
        annualBenefitPerHa: 1300,
        notes: "Surface applied silicon at 2 t/ha."
      }
    ],
    benefits: [
      {
        id: uid(),
        label: "Reduced risk of downgrades",
        category: "Risk reduction",
        frequency: "Annual",
        annualAmount: 15000,
        startYear: thisYear,
        endYear: thisYear + 9,
        linkAdoption: true,
        linkRisk: true,
        notes: "Applies to the whole project, not individual plots."
      }
    ],
    otherCosts: [
      {
        id: uid(),
        label: "Project management and M&E",
        category: "Capital",
        type: "annual",
        annual: 20000,
        startYear: thisYear,
        endYear: thisYear + 4,
        capital: 50000,
        depMethod: "straight",
        depLife: 5,
        depRate: 20
      }
    ],
    sim: {
      n: 1000,
      targetBCR: 2,
      variationPct: 20,
      bcrMode: "all",
      varyOutputs: true,
      varyTreatCosts: true,
      varyInputCosts: false,
      seed: null
    }
  };

  let excelParsed = null;

  // ---------- FIELD BINDINGS ----------

  const fieldMap = {
    projectName: ["project", "name"],
    projectLead: ["project", "lead"],
    analystNames: ["project", "analysts"],
    projectTeam: ["project", "team"],
    organisation: ["project", "organisation"],
    contactEmail: ["project", "contactEmail"],
    contactPhone: ["project", "contactPhone"],
    projectSummary: ["project", "summary"],
    projectGoal: ["project", "goal"],
    withProject: ["project", "withProject"],
    withoutProject: ["project", "withoutProject"],
    projectObjectives: ["project", "objectives"],
    projectActivities: ["project", "activities"],
    stakeholderGroups: ["project", "stakeholders"],
    lastUpdated: ["project", "lastUpdated"],

    startYear: ["time", "startYear"],
    projectStartYear: ["time", "projectStartYear"],
    years: ["time", "years"],
    discBase: ["time", "discBase"],
    discLow: ["time", "discLow"],
    discHigh: ["time", "discHigh"],
    mirrFinance: ["time", "mirrFinance"],
    mirrReinvest: ["time", "mirrReinvest"],
    outputAssumptions: ["outputsMeta", "assumptions"],

    adoptLow: ["adoption", "low"],
    adoptBase: ["adoption", "base"],
    adoptHigh: ["adoption", "high"],

    riskLow: ["risk", "low"],
    riskBase: ["risk", "base"],
    riskHigh: ["risk", "high"],
    rTech: ["risk", "tech"],
    rNonCoop: ["risk", "nonCoop"],
    rSocio: ["risk", "socio"],
    rFin: ["risk", "fin"],
    rMan: ["risk", "man"],

    simN: ["sim", "n"],
    targetBCR: ["sim", "targetBCR"],
    simVarPct: ["sim", "variationPct"],
    bcrMode: ["sim", "bcrMode"],
    simVaryOutputs: ["sim", "varyOutputs"],
    simVaryTreatCosts: ["sim", "varyTreatCosts"],
    simVaryInputCosts: ["sim", "varyInputCosts"],
    randSeed: ["sim", "seed"]
  };

  // ensure outputsMeta exists
  if (!model.outputsMeta) model.outputsMeta = { systemType: "single", assumptions: "" };

  // ---------- CBA CORE CALCULATION ----------

  function computeTreatmentResults() {
    const years = Number(model.time.years) || 1;
    const r = Number(model.time.discBase) || 0;
    const A = annuityFactor(years, r);

    const adoption = clamp(Number(model.adoption.base) || 1, 0, 1);
    const riskFactor = 1 - clamp(Number(model.risk.base) || 0, 0, 1);

    const treatments = model.treatments.map(t => {
      const areaHa = Number(t.areaHa) || 0;
      const capital = Number(t.capitalCost) || 0;
      const labour = Number(t.labourCostPerHa) || 0;
      const materials = Number(t.materialsCostPerHa) || 0;
      const services = Number(t.servicesCostPerHa) || 0;
      const other = Number(t.otherCostPerHa) || 0;
      const annualBenefitPerHa = Number(t.annualBenefitPerHa) || 0;

      const totalAnnualCostPerHa = labour + materials + services + other;
      t.totalCostPerHa = totalAnnualCostPerHa;

      const pvBenefitsRaw = annualBenefitPerHa * areaHa * A;
      const pvCostsRaw = capital + totalAnnualCostPerHa * areaHa * A;

      // apply adoption and risk to benefits only (standard CBA convention)
      const pvBenefits = pvBenefitsRaw * adoption * riskFactor;
      const pvCosts = pvCostsRaw;

      const npv = pvBenefits - pvCosts;
      const bcr = pvCosts > 0 ? pvBenefits / pvCosts : NaN;
      const roi = pvCosts > 0 ? npv / pvCosts : NaN;

      return {
        id: t.id,
        name: t.name,
        isControl: !!t.isControl,
        areaHa,
        pvBenefits,
        pvCosts,
        npv,
        bcr,
        roi
      };
    });

    // ranking: highest BCR first, but keep control in set
    const sorted = [...treatments].sort((a, b) => {
      if (!isFinite(b.bcr) && !isFinite(a.bcr)) return 0;
      if (!isFinite(b.bcr)) return -1;
      if (!isFinite(a.bcr)) return 1;
      return b.bcr - a.bcr;
    });

    sorted.forEach((res, idx) => {
      res.rank = idx + 1;
    });

    // attach ranks back to the map
    const byId = new Map(sorted.map(r => [r.id, r]));
    const results = treatments.map(t => {
      const rObj = byId.get(t.id);
      return { ...t, rank: rObj ? rObj.rank : null };
    });

    // project-level aggregates: sum over all treatments
    const totalPvB = treatments.reduce((s, x) => s + (x.pvBenefits || 0), 0);
    const totalPvC = treatments.reduce((s, x) => s + (x.pvCosts || 0), 0);
    const totalNpv = totalPvB - totalPvC;
    const totalBcr = totalPvC > 0 ? totalPvB / totalPvC : NaN;
    const totalRoi = totalPvC > 0 ? totalNpv / totalPvC : NaN;

    // simple annual averages for project summary
    const gmAnnual = years > 0 ? totalNpv / years : 0;
    const revenueAnnual = years > 0 ? (totalPvB / years) : 0;
    const gmMargin = revenueAnnual > 0 ? (gmAnnual / revenueAnnual) * 100 : NaN;

    return {
      treatments,
      ranked: sorted,
      withRanks: results,
      project: {
        pvBenefits: totalPvB,
        pvCosts: totalPvC,
        npv: totalNpv,
        bcr: totalBcr,
        roi: totalRoi,
        grossMarginAnnual: gmAnnual,
        marginPct: gmMargin
      }
    };
  }

  function computeTimeProjection(baseResults) {
    const horizons = [3, 5, 7, 10, 15, 20];
    const yearsTotal = Number(model.time.years) || 1;
    const discBase = Number(model.time.discBase) || 0;
    const adoption = clamp(Number(model.adoption.base) || 1, 0, 1);
    const riskFactor = 1 - clamp(Number(model.risk.base) || 0, 0, 1);

    const totalAnnualBenefit =
      yearsTotal > 0 ? (baseResults.project.pvBenefits / adoption / riskFactor) / annuityFactor(yearsTotal, discBase) : 0;
    const totalAnnualCost =
      yearsTotal > 0 ? baseResults.project.pvCosts / annuityFactor(yearsTotal, discBase) : 0;

    const rows = horizons
      .filter(h => h <= yearsTotal)
      .map(h => {
        const A = annuityFactor(h, discBase);
        const pvB = totalAnnualBenefit * A * adoption * riskFactor;
        const pvC = totalAnnualCost * A;
        const npv = pvB - pvC;
        const bcr = pvC > 0 ? pvB / pvC : NaN;
        return { years: h, pvBenefits: pvB, pvCosts: pvC, npv, bcr };
      });

    return rows;
  }

  function computeDepreciationSummary() {
    // very simple straight-line approximation
    return (model.otherCosts || [])
      .filter(c => c.capital && Number(c.capital) > 0)
      .map(c => {
        const cap = Number(c.capital) || 0;
        const life = Number(c.depLife) || 5;
        const annualDep = cap / life;
        return {
          label: c.label || "Capital item",
          capital: cap,
          life,
          annualDep
        };
      });
  }

  // ---------- RENDERING ----------

  function bindSimpleFields() {
    Object.keys(fieldMap).forEach(id => {
      const el = document.getElementById(id);
      if (!el) return;
      const path = fieldMap[id];
      // initial value
      let val = model;
      for (const key of path) {
        if (val && Object.prototype.hasOwnProperty.call(val, key)) {
          val = val[key];
        } else {
          val = "";
          break;
        }
      }
      if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
        if (el.type === "number") {
          el.value = val !== undefined && val !== null ? val : "";
        } else if (el.type === "date") {
          el.value = val || "";
        } else {
          el.value = val != null ? val : "";
        }
      } else if (el.tagName === "SELECT") {
        el.value = val != null ? val : "";
      }

      const handler = () => {
        let target = model;
        for (let i = 0; i < path.length - 1; i++) {
          const key = path[i];
          if (!target[key]) target[key] = {};
          target = target[key];
        }
        const last = path[path.length - 1];
        if (el.tagName === "SELECT") {
          target[last] = el.value;
        } else if (el.type === "number") {
          const v = el.value === "" ? null : Number(el.value);
          target[last] = isNaN(v) ? null : v;
        } else {
          target[last] = el.value;
        }
        if (id === "targetBCR") {
          const lbl = document.getElementById("simBcrTargetLabel");
          if (lbl) lbl.textContent = el.value || "2";
        }
      };

      el.addEventListener("change", handler);
      el.addEventListener("input", e => {
        if (el.type === "number") return;
        handler(e);
      });
    });
  }

  function renderOutputs() {
    const container = document.getElementById("outputsList");
    if (!container) return;
    container.innerHTML = "";
    (model.outputs || []).forEach(out => {
      const row = document.createElement("div");
      row.className = "item-row";
      row.innerHTML = `
        <div class="row-4">
          <div class="field">
            <label>Name</label>
            <input type="text" value="${esc(out.name)}" />
          </div>
          <div class="field">
            <label>Unit</label>
            <input type="text" value="${esc(out.unit || "")}" />
          </div>
          <div class="field">
            <label data-tooltip="Monetary value per unit (for example $/t)">Value per unit ($)</label>
            <input type="number" step="0.01" value="${out.unitValue != null ? out.unitValue : ""}" />
          </div>
          <div class="field">
            <label>Source / notes</label>
            <input type="text" value="${esc(out.source || "")}" />
          </div>
        </div>
        <button class="btn small ghost danger remove-btn">Remove</button>
      `;
      const [nameInput, unitInput, valueInput, sourceInput] = row.querySelectorAll("input");
      nameInput.addEventListener("input", () => {
        out.name = nameInput.value;
        renderDbTab();
      });
      unitInput.addEventListener("input", () => {
        out.unit = unitInput.value;
        renderDbTab();
      });
      valueInput.addEventListener("change", () => {
        const v = Number(valueInput.value);
        out.unitValue = isNaN(v) ? null : v;
      });
      sourceInput.addEventListener("input", () => {
        out.source = sourceInput.value;
      });
      row.querySelector(".remove-btn").addEventListener("click", () => {
        model.outputs = model.outputs.filter(o => o.id !== out.id);
        renderOutputs();
        renderDbTab();
      });
      container.appendChild(row);
    });

    const addBtn = document.getElementById("addOutput");
    if (addBtn) {
      addBtn.onclick = () => {
        model.outputs.push({
          id: uid(),
          name: "New output",
          unit: "",
          unitValue: 0,
          source: ""
        });
        renderOutputs();
        renderDbTab();
      };
    }
  }

  function renderTreatments() {
    const container = document.getElementById("treatmentsList");
    if (!container) return;
    container.innerHTML = "";

    (model.treatments || []).forEach(t => {
      const row = document.createElement("div");
      row.className = "item-row treatment-row";
      row.innerHTML = `
        <div class="row-4">
          <div class="field">
            <label>Name</label>
            <input type="text" class="t-name" value="${esc(t.name)}" />
          </div>
          <div class="field">
            <label>Area (ha)</label>
            <input type="number" class="t-area" step="1" value="${t.areaHa != null ? t.areaHa : ""}" />
          </div>
          <div class="field">
            <label>Is control?</label>
            <select class="t-control">
              <option value="false"${!t.isControl ? " selected" : ""}>No</option>
              <option value="true"${t.isControl ? " selected" : ""}>Yes</option>
            </select>
          </div>
          <div class="field">
            <label data-tooltip="Once-off capital cost at year 0 for this treatment (for example machinery share or establishment).">
              Capital cost ($, year 0)
            </label>
            <input type="number" class="t-capital" step="1" value="${t.capitalCost != null ? t.capitalCost : ""}" />
          </div>
        </div>
        <div class="row-4">
          <div class="field">
            <label data-tooltip="Per hectare labour cost linked to this treatment (for example extra passes, handling).">
              Labour cost ($/ha)
            </label>
            <input type="number" class="t-labour" step="0.01" value="${t.labourCostPerHa != null ? t.labourCostPerHa : ""}" />
          </div>
          <div class="field">
            <label data-tooltip="Per hectare cost of inputs or materials (for example amendment, fertiliser, chemicals).">
              Materials cost ($/ha)
            </label>
            <input type="number" class="t-materials" step="0.01" value="${t.materialsCostPerHa != null ? t.materialsCostPerHa : ""}" />
          </div>
          <div class="field">
            <label data-tooltip="Per hectare cost of external services (for example contractors).">
              Services cost ($/ha)
            </label>
            <input type="number" class="t-services" step="0.01" value="${t.servicesCostPerHa != null ? t.servicesCostPerHa : ""}" />
          </div>
          <div class="field">
            <label data-tooltip="Other variable costs per hectare not captured above.">
              Other cost ($/ha)
            </label>
            <input type="number" class="t-other" step="0.01" value="${t.otherCostPerHa != null ? t.otherCostPerHa : ""}" />
          </div>
        </div>
        <div class="row-4">
          <div class="field">
            <label data-tooltip="Automatically calculated total variable cost per hectare (labour + materials + services + other).">
              Total cost ($/ha)
            </label>
            <div class="metric">
              <div class="value t-total-cost">${money(t.totalCostPerHa || 0)}</div>
              <div class="label small muted">Updates when cost components change</div>
            </div>
          </div>
          <div class="field">
            <label data-tooltip="Expected annual monetary benefit per hectare from this treatment (for example yield × price minus baseline).">
              Annual benefit ($/ha)
            </label>
            <input type="number" class="t-benefit" step="0.01" value="${t.annualBenefitPerHa != null ? t.annualBenefitPerHa : ""}" />
          </div>
          <div class="field">
            <label>Notes</label>
            <input type="text" class="t-notes" value="${esc(t.notes || "")}" />
          </div>
          <div class="field">
            <button class="btn small ghost danger remove-btn">Remove</button>
          </div>
        </div>
      `;

      const nameInput = row.querySelector(".t-name");
      const areaInput = row.querySelector(".t-area");
      const controlSelect = row.querySelector(".t-control");
      const capInput = row.querySelector(".t-capital");
      const labourInput = row.querySelector(".t-labour");
      const matInput = row.querySelector(".t-materials");
      const servInput = row.querySelector(".t-services");
      const otherInput = row.querySelector(".t-other");
      const benefitInput = row.querySelector(".t-benefit");
      const notesInput = row.querySelector(".t-notes");
      const totalSpan = row.querySelector(".t-total-cost");

      function updateTotal() {
        const labour = Number(labourInput.value) || 0;
        const materials = Number(matInput.value) || 0;
        const services = Number(servInput.value) || 0;
        const other = Number(otherInput.value) || 0;
        const total = labour + materials + services + other;
        t.labourCostPerHa = labour;
        t.materialsCostPerHa = materials;
        t.servicesCostPerHa = services;
        t.otherCostPerHa = other;
        t.totalCostPerHa = total;
        totalSpan.textContent = money(total);
      }

      nameInput.addEventListener("input", () => {
        t.name = nameInput.value;
        renderDbTab();
      });
      areaInput.addEventListener("change", () => {
        const v = Number(areaInput.value);
        t.areaHa = isNaN(v) ? 0 : v;
      });
      controlSelect.addEventListener("change", () => {
        const isControl = controlSelect.value === "true";
        if (isControl) {
          model.treatments.forEach(tt => {
            if (tt.id !== t.id) tt.isControl = false;
          });
        }
        t.isControl = isControl;
        renderTreatments();
      });
      capInput.addEventListener("change", () => {
        const v = Number(capInput.value);
        t.capitalCost = isNaN(v) ? 0 : v;
      });
      [labourInput, matInput, servInput, otherInput].forEach(inp => {
        inp.addEventListener("change", updateTotal);
        inp.addEventListener("input", updateTotal);
      });
      benefitInput.addEventListener("change", () => {
        const v = Number(benefitInput.value);
        t.annualBenefitPerHa = isNaN(v) ? 0 : v;
      });
      notesInput.addEventListener("input", () => {
        t.notes = notesInput.value;
        renderDbTab();
      });

      row.querySelector(".remove-btn").addEventListener("click", () => {
        if (model.treatments.length <= 1) {
          showToast("At least one treatment is required.");
          return;
        }
        model.treatments = model.treatments.filter(tt => tt.id !== t.id);
        renderTreatments();
        renderDbTab();
        updateAllResults();
      });

      container.appendChild(row);
    });

    const addBtn = document.getElementById("addTreatment");
    if (addBtn) {
      addBtn.onclick = () => {
        model.treatments.push({
          id: uid(),
          name: "New treatment",
          areaHa: 100,
          isControl: false,
          capitalCost: 0,
          labourCostPerHa: 0,
          materialsCostPerHa: 0,
          servicesCostPerHa: 0,
          otherCostPerHa: 0,
          totalCostPerHa: 0,
          annualBenefitPerHa: 0,
          notes: ""
        });
        renderTreatments();
        renderDbTab();
      };
    }
  }

  function renderBenefits() {
    const container = document.getElementById("benefitsList");
    if (!container) return;
    container.innerHTML = "";
    (model.benefits || []).forEach(b => {
      const row = document.createElement("div");
      row.className = "item-row";
      row.innerHTML = `
        <div class="row-4">
          <div class="field">
            <label>Label</label>
            <input type="text" class="b-label" value="${esc(b.label || "")}" />
          </div>
          <div class="field">
            <label>Annual amount ($)</label>
            <input type="number" class="b-amount" step="0.01" value="${b.annualAmount != null ? b.annualAmount : ""}" />
          </div>
          <div class="field">
            <label>Start year</label>
            <input type="number" class="b-start" value="${b.startYear != null ? b.startYear : ""}" />
          </div>
          <div class="field">
            <label>End year</label>
            <input type="number" class="b-end" value="${b.endYear != null ? b.endYear : ""}" />
          </div>
        </div>
        <div class="row-3">
          <div class="field">
            <label>Category</label>
            <input type="text" class="b-cat" value="${esc(b.category || "")}" />
          </div>
          <div class="field">
            <label>Link to adoption?</label>
            <select class="b-adopt">
              <option value="true"${b.linkAdoption ? " selected" : ""}>Yes</option>
              <option value="false"${!b.linkAdoption ? " selected" : ""}>No</option>
            </select>
          </div>
          <div class="field">
            <label>Link to risk?</label>
            <select class="b-risk">
              <option value="true"${b.linkRisk ? " selected" : ""}>Yes</option>
              <option value="false"${!b.linkRisk ? " selected" : ""}>No</option>
            </select>
          </div>
        </div>
        <div class="row-2">
          <div class="field">
            <label>Notes</label>
            <input type="text" class="b-notes" value="${esc(b.notes || "")}" />
          </div>
          <div class="field">
            <button class="btn small ghost danger remove-btn">Remove</button>
          </div>
        </div>
      `;

      row.querySelector(".b-label").addEventListener("input", e => {
        b.label = e.target.value;
      });
      row.querySelector(".b-amount").addEventListener("change", e => {
        const v = Number(e.target.value);
        b.annualAmount = isNaN(v) ? 0 : v;
      });
      row.querySelector(".b-start").addEventListener("change", e => {
        const v = Number(e.target.value);
        b.startYear = isNaN(v) ? null : v;
      });
      row.querySelector(".b-end").addEventListener("change", e => {
        const v = Number(e.target.value);
        b.endYear = isNaN(v) ? null : v;
      });
      row.querySelector(".b-cat").addEventListener("input", e => {
        b.category = e.target.value;
      });
      row.querySelector(".b-adopt").addEventListener("change", e => {
        b.linkAdoption = e.target.value === "true";
      });
      row.querySelector(".b-risk").addEventListener("change", e => {
        b.linkRisk = e.target.value === "true";
      });
      row.querySelector(".b-notes").addEventListener("input", e => {
        b.notes = e.target.value;
      });
      row.querySelector(".remove-btn").addEventListener("click", () => {
        model.benefits = model.benefits.filter(bb => bb.id !== b.id);
        renderBenefits();
      });

      container.appendChild(row);
    });

    const addBtn = document.getElementById("addBenefit");
    if (addBtn) {
      addBtn.onclick = () => {
        model.benefits.push({
          id: uid(),
          label: "New benefit",
          category: "",
          frequency: "Annual",
          annualAmount: 0,
          startYear: thisYear,
          endYear: thisYear + 4,
          linkAdoption: true,
          linkRisk: true,
          notes: ""
        });
        renderBenefits();
      };
    }
  }

  function renderCosts() {
    const container = document.getElementById("costsList");
    if (!container) return;
    container.innerHTML = "";
    (model.otherCosts || []).forEach(c => {
      const row = document.createElement("div");
      row.className = "item-row";
      row.innerHTML = `
        <div class="row-4">
          <div class="field">
            <label>Label</label>
            <input type="text" class="c-label" value="${esc(c.label || "")}" />
          </div>
          <div class="field">
            <label>Annual cost ($)</label>
            <input type="number" class="c-annual" step="0.01" value="${c.annual != null ? c.annual : ""}" />
          </div>
          <div class="field">
            <label>Start year</label>
            <input type="number" class="c-start" value="${c.startYear != null ? c.startYear : ""}" />
          </div>
          <div class="field">
            <label>End year</label>
            <input type="number" class="c-end" value="${c.endYear != null ? c.endYear : ""}" />
          </div>
        </div>
        <div class="row-4">
          <div class="field">
            <label>Capital cost ($)</label>
            <input type="number" class="c-capital" step="1" value="${c.capital != null ? c.capital : ""}" />
          </div>
          <div class="field">
            <label>Depreciation life (years)</label>
            <input type="number" class="c-life" value="${c.depLife != null ? c.depLife : ""}" />
          </div>
          <div class="field">
            <label>Depreciation method</label>
            <select class="c-method">
              <option value="straight"${c.depMethod === "straight" ? " selected" : ""}>Straight-line</option>
              <option value="declining"${c.depMethod === "declining" ? " selected" : ""}>Declining balance</option>
            </select>
          </div>
          <div class="field">
            <button class="btn small ghost danger remove-btn">Remove</button>
          </div>
        </div>
      `;

      row.querySelector(".c-label").addEventListener("input", e => {
        c.label = e.target.value;
      });
      row.querySelector(".c-annual").addEventListener("change", e => {
        const v = Number(e.target.value);
        c.annual = isNaN(v) ? 0 : v;
      });
      row.querySelector(".c-start").addEventListener("change", e => {
        const v = Number(e.target.value);
        c.startYear = isNaN(v) ? null : v;
      });
      row.querySelector(".c-end").addEventListener("change", e => {
        const v = Number(e.target.value);
        c.endYear = isNaN(v) ? null : v;
      });
      row.querySelector(".c-capital").addEventListener("change", e => {
        const v = Number(e.target.value);
        c.capital = isNaN(v) ? 0 : v;
      });
      row.querySelector(".c-life").addEventListener("change", e => {
        const v = Number(e.target.value);
        c.depLife = isNaN(v) ? 5 : v;
      });
      row.querySelector(".c-method").addEventListener("change", e => {
        c.depMethod = e.target.value;
      });
      row.querySelector(".remove-btn").addEventListener("click", () => {
        model.otherCosts = model.otherCosts.filter(cc => cc.id !== c.id);
        renderCosts();
        updateAllResults();
      });

      container.appendChild(row);
    });

    const addBtn = document.getElementById("addCost");
    if (addBtn) {
      addBtn.onclick = () => {
        model.otherCosts.push({
          id: uid(),
          label: "New cost item",
          category: "Capital",
          type: "annual",
          annual: 0,
          startYear: thisYear,
          endYear: thisYear + 4,
          capital: 0,
          depMethod: "straight",
          depLife: 5
        });
        renderCosts();
      };
    }
  }

  function renderDbTab() {
    const dbOut = document.getElementById("dbOutputs");
    const dbTreat = document.getElementById("dbTreatments");
    if (dbOut) {
      dbOut.innerHTML = "";
      (model.outputs || []).forEach(o => {
        const d = document.createElement("div");
        d.className = "db-item";
        d.innerHTML = `<strong>${esc(o.name)}</strong> <span class="small muted">(${esc(
          o.unit || ""
        )}, value ${money(o.unitValue || 0)})</span>`;
        dbOut.appendChild(d);
      });
    }
    if (dbTreat) {
      dbTreat.innerHTML = "";
      (model.treatments || []).forEach(t => {
        const d = document.createElement("div");
        d.className = "db-item";
        d.innerHTML = `<strong>${esc(t.name)}</strong> ${
          t.isControl ? '<span class="badge muted">Control</span>' : ""
        } <span class="small muted">Area ${fmtNumber(t.areaHa || 0)} ha</span>`;
        dbTreat.appendChild(d);
      });
    }
  }

  function renderTreatmentMatrix(calc) {
    const headRow = document.getElementById("treatmentMatrixHead");
    const body = document.getElementById("treatmentMatrixBody");
    if (!headRow || !body) return;

    const treatments = calc.withRanks;

    // control first, then others ordered by rank
    const control = treatments.find(t => t.isControl) || treatments[0];
    const others = treatments.filter(t => !t.isControl).sort((a, b) => (a.rank || 0) - (b.rank || 0));
    const ordered = [control, ...others];

    headRow.innerHTML = "";
    const firstTh = document.createElement("th");
    firstTh.textContent = "Indicator";
    headRow.appendChild(firstTh);
    ordered.forEach(t => {
      const th = document.createElement("th");
      th.innerHTML = esc(t.name) + (t.isControl ? ' <span class="badge muted">Control</span>' : "");
      headRow.appendChild(th);
    });

    const rows = [
      {
        key: "pvBenefits",
        label: "Present value of benefits",
        tooltip: "Total discounted benefits for each treatment scenario.",
        fmt: money
      },
      {
        key: "pvCosts",
        label: "Present value of costs",
        tooltip: "Total discounted costs for each treatment scenario, including capital and operating.",
        fmt: money
      },
      {
        key: "npv",
        label: "Net present value",
        tooltip:
          "Present value of benefits minus costs. Positive values indicate net gain relative to a zero baseline for that treatment.",
        fmt: money
      },
      {
        key: "bcr",
        label: "Benefit–cost ratio",
        tooltip:
          "Present value of benefits divided by present value of costs. Values above 1 mean benefits exceed costs.",
        fmt: v => ratio(v, 2)
      },
      {
        key: "roi",
        label: "Return on investment",
        tooltip:
          "Net present value divided by present value of costs. Indicates net gain per dollar spent, after discounting.",
        fmt: v => ratio(v, 2)
      },
      {
        key: "rank",
        label: "Rank (by BCR)",
        tooltip:
          "Ranking based on benefit–cost ratio using base assumptions. All treatments remain visible, including low-performing options.",
        fmt: v => (v != null ? v : "n/a")
      }
    ];

    body.innerHTML = "";
    rows.forEach(r => {
      const tr = document.createElement("tr");
      const th = document.createElement("th");
      th.innerHTML = `<span class="with-tooltip" data-tooltip="${esc(r.tooltip)}">${esc(r.label)}</span>`;
      tr.appendChild(th);
      ordered.forEach(t => {
        const td = document.createElement("td");
        const val = t[r.key];
        td.textContent = r.fmt(val);
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
  }

  function renderProjectSummary(calc) {
    const proj = calc.project;
    const setText = (id, text) => {
      const el = document.getElementById(id);
      if (el) el.textContent = text;
    };

    setText("pvBenefits", money(proj.pvBenefits));
    setText("pvCosts", money(proj.pvCosts));
    setText("npv", money(proj.npv));
    setText("bcr", ratio(proj.bcr, 2));
    setText("roi", ratio(proj.roi, 2));

    // simple IRR / MIRR and payback approximations
    const years = Number(model.time.years) || 1;
    const A = annuityFactor(years, model.time.discBase || 0);
    const annualNet = years > 0 ? proj.npv / A : 0;

    // approximate IRR: if costs >0 and benefits>0
    let irr = NaN;
    if (proj.pvCosts > 0 && proj.pvBenefits > 0) {
      // simple proportional approximation
      irr = (proj.npv / proj.pvCosts) * (model.time.discBase || 0);
    }

    const mirr = irr; // placeholder, keep simple
    const payback = annualNet > 0 ? proj.pvCosts / annualNet : NaN;

    setText("irr", percent(irr, 1));
    setText("mirr", percent(mirr, 1));
    setText("payback", isFinite(payback) ? fmtNumber(payback, 1) + " years" : "n/a");

    const gm = proj.grossMarginAnnual;
    const marginPct = proj.marginPct;
    setText("grossMargin", money(gm));
    setText("profitMargin", percent(marginPct, 1));
  }

  function renderRanking(calc) {
    const container = document.getElementById("treatmentSummary");
    if (!container) return;
    container.innerHTML = "";

    calc.ranked.forEach(t => {
      const card = document.createElement("div");
      card.className = "card subtle rank-card";
      card.innerHTML = `
        <div class="rank-header">
          <div>
            <span class="badge rank-badge">#${t.rank != null ? t.rank : "–"}</span>
            <strong>${esc(t.name)}</strong> ${t.isControl ? '<span class="badge muted">Control</span>' : ""}
          </div>
          <div class="small muted">
            Area ${fmtNumber(t.areaHa || 0)} ha
          </div>
        </div>
        <div class="row-4 metrics-grid compact">
          <div class="field">
            <label>PV benefits</label>
            <div class="metric"><div class="value">${money(t.pvBenefits)}</div></div>
          </div>
          <div class="field">
            <label>PV costs</label>
            <div class="metric"><div class="value">${money(t.pvCosts)}</div></div>
          </div>
          <div class="field">
            <label>NPV</label>
            <div class="metric"><div class="value">${money(t.npv)}</div></div>
          </div>
          <div class="field">
            <label>BCR</label>
            <div class="metric"><div class="value">${ratio(t.bcr, 2)}</div></div>
          </div>
        </div>
      `;
      container.appendChild(card);
    });
  }

  function renderTimeProjectionTable(rows) {
    const tbody = document.querySelector("#timeProjectionTable tbody");
    if (!tbody) return;
    tbody.innerHTML = "";
    rows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.years}</td>
        <td>${money(r.pvBenefits)}</td>
        <td>${money(r.pvCosts)}</td>
        <td>${money(r.npv)}</td>
        <td>${ratio(r.bcr, 2)}</td>
      `;
      tbody.appendChild(tr);
    });
    // chart left as a simple placeholder; the table carries the key information.
    const canvas = document.getElementById("timeNpvChart");
    if (canvas && canvas.getContext) {
      const ctx = canvas.getContext("2d");
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      if (rows.length === 0) return;
      const xs = rows.map(r => r.years);
      const ys = rows.map(r => r.npv);
      const minY = Math.min(...ys);
      const maxY = Math.max(...ys);
      const pad = 20;
      const w = canvas.width - pad * 2;
      const h = canvas.height - pad * 2;
      ctx.beginPath();
      ctx.moveTo(pad, pad + h);
      ctx.lineTo(pad, pad);
      ctx.lineTo(pad + w, pad);
      ctx.stroke();
      rows.forEach((r, i) => {
        const x = pad + (w * (i / Math.max(1, rows.length - 1)));
        const y =
          pad +
          h -
          (h * (r.npv - minY)) / (maxY - minY || 1);
        ctx.beginPath();
        ctx.arc(x, y, 3, 0, Math.PI * 2);
        ctx.fill();
        if (i > 0) {
          const prev = rows[i - 1];
          const px = pad + (w * ((i - 1) / Math.max(1, rows.length - 1)));
          const py =
            pad +
            h -
            (h * (prev.npv - minY)) / (maxY - minY || 1);
          ctx.beginPath();
          ctx.moveTo(px, py);
          ctx.lineTo(x, y);
          ctx.stroke();
        }
      });
    }
  }

  function renderDepSummary() {
    const container = document.getElementById("depSummary");
    if (!container) return;
    const items = computeDepreciationSummary();
    if (!items.length) {
      container.innerHTML = '<p class="small muted">No capital items recorded in the Other costs tab.</p>';
      return;
    }
    const table = document.createElement("table");
    table.className = "summary-table";
    table.innerHTML = `
      <thead>
        <tr>
          <th>Item</th>
          <th>Capital cost ($)</th>
          <th>Life (years)</th>
          <th>Approx. annual depreciation ($)</th>
        </tr>
      </thead>
      <tbody></tbody>
    `;
    const tbody = table.querySelector("tbody");
    items.forEach(it => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${esc(it.label)}</td>
        <td>${money(it.capital)}</td>
        <td>${fmtNumber(it.life, 0)}</td>
        <td>${money(it.annualDep)}</td>
      `;
      tbody.appendChild(tr);
    });
    container.innerHTML = "";
    container.appendChild(table);
  }

  // ---------- SIMULATION ----------

  function runSimulation() {
    const statusEl = document.getElementById("simStatus");
    const calc = computeTreatmentResults();
    const base = calc.project;

    const n = Number(model.sim.n) || 1000;
    const varPct = Number(model.sim.variationPct) || 0;
    const targetBCR = Number(model.sim.targetBCR) || 2;
    const seed = model.sim.seed != null && model.sim.seed !== "" ? Number(model.sim.seed) : null;

    const rnd = makeRng(seed || undefined);
    const npvs = [];
    const bcrs = [];

    for (let i = 0; i < n; i++) {
      const shock = 1 + (rnd() * 2 - 1) * (varPct / 100);
      const shockCost = 1 + (rnd() * 2 - 1) * (varPct / 100);

      const pvB = base.pvBenefits * shock;
      const pvC = base.pvCosts * shockCost;

      const npv = pvB - pvC;
      const bcr = pvC > 0 ? pvB / pvC : NaN;

      npvs.push(npv);
      bcrs.push(bcr);
    }

    npvs.sort((a, b) => a - b);
    bcrs.sort((a, b) => a - b);

    const quantile = (arr, q) => {
      if (!arr.length) return NaN;
      const pos = (arr.length - 1) * q;
      const baseI = Math.floor(pos);
      const rest = pos - baseI;
      if (arr[baseI + 1] !== undefined) {
        return arr[baseI] + rest * (arr[baseI + 1] - arr[baseI]);
      }
      return arr[baseI];
    };

    const minNpv = npvs[0];
    const maxNpv = npvs[npvs.length - 1];
    const meanNpv = npvs.reduce((s, x) => s + x, 0) / npvs.length;
    const medNpv = quantile(npvs, 0.5);
    const probNpvPos = npvs.filter(x => x > 0).length / npvs.length;

    const minBcr = bcrs[0];
    const maxBcr = bcrs[bcrs.length - 1];
    const meanBcr = bcrs.reduce((s, x) => s + x, 0) / bcrs.length;
    const medBcr = quantile(bcrs, 0.5);
    const probBcr1 = bcrs.filter(x => x > 1).length / bcrs.length;
    const probBcrTarget = bcrs.filter(x => x > targetBCR).length / bcrs.length;

    const setText = (id, text) => {
      const el = document.getElementById(id);
      if (el) el.textContent = text;
    };

    setText("simNpvMin", money(minNpv));
    setText("simNpvMax", money(maxNpv));
    setText("simNpvMean", money(meanNpv));
    setText("simNpvMedian", money(medNpv));
    setText("simNpvProb", percent(probNpvPos * 100, 1));

    setText("simBcrMin", ratio(minBcr, 2));
    setText("simBcrMax", ratio(maxBcr, 2));
    setText("simBcrMean", ratio(meanBcr, 2));
    setText("simBcrMedian", ratio(medBcr, 2));
    setText("simBcrProb1", percent(probBcr1 * 100, 1));
    setText("simBcrProbTarget", percent(probBcrTarget * 100, 1));

    if (statusEl) {
      statusEl.textContent = `Simulation run with ${n} draws at ±${varPct}% variation.`;
    }

    // simple histograms using canvas
    const drawHist = (canvasId, data) => {
      const canvas = document.getElementById(canvasId);
      if (!canvas || !canvas.getContext || !data.length) return;
      const ctx = canvas.getContext("2d");
      ctx.clearRect(0, 0, canvas.width, canvas.height);

      const bins = 20;
      const min = data[0];
      const max = data[data.length - 1];
      const binWidth = (max - min || 1) / bins;
      const counts = new Array(bins).fill(0);
      data.forEach(v => {
        const idx = Math.min(bins - 1, Math.max(0, Math.floor((v - min) / binWidth)));
        counts[idx] += 1;
      });
      const maxCount = Math.max(...counts, 1);

      const pad = 20;
      const w = canvas.width - pad * 2;
      const h = canvas.height - pad * 2;

      counts.forEach((c, i) => {
        const x = pad + (w * i) / bins;
        const barW = w / bins - 2;
        const barH = (c / maxCount) * h;
        ctx.fillRect(x, pad + h - barH, barW, barH);
      });
    };

    drawHist("histNpv", npvs);
    drawHist("histBcr", bcrs);
  }

  // ---------- EXCEL WORKFLOW ----------

  function buildTreatmentsSheetData() {
    const header = [
      "Name",
      "IsControl",
      "AreaHa",
      "CapitalCostYear0",
      "LabourCostPerHa",
      "MaterialsCostPerHa",
      "ServicesCostPerHa",
      "OtherCostPerHa",
      "AnnualBenefitPerHa",
      "Notes"
    ];
    const rows = model.treatments.map(t => [
      t.name,
      t.isControl ? "Yes" : "No",
      t.areaHa || 0,
      t.capitalCost || 0,
      t.labourCostPerHa || 0,
      t.materialsCostPerHa || 0,
      t.servicesCostPerHa || 0,
      t.otherCostPerHa || 0,
      t.annualBenefitPerHa || 0,
      t.notes || ""
    ]);
    return [header, ...rows];
  }

  function downloadTemplate(blank) {
    if (typeof XLSX === "undefined") {
      showToast("XLSX library not loaded. Excel export is not available.");
      return;
    }
    const wb = XLSX.utils.book_new();

    const data = buildTreatmentsSheetData();
    const sheetData = blank ? [data[0]] : data;
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    XLSX.utils.book_append_sheet(wb, ws, "Treatments");

    // simple instruction sheet
    const instr = XLSX.utils.aoa_to_sheet([
      ["Farming CBA Decision Tool 2"],
      [""],
      ["Instructions"],
      [
        "Populate the Treatments sheet with one row per treatment.",
      ],
      [
        "Columns:",
        "Name, IsControl (Yes/No), AreaHa, CapitalCostYear0, LabourCostPerHa, MaterialsCostPerHa, ServicesCostPerHa, OtherCostPerHa, AnnualBenefitPerHa, Notes."
      ],
      [""],
      ["Save the workbook and import it via the Excel tab in the tool."]
    ]);
    XLSX.utils.book_append_sheet(wb, instr, "Instructions");

    const filename = blank ? "farming_cba_tool2_blank_template.xlsx" : "farming_cba_tool2_sample_template.xlsx";
    XLSX.writeFile(wb, filename);
  }

  function downloadSampleWorkbook() {
    downloadTemplate(false);
  }

  function parseExcelFile() {
    if (typeof XLSX === "undefined") {
      showToast("XLSX library not loaded. Excel import is not available.");
      return;
    }
    const fileInput = document.getElementById("excelFile");
    if (!fileInput) return;
    fileInput.value = "";
    fileInput.onchange = e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const data = new Uint8Array(ev.target.result);
          const wb = XLSX.read(data, { type: "array" });
          const sheetName =
            wb.SheetNames.find(n => n.toLowerCase() === "treatments") || wb.SheetNames[0];
          const ws = wb.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json(ws, { defval: null });
          excelParsed = { workbook: wb, treatments: json, sheetName };
          const status = document.getElementById("excelStatus");
          if (status) {
            status.textContent = `Parsed workbook: ${file.name}. Found ${json.length} rows in sheet “${sheetName}”.`;
          }
          showToast("Excel file parsed successfully. Click “Apply parsed Excel data” to update the model.");
        } catch (err) {
          console.error(err);
          showToast("Error parsing Excel file. Please check the template structure.");
        }
      };
      reader.readAsArrayBuffer(file);
    };
    fileInput.click();
  }

  function applyParsedExcel() {
    if (!excelParsed || !excelParsed.treatments) {
      showToast("No parsed Excel data to apply. Please parse a workbook first.");
      return;
    }
    const rows = excelParsed.treatments;
    if (!rows.length) {
      showToast("Parsed treatment sheet is empty.");
      return;
    }

    const requiredCols = [
      "Name",
      "IsControl",
      "AreaHa",
      "CapitalCostYear0",
      "LabourCostPerHa",
      "MaterialsCostPerHa",
      "ServicesCostPerHa",
      "OtherCostPerHa",
      "AnnualBenefitPerHa"
    ];
    const missing = requiredCols.filter(c => !(c in rows[0]));
    if (missing.length) {
      showToast(
        "Missing required columns in Treatments sheet: " + missing.join(", ")
      );
      return;
    }

    model.treatments = rows.map(r => ({
      id: uid(),
      name: r.Name || "Treatment",
      isControl:
        String(r.IsControl || "").toLowerCase().trim().startsWith("y") ||
        String(r.IsControl || "").toLowerCase().trim() === "true" ||
        r.IsControl === 1,
      areaHa: Number(r.AreaHa) || 0,
      capitalCost: Number(r.CapitalCostYear0) || 0,
      labourCostPerHa: Number(r.LabourCostPerHa) || 0,
      materialsCostPerHa: Number(r.MaterialsCostPerHa) || 0,
      servicesCostPerHa: Number(r.ServicesCostPerHa) || 0,
      otherCostPerHa: Number(r.OtherCostPerHa) || 0,
      annualBenefitPerHa: Number(r.AnnualBenefitPerHa) || 0,
      totalCostPerHa:
        (Number(r.LabourCostPerHa) || 0) +
        (Number(r.MaterialsCostPerHa) || 0) +
        (Number(r.ServicesCostPerHa) || 0) +
        (Number(r.OtherCostPerHa) || 0),
      notes: r.Notes || ""
    }));

    if (!model.treatments.some(t => t.isControl) && model.treatments.length) {
      model.treatments[0].isControl = true;
    }

    const status = document.getElementById("excelStatus");
    if (status) {
      status.textContent = `Applied ${model.treatments.length} treatments from Excel sheet “${excelParsed.sheetName}”.`;
    }
    showToast("Excel data applied. Treatments and results have been updated.");
    renderTreatments();
    renderDbTab();
    updateAllResults();
  }

  // ---------- EXPORT CSV / EXCEL / PDF ----------

  function exportCsvAllTreatments() {
    const calc = computeTreatmentResults();
    const headers = [
      "Name",
      "IsControl",
      "AreaHa",
      "PVBenefits",
      "PVCosts",
      "NPV",
      "BCR",
      "ROI",
      "Rank"
    ];
    const rows = calc.withRanks.map(t => [
      `"${t.name.replace(/"/g, '""')}"`,
      t.isControl ? "Yes" : "No",
      t.areaHa || 0,
      t.pvBenefits || 0,
      t.pvCosts || 0,
      t.npv || 0,
      t.bcr || "",
      t.roi || "",
      t.rank || ""
    ]);
    const csv = [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "farming_cba_tool2_treatments.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function exportMatrixToExcel() {
    if (typeof XLSX === "undefined") {
      showToast("XLSX library not loaded. Excel export is not available.");
      return;
    }
    const calc = computeTreatmentResults();
    const treatments = calc.withRanks;
    const control = treatments.find(t => t.isControl) || treatments[0];
    const others = treatments.filter(t => !t.isControl).sort((a, b) => (a.rank || 0) - (b.rank || 0));
    const ordered = [control, ...others];

    const header = ["Indicator", ...ordered.map(t => t.name)];
    const rows = [];

    const pushRow = (label, getter) => {
      rows.push([label, ...ordered.map(getter)]);
    };

    pushRow("PV benefits", t => calc.treatments.find(tt => tt.id === t.id).pvBenefits || 0);
    pushRow("PV costs", t => calc.treatments.find(tt => tt.id === t.id).pvCosts || 0);
    pushRow("NPV", t => calc.treatments.find(tt => tt.id === t.id).npv || 0);
    pushRow("BCR", t => calc.treatments.find(tt => tt.id === t.id).bcr || "");
    pushRow("ROI", t => calc.treatments.find(tt => tt.id === t.id).roi || "");
    pushRow("Rank", t => calc.withRanks.find(tt => tt.id === t.id).rank || "");

    const aoa = [header, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Results matrix");
    XLSX.writeFile(wb, "farming_cba_tool2_results_matrix.xlsx");
  }

  function exportPdf() {
    window.print();
  }

  // ---------- AI HELPER PROMPT ----------

  function buildAiPrompt(calc) {
    const control = calc.withRanks.find(t => t.isControl) || calc.withRanks[0];
    const treatments = calc.withRanks.map(t => ({
      name: t.name,
      isControl: !!t.isControl,
      areaHa: t.areaHa,
      pvBenefits: t.pvBenefits,
      pvCosts: t.pvCosts,
      npv: t.npv,
      bcr: t.bcr,
      roi: t.roi,
      rank: t.rank,
      notes: (model.treatments.find(tt => tt.id === t.id) || {}).notes || ""
    }));

    const payload = {
      tool_name: "Farming CBA Decision Tool 2",
      project_name: model.project.name,
      organisation: model.project.organisation,
      analysis_years: model.time.years,
      discount_rate_base_percent: model.time.discBase,
      adoption_base_multiplier: model.adoption.base,
      risk_base_fraction: model.risk.base,
      summary: model.project.summary,
      goal: model.project.goal,
      with_project_story: model.project.withProject,
      without_project_story: model.project.withoutProject,
      control_treatment_name: control ? control.name : null,
      treatments,
      instructions_for_ai: [
        "You are asked to interpret a farm cost–benefit analysis for a farmer or on-farm manager.",
        "Explain in plain language what the key indicators mean: PV benefits, PV costs, NPV, benefit–cost ratio (BCR), and return on investment (ROI).",
        "Treat this as decision support. Do NOT tell the farmer what to choose and do NOT impose decision rules or thresholds.",
        "Compare each treatment to the control, describing which treatments perform better or worse economically and by how much.",
        "Highlight what drives performance: capital costs, annual costs, and annual benefits (for example yield and price).",
        "When BCR or ROI is low for a treatment, suggest practical, realistic ways the farmer could improve performance, such as reducing input costs, improving yield, adjusting prices, or changing agronomic practices.",
        "Frame all suggestions as guidance and reflection, not rules. Acknowledge uncertainty and farm-specific factors.",
        "If relevant, comment on risk and payback time in intuitive terms.",
        "Produce a two to three page narrative (around 1,200–1,800 words) suitable for a farmer or on-farm manager.",
        "Keep the tone neutral and supportive. Emphasise learning: why some treatments underperform and what would need to change for them to become attractive."
      ]
    };

    return JSON.stringify(payload, null, 2);
  }

  function renderAiPrompt(calc) {
    const preview = document.getElementById("copilotPreview");
    if (!preview) return;
    preview.value = buildAiPrompt(calc);
  }

  function copyAiPrompt() {
    const preview = document.getElementById("copilotPreview");
    if (!preview) return;
    preview.select();
    try {
      document.execCommand("copy");
      showToast("AI prompt copied to clipboard. Paste into Copilot or ChatGPT.");
    } catch (e) {
      showToast("Select the text and copy it manually (Ctrl+C or Cmd+C).");
    }
  }

  function downloadPolicyBrief() {
    const txt = document.getElementById("policyBriefText");
    if (!txt) return;
    const content = txt.value || "";
    if (!content.trim()) {
      showToast("Policy brief text is empty. Paste an AI-generated brief first.");
      return;
    }
    const blob = new Blob([content], { type: "text/plain;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const safeName = (model.project.name || "farming_cba_tool2_policy_brief")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_");
    a.download = safeName + "_brief.txt";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ---------- TABS & NAV ----------

  function setupTabs() {
    const tabs = document.querySelectorAll(".tab-link");
    const panels = document.querySelectorAll(".tab-panel");

    function activateTab(tabId) {
      tabs.forEach(btn => {
        const active = btn.getAttribute("data-tab") === tabId;
        btn.classList.toggle("active", active);
        btn.setAttribute("aria-selected", active ? "true" : "false");
      });
      panels.forEach(p => {
        const active = p.getAttribute("data-tab-panel") === tabId;
        p.classList.toggle("show", active);
        p.classList.toggle("active", active);
        p.setAttribute("aria-hidden", active ? "false" : "true");
      });
    }

    tabs.forEach(btn => {
      btn.addEventListener("click", () => {
        const id = btn.getAttribute("data-tab");
        if (id) activateTab(id);
      });
    });

    document.querySelectorAll("[data-tab-jump]").forEach(btn => {
      btn.addEventListener("click", () => {
        const id = btn.getAttribute("data-tab-jump");
        if (id) activateTab(id);
      });
    });

    const startBtn = document.getElementById("startBtn");
    const startBtnDup = document.getElementById("startBtn-duplicate");
    if (startBtn) startBtn.onclick = () => activateTab("project");
    if (startBtnDup) startBtnDup.onclick = () => activateTab("project");
  }

  // ---------- JSON SAVE / LOAD ----------

  function saveProjectJson() {
    const blob = new Blob([JSON.stringify(model, null, 2)], {
      type: "application/json;charset=utf-8;"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const safeName = (model.project.name || "farming_cba_tool2_project")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_");
    a.href = url;
    a.download = safeName + ".json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function loadProjectJson() {
    const input = document.getElementById("loadFile");
    if (!input) return;
    input.value = "";
    input.onchange = e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        try {
          const obj = JSON.parse(ev.target.result);
          Object.assign(model, obj);
          renderAll();
          updateAllResults();
          showToast("Project JSON loaded.");
        } catch (err) {
          console.error(err);
          showToast("Error loading JSON file.");
        }
      };
      reader.readAsText(file);
    };
    input.click();
  }

  // ---------- RISK COMBINATION ----------

  function calcCombinedRisk() {
    const tech = clamp(Number(model.risk.tech) || 0, 0, 1);
    const nonCoop = clamp(Number(model.risk.nonCoop) || 0, 0, 1);
    const socio = clamp(Number(model.risk.socio) || 0, 0, 1);
    const fin = clamp(Number(model.risk.fin) || 0, 0, 1);
    const man = clamp(Number(model.risk.man) || 0, 0, 1);
    const combined = 1 - (1 - tech) * (1 - nonCoop) * (1 - socio) * (1 - fin) * (1 - man);
    model.risk.base = combined;
    const out = document.getElementById("combinedRiskOut");
    if (out) {
      const valEl = out.querySelector(".value");
      if (valEl) valEl.textContent = percent(combined * 100, 1);
    }
    const riskBaseInput = document.getElementById("riskBase");
    if (riskBaseInput) {
      riskBaseInput.value = combined.toFixed(2);
    }
  }

  // ---------- GLOBAL RENDER / UPDATE ----------

  function renderAll() {
    bindSimpleFields();
    renderOutputs();
    renderTreatments();
    renderBenefits();
    renderCosts();
    renderDbTab();
  }

  function updateAllResults() {
    const calc = computeTreatmentResults();
    renderTreatmentMatrix(calc);
    renderProjectSummary(calc);
    renderRanking(calc);
    renderTimeProjectionTable(computeTimeProjection(calc));
    renderDepSummary();
    renderAiPrompt(calc);
  }

  // ---------- INIT ----------

  document.addEventListener("DOMContentLoaded", () => {
    setupTabs();
    renderAll();
    updateAllResults();

    const recalcBtn = document.getElementById("recalc");
    if (recalcBtn) recalcBtn.onclick = updateAllResults;

    const simBtn = document.getElementById("runSim");
    if (simBtn) simBtn.onclick = runSimulation;

    const saveBtn = document.getElementById("saveProject");
    if (saveBtn) saveBtn.onclick = saveProjectJson;

    const loadBtn = document.getElementById("loadProject");
    if (loadBtn) loadBtn.onclick = loadProjectJson;

    const exportCsvBtn = document.getElementById("exportCsv");
    const exportCsvFoot = document.getElementById("exportCsvFoot");
    if (exportCsvBtn) exportCsvBtn.onclick = exportCsvAllTreatments;
    if (exportCsvFoot) exportCsvFoot.onclick = exportCsvAllTreatments;

    const exportPdfBtn = document.getElementById("exportPdf");
    const exportPdfFoot = document.getElementById("exportPdfFoot");
    if (exportPdfBtn) exportPdfBtn.onclick = exportPdf;
    if (exportPdfFoot) exportPdfFoot.onclick = exportPdf;

    const exportMatrixBtn = document.getElementById("exportExcelMatrix");
    if (exportMatrixBtn) exportMatrixBtn.onclick = exportMatrixToExcel;

    const tmplBtn = document.getElementById("downloadTemplate");
    if (tmplBtn) tmplBtn.onclick = () => downloadTemplate(true);
    const sampleBtn = document.getElementById("downloadSample");
    if (sampleBtn) sampleBtn.onclick = downloadSampleWorkbook;

    const parseExcelBtn = document.getElementById("parseExcel");
    if (parseExcelBtn) parseExcelBtn.onclick = parseExcelFile;
    const importExcelBtn = document.getElementById("importExcel");
    if (importExcelBtn) importExcelBtn.onclick = applyParsedExcel;

    const aiBtn = document.getElementById("openCopilot");
    if (aiBtn) aiBtn.onclick = copyAiPrompt;

    const briefBtn = document.getElementById("downloadBrief");
    if (briefBtn) briefBtn.onclick = downloadPolicyBrief;

    const riskBtn = document.getElementById("calcCombinedRisk");
    if (riskBtn) riskBtn.onclick = calcCombinedRisk;
  });
})();
