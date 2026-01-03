// app.js – Farming CBA Decision Tool 2
// Tabs, data upload (XLSX/CSV), aggregation by Amendment, CBA, additional variables,
// simulations, AI prompt, Excel export, print.

/* ---------- Global state ---------- */

const state = {
  pricePerTonne: 600,      // $ per tonne
  discountRate: 0.07,      // annual rate in decimal
  horizonYears: 10,        // years
  treatments: [],          // [{ id, name, avgYieldPerHa, annualCostPerHa, capitalCostY0, isControl, extraAggregates }]
  additionalFields: [],    // [{ key, label, colIndex, type }]
  results: [],             // computed CBA results sorted by NPV
  simulation: {
    lowPricePct: 80,
    highPricePct: 120,
    highCostPct: 120,
    lowCostPct: 80
  }
};

/* ---------- Utilities ---------- */

function normaliseHeader(str) {
  return String(str || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function cellToNumber(v) {
  if (typeof v === "number") {
    return Number.isFinite(v) ? v : 0;
  }
  if (typeof v === "string") {
    const cleaned = v.replace(/[^0-9.\-]/g, "");
    if (!cleaned) return 0;
    const num = parseFloat(cleaned);
    return Number.isFinite(num) ? num : 0;
  }
  return 0;
}

function roundTo(x, decimals) {
  const n = Number(x);
  if (!Number.isFinite(n)) return 0;
  const p = Math.pow(10, decimals);
  return Math.round(n * p) / p;
}

function pvOfAnnual(annualValue, r, T) {
  // PV of a constant annual stream from t=1..T
  let sum = 0;
  for (let t = 1; t <= T; t++) {
    sum += annualValue / Math.pow(1 + r, t);
  }
  return sum;
}

function formatCurrency(x) {
  const n = Number(x);
  if (!Number.isFinite(n)) return "–";
  const abs = Math.abs(n);
  const decimals = abs >= 1000 ? 0 : 2;
  return (n < 0 ? "-" : "") + "$" + Math.abs(n).toFixed(decimals);
}

function formatNumber(x, decimals) {
  const n = Number(x);
  if (!Number.isFinite(n)) return "–";
  return n.toFixed(decimals);
}

function formatMetricValue(key, value) {
  if (value === null || value === undefined || !Number.isFinite(Number(value))) {
    if (key === "rank") return "–";
    return "–";
  }
  switch (key) {
    case "npv":
    case "pvBenefits":
    case "pvCosts":
    case "annualCost":
      return formatCurrency(value);
    case "bcr":
    case "roi":
      return formatNumber(value, 2);
    case "deltaYield":
      return formatNumber(value, 2);
    case "rank":
      return String(Math.round(value));
    default:
      return String(value);
  }
}

function classifyAdditionalField(normHeader) {
  const costKeywords = [
    "cost", "labour", "labor", "machine", "machinery",
    "tractor", "header", "ute", "truck", "spray", "seeder",
    "ripper", "tiller", "transport", "capital"
  ];
  const benefitKeywords = [
    "yield", "biomass", "plants", "plant", "protein",
    "moisture", "anthesis", "harvest"
  ];

  for (const kw of costKeywords) {
    if (normHeader.includes(kw)) return "cost";
  }
  for (const kw of benefitKeywords) {
    if (normHeader.includes(kw)) return "benefit";
  }
  return "other";
}

/* ---------- Metrics definition (for table & Excel) ---------- */

const METRICS = [
  {
    key: "npv",
    label: "Net present value (NPV)",
    tooltip: "NPV = PV benefits minus PV costs over the chosen time horizon."
  },
  {
    key: "pvBenefits",
    label: "PV of benefits",
    tooltip: "PV of benefits is the discounted value of extra yield revenue compared with the control."
  },
  {
    key: "pvCosts",
    label: "PV of costs",
    tooltip: "PV of costs is the upfront capital cost plus the discounted stream of annual treatment costs."
  },
  {
    key: "bcr",
    label: "Benefit–cost ratio (BCR)",
    tooltip: "BCR = PV benefits divided by PV costs."
  },
  {
    key: "roi",
    label: "Return on investment (ROI)",
    tooltip: "ROI = NPV divided by PV costs (net gain per dollar of PV cost)."
  },
  {
    key: "deltaYield",
    label: "Yield gain vs control (t/ha)",
    tooltip: "Difference in average yield compared with the control treatment."
  },
  {
    key: "annualCost",
    label: "Annual treatment cost ($/ha)",
    tooltip: "Average yearly cost of the treatment inputs per hectare."
  },
  {
    key: "rank",
    label: "Rank (by NPV)",
    tooltip: "1 = highest NPV; larger rank means weaker performance in economic terms."
  }
];

/* ---------- DOM setup ---------- */

document.addEventListener("DOMContentLoaded", () => {
  setupTabs();
  setupInputs();
  setupUpload();
  setupCopyPrompt();
  setupExports();
  setupSimulationInputs();
  updateScenarioSummary();
  computeAndRenderAll();
});

/* ---------- Tabs ---------- */

function setupTabs() {
  const buttons = document.querySelectorAll(".tab-button");
  const panes = document.querySelectorAll(".tab-pane");

  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-tab");
      buttons.forEach((b) => b.classList.toggle("active", b === btn));
      panes.forEach((pane) => {
        pane.classList.toggle("active", pane.id === targetId);
      });
    });
  });
}

/* ---------- Inputs: price, discount, horizon ---------- */

function setupInputs() {
  const priceInput = document.getElementById("price-per-tonne");
  const discountInput = document.getElementById("discount-rate");
  const horizonInput = document.getElementById("time-horizon");

  if (priceInput) {
    if (!priceInput.value) priceInput.value = state.pricePerTonne;
    priceInput.addEventListener("change", () => {
      const v = parseFloat(priceInput.value);
      if (Number.isFinite(v) && v > 0) {
        state.pricePerTonne = v;
      }
      priceInput.value = state.pricePerTonne;
      updateScenarioSummary();
      computeAndRenderAll();
    });
  }

  if (discountInput) {
    if (!discountInput.value) discountInput.value = (state.discountRate * 100).toFixed(1);
    discountInput.addEventListener("change", () => {
      let v = parseFloat(discountInput.value);
      if (!Number.isFinite(v) || v < 0) v = 7;
      state.discountRate = v / 100;
      discountInput.value = v.toFixed(1);
      updateScenarioSummary();
      computeAndRenderAll();
    });
  }

  if (horizonInput) {
    if (!horizonInput.value) horizonInput.value = state.horizonYears;
    horizonInput.addEventListener("change", () => {
      let v = parseInt(horizonInput.value, 10);
      if (!Number.isFinite(v) || v < 1) v = 10;
      state.horizonYears = v;
      horizonInput.value = v;
      updateScenarioSummary();
      computeAndRenderAll();
    });
  }
}

function updateScenarioSummary() {
  const sPrice = document.getElementById("summary-price");
  const sDisc = document.getElementById("summary-discount");
  const sHor = document.getElementById("summary-horizon");

  if (sPrice) sPrice.textContent = state.pricePerTonne;
  if (sDisc) sDisc.textContent = (state.discountRate * 100).toFixed(1);
  if (sHor) sHor.textContent = state.horizonYears;
}

/* ---------- File upload & parsing ---------- */

function setupUpload() {
  const input = document.getElementById("file-input");
  if (!input) return;

  input.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    handleDataFile(file);
  });
}

function updateFileStatus(msg, isError) {
  const el = document.getElementById("file-status");
  if (!el) return;
  el.textContent = msg;
  el.style.color = isError ? "#b91c1c" : "";
}

function handleDataFile(file) {
  const fileName = file.name || "";
  const lower = fileName.toLowerCase();
  const isCSV = lower.endsWith(".csv");

  if (typeof XLSX === "undefined") {
    updateFileStatus("XLSX library not loaded; cannot read file.", true);
    return;
  }

  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      let workbook;
      if (isCSV) {
        const text = event.target.result;
        workbook = XLSX.read(text, { type: "string" });
      } else {
        const data = new Uint8Array(event.target.result);
        workbook = XLSX.read(data, { type: "array" });
      }

      const sheetName = workbook.SheetNames[0];
      const ws = workbook.Sheets[sheetName];
      if (!ws) {
        updateFileStatus("No sheet found in the file.", true);
        return;
      }

      // Get raw matrix: rows as arrays
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!matrix || !matrix.length) {
        updateFileStatus("The sheet is empty.", true);
        return;
      }

      // Find header row that contains all required columns
      const requiredHeaders = [
        "amendment",
        "yield t/ha",
        "treatment input cost only /ha"
      ];

      let headerIndex = -1;
      let headerRow = null;

      outer: for (let i = 0; i < matrix.length; i++) {
        const row = matrix[i];
        if (!row || !row.length) continue;
        const rowNorm = row.map((cell) => normaliseHeader(cell));
        let allFound = true;
        for (const req of requiredHeaders) {
          if (!rowNorm.includes(req)) {
            allFound = false;
            break;
          }
        }
        if (allFound) {
          headerIndex = i;
          headerRow = row;
          break outer;
        }
      }

      if (headerIndex === -1 || !headerRow) {
        updateFileStatus(
          "Could not locate a header row containing Amendment, Yield t/ha, and Treatment Input Cost Only /Ha. Please check the file.",
          true
        );
        return;
      }

      // Build mapping from normalised header to column index
      const headerMap = {};
      headerRow.forEach((cell, idx) => {
        const keyNorm = normaliseHeader(cell);
        if (keyNorm) headerMap[keyNorm] = idx;
      });

      function idxFor(labelNorm) {
        if (Object.prototype.hasOwnProperty.call(headerMap, labelNorm)) {
          return headerMap[labelNorm];
        }
        throw new Error(
          "Column " +
            labelNorm +
            " not found in header after detection; please check the dataset."
        );
      }

      const idxAmendment = idxFor("amendment");
      const idxYield = idxFor("yield t/ha");
      const idxCost = idxFor("treatment input cost only /ha");

      // Additional fields: all other columns with non-empty labels
      const additionalFieldDefs = [];
      headerRow.forEach((cell, idx) => {
        if (idx === idxAmendment || idx === idxYield || idx === idxCost) return;
        const labelRaw = String(cell || "").trim();
        if (!labelRaw) return;
        const norm = normaliseHeader(labelRaw);
        if (!norm) return;
        const type = classifyAdditionalField(norm);
        additionalFieldDefs.push({
          key: labelRaw,   // use original header text as key
          label: labelRaw,
          colIndex: idx,
          type
        });
      });
      state.additionalFields = additionalFieldDefs;

      // Aggregate by amendment
      const groups = {}; // { name: { name, yieldSum, costSum, count, extraSums } }
      let usedRowCount = 0;

      for (let i = headerIndex + 1; i < matrix.length; i++) {
        const row = matrix[i];
        if (!row || row.length === 0) continue;

        const rawName = row[idxAmendment];
        const name = String(rawName || "").trim();
        if (!name) continue;

        const yVal = cellToNumber(row[idxYield]);
        const cVal = cellToNumber(row[idxCost]);

        if (!groups[name]) {
          groups[name] = {
            name,
            yieldSum: 0,
            costSum: 0,
            count: 0,
            extraSums: {} // key -> {sum, count}
          };
        }
        const g = groups[name];
        g.yieldSum += yVal;
        g.costSum += cVal;
        g.count += 1;
        usedRowCount += 1;

        // Additional fields
        additionalFieldDefs.forEach((field) => {
          const val = cellToNumber(row[field.colIndex]);
          const key = field.key;
          if (!g.extraSums[key]) {
            g.extraSums[key] = { sum: 0, count: 0 };
          }
          g.extraSums[key].sum += val;
          g.extraSums[key].count += 1;
        });
      }

      const treatments = Object.values(groups).map((g) => {
        const avgYield = g.count > 0 ? g.yieldSum / g.count : 0;
        const avgCost = g.count > 0 ? g.costSum / g.count : 0;
        const extraAggregates = {};
        if (g.extraSums) {
          additionalFieldDefs.forEach((field) => {
            const stats = g.extraSums[field.key];
            if (stats && stats.count > 0) {
              extraAggregates[field.key] = stats.sum / stats.count;
            }
          });
        }
        return {
          id: g.name,
          name: g.name,
          avgYieldPerHa: roundTo(avgYield, 3),
          annualCostPerHa: roundTo(avgCost, 2),
          capitalCostY0: 0,
          isControl: /control/i.test(g.name),
          extraAggregates
        };
      });

      if (!treatments.length) {
        updateFileStatus(
          "No treatments could be read from the file (no Amendment values found).",
          true
        );
        return;
      }

      // Ensure exactly one control
      if (!treatments.some((t) => t.isControl)) {
        treatments[0].isControl = true;
      } else {
        let controlFound = false;
        treatments.forEach((t) => {
          if (t.isControl && !controlFound) {
            controlFound = true;
          } else {
            t.isControl = false;
          }
        });
        if (!controlFound) treatments[0].isControl = true;
      }

      state.treatments = treatments;
      updateFileStatus(
        fileName +
          " loaded: " +
          usedRowCount +
          " plot rows, " +
          treatments.length +
          " treatments.",
        false
      );

      computeAndRenderAll();
    } catch (err) {
      console.error(err);
      updateFileStatus(
        "Error reading file: " + (err && err.message ? err.message : String(err)),
        true
      );
    }
  };

  if (isCSV) {
    reader.readAsText(file);
  } else {
    reader.readAsArrayBuffer(file);
  }
}

/* ---------- Treatments table rendering ---------- */

function renderTreatmentsConfig() {
  const tbody = document.getElementById("treatments-table-body");
  const noData = document.getElementById("treatments-no-data");

  if (!tbody) return;

  tbody.innerHTML = "";

  if (!state.treatments || state.treatments.length === 0) {
    if (noData) noData.style.display = "block";
    return;
  } else if (noData) {
    noData.style.display = "none";
  }

  state.treatments.forEach((t, index) => {
    const tr = document.createElement("tr");

    // Name + control pill
    const nameTd = document.createElement("td");
    const nameSpan = document.createElement("span");
    nameSpan.textContent = t.name;
    nameTd.appendChild(nameSpan);
    if (t.isControl) {
      const br = document.createElement("br");
      const pill = document.createElement("span");
      pill.className = "pill pill-control";
      pill.textContent = "Control";
      nameTd.appendChild(br);
      nameTd.appendChild(pill);
    }
    tr.appendChild(nameTd);

    // Yield input
    const yieldTd = document.createElement("td");
    const yieldInput = document.createElement("input");
    yieldInput.type = "number";
    yieldInput.step = "0.01";
    yieldInput.value = Number.isFinite(t.avgYieldPerHa)
      ? t.avgYieldPerHa
      : "";
    yieldInput.dataset.index = String(index);
    yieldInput.addEventListener("change", (ev) => {
      const idx = parseInt(ev.target.dataset.index, 10);
      if (!Number.isFinite(idx)) return;
      const v = parseFloat(ev.target.value);
      state.treatments[idx].avgYieldPerHa = Number.isFinite(v) ? v : 0;
      ev.target.value = state.treatments[idx].avgYieldPerHa;
      computeAndRenderAll();
    });
    yieldTd.appendChild(yieldInput);
    tr.appendChild(yieldTd);

    // Capital cost input
    const capTd = document.createElement("td");
    const capInput = document.createElement("input");
    capInput.type = "number";
    capInput.step = "1";
    capInput.value = Number.isFinite(t.capitalCostY0)
      ? t.capitalCostY0
      : 0;
    capInput.dataset.index = String(index);
    capInput.addEventListener("change", (ev) => {
      const idx = parseInt(ev.target.dataset.index, 10);
      if (!Number.isFinite(idx)) return;
      const v = parseFloat(ev.target.value);
      state.treatments[idx].capitalCostY0 = Number.isFinite(v) ? v : 0;
      ev.target.value = state.treatments[idx].capitalCostY0;
      computeAndRenderAll();
    });
    capTd.appendChild(capInput);
    tr.appendChild(capTd);

    // Annual cost input
    const costTd = document.createElement("td");
    const costInput = document.createElement("input");
    costInput.type = "number";
    costInput.step = "1";
    costInput.value = Number.isFinite(t.annualCostPerHa)
      ? t.annualCostPerHa
      : 0;
    costInput.dataset.index = String(index);
    costInput.addEventListener("change", (ev) => {
      const idx = parseInt(ev.target.dataset.index, 10);
      if (!Number.isFinite(idx)) return;
      const v = parseFloat(ev.target.value);
      state.treatments[idx].annualCostPerHa = Number.isFinite(v) ? v : 0;
      ev.target.value = state.treatments[idx].annualCostPerHa;
      computeAndRenderAll();
    });
    costTd.appendChild(costInput);
    tr.appendChild(costTd);

    // Control radio
    const ctrlTd = document.createElement("td");
    const ctrlInput = document.createElement("input");
    ctrlInput.type = "radio";
    ctrlInput.name = "control-treatment";
    ctrlInput.checked = !!t.isControl;
    ctrlInput.dataset.index = String(index);
    ctrlInput.addEventListener("change", (ev) => {
      const idx = parseInt(ev.target.dataset.index, 10);
      if (!Number.isFinite(idx)) return;
      state.treatments.forEach((tt, j) => {
        tt.isControl = j === idx;
      });
      renderTreatmentsConfig(); // refresh control pill
      computeAndRenderAll();
    });
    ctrlTd.appendChild(ctrlInput);
    tr.appendChild(ctrlTd);

    tbody.appendChild(tr);
  });
}

/* ---------- Additional variables rendering ---------- */

function renderAdditionalVariables() {
  const empty = document.getElementById("additional-empty");
  const benefitsWrapper = document.getElementById("additional-benefits-wrapper");
  const costsWrapper = document.getElementById("additional-costs-wrapper");
  const otherWrapper = document.getElementById("additional-other-wrapper");
  const benefitsTable = document.getElementById("additional-benefits-table");
  const costsTable = document.getElementById("additional-costs-table");
  const otherTable = document.getElementById("additional-other-table");

  if (!benefitsTable || !costsTable || !otherTable) return;

  benefitsTable.innerHTML = "";
  costsTable.innerHTML = "";
  otherTable.innerHTML = "";

  const hasTreatments = state.treatments && state.treatments.length > 0;
  const hasFields = state.additionalFields && state.additionalFields.length > 0;

  if (!hasTreatments || !hasFields) {
    if (empty) empty.style.display = "block";
    if (benefitsWrapper) benefitsWrapper.style.display = "none";
    if (costsWrapper) costsWrapper.style.display = "none";
    if (otherWrapper) otherWrapper.style.display = "none";
    return;
  }

  if (empty) empty.style.display = "none";

  const treatments = state.treatments;
  const benefitsFields = state.additionalFields.filter((f) => f.type === "benefit");
  const costFields = state.additionalFields.filter((f) => f.type === "cost");
  const otherFields = state.additionalFields.filter((f) => f.type === "other");

  function buildTable(fields, tableEl, wrapperEl) {
    if (!wrapperEl) return;
    if (!fields.length) {
      wrapperEl.style.display = "none";
      return;
    }
    wrapperEl.style.display = "block";

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");

    const firstTh = document.createElement("th");
    firstTh.textContent = "Variable";
    headRow.appendChild(firstTh);

    treatments.forEach((t) => {
      const th = document.createElement("th");
      th.textContent = t.name;
      if (t.isControl) th.classList.add("col-control");
      headRow.appendChild(th);
    });

    thead.appendChild(headRow);
    tableEl.appendChild(thead);

    const tbody = document.createElement("tbody");

    fields.forEach((field) => {
      const tr = document.createElement("tr");
      const labelTh = document.createElement("th");
      labelTh.scope = "row";
      labelTh.textContent = field.label;
      tr.appendChild(labelTh);

      treatments.forEach((t) => {
        const td = document.createElement("td");
        if (t.isControl) td.classList.add("col-control");
        const val =
          t.extraAggregates && Object.prototype.hasOwnProperty.call(t.extraAggregates, field.key)
            ? t.extraAggregates[field.key]
            : null;
        td.textContent = Number.isFinite(Number(val)) ? formatNumber(val, 2) : "–";
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    tableEl.appendChild(tbody);
  }

  buildTable(benefitsFields, benefitsTable, benefitsWrapper);
  buildTable(costFields, costsTable, costsWrapper);
  buildTable(otherFields, otherTable, otherWrapper);
}

/* ---------- CBA computation & rendering ---------- */

function computeAndRenderAll() {
  const hasTreatments = state.treatments && state.treatments.length > 0;

  // Treatments grid
  renderTreatmentsConfig();
  // Additional variables
  renderAdditionalVariables();

  if (!hasTreatments) {
    renderResultsTable(null);
    renderSimulationsTable(null, null);
    updateAIPrompt(null, null);
    return;
  }

  const treatments = state.treatments;
  const control =
    treatments.find((t) => t.isControl) || treatments[0];

  const baseYield = Number(control.avgYieldPerHa) || 0;
  const baseAnnualCost = 0; // we treat treatment cost as extra vs control

  const r = state.discountRate;
  const T = state.horizonYears;
  const price = state.pricePerTonne;

  const results = treatments.map((t) => {
    const avgYield = Number(t.avgYieldPerHa) || 0;
    const annualCost = Number(t.annualCostPerHa) || 0;
    const capitalCost = Number(t.capitalCostY0) || 0;

    const deltaYield = avgYield - baseYield;
    const annualBenefit = deltaYield * price;
    const pvBenefits = pvOfAnnual(annualBenefit, r, T);
    const pvCosts = capitalCost + pvOfAnnual(annualCost - baseAnnualCost, r, T);
    const npv = pvBenefits - pvCosts;

    let bcr = null;
    let roi = null;
    if (pvCosts > 0) {
      bcr = pvBenefits / pvCosts;
      roi = npv / pvCosts;
    }

    return {
      id: t.id,
      name: t.name,
      isControl: !!t.isControl,
      avgYieldPerHa: avgYield,
      annualCostPerHa: annualCost,
      capitalCostY0: capitalCost,
      deltaYield,
      annualCost,
      pvBenefits,
      pvCosts,
      npv,
      bcr,
      roi
    };
  });

  // Ranking by NPV (highest first)
  const sorted = [...results].sort((a, b) => b.npv - a.npv);
  sorted.forEach((res, idx) => {
    res.rank = idx + 1;
  });

  state.results = sorted;

  renderResultsTable(sorted);
  computeAndRenderSimulations();
  updateAIPrompt(sorted, control);
}

/* ---------- Results table ---------- */

function renderResultsTable(resultsSorted) {
  const table = document.getElementById("results-table");
  const wrapper = document.getElementById("results-table-wrapper");
  const empty = document.getElementById("results-empty");

  if (!table) return;

  table.innerHTML = "";

  if (!resultsSorted || !resultsSorted.length) {
    if (wrapper) wrapper.style.display = "none";
    if (empty) empty.style.display = "block";
    return;
  }

  if (wrapper) wrapper.style.display = "block";
  if (empty) empty.style.display = "none";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  // First column header
  const firstTh = document.createElement("th");
  firstTh.textContent = "Indicator";
  headRow.appendChild(firstTh);

  // Treatment columns
  resultsSorted.forEach((res) => {
    const th = document.createElement("th");
    th.textContent = res.name;
    if (res.isControl) {
      th.classList.add("col-control");
    }
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  METRICS.forEach((metric) => {
    const tr = document.createElement("tr");

    const labelTh = document.createElement("th");
    labelTh.scope = "row";

    const labelDiv = document.createElement("div");
    labelDiv.className = "metric-label";

    const labelSpan = document.createElement("span");
    labelSpan.textContent = metric.label;
    labelDiv.appendChild(labelSpan);

    const tooltipSpan = document.createElement("span");
    tooltipSpan.className = "tooltip";

    const iconSpan = document.createElement("span");
    iconSpan.className = "tooltip-icon";
    iconSpan.textContent = "?";

    const textSpan = document.createElement("span");
    textSpan.className = "tooltip-text";
    textSpan.textContent = metric.tooltip;

    tooltipSpan.appendChild(iconSpan);
    tooltipSpan.appendChild(textSpan);
    labelDiv.appendChild(tooltipSpan);

    labelTh.appendChild(labelDiv);
    tr.appendChild(labelTh);

    resultsSorted.forEach((res) => {
      const td = document.createElement("td");
      if (res.isControl) td.classList.add("col-control");

      const rawVal = res[metric.key];
      td.textContent = formatMetricValue(metric.key, rawVal);

      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
}

/* ---------- Simulations ---------- */

function setupSimulationInputs() {
  const lowPrice = document.getElementById("sim-low-price");
  const highPrice = document.getElementById("sim-high-price");
  const highCost = document.getElementById("sim-high-cost");
  const lowCost = document.getElementById("sim-low-cost");

  if (lowPrice) {
    lowPrice.value = state.simulation.lowPricePct;
    lowPrice.addEventListener("change", () => {
      let v = parseFloat(lowPrice.value);
      if (!Number.isFinite(v) || v < 0) v = 80;
      state.simulation.lowPricePct = v;
      lowPrice.value = v;
      computeAndRenderSimulations();
    });
  }

  if (highPrice) {
    highPrice.value = state.simulation.highPricePct;
    highPrice.addEventListener("change", () => {
      let v = parseFloat(highPrice.value);
      if (!Number.isFinite(v) || v < 0) v = 120;
      state.simulation.highPricePct = v;
      highPrice.value = v;
      computeAndRenderSimulations();
    });
  }

  if (highCost) {
    highCost.value = state.simulation.highCostPct;
    highCost.addEventListener("change", () => {
      let v = parseFloat(highCost.value);
      if (!Number.isFinite(v) || v < 0) v = 120;
      state.simulation.highCostPct = v;
      highCost.value = v;
      computeAndRenderSimulations();
    });
  }

  if (lowCost) {
    lowCost.value = state.simulation.lowCostPct;
    lowCost.addEventListener("change", () => {
      let v = parseFloat(lowCost.value);
      if (!Number.isFinite(v) || v < 0) v = 80;
      state.simulation.lowCostPct = v;
      lowCost.value = v;
      computeAndRenderSimulations();
    });
  }
}

function computeAndRenderSimulations() {
  const table = document.getElementById("sim-results-table");
  const wrapper = document.getElementById("sim-results-wrapper");
  const empty = document.getElementById("sim-empty");

  if (!table) return;

  table.innerHTML = "";

  if (!state.treatments || !state.treatments.length) {
    if (wrapper) wrapper.style.display = "none";
    if (empty) empty.style.display = "block";
    return;
  }

  const treatments = state.treatments;
  const control =
    treatments.find((t) => t.isControl) || treatments[0];
  const baseYield = Number(control.avgYieldPerHa) || 0;

  const r = state.discountRate;
  const T = state.horizonYears;
  const basePrice = state.pricePerTonne;

  const pLow = basePrice * (state.simulation.lowPricePct / 100);
  const pHigh = basePrice * (state.simulation.highPricePct / 100);
  const cHighMult = state.simulation.highCostPct / 100;
  const cLowMult = state.simulation.lowCostPct / 100;

  const scenarios = {
    worst: {
      key: "worst",
      label: "Worst case (low price, high cost)",
      price: pLow,
      costMult: cHighMult
    },
    base: {
      key: "base",
      label: "Base case",
      price: basePrice,
      costMult: 1
    },
    best: {
      key: "best",
      label: "Best case (high price, low cost)",
      price: pHigh,
      costMult: cLowMult
    }
  };

  const rows = treatments.map((t) => {
    const avgYield = Number(t.avgYieldPerHa) || 0;
    const annualCost = Number(t.annualCostPerHa) || 0;
    const capitalCost = Number(t.capitalCostY0) || 0;
    const deltaYield = avgYield - baseYield;
    const baseBenefitMultiplier = deltaYield; // NPV benefits scale with price

    const npvWorst = computeNPVScenario(
      baseBenefitMultiplier,
      scenarios.worst.price,
      annualCost,
      capitalCost,
      r,
      T,
      scenarios.worst.costMult
    );
    const npvBase = computeNPVScenario(
      baseBenefitMultiplier,
      scenarios.base.price,
      annualCost,
      capitalCost,
      r,
      T,
      scenarios.base.costMult
    );
    const npvBest = computeNPVScenario(
      baseBenefitMultiplier,
      scenarios.best.price,
      annualCost,
      capitalCost,
      r,
      T,
      scenarios.best.costMult
    );

    return {
      name: t.name,
      isControl: !!t.isControl,
      npvWorst,
      npvBase,
      npvBest
    };
  });

  renderSimulationsTable(rows, scenarios);
}

function computeNPVScenario(baseBenefitMultiplier, price, annualCost, capitalCost, r, T, costMult) {
  const annualBenefit = baseBenefitMultiplier * price;
  const pvBenefits = pvOfAnnual(annualBenefit, r, T);
  const pvCosts = capitalCost + pvOfAnnual(annualCost * costMult, r, T);
  return pvBenefits - pvCosts;
}

function renderSimulationsTable(rows, scenarios) {
  const table = document.getElementById("sim-results-table");
  const wrapper = document.getElementById("sim-results-wrapper");
  const empty = document.getElementById("sim-empty");

  if (!table) return;

  table.innerHTML = "";

  if (!rows || !rows.length || !scenarios) {
    if (wrapper) wrapper.style.display = "none";
    if (empty) empty.style.display = "block";
    return;
  }

  if (wrapper) wrapper.style.display = "block";
  if (empty) empty.style.display = "none";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  const thName = document.createElement("th");
  thName.textContent = "Treatment";
  headRow.appendChild(thName);

  const thWorst = document.createElement("th");
  thWorst.textContent = "NPV – worst case";
  headRow.appendChild(thWorst);

  const thBase = document.createElement("th");
  thBase.textContent = "NPV – base case";
  headRow.appendChild(thBase);

  const thBest = document.createElement("th");
  thBest.textContent = "NPV – best case";
  headRow.appendChild(thBest);

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  rows.forEach((row) => {
    const tr = document.createElement("tr");

    const tdName = document.createElement("td");
    tdName.textContent = row.name;
    if (row.isControl) tdName.classList.add("col-control");
    tr.appendChild(tdName);

    const tdWorst = document.createElement("td");
    tdWorst.textContent = formatCurrency(row.npvWorst);
    tr.appendChild(tdWorst);

    const tdBase = document.createElement("td");
    tdBase.textContent = formatCurrency(row.npvBase);
    tr.appendChild(tdBase);

    const tdBest = document.createElement("td");
    tdBest.textContent = formatCurrency(row.npvBest);
    tr.appendChild(tdBest);

    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
}

/* ---------- AI helper prompt ---------- */

function setupCopyPrompt() {
  const btn = document.getElementById("btn-copy-prompt");
  const status = document.getElementById("copy-status");
  const textarea = document.getElementById("ai-prompt");

  if (!btn || !textarea) return;

  btn.addEventListener("click", async () => {
    if (!textarea.value) return;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(textarea.value);
      } else {
        textarea.select();
        document.execCommand("copy");
      }
      if (status) {
        status.textContent = "Prompt copied.";
        setTimeout(() => {
          status.textContent = "";
        }, 2000);
      }
    } catch (err) {
      console.error(err);
      if (status) status.textContent = "Could not copy prompt.";
    }
  });
}

function updateAIPrompt(resultsSorted, controlTreatment) {
  const textarea = document.getElementById("ai-prompt");
  if (!textarea) return;

  if (!resultsSorted || !resultsSorted.length) {
    textarea.value =
      "Upload your dataset and configure the scenario to generate an AI interpretation prompt.";
    return;
  }

  const treatmentsExport = resultsSorted.map((r) => {
    const extra = {};
    const t = state.treatments.find((tt) => tt.name === r.name);
    if (t && t.extraAggregates) {
      Object.keys(t.extraAggregates).forEach((k) => {
        extra[k] = t.extraAggregates[k];
      });
    }
    return {
      name: r.name,
      is_control: !!r.isControl,
      avg_yield_t_per_ha: r.avgYieldPerHa,
      annual_cost_per_ha: r.annualCost,
      capital_cost_year0: r.capitalCostY0,
      delta_yield_vs_control: r.deltaYield,
      pv_benefits: r.pvBenefits,
      pv_costs: r.pvCosts,
      npv: r.npv,
      bcr: r.bcr,
      roi: r.roi,
      rank_by_npv: r.rank,
      additional_variables: extra
    };
  });

  const exportObj = {
    tool_name: "Farming CBA Decision Tool 2",
    currency: "AUD",
    price_per_tonne: state.pricePerTonne,
    discount_rate_decimal: state.discountRate,
    discount_rate_percent: state.discountRate * 100,
    time_horizon_years: state.horizonYears,
    control_treatment_name: controlTreatment ? controlTreatment.name : null,
    definitions: {
      npv: "NPV = PV benefits − PV costs, per hectare, relative to the control.",
      pv_benefits:
        "PV benefits = discounted value of extra yield revenue relative to the control.",
      pv_costs:
        "PV costs = discounted value of treatment-related costs, including any capital cost in year 0.",
      bcr: "BCR = PV benefits ÷ PV costs.",
      roi: "ROI = NPV ÷ PV costs.",
      simulation_scenarios:
        "Worst case: low grain price and high treatment costs. Best case: high grain price and low treatment costs."
    },
    simulation_settings: {
      low_price_pct_of_base: state.simulation.lowPricePct,
      high_price_pct_of_base: state.simulation.highPricePct,
      high_cost_pct_of_base: state.simulation.highCostPct,
      low_cost_pct_of_base: state.simulation.lowCostPct
    },
    treatments: treatmentsExport
  };

  const lines = [
    "You are interpreting results from a farm cost–benefit analysis tool called \"Farming CBA Decision Tool 2\".",
    "Use plain language suitable for a farmer or on-farm manager. Avoid jargon. Focus on what drives results and what could be changed.",
    "",
    "Important constraints:",
    "- Do not tell the user which treatment to choose.",
    "- Do not impose decision rules or hard thresholds (for example, do not say \"always choose BCR > 1\").",
    "- Treat this as decision support: explain trade-offs, risks, and practical ways to improve low-performing options.",
    "",
    "Definitions for the indicators:",
    "- NPV = PV benefits − PV costs. Positive NPV indicates net economic gain compared with the control.",
    "- PV benefits = discounted value of extra yield revenue compared with the control.",
    "- PV costs = discounted value of treatment-related costs, including any upfront capital cost.",
    "- BCR = PV benefits ÷ PV costs.",
    "- ROI = NPV ÷ PV costs (net gain per dollar of PV cost).",
    "",
    "TASK 1 – Farmer-facing interpretation:",
    "Write a 2–3 page narrative (about 1,200–1,800 words) that:",
    "1. Summarises which treatments perform better or worse in economic terms and why (linking to yield, costs, additional variables, and the discounting assumptions).",
    "2. Explains what each indicator (NPV, PV benefits, PV costs, BCR, ROI) means in practice for an on-farm decision.",
    "3. For any treatment with weak performance (low or negative NPV, low BCR, or clearly dominated by others), discusses practical ways the farmer might improve it, such as reducing costs, improving yields, or altering agronomic practices. Frame these as possibilities, not instructions.",
    "",
    "TASK 2 – Policy- and investor-facing brief:",
    "Using the same data, prepare a structured policy brief suitable for government or investors that:",
    "1. Opens with a half-page summary of the key economic findings (no recommendations, just what the numbers show).",
    "2. Includes a clear table comparing treatments on NPV, PV benefits, PV costs, BCR, ROI, and yield differences versus the control (in a format that can be pasted into Word or Excel).",
    "3. Describes where the results are sensitive to assumptions about prices, costs and discount rates, drawing on the worst/base/best scenarios.",
    "4. Highlights where additional agronomic indicators (for example biomass, plants per m²) or cost components (for example labour, machinery, transport) help explain differences in performance.",
    "5. Ends with a short section on questions that decision makers should ask before scaling up any treatment (for example robustness to input prices, water availability, or operational capacity).",
    "",
    "SCENARIO DATA (JSON):",
    JSON.stringify(exportObj, null, 2)
  ];

  textarea.value = lines.join("\n");
}

/* ---------- Export & print ---------- */

function setupExports() {
  const btnExcel = document.getElementById("btn-export-excel");
  const btnPrint = document.getElementById("btn-print");

  if (btnExcel) {
    btnExcel.addEventListener("click", exportResultsToExcel);
  }

  if (btnPrint) {
    btnPrint.addEventListener("click", () => {
      window.print();
    });
  }
}

function exportResultsToExcel() {
  if (typeof XLSX === "undefined") {
    alert("XLSX library not available in this page.");
    return;
  }
  if (!state.results || !state.results.length) {
    alert("No results to export. Upload data and compute the CBA first.");
    return;
  }

  const treatments = state.results;

  const aoa = [];

  // Header row
  const headerRow = ["Indicator"].concat(treatments.map((t) => t.name));
  aoa.push(headerRow);

  // Metric rows
  METRICS.forEach((metric) => {
    const row = [metric.label];
    treatments.forEach((t) => {
      const val = t[metric.key];
      row.push(formatMetricValue(metric.key, val));
    });
    aoa.push(row);
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "CBA Results");

  const fileName = "Farming_CBA_Results.xlsx";
  XLSX.writeFile(wb, fileName);
}
