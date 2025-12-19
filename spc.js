// Simple SPC logic + wiring for Run chart and XmR chart + MR chart
// Features:
//  - CSV upload + column selection
//  - Run chart with run rule (>=8 points on one side of median)
//  - XmR chart with mean, UCL, LCL, +/-1σ, +/-2σ
//  - MR chart paired with XmR
//  - Baseline: use first N points for centre line & limits (optional)
//  - Target line (goal value) optional
//  - Summary panel with basic NHS-style interpretation
//  - Download main chart as PNG	
//  - Custom chart title and axis labels

let rawRows = [];
let currentChart = null;   // main I / run chart
let mrChart = null;        // moving range chart
let annotations = [];      // { date: 'YYYY-MM-DD', label: 'text' }
let splits = [];   // indices where a new XmR segment starts (split AFTER index)
let lastXmRAnalysis = null;

const fileInput         = document.getElementById("fileInput");
const columnSelectors   = document.getElementById("columnSelectors");
const dateSelect        = document.getElementById("dateColumn");
const valueSelect       = document.getElementById("valueColumn");
const baselineInput     = document.getElementById("baselinePoints");
const chartTitleInput   = document.getElementById("chartTitle");
const xAxisLabelInput   = document.getElementById("xAxisLabel");
const yAxisLabelInput   = document.getElementById("yAxisLabel");
const targetInput       = document.getElementById("targetValue");
const targetDirectionInput = document.getElementById("targetDirection");
const capabilityDiv     = document.getElementById("capability");
const annotationDateInput  = document.getElementById("annotationDate");
const annotationLabelInput = document.getElementById("annotationLabel");
const addAnnotationBtn     = document.getElementById("addAnnotationButton");
const clearAnnotationsBtn  = document.getElementById("clearAnnotationsButton");
const toggleSidebarButton = document.getElementById("toggleSidebarButton");
const splitPointSelect  = document.getElementById("splitPointSelect");
const addSplitButton    = document.getElementById("addSplitButton");
const clearSplitsButton = document.getElementById("clearSplitsButton");
const showMRCheckbox   = document.getElementById("showMRCheckbox");


const generateButton    = document.getElementById("generateButton");
const errorMessage      = document.getElementById("errorMessage");
const chartCanvas       = document.getElementById("spcChart");
const summaryDiv        = document.getElementById("summary");
const downloadBtn       = document.getElementById("downloadPngButton");
const downloadPdfBtn    = document.getElementById("downloadPdfButton");
const openDataEditorButton   = document.getElementById("openDataEditorButton");
const dataEditorOverlay      = document.getElementById("dataEditorOverlay");
const dataEditorTextarea     = document.getElementById("dataEditorTextarea");
const dataEditorApplyButton  = document.getElementById("dataEditorApplyButton");
const dataEditorCancelButton = document.getElementById("dataEditorCancelButton");
const aiQuestionInput   = document.getElementById("aiQuestionInput");
const aiAskButton       = document.getElementById("aiAskButton");
const spcHelperPanel    = document.getElementById("spcHelperPanel");

const spcHelperIntro    = document.getElementById("spcHelperIntro");
const spcHelperChipsGeneral = document.getElementById("spcHelperChipsGeneral");
const spcHelperChipsChart   = document.getElementById("spcHelperChipsChart");
const spcHelperOutput   = document.getElementById("spcHelperOutput");

const shiftRulePointsInput = document.getElementById("shiftRulePoints");
const trendRulePointsInput = document.getElementById("trendRulePoints");
const flagSpecialCauseOnChartCheckbox = document.getElementById("flagSpecialCauseOnChart");
const lclClampRow = document.getElementById("lclClampRow");
const clampLclAtZeroCheckbox = document.getElementById("clampLclAtZero");


const mrPanel           = document.getElementById("mrPanel");
const mrChartCanvas     = document.getElementById("mrChartCanvas");


function guessColumns(rows) {
  if (!rows || rows.length === 0) return { dateCol: null, valueCol: null, hasDateCandidate: false };

  const sample = rows.slice(0, Math.min(rows.length, 50));
  const cols = Object.keys(sample[0] || {});
  if (cols.length === 0) return { dateCol: null, valueCol: null, hasDateCandidate: false };

  function dateScore(col) {
    let valid = 0, total = 0;
    for (const r of sample) {
      const raw = r[col];
      if (raw === null || raw === undefined || String(raw).trim() === "") continue;
      total++;
      const d = parseDateValue(raw);
      if (isFinite(d.getTime())) valid++;
    }
    return total === 0 ? 0 : valid / total;
  }

  function numericScore(col) {
    let valid = 0, total = 0;
    for (const r of sample) {
      const raw = r[col];
      if (raw === null || raw === undefined || String(raw).trim() === "") continue;
      total++;
      const y = toNumericValue(raw);
      if (isFinite(y)) valid++;
    }
    return total === 0 ? 0 : valid / total;
  }

  const scored = cols.map(c => ({
    col: c,
    d: dateScore(c),
    n: numericScore(c)
  }));

  // Prefer date-like columns that are NOT strongly numeric (prevents "Value" being treated as a date)
  const bestDate = scored
    .filter(s => s.d > 0 && s.n < 0.5)
    .sort((a, b) => b.d - a.d)[0];

  const bestNum = scored.slice().sort((a, b) => b.n - a.n)[0];

  // Use let (we may adjust later)
  let dateCol = bestDate && bestDate.d >= 0.4 ? bestDate.col : null;

  // Pick numeric value column
  let valueCol = bestNum && bestNum.n >= 0.4 ? bestNum.col : null;

  // If the best numeric happens to be the same as dateCol, pick the next best numeric column
  if (dateCol && valueCol === dateCol) {
    const nextBestNum = scored
      .filter(s => s.col !== dateCol)
      .sort((a, b) => b.n - a.n)[0];
    if (nextBestNum && nextBestNum.n >= 0.4) valueCol = nextBestNum.col;
  }

  const hasDateCandidate = !!dateCol;

  return { dateCol, valueCol, hasDateCandidate };
}

function hideMrPanelNow() {
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }
  if (mrPanel) {
    mrPanel.style.display = "none";
  }
}

if (showMRCheckbox) {
  showMRCheckbox.addEventListener("change", () => {
    const chartType = getSelectedChartType ? getSelectedChartType() : "run";

    // If you're not on XmR, MR chart isn't relevant anyway
    if (chartType !== "xmr") {
      hideMrPanelNow();
      return;
    }

    // If you already have a chart, just regenerate to show/hide MR
    if (currentChart) {
      generateButton.click();
    } else {
      hideMrPanelNow();
    }
  });
}


const targetToggleBtn = document.getElementById("targetToggleBtn");
let targetEnabled = true;

function updateTargetToggleBtn() {
  if (!targetToggleBtn) return;
  targetToggleBtn.textContent = targetEnabled ? "Hide target line" : "Show target line";
}

function applyPresentationEditsLive() {
  if (!currentChart) return;

  const title = (chartTitleInput?.value || "").trim();
  const xLabel = (xAxisLabelInput?.value || "").trim();
  const yLabel = (yAxisLabelInput?.value || "").trim();

  // Title
  if (currentChart.options?.plugins?.title) {
    currentChart.options.plugins.title.display = !!title;
    currentChart.options.plugins.title.text = title;
  }

  // Axes
  if (currentChart.options?.scales?.x?.title) {
    currentChart.options.scales.x.title.display = !!xLabel;
    currentChart.options.scales.x.title.text = xLabel;
  }
  if (currentChart.options?.scales?.y?.title) {
    currentChart.options.scales.y.title.display = !!yLabel;
    currentChart.options.scales.y.title.text = yLabel;
  }

  // Update without animation for a crisp “as you type” feel
  currentChart.update("none");
}

function hasValidTargetInput() {
  if (!targetInput) return false;
  const v = targetInput.value.trim();
  if (v === "") return false;
  const num = Number(v);
  return isFinite(num);
}

function updateTargetToggleVisibility() {
  if (!targetToggleBtn) return;

  if (hasValidTargetInput()) {
    targetToggleBtn.style.display = "inline-flex";
  } else {
    // No target defined: hide button and force target OFF
    targetToggleBtn.style.display = "none";
    targetEnabled = false;              // assumes you use the button toggle model
    if (typeof updateTargetToggleBtn === "function") updateTargetToggleBtn();
  }
}

// When user types target value: show/hide button and (optionally) redraw
if (targetInput) {
  targetInput.addEventListener("input", () => {
    updateTargetToggleVisibility();

    // If user clears the target, redraw to remove the line immediately
    if (!hasValidTargetInput() && currentChart) {
      generateButton.click();
    }
  });
}

// Call once on load
updateTargetToggleVisibility();
	

function debounce(fn, ms = 80) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

const applyPresentationEditsLiveDebounced = debounce(applyPresentationEditsLive, 60);

if (chartTitleInput) chartTitleInput.addEventListener("input", applyPresentationEditsLiveDebounced);
if (xAxisLabelInput) xAxisLabelInput.addEventListener("input", applyPresentationEditsLiveDebounced);
if (yAxisLabelInput) yAxisLabelInput.addEventListener("input", applyPresentationEditsLiveDebounced);



function loadRows(rows) {
  if (!rows || rows.length === 0) {
    errorMessage.textContent = "No rows found in the data.";
    return false;
  }

  rawRows = rows;
  const firstRow = rows[0];
  const columns = Object.keys(firstRow);

  if (!columns || columns.length === 0) {
    errorMessage.textContent = "Could not detect any columns in the data.";
    return false;
  }

  dateSelect.innerHTML = "";
  valueSelect.innerHTML = "";

  columns.forEach(col => {
    const opt1 = document.createElement("option");
    opt1.value = col;
    opt1.textContent = col;
    dateSelect.appendChild(opt1);

    const opt2 = document.createElement("option");
    opt2.value = col;
    opt2.textContent = col;
    valueSelect.appendChild(opt2);
  });

// --- NEW: auto-guess defaults ---
const guessed = guessColumns(rows);

// Choose X column (date/sequence)
if (guessed.dateCol) {
  dateSelect.value = guessed.dateCol;
} else {
  // No date column detected: fall back to first column and switch axis to sequence
  dateSelect.selectedIndex = 0;

  // If your UI has axisType radios (date/sequence), ensure it is set to sequence
  if (typeof setAxisType === "function") {
    setAxisType("sequence");
  }

  showError("Tip: No date column detected. I’ll treat the data as a simple sequence (run chart by order).");
}

// Choose Y (value) column
if (guessed.valueCol) {
  valueSelect.value = guessed.valueCol;
} else {
  // Fall back to second column if available, otherwise first
  valueSelect.selectedIndex = Math.min(1, valueSelect.options.length - 1);

  // Only show this hint if we haven't already shown the "no date" hint above
  // (do not clear errors here)
  if (errorMessage && !errorMessage.textContent) {
    showError("Tip: I couldn’t confidently detect a numeric value column. Please check the Value dropdown before generating a chart.");
  }
}

// Do NOT call clearError() here — we want hints to remain visible


  columnSelectors.style.display = "block";
  return true;
}

function showError(msg) {
  if (errorMessage) errorMessage.textContent = msg;
}
function clearError() {
  if (errorMessage) errorMessage.textContent = "";
}


function getTargetValue() {
  if (!targetEnabled) return null;
  if (!targetInput) return null;

  const v = targetInput.value.trim();
  if (v === "") return null;

  const num = Number(v);
  return isFinite(num) ? num : null;
}



if (targetToggleBtn) {
  updateTargetToggleBtn();
  targetToggleBtn.addEventListener("click", () => {
    targetEnabled = !targetEnabled;
    updateTargetToggleBtn();
    if (currentChart) generateButton.click();
  });
}

const debouncedRegen = debounce(() => {
  if (rawRows && rawRows.length) generateButton.click();
}, 250);

if (baselineInput) {
  baselineInput.addEventListener("input", debouncedRegen);
  baselineInput.addEventListener("change", debouncedRegen);
}

if (shiftRulePointsInput) {
  shiftRulePointsInput.addEventListener("input", debouncedRegen);
  shiftRulePointsInput.addEventListener("change", debouncedRegen);
}
if (trendRulePointsInput) {
  trendRulePointsInput.addEventListener("input", debouncedRegen);
  trendRulePointsInput.addEventListener("change", debouncedRegen);
}
if (flagSpecialCauseOnChartCheckbox) {
  flagSpecialCauseOnChartCheckbox.addEventListener("change", () => {
    if (rawRows && rawRows.length) generateButton.click();
  });
}
if (clampLclAtZeroCheckbox) {
  clampLclAtZeroCheckbox.addEventListener("change", () => {
    if (rawRows && rawRows.length) generateButton.click();
  });
}


let dataModelDirty = false;

function markDataModelDirty() {
  dataModelDirty = true;
  showError("Data changed — click Generate chart to refresh the chart and analysis.");
}
function clearDataModelDirty() {
  dataModelDirty = false;
  // don’t clearError() automatically; user may still want to see tips
}


//---- Add annotations button

if (addAnnotationBtn) {
  addAnnotationBtn.addEventListener("click", () => {
    if (!annotationDateInput || !annotationLabelInput) return;

    const dateVal = annotationDateInput.value;
    const labelVal = annotationLabelInput.value.trim();

    if (!dateVal || !labelVal) {
      alert("Please enter both a date and a label for the annotation.");
      return;
    }

    // Dates from <input type="date"> are already 'YYYY-MM-DD'
    annotations.push({ date: dateVal, label: labelVal });

	// Clear just the label field, keep the date selection
	annotationLabelInput.value = "";

    // Re-generate the chart with the new annotation
    generateButton.click();
  });
}

//---- Clear annotations button
if (clearAnnotationsBtn) {
  clearAnnotationsBtn.addEventListener("click", () => {
    annotations = [];

    if (annotationDateInput) annotationDateInput.value = "";
    if (annotationLabelInput) annotationLabelInput.value = "";

    // If a chart already exists, re-generate it to remove the lines
    if (currentChart) {
      generateButton.click();
    }
  });
}

// ---- Toggle sidebar button ----
if (toggleSidebarButton) {
  toggleSidebarButton.addEventListener("click", () => {
    const collapsed = document.body.classList.toggle("sidebar-collapsed");
    toggleSidebarButton.textContent = collapsed ? "Show controls" : "Hide controls";
  });
}

// ---- CSV upload & column selection ----
fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;

  clearError();
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  try {
    const text = await file.text();
    const parsed = parseTabularTextWithHeaderDetection(text);

    if (!parsed.ok) {
      showError("Error parsing CSV: " + parsed.message);
      return;
    }

    // If header detection said "no header", but the first two rows are identical header-like rows
    // (e.g. Date,Value repeated), treat it as header mode and just remove the duplicate header row.
    if (!parsed.hadHeader && parsed.rows2D && parsed.rows2D.length >= 2) {
      const r0 = parsed.rows2D[0];
      const r1 = parsed.rows2D[1];

      const score0 = rowDataLikenessScore(r0);
      const duplicateHeaderRow = rowsEqualNormalized(r0, r1) && score0 <= 0.2;

      if (duplicateHeaderRow) {
        // Parse as headered CSV so fields are created, then strip the duplicate header row
        const results = Papa.parse(text, { header: true, dynamicTyping: true, skipEmptyLines: true });

        if (results.errors && results.errors.length > 0) {
          console.error(results.errors);
          showError("Error parsing CSV: " + results.errors[0].message);
          return;
        }

        let rows = results.data || [];
        const headers = results.meta && results.meta.fields ? results.meta.fields : null;
        rows = stripDuplicateHeaderRow(rows, headers);

        if (!loadRows(rows)) return;

        // Reset annotations/splits because the data changed
        annotations = [];
        if (annotationDateInput) annotationDateInput.value = "";
        if (annotationLabelInput) annotationLabelInput.value = "";
        splits = [];
        if (splitPointSelect) splitPointSelect.innerHTML = "";
        return;
      }
    }

    if (parsed.hadHeader) {
      // Normal case: CSV has headers (already stripped of duplicate header row inside parser)
      if (!loadRows(parsed.rows)) return;

    } else {
      // No headers detected — ask the user
      const ok = confirm(
        "It looks like your CSV does not include column headings.\n\n" +
        "Click OK to treat the first row as DATA (I will create Column1, Column2...).\n" +
        "Click Cancel if the first row IS a header row (then add headings and upload again)."
      );

      if (!ok) {
        showError("Please add a header row (e.g. Date,Value) and upload again.");
        return;
      }

      const data2D = parsed.rows2D;
      const colCount = Math.max(...data2D.map(r => r.length));
      const headers = Array.from({ length: colCount }, (_, i) => `Column${i + 1}`);

      const objRows = data2D.map(r => {
        const o = {};
        headers.forEach((h, i) => (o[h] = r[i]));
        return o;
      });

      if (!loadRows(objRows)) return;
    }

	markDataModelDirty();


    // Reset annotations and splits because the data changed
    annotations = [];
    if (annotationDateInput) annotationDateInput.value = "";
    if (annotationLabelInput) annotationLabelInput.value = "";
    splits = [];
    if (splitPointSelect) splitPointSelect.innerHTML = "";

  } catch (err) {
    console.error(err);
    showError("Unexpected error reading the CSV file.");
  }
});


function resetAll() {
  // --- Clear stored data ---
  rawRows = [];
  annotations = [];
  splits = [];
  lastXmRAnalysis = null;

  // --- Reset file input ---
  if (fileInput) fileInput.value = "";

  // --- Hide column selectors ---
  if (columnSelectors) columnSelectors.style.display = "none";

  // --- Reset dropdowns ---
  if (dateSelect) dateSelect.innerHTML = "";
  if (valueSelect) valueSelect.innerHTML = "";
  if (splitPointSelect) splitPointSelect.innerHTML = "";

  // --- Reset text inputs ---
  if (baselineInput) baselineInput.value = "";
  if (chartTitleInput) chartTitleInput.value = "";
  if (xAxisLabelInput) xAxisLabelInput.value = "";
  if (yAxisLabelInput) yAxisLabelInput.value = "";
  if (targetInput) targetInput.value = "";
  if (annotationDateInput) annotationDateInput.value = "";
  if (annotationLabelInput) annotationLabelInput.value = "";

  // --- Reset target direction dropdown ---
  if (targetDirectionInput) targetDirectionInput.value = "above";

  // --- Clear any error message ---
  if (errorMessage) errorMessage.textContent = "";

  // --- Clear summary & capability output ---
  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  // --- Destroy main chart ---
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }

  // --- Destroy MR chart ---
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }

  // --- Hide MR panel ---
  if (mrPanel) mrPanel.style.display = "none";

    // --- Reset AI helper ---
  if (aiQuestionInput) aiQuestionInput.value = "";
  if (spcHelperOutput) spcHelperOutput.innerHTML = "";
  if (spcHelperPanel) spcHelperPanel.classList.remove("visible"); // keep consistent with toggleHelpSection()
  renderHelperState();

  // --- Reset data editor ---
  if (dataEditorTextarea) dataEditorTextarea.value = "";
  if (dataEditorOverlay) dataEditorOverlay.style.display = "none";

  console.log("All elements reset.");
}


function validateBeforeGenerate() {
  if (!rawRows || rawRows.length === 0) {
    showError("No data loaded yet. Upload a CSV or use the data editor first.");
    return false;
  }

  const dateCol = dateSelect?.value;
  const valueCol = valueSelect?.value;

  if (!dateCol || !valueCol) {
    showError("Please choose both an X-axis column and a value column.");
    return false;
  }

  // Check at least 3 valid numeric points
  let good = 0;
  for (const row of rawRows) {
    const y = toNumericValue(row[valueCol]);
    if (isFinite(y)) good++;
  }

  if (good < 3) {
    showError(
      "I can’t create a chart yet: I need at least 3 numeric values in the selected value column. " +
      "Check the column selection and make sure the values are numbers (e.g. 12.3 not '12,3' or text)."
    );
    return false;
  }

  clearError();
  return true;
}



// ---- Helpers ----

function getSelectedChartType() {
  const radios = document.querySelectorAll("input[name='chartType']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "run";
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function getRuleSettings() {
  const shift = shiftRulePointsInput ? parseInt(shiftRulePointsInput.value, 10) : NaN;
  const trend = trendRulePointsInput ? parseInt(trendRulePointsInput.value, 10) : NaN;

  return {
    shiftLength: Number.isFinite(shift) && shift >= 3 ? shift : 8,
    trendLength: Number.isFinite(trend) && trend >= 3 ? trend : 6
  };
}

function shouldFlagSpecialCauseOnChart() {
  return flagSpecialCauseOnChartCheckbox ? !!flagSpecialCauseOnChartCheckbox.checked : true;
}

function shouldClampLclAtZero() {
  // only allow if UI row is visible
  if (!lclClampRow || lclClampRow.style.display === "none") return false;
  return clampLclAtZeroCheckbox ? !!clampLclAtZeroCheckbox.checked : false;
}

function setLclClampVisibility(shouldShow) {
  if (!lclClampRow) return;
  lclClampRow.style.display = shouldShow ? "block" : "none";

  // if the option disappears, clear it to avoid “sticky” state
  if (!shouldShow && clampLclAtZeroCheckbox) clampLclAtZeroCheckbox.checked = false;
}

function findLongRunRanges(values, centre, runLength) {
  const ranges = [];
  let start = 0;

  while (start < values.length) {
    const v = values[start];
    const side = v > centre ? "above" : v < centre ? "below" : "on";
    if (side === "on") { start++; continue; }

    let end = start + 1;
    while (end < values.length) {
      const v2 = values[end];
      const side2 = v2 > centre ? "above" : v2 < centre ? "below" : "on";
      if (side2 !== side) break;
      end++;
    }

    const len = end - start;
    if (len >= runLength) ranges.push({ start, end: end - 1, side, len });

    start = end;
  }
  return ranges;
}

function flagFromRanges(n, ranges) {
  const flags = new Array(n).fill(false);
  ranges.forEach(r => {
    for (let i = r.start; i <= r.end; i++) flags[i] = true;
  });
  return flags;
}

function findTrendRanges(values, length) {
  const ranges = [];
  if (values.length < length) return ranges;

  let incStart = 0, incLen = 1;
  let decStart = 0, decLen = 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i] > values[i - 1]) {
      incLen++; decLen = 1; decStart = i;
    } else if (values[i] < values[i - 1]) {
      decLen++; incLen = 1; incStart = i;
    } else {
      incLen = 1; decLen = 1; incStart = i; decStart = i;
    }

    if (incLen >= length) {
      const start = i - incLen + 1;
      ranges.push({ start, end: i, direction: "increasing", len: incLen });
      incLen = 1; // avoid overlapping spam; simple approach
      incStart = i;
    }
    if (decLen >= length) {
      const start = i - decLen + 1;
      ranges.push({ start, end: i, direction: "decreasing", len: decLen });
      decLen = 1;
      decStart = i;
    }
  }

  return ranges;
}


function parseTabularTextWithHeaderDetection(text) {
  const preview = Papa.parse(text, {
    header: false,
    dynamicTyping: false,
    skipEmptyLines: true
  });

  if (preview.errors && preview.errors.length) {
    return { ok: false, message: preview.errors[0].message };
  }

  const rows2D = preview.data || [];
  if (rows2D.length < 2) {
    return { ok: false, message: "Please provide at least 2 rows." };
  }

  const r0 = rows2D[0];
  const r1 = rows2D[1];

  // Same scoring functions you already added for the data editor:
  const score0 = rowDataLikenessScore(r0);
  const score1 = rowDataLikenessScore(r1);
  const looksLikeHeader = (score1 - score0) >= 0.35;

  if (looksLikeHeader) {
    const results = Papa.parse(text, { header: true, dynamicTyping: true, skipEmptyLines: true });
    if (results.errors && results.errors.length) {
      return { ok: false, message: results.errors[0].message };
    }

    let rows = results.data || [];
    const headers = results.meta && results.meta.fields ? results.meta.fields : null;
    rows = stripDuplicateHeaderRow(rows, headers);

    return { ok: true, rows, hadHeader: true };
  }

  // No header detected
  return { ok: true, rows2D, hadHeader: false };
}

function computeMAD(values, centre) {
  const absDevs = values.map(v => Math.abs(v - centre));
  return computeMedian(absDevs);
}

/**
 * Astronomical point detection using modified z-score (MAD-based).
 * Common robust rule of thumb: |z| > 3.5
 * Returns { indices: number[], flags: boolean[] }
 */
function findAstronomicalPoints(values, centre, referenceValues = null, threshold = 3.5) {
  const ref = (Array.isArray(referenceValues) && referenceValues.length >= 3) ? referenceValues : values;
  const refMedian = centre;
  const mad = computeMAD(ref, refMedian);

  const flags = new Array(values.length).fill(false);
  const indices = [];

  // If MAD is 0 (flat data), there is no sensible astronomical rule
  if (!mad || mad === 0 || !Number.isFinite(mad)) return { indices, flags, mad: 0 };

  // modified z-score constant
  const c = 0.6745;

  for (let i = 0; i < values.length; i++) {
    const z = (c * (values[i] - refMedian)) / mad;
    if (Math.abs(z) > threshold) {
      flags[i] = true;
      indices.push(i);
    }
  }

  return { indices, flags, mad };
}


function computeMedian(values) {
  const sorted = [...values].sort((a, b) => a - b);
  const n = sorted.length;
  if (n === 0) return NaN;
  if (n % 2 === 1) return sorted[(n - 1) / 2];
  return (sorted[n / 2 - 1] + sorted[n / 2]) / 2;
}

/**
 * Detect runs of >= runLength points on the same side of the centre line.
 */
function detectLongRuns(values, centre, runLength = 8) {
  const flags = new Array(values.length).fill(false);

  let start = 0;
  while (start < values.length) {
    const v = values[start];
    const side = v > centre ? "above" : v < centre ? "below" : "on";

    if (side === "on") {
      start++;
      continue;
    }

    // extend this run while points stay on the same side
    let end = start + 1;
    while (end < values.length) {
      const v2 = values[end];
      const side2 = v2 > centre ? "above" : v2 < centre ? "below" : "on";
      if (side2 !== side) break;
      end++;
    }

    const length = end - start;
    if (length >= runLength) {
      for (let i = start; i < end; i++) {
        flags[i] = true;
      }
    }

    start = end;
  }

  return flags;
}

/**
 * Detect simple trend: >= length points all increasing or all decreasing
 */
function detectTrend(values, length = 6) {
  if (values.length < length) return false;

  let incRun = 1;
  let decRun = 1;

  for (let i = 1; i < values.length; i++) {
    if (values[i] > values[i - 1]) {
      incRun++;
      decRun = 1;
    } else if (values[i] < values[i - 1]) {
      decRun++;
      incRun = 1;
    } else {
      incRun = 1;
      decRun = 1;
    }

    if (incRun >= length || decRun >= length) {
      return true;
    }
  }
  return false;
}

function populateSplitOptions(labels) {
  if (!splitPointSelect) return;

  splitPointSelect.innerHTML = "";

  if (!labels || labels.length <= 1) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "Not enough points to split";
    splitPointSelect.appendChild(opt);
    splitPointSelect.disabled = true;
    if (addSplitButton) addSplitButton.disabled = true;
    return;
  }

  splitPointSelect.disabled = false;
  if (addSplitButton) addSplitButton.disabled = false;

  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select point…";
  splitPointSelect.appendChild(placeholder);

  // You can split after any point except the last one
  for (let i = 0; i < labels.length - 1; i++) {
    const opt = document.createElement("option");
    opt.value = String(i); // index of the point AFTER which we split
    opt.textContent = `After ${labels[i]} (point ${i + 1})`;
    splitPointSelect.appendChild(opt);
  }
}



/**
 * Compute XmR statistics and MR values.
 */
function computeXmR(points, baselineCount, clampLclAtZero = false) {
  const pts = [...points].sort((a, b) => a.x - b.x);

  let baselineCountUsed;
  if (baselineCount && baselineCount >= 2) {
    baselineCountUsed = Math.min(baselineCount, pts.length);
  } else {
    baselineCountUsed = pts.length;
  }

  const baseline = pts.slice(0, baselineCountUsed);

  const mean = baseline.reduce((sum, p) => sum + p.y, 0) / baseline.length;

  const baselineMRs = [];
  for (let i = 1; i < baseline.length; i++) {
    baselineMRs.push(Math.abs(baseline[i].y - baseline[i - 1].y));
  }

  const avgMR = baselineMRs.length
    ? baselineMRs.reduce((sum, v) => sum + v, 0) / baselineMRs.length
    : 0;

  const sigma = avgMR === 0 ? 0 : avgMR / 1.128;

  const ucl = mean + 3 * sigma;
  const rawLcl = mean - 3 * sigma;
  const lcl = (clampLclAtZero && rawLcl < 0) ? 0 : rawLcl;

  const mrValues = [];
  for (let i = 1; i < pts.length; i++) {
    mrValues.push(Math.abs(pts[i].y - pts[i - 1].y));
  }

  const flagged = pts.map(p => ({
    ...p,
    beyondLimits: sigma > 0 && (p.y > ucl || p.y < lcl)
  }));

  return {
    points: flagged,
    mean,
    ucl,
    lcl,
    rawLcl,
    sigma,
    avgMR,
    baselineCountUsed,
    mrValues
  };
}
	

// Get title / axis labels with fallbacks
function getChartLabels(defaultTitle, defaultX, defaultY) {
  const title = chartTitleInput && chartTitleInput.value.trim()
    ? chartTitleInput.value.trim()
    : defaultTitle;

  const xLabel = xAxisLabelInput && xAxisLabelInput.value.trim()
    ? xAxisLabelInput.value.trim()
    : defaultX;

  const yLabel = yAxisLabelInput && yAxisLabelInput.value.trim()
    ? yAxisLabelInput.value.trim()
    : defaultY;

  return { title, xLabel, yLabel };
}

function populateAnnotationDateOptions(labels) {
  if (!annotationDateInput) return;

  // Clear existing options
  annotationDateInput.innerHTML = "";

  // Placeholder option
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select date…";
  annotationDateInput.appendChild(placeholder);

  // Add one option per label (these are your x-axis dates like "2024-06-01")
  labels.forEach((lbl) => {
    const opt = document.createElement("option");
    opt.value = lbl;
    opt.textContent = lbl;
    annotationDateInput.appendChild(opt);
  });

  // Reset selection
  annotationDateInput.value = "";
}


function getAxisType() {
  const radios = document.querySelectorAll("input[name='axisType']");
  for (const r of radios) {
    if (r.checked) return r.value;
  }
  return "date"; // sensible default
}

function setAxisType(type) {
  const radios = document.querySelectorAll("input[name='axisType']");
  for (const r of radios) {
    r.checked = (r.value === type);
  }
}

function buildAnnotationConfig(labels) {
  if (!annotations || annotations.length === 0) {
    return {};
  }

  const cfg = {};
  annotations.forEach((a, idx) => {
    const xVal = a.date; // 'YYYY-MM-DD' from <input type="date">
    if (!labels.includes(xVal)) {
      return; // skip if this date isn't on the x-axis
    }

    cfg["annot" + idx] = {
      type: "line",
      xMin: xVal,
      xMax: xVal,
      borderColor: "#000000",
      borderWidth: 1,
      borderDash: [2, 2],
      label: {
        display: true,
        content: a.label,
        backgroundColor: "rgba(255,255,255,0.9)",
        color: "#000000",
        borderColor: "#000000",
        borderWidth: 0.5,
        font: {
          size: 10,
          weight: "bold"
        },
        position: "end",   // near the top of the line
        yAdjust: -6        // nudge it up a little
        // no rotation – keep it horizontal so it's easy to read
      }
    };
  });

  return cfg;
}

function openDataEditor() {
  if (!dataEditorOverlay || !dataEditorTextarea) return;

  // If we already have data, show it as CSV; otherwise, give a skeleton
  if (rawRows && rawRows.length > 0) {
    try {
      dataEditorTextarea.value = Papa.unparse(rawRows);
    } catch (e) {
      // Fallback to blank if unparse fails for any reason
      dataEditorTextarea.value = "";
    }
  } else {
    dataEditorTextarea.value = "Date,Value\n";
  }

  dataEditorOverlay.style.display = "flex";
}

function closeDataEditor() {
  if (dataEditorOverlay) {
    dataEditorOverlay.style.display = "none";
  }
}

if (openDataEditorButton) {
  openDataEditorButton.addEventListener("click", () => {
    openDataEditor();
  });
}

if (dataEditorCancelButton) {
  dataEditorCancelButton.addEventListener("click", () => {
    closeDataEditor();
  });
}

function rowsEqualNormalized(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b)) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    const sa = String(a[i] ?? "").trim().toLowerCase();
    const sb = String(b[i] ?? "").trim().toLowerCase();
    if (sa !== sb) return false;
  }
  return true;
}


function rowDataLikenessScore(rowArr) {
  // Score = fraction of cells that look like a date OR a number
  if (!Array.isArray(rowArr) || rowArr.length === 0) return 0;

  let total = 0;
  let looksLikeData = 0;

  for (const cell of rowArr) {
    const s = String(cell ?? "").trim();
    if (!s) continue;
    total++;

    // numeric?
    const num = toNumericValue(s);
    const isNum = isFinite(num);

    // date?
    const d = parseDateValue(s);
    const isDate = isFinite(d.getTime());

    if (isNum || isDate) looksLikeData++;
  }

  return total === 0 ? 0 : looksLikeData / total;
}

function stripDuplicateHeaderRow(rows, headers) {
  // If first "data row" repeats the headers (common after accidental double header),
  // remove it.
  if (!rows || rows.length === 0) return rows;
  const first = rows[0];
  if (!first) return rows;

  const keys = headers || Object.keys(first);
  if (!keys || keys.length === 0) return rows;

  let matches = 0;
  let checked = 0;
  for (const k of keys) {
    const v = first[k];
    if (v === null || v === undefined) continue;
    checked++;
    if (String(v).trim().toLowerCase() === String(k).trim().toLowerCase()) matches++;
  }

  // If most columns match their own header text, treat it as a duplicate header row
  if (checked > 0 && matches / checked >= 0.7) {
    return rows.slice(1);
  }
  return rows;
}

if (dataEditorApplyButton) {
  dataEditorApplyButton.addEventListener("click", () => {
    if (!dataEditorTextarea) return;

    const text = dataEditorTextarea.value.trim();
    if (!text) {
      alert("Please paste or type some data first.");
      return;
    }

    try {
      // Preview without headers
      const preview = Papa.parse(text, {
        header: false,
        dynamicTyping: false,
        skipEmptyLines: true
      });

      if (preview.errors && preview.errors.length > 0) {
        console.error(preview.errors);
        showError("Error parsing pasted data: " + preview.errors[0].message);
        return;
      }

      const rows2D = preview.data;
      if (!rows2D || rows2D.length < 2) {
        showError("Please paste at least 2 rows.");
        return;
      }

      const r0 = rows2D[0];
      const r1 = rows2D[1];

      const score0 = rowDataLikenessScore(r0);
      const score1 = rowDataLikenessScore(r1);

      // Header if row0 looks much LESS like data than row1
      // (e.g., "Date,Value" vs "2024-01-01,12")
      const duplicateHeaderRow = rowsEqualNormalized(r0, r1) && score0 <= 0.2;

	const looksLikeHeader =  (score1 - score0) >= 0.35 ||  duplicateHeaderRow;

      if (looksLikeHeader) {
        // Parse with header row
        const results = Papa.parse(text, {
          header: true,
          dynamicTyping: true,
          skipEmptyLines: true
        });

        if (results.errors && results.errors.length > 0) {
          console.error(results.errors);
          showError("Error parsing pasted data: " + results.errors[0].message);
          return;
        }

        let rows = results.data || [];
        // Strip accidental duplicate header row if present
        const headers = results.meta && results.meta.fields ? results.meta.fields : null;
        rows = stripDuplicateHeaderRow(rows, headers);

        if (!loadRows(rows)) return;


      } else {
        // No headers -> ask user
        const ok = confirm(
          "It looks like your pasted data does not include column headings.\n\n" +
          "Click OK to treat the first row as DATA (I will create Column1, Column2...).\n" +
          "Click Cancel if the first row IS a header row (then please add headings and try again)."
        );

        if (!ok) {
          showError("Please include a header row (e.g. Date,Value) then click Apply again.");
          return;
        }

        // Convert 2D array into objects Column1/Column2...
        const data2D = rows2D;
        const colCount = Math.max(...data2D.map(r => r.length));
        const headers = Array.from({ length: colCount }, (_, i) => `Column${i + 1}`);

        const objRows = data2D.map(r => {
          const o = {};
          headers.forEach((h, i) => (o[h] = r[i]));
          return o;
        });

        if (!loadRows(objRows)) return;
      }

      // Reset annotations/splits because the data changed
      annotations = [];
      if (annotationDateInput) annotationDateInput.value = "";
      if (annotationLabelInput) annotationLabelInput.value = "";
      splits = [];
      if (splitPointSelect) splitPointSelect.innerHTML = "";

      closeDataEditor();
      clearError();

    } catch (e) {
      console.error(e);
      showError("Unexpected error parsing pasted data.");
    }
  });
}



// ---- Summary helpers ----

function updateRunSummary(points, median, ruleHits, baselineCountUsed) {
  if (!summaryDiv) return;

  const { shiftLength, trendLength } = getRuleSettings();
  const runRanges = ruleHits?.runRanges || [];
  const trendRanges = ruleHits?.trendRanges || [];
  const astro = ruleHits?.astro || { indices: [] };

  const n = points.length;

  let html = `<h3>Summary (Run chart)</h3>`;
  html += `<ul>`;
  html += `<li>Number of points: <strong>${n}</strong></li>`;
  html += baselineCountUsed && baselineCountUsed < n
    ? `<li>Baseline: first <strong>${baselineCountUsed}</strong> points used to calculate median.</li>`
    : `<li>Baseline: all points used to calculate median.</li>`;
  html += `<li>Median: <strong>${median.toFixed(3)}</strong></li>`;
  html += `</ul>`;

  html += `<h4>Rules identified</h4>`;
  html += `<ul>`;

  // Rule 1: Shift
  if (runRanges.length) {
    const parts = runRanges.map(r => `points ${r.start + 1}–${r.end + 1} (${r.side}, length ${r.len})`);
    html += `<li><strong>Rule 1 — Shift:</strong> YES (≥${shiftLength} on one side). ${parts.join("; ")}.</li>`;
  } else {
    html += `<li><strong>Rule 1 — Shift:</strong> No (≥${shiftLength} on one side).</li>`;
  }

  // Rule 2: Trend
  if (trendRanges.length) {
    const parts = trendRanges.map(r => `points ${r.start + 1}–${r.end + 1} (${r.direction}, length ${r.len})`);
    html += `<li><strong>Rule 2 — Trend:</strong> YES (≥${trendLength} consecutive moves). ${parts.join("; ")}.</li>`;
  } else {
    html += `<li><strong>Rule 2 — Trend:</strong> No (≥${trendLength} consecutive moves).</li>`;
  }

  // Rule 3: Astronomical point
  if (astro.indices && astro.indices.length) {
    const pts = astro.indices.map(i => `point ${i + 1}`).join(", ");
    html += `<li><strong>Rule 3 — Astronomical point:</strong> YES (${pts}).</li>`;
  } else {
    html += `<li><strong>Rule 3 — Astronomical point:</strong> No.</li>`;
  }

  html += `</ul>`;

  const anySignals = runRanges.length || trendRanges.length || (astro.indices && astro.indices.length);
  html += anySignals
    ? `<p><strong>Interpretation:</strong> Special-cause signals are present. See the labelled rules above and consider what changed at those times.</p>`
    : `<p><strong>Interpretation:</strong> No rule breaches detected from shift/trend/astronomical rules. Variation looks consistent with common-cause only (interpret in context).</p>`;

  summaryDiv.innerHTML = html;
}



// ---- Summary helpers ----

// Multi-period XmR summary (handles baseline + splits) — lay-user interpretation + astronomical points
function updateXmRMultiSummary(segments, totalPoints) {
  if (!summaryDiv) return;

  if (!segments || segments.length === 0) {
    summaryDiv.innerHTML = "";
    if (capabilityDiv) capabilityDiv.innerHTML = "";
    return;
  }

  const target = getTargetValue();
  const direction = targetDirectionInput ? targetDirectionInput.value : "above";

  // Use configured thresholds if available (defaults stay 8 and 6)
  const { shiftLength, trendLength } =
    (typeof getRuleSettings === "function")
      ? getRuleSettings()
      : { shiftLength: 8, trendLength: 6 };

  let html = `<h3>Summary (XmR chart)</h3>`;
  html += `<p>Total number of points: <strong>${totalPoints}</strong>. `;
  html += `The chart is divided into <strong>${segments.length}</strong> period${segments.length > 1 ? "s" : ""} `;
  html += `(based on the baseline and any splits).</p>`;

  // For capability badge (last period only)
  let lastPeriodSignals = [];
  let lastPeriodCapability = null;
  let lastPeriodHasCapability = false;

  segments.forEach((seg, idx) => {
    const { startIndex, endIndex, labelStart, labelEnd, result } = seg;
    const { mean, ucl, lcl, sigma, avgMR, baselineCountUsed } = result;

    const points = result.points || [];
    const n = points.length;
    const values = points.map(p => p.y);

    // --- Special-cause detection (simple, lay-focused labels) ---
    // 1) Points beyond limits
    const beyondIdx = [];
    points.forEach((p, i) => {
      if (p.beyondLimits) beyondIdx.push(i);
    });

    // 2) Sustained shift (run on one side of mean)
    let runRanges = [];
    if (typeof findLongRunRanges === "function") {
      runRanges = findLongRunRanges(values, mean, shiftLength) || [];
    } else {
      // fallback: your existing boolean flags
      const runFlags = detectLongRuns(values, mean, shiftLength);
      let any = runFlags.some(Boolean);
      if (any) runRanges = [{ start: 0, end: 0 }]; // placeholder (we won't list ranges in fallback)
    }

    // 3) Trend
    let trendRanges = [];
    if (typeof findTrendRanges === "function") {
      trendRanges = findTrendRanges(values, trendLength) || [];
    } else {
      const hasTrend = detectTrend(values, trendLength);
      if (hasTrend) trendRanges = [{ start: 0, end: 0 }]; // placeholder
    }

    // 4) Astronomical point (robust outlier)
    // Use baseline of this *period* to set the reference for outlier detection where possible.
    let astro = { indices: [], flags: [] };
    if (typeof findAstronomicalPoints === "function") {
      const periodBaselineCount = (baselineCountUsed && baselineCountUsed >= 3) ? baselineCountUsed : n;
      const refValues = values.slice(0, Math.min(periodBaselineCount, values.length));
      astro = findAstronomicalPoints(values, mean, refValues, 3.5) || { indices: [], flags: [] };
    }

    // Build simple signals list
    const signals = [];

    if (beyondIdx.length > 0) {
      signals.push("one or more points are outside the control limits");
    }

    if (runRanges.length > 0) {
      signals.push("a sustained shift (many points on the same side of the mean)");
    }

    if (trendRanges.length > 0) {
      signals.push("a sustained trend (steady increase or decrease)");
    }

    if (astro.indices && astro.indices.length > 0) {
      signals.push("an unusual outlier (an ‘astronomical’ point)");
    }

    // Capability (only if target exists and sigma > 0)
    let capability = null;
    if (target !== null && sigma > 0) {
      capability = computeTargetCapability(mean, sigma, target, direction);
    }

    // Target coverage in this period
    let targetCoverageText = "";
    if (target !== null && n > 0) {
      let hits = 0;
      points.forEach(p => {
        if (direction === "above") {
          if (p.y >= target) hits++;
        } else {
          if (p.y <= target) hits++;
        }
      });
      const prop = hits / n;
      targetCoverageText = `${(prop * 100).toFixed(1)}% of points in this period meet the target (${hits}/${n}).`;
    }

    const periodLabel =
      segments.length === 1
        ? "Single period"
        : idx === 0
          ? "Period 1 (initial segment / baseline)"
          : `Period ${idx + 1}`;

    const rangeText =
      labelStart !== undefined && labelEnd !== undefined
        ? `points ${startIndex + 1}–${endIndex + 1} (${labelStart} to ${labelEnd})`
        : `points ${startIndex + 1}–${endIndex + 1}`;

    html += `<h4>${periodLabel}</h4>`;
    html += `<ul>`;
    html += `<li>Coverage: <strong>${rangeText}</strong> – ${n} point${n !== 1 ? "s" : ""}.</li>`;

    if (baselineCountUsed && baselineCountUsed < n) {
      html += `<li>Baseline for this period: first <strong>${baselineCountUsed}</strong> point${baselineCountUsed !== 1 ? "s" : ""} used to calculate mean and limits.</li>`;
    } else {
      html += `<li>Baseline for this period: all points in this period used to calculate mean and limits.</li>`;
    }

    html += `<li>Mean: <strong>${mean.toFixed(3)}</strong>; control limits: <strong>LCL = ${lcl.toFixed(3)}</strong>, <strong>UCL = ${ucl.toFixed(3)}</strong>.</li>`;
    html += `<li>Estimated σ (from MR): <strong>${sigma.toFixed(3)}</strong> (average MR = ${avgMR.toFixed(3)}).</li>`;

    if (target !== null) {
      html += `<li>Target: <strong>${target}</strong> (${direction === "above" ? "at or above is better" : "at or below is better"}). `;
      html += targetCoverageText ? (targetCoverageText + `</li>`) : `Target coverage not calculated for this period.</li>`;
    }

    // Simple, clearly labelled interpretation
    if (signals.length === 0) {
      html += `<li><strong>Interpretation:</strong> No clear special-cause signals were detected in this period. The pattern is consistent with natural/common variation (still interpret in clinical context).</li>`;
    } else {
      html += `<li><strong>Interpretation:</strong> This period shows special-cause signals: ${signals.join("; ")}.</li>`;

      // Optional: very short “where” hints (kept minimal)
      const whereBits = [];

      if (beyondIdx.length > 0) {
        const shown = beyondIdx.slice(0, 3).map(i => (startIndex + i + 1));
        whereBits.push(`outside limits at point${shown.length > 1 ? "s" : ""} ${shown.join(", ")}${beyondIdx.length > 3 ? ", …" : ""}`);
      }

      if (astro.indices && astro.indices.length > 0) {
        const shown = astro.indices.slice(0, 3).map(i => (startIndex + i + 1));
        whereBits.push(`outlier at point${shown.length > 1 ? "s" : ""} ${shown.join(", ")}${astro.indices.length > 3 ? ", …" : ""}`);
      }

      // Only add “where” if we actually have something specific to show
      if (whereBits.length > 0) {
        html += `<li><strong>Where to look:</strong> ${whereBits.join("; ")}.</li>`;
      }
    }

    if (capability && sigma > 0) {
      if (signals.length === 0) {
        html += `<li><strong>Estimated capability (this period):</strong> if the process remains stable, about <strong>${(capability.prob * 100).toFixed(1)}%</strong> of future points are expected to meet the target.</li>`;
      } else {
        html += `<li><strong>Capability:</strong> a target has been set, but because special-cause signals are present in this period, any capability estimate would be unreliable.</li>`;
      }
    }

    html += `</ul>`;

    // Remember last period for badge + helper (store structured information)
    if (idx === segments.length - 1) {
      lastPeriodSignals = signals;
      lastPeriodCapability = capability;
      lastPeriodHasCapability = sigma > 0 && !!capability;

      const hasTrend = trendRanges.length > 0;
      const hasRunViolation = runRanges.length > 0;
      const hasAstronomical = !!(astro.indices && astro.indices.length > 0);
      const nBeyond = beyondIdx.length;

      lastXmRAnalysis = {
        mean,
        ucl,
        lcl,
        sigma,
        avgMR,
        n,
        signals: signals.slice(),
        hasTrend,
        hasRunViolation,
        hasAstronomical,
        nBeyond,
        baselineCountUsed,
        target,
        direction,
        capability,
        isStable: signals.length === 0,
        // thresholds used (handy for helper explanations)
        shiftLength,
        trendLength
      };
    }
  });

  if (target !== null && segments.length > 1) {
    html += `<p><em>Note:</em> comparing means, limits and target performance between periods can indicate whether the process changed after interventions.</p>`;
  }

  summaryDiv.innerHTML = html;

  // Capability badge – last period only
  if (!capabilityDiv) return;

  if (target === null || !lastPeriodHasCapability) {
    capabilityDiv.innerHTML = "";
    return;
  }

  const hasAnySignals = lastPeriodSignals && lastPeriodSignals.length > 0;

  if (!hasAnySignals && lastPeriodCapability) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#fff59d;
        border:1px solid #ccc;
        border-radius:0.25rem;
      ">
        <div style="font-weight:bold; text-align:center;">PROCESS CAPABILITY (last period)</div>
        <div style="font-size:1.4rem; font-weight:bold; text-align:center; margin-top:0.2rem;">
          ${(lastPeriodCapability.prob * 100).toFixed(1)}%
        </div>
        <div style="font-size:0.8rem; margin-top:0.2rem;">
          (Estimated probability of meeting the target in the final period, assuming a stable process and approximate normality.)
        </div>
      </div>
    `;
  } else if (target !== null && hasAnySignals) {
    capabilityDiv.innerHTML = `
      <div style="
        display:inline-block;
        padding:0.6rem 1.2rem;
        background:#ffe0b2;
        border:1px solid #ccc;
        border-radius:0.25rem;
        max-width:32rem;
      ">
        <strong>Process not stable in the last period:</strong> special-cause signals are present.
        Focus on understanding and addressing these causes before relying on capability estimates.
      </div>
    `;
  } else {
    capabilityDiv.innerHTML = "";
  }
}

// Approximate standard normal CDF Φ(z)
function normalCdf(z) {
  // Abramowitz & Stegun approximation
  const t = 1 / (1 + 0.2316419 * Math.abs(z));
  const d = 0.3989423 * Math.exp(-0.5 * z * z);
  let prob = d * t * (0.3193815 +
    t * (-0.3565638 +
    t * (1.781478 +
    t * (-1.821256 +
    t * 1.330274))));
  if (z > 0) prob = 1 - prob;
  return prob;
}

// mean, sigma from XmR; target number; direction "above"/"below"
function computeTargetCapability(mean, sigma, target, direction) {
  if (!isFinite(mean) || !isFinite(sigma) || sigma <= 0 || !isFinite(target)) {
    return null;
  }
  const z = (target - mean) / sigma;
  let p;
  if (direction === "above") {
    // P(X >= target)
    p = 1 - normalCdf(z);
  } else {
    // P(X <= target)
    p = normalCdf(z);
  }
  return { prob: p, z };
}

// Parse dates safely, supporting NHS-style dd/mm/yyyy as well as ISO yyyy-mm-dd
function parseDateValue(xRaw) {
  if (xRaw instanceof Date && !isNaN(xRaw)) {
    return xRaw;
  }

  if (xRaw === null || xRaw === undefined) {
    return new Date(NaN);
  }

  const s = String(xRaw).trim();
  if (!s) return new Date(NaN);

  // ISO style: 2025-10-02 or 2025-10-02T...
  const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    const y = Number(isoMatch[1]);
    const m = Number(isoMatch[2]);
    const d = Number(isoMatch[3]);
    return new Date(y, m - 1, d);
  }

  // NHS-style day-first: dd/mm/yyyy or dd-mm-yyyy
  const dmMatch = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (dmMatch) {
    let day   = Number(dmMatch[1]);
    let month = Number(dmMatch[2]);
    let year  = Number(dmMatch[3]);
    if (year < 100) year += 2000; // e.g. 25 -> 2025
    return new Date(year, month - 1, day);
  }

  // Fallback: let the browser try
  return new Date(s);
}

// Parse numeric cells, including percentages like "55.17%"
function toNumericValue(raw) {
  if (raw === null || raw === undefined) return NaN;

  if (typeof raw === "number") return raw;

  const s = String(raw).trim();
  if (!s) return NaN;

  // Handle simple percentages, e.g. "55.17%" or "55.17 %"
  const percentMatch = s.match(/^(-?\d+(?:\.\d+)?)\s*%$/);
  if (percentMatch) {
    return Number(percentMatch[1]); // return 55.17
  }

  const num = Number(s);
  return isFinite(num) ? num : NaN;
}


// ---- Generate chart button ----
generateButton.addEventListener("click", () => {
  // Don’t auto-clear tips; only clear “hard errors”
  // If you want to preserve tips, comment out clearError().
  clearError();

  if (summaryDiv) summaryDiv.innerHTML = "";
  if (capabilityDiv) capabilityDiv.innerHTML = "";

  if (!validateBeforeGenerate()) return;

  const dateCol = dateSelect.value;
  const valueCol = valueSelect.value;
  const axisType = getAxisType();

  // --- 1) Build points depending on axis type ---
  let parsedPoints;

  if (axisType === "date") {
    parsedPoints = rawRows
      .map((row) => {
        const d = parseDateValue(row[dateCol]);
        const y = toNumericValue(row[valueCol]);
        if (!isFinite(d.getTime()) || !isFinite(y)) return null;
        return { x: d, y };
      })
      .filter(Boolean);
  } else {
    // sequence/category axis
    parsedPoints = rawRows
      .map((row, idx) => {
        const y = toNumericValue(row[valueCol]);
        if (!isFinite(y)) return null;

        const rawLabel = row[dateCol];
        const label =
          rawLabel !== undefined && rawLabel !== null && String(rawLabel).trim() !== ""
            ? String(rawLabel)
            : `Point ${idx + 1}`;

        return { x: idx, y, label };
      })
      .filter(Boolean);
  }

  // You can lower this if you want charts from fewer points
  if (parsedPoints.length < 3) {
    showError("Not enough valid data points after parsing. Check your column choices.");
    return;
  }

  // --- 2) Create points + labels for the chart ---
  let points, labels;

  if (axisType === "date") {
    points = [...parsedPoints].sort((a, b) => a.x - b.x);
    labels = points.map((p) => p.x.toISOString().slice(0, 10));
  } else {
    points = parsedPoints;
    labels = points.map((p) => p.label);
  }

  // --- baseline interpretation ---
  let baselineCount = null;
  if (baselineInput && baselineInput.value.trim() !== "") {
    const n = parseInt(baselineInput.value, 10);
    if (!isNaN(n) && n >= 2) baselineCount = Math.min(n, points.length);
  }

  const chartType = getSelectedChartType();

  // clear existing charts
  if (currentChart) {
    currentChart.destroy();
    currentChart = null;
  }
  if (mrChart) {
    mrChart.destroy();
    mrChart = null;
  }
  if (mrPanel) mrPanel.style.display = "none";

  if (chartType === "run") {
    drawRunChart(points, baselineCount, labels);
  } else {
    drawXmRChart(points, baselineCount, labels);
  }

  // clear “data changed” banner once we’ve successfully regenerated
  if (typeof clearDataModelDirty === "function") clearDataModelDirty();

  renderHelperState();
});


// ---- Chart drawing ----

function drawRunChart(points, baselineCount, labels) {
  const n = points.length;

  // Baseline count used for median
  let baselineCountUsed;
  if (baselineCount && baselineCount >= 2) {
    baselineCountUsed = Math.min(baselineCount, n);
  } else {
    baselineCountUsed = n;
  }

  const baselineValues = points.slice(0, baselineCountUsed).map(p => p.y);
  const values = points.map(p => p.y);
  const median = computeMedian(baselineValues);

  // Keep annotation date dropdown in sync with the current chart dates
  populateAnnotationDateOptions(labels);

  // Rule settings (defaults to 8 + 6 if inputs are missing/invalid)
  const { shiftLength, trendLength } = (typeof getRuleSettings === "function")
    ? getRuleSettings()
    : { shiftLength: 8, trendLength: 6 };

  // Detect rule hits (ranges)
  const runRanges = (typeof findLongRunRanges === "function")
    ? findLongRunRanges(values, median, shiftLength)
    : [];

  const trendRanges = (typeof findTrendRanges === "function")
    ? findTrendRanges(values, trendLength)
    : [];

  // Convert ranges -> per-point flags
  const runFlags = (typeof flagFromRanges === "function")
    ? flagFromRanges(values.length, runRanges)
    : new Array(values.length).fill(false);

  const trendFlags = (typeof flagFromRanges === "function")
    ? flagFromRanges(values.length, trendRanges)
    : new Array(values.length).fill(false);

   // Astronomical points (MAD-based), using baseline values as the reference
   const astro = findAstronomicalPoints(values, median, baselineValues, 3.5);
   const astroFlags = astro.flags;


  // Point colours (optional special-cause highlighting)
  const flagOnChart = (typeof shouldFlagSpecialCauseOnChart === "function")
    ? shouldFlagSpecialCauseOnChart()
    : true;

  const pointColours = values.map((_, i) => {
  if (!flagOnChart) return "#003f87";
  if (astroFlags[i]) return "#d73027"; // red for astronomical
  return (runFlags[i] || trendFlags[i]) ? "#ff8c00" : "#003f87";
});


  const { title, xLabel, yLabel } = getChartLabels(
    "Run Chart",
    "Date",
    "Value"
  );

  const target = getTargetValue();

  const datasets = [
    {
      label: "Value",
      data: values,
      pointRadius: 4,
      pointBackgroundColor: pointColours,
      borderColor: "#003f87",
      borderWidth: 2,
      fill: false
    },
    {
      label: "Median",
      data: values.map(() => median),
      borderDash: [6, 4],
      borderWidth: 2,
      borderColor: "#e41a1c",
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    }
  ];

  if (target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderDash: [4, 2],
      borderWidth: 2,
      borderColor: "#fdae61",
      pointRadius: 0,
      pointHoverRadius: 0,
      fill: false
    });
  }

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: {
      labels: labels,
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: title,
          font: { size: 16, weight: "bold" }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        },
        annotation: {
          annotations: buildAnnotationConfig(labels)
        }
      },
      elements: {
        point: {
          radius: 0,
          hoverRadius: 0
        }
      },
      scales: {
        x: {
          grid: { display: false },
          title: { display: !!xLabel, text: xLabel }
        },
        y: {
          grid: { display: false },
          title: { display: !!yLabel, text: yLabel }
        }
      }
    }
  });

  clearDataModelDirty();

  // Pass structured rule hits so the summary can label them clearly
  updateRunSummary(points, median, { runRanges, trendRanges, astro }, baselineCountUsed);

}

function drawXmRChart(points, baselineCount, labels) {
  if (!chartCanvas) return;

  const n = points.length;
  if (n < 12) {
    if (errorMessage) errorMessage.textContent = "XmR chart needs at least 12 points.";
    return;
  }

  // ---- Read “rules & interpretation” settings (with safe fallbacks) ----
  const { shiftLength, trendLength } =
    (typeof getRuleSettings === "function")
      ? getRuleSettings()
      : { shiftLength: 8, trendLength: 6 };

  const flagOnChart =
    (typeof shouldFlagSpecialCauseOnChart === "function")
      ? shouldFlagSpecialCauseOnChart()
      : true;

  const clampLcl =
    (typeof shouldClampLclAtZero === "function")
      ? shouldClampLclAtZero()
      : false;

  // ----- Segment definition from splits -----
  let effectiveSplits = Array.isArray(splits) ? splits.slice() : [];
  effectiveSplits = effectiveSplits
    .filter(i => Number.isInteger(i) && i >= 0 && i < n - 1)
    .sort((a, b) => a - b);

  const segmentStarts = [0];
  const segmentEnds = [];
  effectiveSplits.forEach(idx => {
    segmentEnds.push(idx);
    segmentStarts.push(idx + 1);
  });
  segmentEnds.push(n - 1);

  // Compute a "global" XmR as a fallback (no splits)
  // (computeXmR should accept clampLcl as third arg; if not, it’ll just ignore it)
  const globalResult = computeXmR(points, baselineCount, clampLcl);

  // ----- Global arrays for plotting -----
  const values = points.map(p => p.y);

  const meanLine     = new Array(n).fill(NaN);
  const uclLine      = new Array(n).fill(NaN);
  const lclLine      = new Array(n).fill(NaN);
  const oneSigmaUp   = new Array(n).fill(NaN);
  const oneSigmaDown = new Array(n).fill(NaN);
  const twoSigmaUp   = new Array(n).fill(NaN);
  const twoSigmaDown = new Array(n).fill(NaN);

  const pointColours = new Array(n).fill("#003f87");

  let anySigma = false;

  // We'll collect per-period results for the summary
  const segmentSummaries = [];

  // Track whether any raw LCL would be below 0 (so we can show the option conditionally)
  let anyRawLclBelowZero = false;

  // ----- Per-segment XmR -----
  for (let s = 0; s < segmentStarts.length; s++) {
    const start = segmentStarts[s];
    const end   = segmentEnds[s];

    const segPoints = points.slice(start, end + 1);

    // Only the first segment uses the user baseline; later segments use all points as baseline.
    const segBaseline = s === 0 ? baselineCount : null;

    const segResult = computeXmR(segPoints, segBaseline, clampLcl);
    const segPts    = segResult.points;

    const mean  = segResult.mean;
    const ucl   = segResult.ucl;
    const lcl   = segResult.lcl;
    const sigma = segResult.sigma;

    // If computeXmR returns rawLcl, use it to decide whether to show the clamp option
    if (typeof segResult.rawLcl === "number" && segResult.rawLcl < 0) {
      anyRawLclBelowZero = true;
    }

    // Store for multi-period summary
    segmentSummaries.push({
      startIndex: start,
      endIndex: end,
      labelStart: labels[start],
      labelEnd: labels[end],
      result: segResult
    });

    // Extra rule detection for colouring (shift/trend relative to MEAN within this segment)
    const segValues = segPts.map(p => p.y);

    const runRanges = (typeof findLongRunRanges === "function")
      ? findLongRunRanges(segValues, mean, shiftLength)
      : [];

    const trendRanges = (typeof findTrendRanges === "function")
      ? findTrendRanges(segValues, trendLength)
      : [];

    const runFlags = (typeof flagFromRanges === "function")
      ? flagFromRanges(segValues.length, runRanges)
      : new Array(segValues.length).fill(false);

    const trendFlags = (typeof flagFromRanges === "function")
      ? flagFromRanges(segValues.length, trendRanges)
      : new Array(segValues.length).fill(false);

    for (let i = 0; i < segPts.length; i++) {
      const globalIdx = start + i;

      // Colouring:
      // - beyond limits = red
      // - shift/trend = orange
      // - otherwise blue
      if (flagOnChart) {
        if (segPts[i].beyondLimits) {
          pointColours[globalIdx] = "#d73027";
        } else if (runFlags[i] || trendFlags[i]) {
          pointColours[globalIdx] = "#ff8c00";
        }
      }

      // Centre line & limits
      meanLine[globalIdx] = mean;
      uclLine[globalIdx]  = ucl;
      lclLine[globalIdx]  = lcl;

      // Sigma lines (only if sigma is valid)
      if (sigma && sigma > 0) {
        anySigma = true;
        oneSigmaUp[globalIdx]   = mean + sigma;
        oneSigmaDown[globalIdx] = mean - sigma;
        twoSigmaUp[globalIdx]   = mean + 2 * sigma;
        twoSigmaDown[globalIdx] = mean - 2 * sigma;
      }
    }
  }

  // ---- Show/hide the “Fix LCL at 0” option only when relevant ----
  if (typeof setLclClampVisibility === "function") {
    setLclClampVisibility(anyRawLclBelowZero);
  } else {
    // Fallback if you haven't added the helper yet
    const row = document.getElementById("lclClampRow");
    if (row) row.style.display = anyRawLclBelowZero ? "block" : "none";
  }

  // ----- Build datasets -----
  const datasets = [];

  // Main values
  datasets.push({
    label: "Value",
    data: values,
    borderColor: "#003f87",
    backgroundColor: "#003f87",
    pointRadius: 3,
    pointHoverRadius: 4,
    pointBackgroundColor: pointColours,
    pointBorderColor: "#ffffff",
    pointBorderWidth: 1,
    tension: 0,
    yAxisID: "y"
  });

  // Mean + limits
  datasets.push(
    {
      label: "Mean",
      data: meanLine,
      borderColor: "#d73027",
      borderDash: [6, 4],
      pointRadius: 0
    },
    {
      label: "UCL (3σ)",
      data: uclLine,
      borderColor: "#2ca25f",
      borderDash: [4, 4],
      pointRadius: 0
    },
    {
      label: "LCL (3σ)",
      data: lclLine,
      borderColor: "#2ca25f",
      borderDash: [4, 4],
      pointRadius: 0
    }
  );

  // Optional sigma reference lines
  if (anySigma) {
    const sigmaStyle = {
      borderColor: "rgba(0,0,0,0.12)",
      borderWidth: 1,
      borderDash: [2, 2],
      pointRadius: 0
    };

    datasets.push(
      { label: "+1σ", data: oneSigmaUp,   ...sigmaStyle },
      { label: "-1σ", data: oneSigmaDown, ...sigmaStyle },
      { label: "+2σ", data: twoSigmaUp,   ...sigmaStyle },
      { label: "-2σ", data: twoSigmaDown, ...sigmaStyle }
    );
  }

  // Target line (optional)
  const target = getTargetValue();
  if (target !== null) {
    datasets.push({
      label: "Target",
      data: values.map(() => target),
      borderColor: "#fdae61",
      borderWidth: 2,
      borderDash: [4, 2],
      pointRadius: 0,
      tension: 0
    });
  }

  // Update annotation and split dropdowns
  populateAnnotationDateOptions(labels);
  populateSplitOptions(labels);

  // ----- Create chart -----
  if (currentChart) currentChart.destroy();

  const { title, xLabel, yLabel } = getChartLabels("I-MR Chart", "Date", "Value");

  currentChart = new Chart(chartCanvas, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: title,
          font: { size: 16, weight: "bold" }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        },
        annotation: {
          annotations: buildAnnotationConfig(labels)
        }
      },
      elements: {
        point: { radius: 0, hoverRadius: 0 }
      },
      scales: {
        x: {
          grid: { display: false },
          title: { display: !!xLabel, text: xLabel }
        },
        y: {
          grid: { display: false },
          title: { display: !!yLabel, text: yLabel }
        }
      }
    }
  });

  // ----- Summary -----
  if (segmentSummaries.length > 0) {
    updateXmRMultiSummary(segmentSummaries, points.length);
  } else {
    updateXmRMultiSummary(
      [{
        startIndex: 0,
        endIndex: n - 1,
        labelStart: labels[0],
        labelEnd: labels[n - 1],
        result: globalResult
      }],
      points.length
    );
  }

  // ----- Show / hide MR chart depending on checkbox -----
  const showMR = showMRCheckbox ? showMRCheckbox.checked : true;

  // Use the last period for the MR chart (as a simple, focused view)
  const lastSegmentResult =
    segmentSummaries.length > 0
      ? segmentSummaries[segmentSummaries.length - 1].result
      : globalResult;

  if (showMR && lastSegmentResult) {
    drawMRChart(lastSegmentResult, labels);
  } else {
    if (mrChart) {
      mrChart.destroy();
      mrChart = null;
    }
    if (mrPanel) {
      mrPanel.style.display = "none";
    }
  }
}


// MR chart: average MR as centre, UCL = 3.268 * avgMR, LCL = 0
function drawMRChart(result, labels) {
  if (!mrPanel || !mrChartCanvas) return;

  const mrValues = result.mrValues;
  const mrLabels = labels.slice(1);

  if (!mrValues || mrValues.length === 0) {
    mrPanel.style.display = "none";
    return;
  }

  const avgMR = result.avgMR;
  const centre = avgMR;
  const uclMR = avgMR * 3.268; // D4 for n=2
  const lclMR = 0;

  mrPanel.style.display = "block";

  mrChart = new Chart(mrChartCanvas, {
    type: "line",
    data: {
      labels: mrLabels,
      datasets: [
        {
          label: "Moving range",
          data: mrValues,
          pointRadius: 4,
          pointBackgroundColor: "#003f87",
          borderColor: "#003f87",
          borderWidth: 2,
          fill: false
        },
        {
          label: "Average MR",
          data: mrValues.map(() => centre),
          borderDash: [6, 4],
          borderWidth: 2,
          borderColor: "#e41a1c",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        },
        {
          label: "UCL (MR)",
          data: mrValues.map(() => uclMR),
          borderDash: [4, 4],
          borderWidth: 2,
          borderColor: "#1a9850",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        },
        {
          label: "LCL (MR)",
          data: mrValues.map(() => lclMR),
          borderDash: [4, 4],
          borderWidth: 2,
          borderColor: "#1a9850",
          pointRadius: 0,
          pointHoverRadius: 0,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
	maintainAspectRatio: false,
      plugins: {
        title: {
          display: true,
          text: "Moving Range chart",
          font: {
            size: 14,
            weight: "bold"
          }
        },
        legend: {
          display: true,
          position: "bottom",
          align: "center"
        }
      },
      elements: {
        point: {
          radius: 0,
          hoverRadius: 0
        }
      },
      scales: {
        x: {
          grid: { display: false },
          title: {
            display: true,
            text: "Date (second and subsequent points)"
          }
        },
        y: {
          grid: { display: false },
          title: {
            display: true,
            text: "Moving range"
          }
        }
      }
    }
  });
}

// ---- AI helper function  -----

function answerSpcQuestion(question) {
  const q = (question || "").trim().toLowerCase();
  if (!q) {
    return "Please type a question about SPC or your chart (for example: \"Is the process stable?\" or \"What is a moving range chart?\").";
  }

  // Helper to match simple keyword / phrase based FAQs
  function matchFaq(items, text) {
    for (const item of items) {
      if (
        item.keywords.some(k =>
          Array.isArray(k)
            ? k.every(word => text.includes(word))
            : text.includes(k)
        )
      ) {
        return item.answer;
      }
    }
    return null;
  }

  // ----- 0. Conceptual SPC knowledge (no chart needed at all) -----
  const conceptualFaq = [
    {
      keywords: [
        "what is spc",
        "what is an spc",
        "what is a spc",
        "what is an spc chart",
        "what is a spc chart",
        "what are spc charts",
        "what is statistical process control",
        "spc chart",
        "control chart",
        ["statistical", "process", "control"]
      ],
      answer:
        "Statistical Process Control (SPC) is a way of using time-series charts to separate routine \"common-cause\" variation from unusual \"special-cause\" variation. " +
        "An SPC or control chart plots your measure over time, shows a typical average, and adds upper and lower control limits that represent the range you would expect if the system is stable. " +
        "When the pattern of points breaks simple rules (for example a point outside the limits or a long run of points on one side of the average), this is treated as a signal that the system may have changed."
    },
    {
      keywords: [
        "what is a run chart",
        "what is run chart",
        "run chart",
        ["what", "run chart"]
      ],
      answer:
        "A run chart is a simple time-series chart that shows your data in order with a median line. " +
        "It uses basic run and trend rules (such as long runs of points on one side of the median or steady upward or downward trends) to highlight possible special-cause variation even without formal control limits."
    }
  ];

  const conceptualHit = matchFaq(conceptualFaq, q);
  if (conceptualHit) return conceptualHit;

  // ----- 1. General SPC FAQs (can be answered without your specific chart) -----
  const generalFaq = [
    {
      keywords: [
        "moving range", "mr chart", "m-r chart",
        "use the moving range", "interpret the moving range",
        ["moving", "range"]
      ],
      answer:
        "A moving range (MR) chart shows how much each value changes from one point to the next. " +
        "On an XmR chart, the X chart shows the individual values over time and the MR chart shows the absolute difference between consecutive values. " +
        "If the moving ranges are mostly small and within their limits, the short-term variation looks stable. Large spikes in the moving range can indicate a one-off shock or a change in how the process behaves."
    },
    {
      keywords: ["xmr", "xm r", "i-mr", "individuals chart", "individual chart"],
      answer:
        "An XmR chart (also called an Individuals and Moving Range chart, or I-MR) is used when you have one value per time period, such as daily admissions, length of stay per day, or time for a single patient. " +
        "The X chart shows the individual values with a centre line and control limits. The MR chart shows the size of the step between each pair of consecutive points. " +
        "The average moving range is used to estimate the underlying variation (sigma), which then gives the control limits on the X chart."
    },
    {
      keywords: ["control limit", "control limits", "ucl", "lcl"],
      answer:
        "Control limits show the range of values you would expect to see from a stable process just due to routine variation. " +
        "They are not targets and they are not hard performance thresholds. Points outside the limits or unusual patterns inside the limits suggest special-cause variation that may be worth investigating."
    },
    {
      keywords: ["sigma", "standard deviation", "spread of the data", "variation"],
      answer:
        "In SPC, sigma is an estimate of the usual spread of the process. On an XmR chart, sigma is estimated from the average moving range between consecutive points. " +
        "Control limits are typically placed at plus or minus three sigma from the mean. A larger sigma means a wider spread of routine variation."
    },
    {
      keywords: ["common cause", "special cause"],
      answer:
        "Common-cause variation is the natural background noise of a stable system. Special-cause variation is a signal that the system may have changed, for example due to a new policy, a change in demand, or a data issue. " +
        "SPC helps you distinguish common-cause from special-cause variation so that you can avoid over-reacting to noise while still spotting real changes."
    },
    {
      keywords: ["run rule", "run rules", "spc rule", "spc rules", "signal", "signals"],
      answer:
        "SPC rules are simple patterns that are unlikely to occur if the process is stable. Examples include a point outside the control limits, a long run of points on one side of the mean, or a long trend of points steadily increasing or decreasing. " +
        "When one of these patterns appears, it is treated as a potential special-cause signal that may be worth investigating."
    },
    {
      keywords: ["capability", "capable process", "process capability"],
      answer:
        "Capability in this context is about the chance that future points will meet a chosen target, assuming the process stays as it is now. " +
        "If the process is stable, we can estimate the mean and sigma and then work out the percentage of future points likely to fall above or below a target threshold."
    },
    {
      keywords: ["baseline", "phase", "segment", "split the chart"],
      answer:
        "Splitting an SPC chart into phases (baselines) lets you compare the process before and after a known change, such as a new pathway or intervention. " +
        "Each segment gets its own mean and control limits so you can see whether the system has shifted, rather than averaging everything together."
    }
  ];

  const generalHit = matchFaq(generalFaq, q);
  if (generalHit) return generalHit;

  // ----- 2. Chart-specific interpretation (XmR only) -----
  const chartType = (typeof getSelectedChartType === "function")
    ? getSelectedChartType()
    : "xmr";

  if (chartType !== "xmr") {
    return (
      "I can answer general SPC questions for any chart, but the automatic detailed interpretation currently applies only to XmR charts. " +
      "Please switch to an XmR chart if you want automated interpretation of stability, signals, limits or capability."
    );
  }

  if (!lastXmRAnalysis) {
    return (
      "I can only interpret your chart once an XmR chart has been generated. " +
      "Please create an XmR chart first, then ask me about stability, signals, control limits, target performance or capability."
    );
  }

  // ----- Special: “My chart” standard questions -----
  const isMyChartQ =
    q.includes("what is my chart telling") ||
    q.includes("what's my chart telling") ||
    q.includes("what is this chart telling") ||
    q.includes("what decision should i make") ||
    q.includes("what should i do") ||
    (q.includes("decision") && q.includes("make")) ||
    q.includes("what about my target") ||
    (q.includes("my target") && q.includes("what about"));

  if (isMyChartQ) {
    const a = lastXmRAnalysis;

    // If we somehow got here without analysis
    if (!a) {
      return "Please generate an XmR chart first, then ask one of the “My chart” questions.";
    }

    const signals = Array.isArray(a.signals) ? a.signals : [];
    const stable = !!a.isStable;

    // Helpful phrasing for signals list
    const signalsText =
      signals.length === 0
        ? "No special-cause signals detected."
        : `Signals detected: ${signals.join("; ")}.`;

    // Target summary (if present)
    let targetText = "No target is set on this chart.";
    if (a.target != null && a.direction) {
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      targetText = `Target is ${a.target} (${dirText} is better).`;

      if (stable && a.capability && typeof a.capability.prob === "number") {
        targetText += ` If the process stays stable, about ${(a.capability.prob * 100).toFixed(1)}% of future points are expected to meet the target.`;
      } else if (!stable) {
        targetText += " Because special-cause signals are present, any capability estimate is unreliable until the process is stable.";
      }
    }

    // 1) “What is my chart telling me?”
    if (q.includes("what is my chart telling") || q.includes("what is this chart telling") || q.includes("what's my chart telling")) {
      const meanText = (typeof a.mean === "number") ? a.mean.toFixed(2) : "n/a";
      const uclText  = (typeof a.ucl === "number") ? a.ucl.toFixed(2) : "n/a";
      const lclText  = (typeof a.lcl === "number") ? a.lcl.toFixed(2) : "n/a";

      return (
        `Your chart summary: mean ≈ ${meanText}, limits ≈ [${lclText}, ${uclText}]. ` +
        (stable
          ? "The process looks stable (common-cause variation). "
          : "The process does not look stable (special-cause variation). ") +
        signalsText + " " +
        targetText
      );
    }

    // 2) “What decision should I make?”
    if (q.includes("what decision should i make") || q.includes("what should i do") || (q.includes("decision") && q.includes("make"))) {
      if (!stable) {
        return (
          "Decision guidance: don’t react to individual points as if they are “performance”. " +
          "Because special-cause signals are present, treat this as evidence the system may have changed. " +
          "Investigate the timing of the signals (what changed in the process/data), confirm the change is real, and then re-baseline (use a split) once the new system is established. " +
          targetText
        );
      }

      // Stable process
      if (a.target == null) {
        return (
          "Decision guidance: the process looks stable, so most up-and-down movement is routine variation. " +
          "If performance is not good enough, the decision is to change the system (not chase individual points), then use the chart to see whether a real shift occurs. " +
          "If performance is acceptable, the decision is to hold the system steady and continue monitoring."
        );
      }

      // Stable + target set
      if (a.capability && typeof a.capability.prob === "number") {
        const pct = (a.capability.prob * 100);
        if (pct >= 90) {
          return (
            `Decision guidance: the process is stable and is very likely to meet the target (~${pct.toFixed(1)}%). ` +
            "Hold the gains, standardise the current approach, and keep monitoring for any new special-cause signals."
          );
        }
        if (pct >= 50) {
          return (
            `Decision guidance: the process is stable but only sometimes meets the target (~${pct.toFixed(1)}%). ` +
            "If the target matters, you’ll need a system change to shift the mean and/or reduce variation. " +
            "Use improvement cycles and watch for a sustained shift before re-baselining."
          );
        }
        return (
          `Decision guidance: the process is stable but unlikely to meet the target (~${pct.toFixed(1)}%). ` +
          "A system redesign is needed (shift the mean and/or reduce variation). Consider stratifying data, reviewing drivers of variation, and testing changes."
        );
      }

      return (
        "Decision guidance: the process looks stable. With a target set, the key question is whether the mean is on the right side of the target and whether variation frequently crosses it. " +
        "If it does, you’ll likely need a system change to make target achievement more reliable."
      );
    }

    // 3) “What about my target?”
    if (q.includes("what about my target") || (q.includes("my target") && q.includes("what about"))) {
      return targetText;
    }
  }



  const a = lastXmRAnalysis;
  const lines = [];

  // ----- 2a. Stability / signals -----
  if (
    q.includes("stable") || q.includes("stability") ||
    q.includes("in control") || q.includes("out of control") ||
    q.includes("special cause") || q.includes("signal") ||
    q.includes("run rule") || q.includes("rule broken") ||
    q.includes("any signals")
  ) {
    if (a.isStable) {
      lines.push(
        "This segment of the XmR chart appears stable: no SPC rules are triggered and the points fluctuate randomly around the mean within the control limits."
      );
    } else if (Array.isArray(a.signals) && a.signals.length > 0) {
      const count = a.signals.length;
      const labels = a.signals.map(s => s.description || s.type || "signal").join("; ");
      lines.push(
        `This XmR chart shows evidence of special-cause variation. I can see ${count} signal${count > 1 ? "s" : ""}: ${labels}. ` +
        "These patterns are unlikely to arise from common-cause variation alone and suggest that the system may have changed."
      );
    } else {
      lines.push(
        "The chart does not look completely stable, but no specific SPC signals have been recorded. Check for obvious shifts, trends or outlying points."
      );
    }
  }

  // ----- 2b. Mean and control limits -----
  if (
    q.includes("mean") || q.includes("average") ||
    q.includes("ucl") || q.includes("lcl") ||
    q.includes("control limit") || q.includes("limits")
  ) {
    if (typeof a.mean === "number" && typeof a.sigma === "number") {
      const meanText = a.mean.toFixed(2);
      const uclText = (a.ucl != null ? a.ucl.toFixed(2) : "not calculated");
      const lclText = (a.lcl != null ? a.lcl.toFixed(2) : "not calculated");
      lines.push(
        `The current segment has an estimated mean of ${meanText}. The upper control limit (UCL) is ${uclText} and the lower control limit (LCL) is ${lclText}. ` +
        "These are based on the average moving range and represent the range you would expect from common-cause variation in this period."
      );
    } else {
      lines.push(
        "Mean and control limits could not be calculated for this chart. Check that there are enough data points and that the values are numeric."
      );
    }
  }

  // ----- 2c. Short-term variation / sigma -----
  if (
    q.includes("sigma") || q.includes("variation") ||
    q.includes("spread") || q.includes("variability")
  ) {
    if (typeof a.sigma === "number" && typeof a.avgMR === "number") {
      lines.push(
        `The estimated sigma (spread) for this segment is approximately ${a.sigma.toFixed(2)}, based on an average moving range of ${a.avgMR.toFixed(2)}. ` +
        "This captures the usual short-term variation between consecutive points and is used to set the control limits."
      );
    } else {
      lines.push(
        "An estimate of sigma could not be calculated. This usually happens if there are too few points or no variation in the data."
      );
    }
  }

  // ----- 2d. Target / direction / performance relative to target -----
  if (
    q.includes("target") || q.includes("goal") ||
    q.includes("above target") || q.includes("below target") ||
    q.includes("better") || q.includes("worse") ||
    q.includes("improve") || q.includes("improvement")
  ) {
    if (a.target == null || !a.direction) {
      lines.push(
        "A target has not been set for this chart, or the direction of improvement (above or below the target) is not defined. " +
        "Set a target value and specify whether higher or lower is better to get a clearer view of performance."
      );
    } else if (!a.isStable) {
      lines.push(
        "Because the process is not yet stable, performance against the target may change unpredictably. " +
        "Stabilise the process first, then reassess how reliably the target is being met."
      );
    } else if (a.capability && typeof a.capability.prob === "number") {
      const prob = (a.capability.prob * 100).toFixed(1);
      const dirText = a.direction === "above" ? "at or above" : "at or below";
      lines.push(
        `Given the current stable process and a target of ${a.target}, about ${prob}% of future points are expected to fall ${dirText} the target. ` +
        "This assumes that the underlying process does not change."
      );
    } else {
      lines.push(
        "A formal capability estimate against the target could not be calculated, but you can still use the chart to see whether the mean is comfortably on the desired side of the target and how often points cross it."
      );
    }
  }

  // ----- 2e. Splits / baselines (using the global splits array) -----
  if (
    q.includes("split") || q.includes("baseline") ||
    q.includes("phase") || q.includes("segment")
  ) {
    if (!Array.isArray(splits) || splits.length === 0) {
      lines.push(
        "No splits have been added, so the whole series is treated as one baseline. " +
        "If a known system change occurred, you can add a split so that before-and-after periods each get their own mean and control limits."
      );
    } else {
      lines.push(
        `You have added ${splits.length} split${splits.length > 1 ? "s" : ""} to this chart. ` +
        "Each split marks a point where a new baseline begins with its own mean and limits, allowing you to compare periods before and after key changes."
      );
    }
  }

  // ----- 2f. Moving range chart questions (chart-specific) -----
  if (
    q.includes("moving range") || q.includes("mr chart") || q.includes("m-r chart") ||
    q.includes("mr line") || q.includes("mr panel")
  ) {
    lines.push(
      "The moving range (MR) chart under the main XmR chart shows the size of the jump between one point and the next. " +
      "Large spikes in the MR chart indicate abrupt changes between consecutive observations, while a stable band of small ranges suggests consistent short-term behaviour."
    );
  }

  // ----- 3. Fallback if nothing matched in chart-specific logic -----
  if (lines.length === 0) {
    return (
      "I could not match that question to a specific SPC topic. " +
      "Try asking about stability, signals, control limits, sigma (variation), target performance, capability, moving range, or splits/baselines."
    );
  }

  // ----- 4. Final reminder -----
  lines.push(
    "Always interpret SPC charts alongside clinical or operational context, rather than in isolation. " +
    "Use the signals as prompts for discussion, not as automatic proof that a change has worked."
  );

  return lines.join(" ");
}

function renderHelperState() {
  if (!spcHelperIntro) return;

  const hasChart = !!lastXmRAnalysis;

  // 1) Intro text
  if (!hasChart) {
    spcHelperIntro.innerHTML = `
      <div><strong>SPC helper</strong></div>
      <div>Ask a general question before you load any data, or use a suggested prompt below.</div>
    `;
  } else {
    spcHelperIntro.innerHTML = `
      <div><strong>Chart helper</strong></div>
      <div>Use the <strong>My chart</strong> questions for a tailored interpretation.</div>
    `;
  }

  // 2) General chips (always available)
  const generalQs = [
    "What is an SPC chart?",
    "What is a run chart?",
    "What is an XmR chart?",
    "What is common cause vs special cause variation?",
    "How do control limits work?"
  ];

  if (spcHelperChipsGeneral) {
    spcHelperChipsGeneral.innerHTML = generalQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");
    spcHelperChipsGeneral.classList.remove("is-disabled");
  }

  // 3) My chart chips (available only when a chart exists)
  const chartQs = [
    "What is my chart telling me?",
    "What decision should I make?",
    "What about my target?"
  ];

  if (spcHelperChipsChart) {
    spcHelperChipsChart.innerHTML = chartQs
      .map(q => `<button type="button" class="spc-chip" data-q="${escapeHtml(q)}">${escapeHtml(q)}</button>`)
      .join("");

    // Optional: disable interaction until a chart exists
    if (!hasChart) spcHelperChipsChart.classList.add("is-disabled");
    else spcHelperChipsChart.classList.remove("is-disabled");
  }
}

const mrToggleRow = document.getElementById("mrToggleRow");

function updateMrToggleVisibility() {
  const chartType = getSelectedChartType();

  if (mrToggleRow) {
    mrToggleRow.style.display = (chartType === "xmr") ? "flex" : "none";
  }

  // If we're not in XmR mode, ensure MR panel is hidden (and chart destroyed)
  if (chartType !== "xmr") {
    hideMrPanelNow();
  } else {
    // On first load into XmR, default ON if checkbox exists
    if (showMRCheckbox && showMRCheckbox.checked !== true) {
      // don’t force it on if user explicitly turned it off later;
      // this line only matters for initial state where it should be checked in HTML anyway
    }
  }
}



// ===============================
// Chart context menu + export tools
// ===============================

const chartContextMenu = document.getElementById("chartContextMenu");

// Track which point index was right-clicked
let contextMenuPointIndex = null;

// Helper: hide menu
function hideChartContextMenu() {
  if (!chartContextMenu) return;
  chartContextMenu.style.display = "none";
  contextMenuPointIndex = null;
}

// Helper: show menu at cursor, clamped to viewport
function showChartContextMenu(clientX, clientY, pointIndex) {
  if (!chartContextMenu) return;

  contextMenuPointIndex = pointIndex;

  chartContextMenu.style.display = "block";
  chartContextMenu.style.left = "0px";
  chartContextMenu.style.top = "0px";

  // Clamp so it stays on-screen
  const menuRect = chartContextMenu.getBoundingClientRect();
  const pad = 8;
  let x = clientX;
  let y = clientY;

  if (x + menuRect.width + pad > window.innerWidth) x = window.innerWidth - menuRect.width - pad;
  if (y + menuRect.height + pad > window.innerHeight) y = window.innerHeight - menuRect.height - pad;
  if (x < pad) x = pad;
  if (y < pad) y = pad;

  chartContextMenu.style.left = `${x}px`;
  chartContextMenu.style.top = `${y}px`;
}

// Helper: get the nearest chart point index from a mouse event
function getNearestPointIndexFromEvent(evt) {
  if (!currentChart) return null;

  const elements = currentChart.getElementsAtEventForMode(
    evt,
    "nearest",
    { intersect: true },
    true
  );

  if (!elements || elements.length === 0) return null;

  // Chart.js v3+: element has .index
  const idx = elements[0].index;
  return Number.isFinite(idx) ? idx : null;
}

// ---- Split helpers ----
function addSplitAfterIndex(splitAfterIndex) {
  if (!Number.isFinite(splitAfterIndex)) return;

  // can’t split after last point
  const labels = currentChart?.data?.labels || [];
  if (labels.length === 0) return;
  if (splitAfterIndex < 0 || splitAfterIndex >= labels.length - 1) {
    alert("You can’t split after the last point.");
    return;
  }

  // avoid duplicates
  if (!splits.includes(splitAfterIndex)) {
    splits.push(splitAfterIndex);
    splits.sort((a, b) => a - b);
  }

  // keep the dropdown in sync (if present)
  if (labels && labels.length) {
    populateSplitOptions(labels);
  }

  // redraw with new split
  if (generateButton) generateButton.click();
}

// ---- Export helpers ----

// Return the canvases to export (main + MR if shown)
function getExportCanvases() {
  const canvases = [];
  if (chartCanvas) canvases.push(chartCanvas);

  const showMR = showMRCheckbox ? showMRCheckbox.checked : false;
  const mrVisible = mrPanel && mrPanel.style.display !== "none";
  if (showMR && mrVisible && mrChartCanvas) canvases.push(mrChartCanvas);

  return canvases;
}

// Build one combined image from multiple canvases (stacked vertically).
// Optionally add summary text under the charts.
function buildCompositeCanvas({ includeSummaryText }) {
  const canvases = getExportCanvases();
  if (!canvases.length) return null;

  // widths/heights from actual rendered pixels
  const widths = canvases.map(c => c.width);
  const heights = canvases.map(c => c.height);

  const outWidth = Math.max(...widths);
  const chartsHeight = heights.reduce((a, b) => a + b, 0);

  // Summary text (plain) – keep it simple for clipboard/export
  let summaryText = "";
  if (includeSummaryText && summaryDiv) {
    summaryText = (summaryDiv.innerText || "").trim();
  }

  // Basic text layout
  const fontSize = 14;
  const lineHeight = Math.round(fontSize * 1.35);
  const padding = 16;

  // Wrap text to fit image width
  function wrapLines(ctx, text, maxWidth) {
    if (!text) return [];
    const words = text.replace(/\s+/g, " ").split(" ");
    const lines = [];
    let line = "";

    for (const w of words) {
      const test = line ? `${line} ${w}` : w;
      if (ctx.measureText(test).width <= maxWidth) {
        line = test;
      } else {
        if (line) lines.push(line);
        line = w;
      }
    }
    if (line) lines.push(line);
    return lines;
  }

  // Create output canvas
  const out = document.createElement("canvas");
  const ctx = out.getContext("2d");

  // temp set font for measuring
  ctx.font = `${fontSize}px system-ui, -apple-system, Segoe UI, sans-serif`;

  const textMaxWidth = Math.max(200, outWidth - padding * 2);
  const textLines = wrapLines(ctx, summaryText, textMaxWidth);
  const textHeight = summaryText ? (padding + textLines.length * lineHeight + padding) : 0;

  out.width = outWidth;
  out.height = chartsHeight + (summaryText ? textHeight : 0);

  // White background
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, out.width, out.height);

  // Draw charts
  let y = 0;
  canvases.forEach((c, i) => {
    // center each canvas if narrower
    const x = Math.round((outWidth - c.width) / 2);
    ctx.drawImage(c, x, y);
    y += c.height;
  });

  // Draw summary text
  if (summaryText) {
    // separator line
    ctx.fillStyle = "#eef2f6";
    ctx.fillRect(0, y, out.width, 1);
    y += padding;

    ctx.fillStyle = "#111111";
    ctx.font = `${fontSize}px system-ui, -apple-system, Segoe UI, sans-serif`;

    let ty = y + lineHeight;
    for (const line of textLines) {
      ctx.fillText(line, padding, ty);
      ty += lineHeight;
    }
  }

  return out;
}

async function copyCanvasToClipboard(canvas) {
  if (!canvas) return;

  // Modern clipboard image API
  if (navigator.clipboard && window.ClipboardItem) {
    const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/png"));
    if (!blob) throw new Error("Failed to create image.");
    await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
    return;
  }

  // Fallback
  alert("Copy to clipboard is not supported in this browser. Try 'Save chart(s) as…' instead.");
}

function downloadCanvasAsPng(canvas, filename) {
  if (!canvas) return;
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = filename;
  link.click();
}

// ---- Existing top button: Download chart as PNG ----
// Update to download chart(s) (main + MR if shown), as one image.
if (downloadBtn) {
  downloadBtn.addEventListener("click", () => {
    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }
    const composite = buildCompositeCanvas({ includeSummaryText: false });
    downloadCanvasAsPng(composite, "spc-charts.png");
  });
}

// ---- Existing split dropdown button still works ----
if (addSplitButton) {
  addSplitButton.addEventListener("click", () => {
    if (!splitPointSelect) return;

    const value = splitPointSelect.value;
    if (value === "") {
      alert("Please choose a point to split after.");
      return;
    }

    const idx = parseInt(value, 10);
    if (!Number.isFinite(idx)) return;

    addSplitAfterIndex(idx);
  });
}

// ---- Right-click on chart: show menu ----
if (chartCanvas) {
  chartCanvas.addEventListener("contextmenu", (evt) => {
    // Only show our menu when user right-clicks a point
    const idx = getNearestPointIndexFromEvent(evt);
    if (idx === null) return; // allow normal browser menu if not on a point

    evt.preventDefault();
    showChartContextMenu(evt.clientX, evt.clientY, idx);
  });
}

// Hide menu on click elsewhere / escape / scroll
document.addEventListener("click", () => hideChartContextMenu());
document.addEventListener("keydown", (e) => { if (e.key === "Escape") hideChartContextMenu(); });
document.addEventListener("scroll", () => hideChartContextMenu(), true);

// Menu actions
if (chartContextMenu) {
  chartContextMenu.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;

    const action = btn.getAttribute("data-action");
    hideChartContextMenu();

    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }

    try {
      if (action === "addSplit") {
        if (contextMenuPointIndex === null) return;
        addSplitAfterIndex(contextMenuPointIndex);
      }

      if (action === "copyCharts") {
        const composite = buildCompositeCanvas({ includeSummaryText: false });
        await copyCanvasToClipboard(composite);
        alert("Chart image copied to clipboard.");
      }

      if (action === "copyChartsAndAnalysis") {
        const composite = buildCompositeCanvas({ includeSummaryText: true });
        await copyCanvasToClipboard(composite);
        alert("Chart + analysis image copied to clipboard.");
      }

      if (action === "saveChartsAs") {
        const name = prompt("Save as file name (PNG):", "spc-charts.png") || "spc-charts.png";
        const safe = name.toLowerCase().endsWith(".png") ? name : `${name}.png`;
        const composite = buildCompositeCanvas({ includeSummaryText: false });
        downloadCanvasAsPng(composite, safe);
      }
    } catch (err) {
      console.error(err);
      alert("Sorry — that action failed in this browser. Try 'Save chart(s) as…' instead.");
    }
  });
}

function showHelperAnswer(questionText) {
  if (!spcHelperOutput) return;

  const q = (questionText ?? aiQuestionInput?.value ?? "").trim();
  if (!q) {
    spcHelperOutput.innerHTML = `<p>${escapeHtml("Type a question (or click a suggestion) to get started.")}</p>`;
    return;
  }

  const ans = answerSpcQuestion(q);
  spcHelperOutput.innerHTML = `<p>${escapeHtml(ans)}</p>`;
}

if (aiAskButton && aiQuestionInput) {
  aiAskButton.addEventListener("click", () => {
    showHelperAnswer();
  });

  aiQuestionInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      showHelperAnswer();
    }
  });
}

function handleChipClick(e) {
  const btn = e.target.closest("button[data-q]");
  if (!btn) return;

  const q = btn.getAttribute("data-q") || "";
  if (aiQuestionInput) aiQuestionInput.value = q;

  showHelperAnswer(q);
}

if (spcHelperChipsGeneral) {
  spcHelperChipsGeneral.addEventListener("click", handleChipClick);
}
if (spcHelperChipsChart) {
  spcHelperChipsChart.addEventListener("click", handleChipClick);
}



if (clearSplitsButton) {
  clearSplitsButton.addEventListener("click", () => {
    splits = [];

    if (splitPointSelect) {
      splitPointSelect.value = "";
    }

    if (getSelectedChartType() === "xmr") {
      generateButton.click();
    }
  });
}

// -----------------------------
// Help section toggle
// -----------------------------
function toggleHelpSection() {
  const help = document.getElementById("helpSection");
  if (!help) return;

  const isHidden = help.style.display === "none" || help.style.display === "";

  if (isHidden) {
    help.style.display = "block";
    help.scrollIntoView({ behavior: "smooth" });
  } else {
    help.style.display = "none";
  }
}

let spcHelperHasBeenOpened = false;

function toggleSpcHelper() {
  const panel = document.getElementById("spcHelperPanel");
  if (!panel) return;

  const isVisible = panel.classList.toggle("visible");

  // Populate chips / intro once, when the helper is first opened
  if (isVisible && !spcHelperHasBeenOpened) {
    if (typeof renderHelperState === "function") renderHelperState();
    spcHelperHasBeenOpened = true;
  }
}




const spcHelperCloseBtn = document.getElementById("spcHelperCloseBtn");

if (spcHelperCloseBtn) {
  spcHelperCloseBtn.addEventListener("click", () => {
    if (spcHelperPanel) {
      spcHelperPanel.classList.remove("visible");
    }
  });
}


const resetButton = document.getElementById("resetButton");

if (resetButton) {
  resetButton.addEventListener("click", resetAll);
}

// Allow Escape key to close the SPC helper
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (spcHelperPanel && spcHelperPanel.classList.contains("visible")) {
      spcHelperPanel.classList.remove("visible");
    }
  }
});

function countValidNumericPoints() {
  if (!rawRows || !rawRows.length) return 0;
  const valueCol = valueSelect?.value;
  if (!valueCol) return 0;

  let n = 0;
  for (const row of rawRows) {
    const y = toNumericValue(row[valueCol]);
    if (isFinite(y)) n++;
  }
  return n;
}

function enforceChartTypeSuitabilityAndRegen() {
  if (!rawRows || !rawRows.length) return;

  const chartType = getSelectedChartType();
  const valueCol = valueSelect?.value;

  let validPoints = 0;
  if (valueCol) {
    for (const row of rawRows) {
      const y = toNumericValue(row[valueCol]);
      if (isFinite(y)) validPoints++;
    }
  }

  const minXmr = 12;

  if (chartType === "xmr" && validPoints < minXmr) {
    showError(
      `XmR charts need at least ${minXmr} valid numeric points. ` +
      `You currently have ${validPoints}. Switching back to a run chart.`
    );

    // revert to run chart
    const runRadio = document.querySelector(
      "input[name='chartType'][value='run']"
    );
    if (runRadio) runRadio.checked = true;

    return;
  }

  // Suitable → regenerate immediately
  generateButton.click();
}

// ---- Auto-regenerate when chart type or axis type changes ----
function wireAutoRedrawControls() {
  // Chart type radios (run / xmr)
  document.querySelectorAll("input[name='chartType']").forEach(radio => {
    radio.addEventListener("change", () => {
      // show/hide MR toggle row + MR panel
      if (typeof updateMrToggleVisibility === "function") {
        updateMrToggleVisibility();
      }

      // only regenerate if data exists
      if (rawRows && rawRows.length) {
        // Enforce minimum points for XmR + regen
        if (typeof enforceChartTypeSuitabilityAndRegen === "function") {
          enforceChartTypeSuitabilityAndRegen();
        } else {
          generateButton.click();
        }
      }
    });
  });

  // Axis type radios (date / sequence)
  document.querySelectorAll("input[name='axisType']").forEach(radio => {
    radio.addEventListener("change", () => {
      if (rawRows && rawRows.length) {
        generateButton.click();
      }
    });
  });

  // Run once on load so MR toggle visibility matches initial selection
  if (typeof updateMrToggleVisibility === "function") {
    updateMrToggleVisibility();
  }
}

// Call after the DOM is available (safe even if script is at bottom, but robust)
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", wireAutoRedrawControls);
} else {
  wireAutoRedrawControls();
}


if (downloadPdfBtn) {
  downloadPdfBtn.addEventListener("click", () => {
    const reportElement = document.getElementById("reportContent");
    if (!reportElement) {
      alert("Report content not found.");
      return;
    }
    if (!currentChart) {
      alert("Please generate a chart first.");
      return;
    }

    // Basic options – you can tweak orientation/format later
    const opt = {
      margin:       10,
      filename:     "spc-report.pdf",
      image:        { type: "jpeg", quality: 0.98 },
      html2canvas:  { scale: 2, scrollY: -window.scrollY },
      jsPDF:        { unit: "mm", format: "a4", orientation: "landscape" }
    };


    html2pdf().set(opt).from(reportElement).save();
  });
}

renderHelperState();
