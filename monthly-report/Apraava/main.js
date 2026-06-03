// Shared constants — update here instead of hunting through export functions
const CONFIG = {
  PPTX_NAVY: "12284A",
  PPTX_BLUE: "1E4F9A",
  PPTX_ROWS_PER_SLIDE: 6,
  REPORT_CONTENT_FONT_SIZE: 14,
  CONTACT_SLIDE_TEXT: "FOR FURTHER DETAILS:\nDiptesh Saha\nCISO & Practice Head - Cyber Security & Managed Security\nContact No. 7338882888\ndiptesh.s@snsin.com",
  CONTACT_FOOTER_TEXT: "For further details: Diptesh Saha | diptesh.s@snsin.com",
  COMPANY_NAME: "Secure Network Solutions India Pvt Ltd"
};

let slideNavController = null;
let formSectionsOrdered = false;

const APRAAVA_COPY = {
  summaryTitle: "Potential Incidents (severity summary)",
  summarySubtitle: "Potential incidents and alert activity during",
  trendTitle: "Potential Incident Tickets Trend",
  riskTitle: "Potential Incidents – Risks Mitigated",
  potIncTitle: "Potential Incidents"
};

const APRAAVA_INITIAL = {};

const fields = {
  customer: "customerName",
  range: "dateRange",
  preparedBy: "preparedBy",
  submittedOn: "submittedOn",
  executiveSummary: "executiveSummary",
  trendNote: "trendNote",
  trendNarrative: "trendNarrative",
  potIncidentsNarrative: "potIncidentsNarrative",
  riskNarrative: "riskNarrative",
  inventoryNote: "inventoryNote",
  kpEnhancement: "kpEnhancement",
  kpDecommissioning: "kpDecommissioning",
  kpRuleImplementation: "kpRuleImplementation"
};

const STATIC_REVIEWED_BY = "Kishore Kumar";
const STATIC_APPROVED_BY = "Diptesh Saha";

/** Apraava slide nav labels (matches PPT titles). */
const APRAAVA_SLIDE_NAV_LABELS = {
  cover: "Cover Page",
  revision: "Document Revision History",
  engagement: "The Engagement",
  "alert-quad": "Potential Incidents and Alert Summary",
  trend: "Potential Incident Tickets Trend",
  "pot-incidents": "Potential Incidents",
  eps: "EPS Trend Plot",
  "top-eps": "Top 10 Devices Contributing To Highest EPS",
  risks: "Potential Incidents – Risks Mitigated",
  "sla-monthly": "Response Time SLA",
  inventory: "Integrated Device Inventory",
  "key-points": "Key Points - Overall Summary",
  contact: "Your Trusted Security Advisor"
};

function configureSlideNavLabels() {
  if (typeof window === "undefined") return;
  window.SLIDE_NAV_LABELS = { ...(window.SLIDE_NAV_LABELS || {}), ...APRAAVA_SLIDE_NAV_LABELS };
}

let alertChart;
let tpfpChart;
let trendChart;
let potIncidentsChart;
let epsTrendChart;
let topEpsChart;
let slaPctChart;
let apraavaQuadCharts = [];
let snsLogoDataUrl = "";
let clientLogoDataUrl = "";
let puzzleLogoDataUrl = "";
let defaultPuzzleDataUrl = "";
let exportRestoreCallback = null;

const PPT_LAYOUT_KEY = "LAYOUT_16X9";

const PPT_SLIDE_IN = { w: 10, h: 5.625 };
/** Usable body area below title bar and above footer on 16:9 slides (inches). */
const PPT_BODY = { top: 0.5, bottom: 4.85, left: 0.35, right: 9.65 };
PPT_BODY.height = PPT_BODY.bottom - PPT_BODY.top;
PPT_BODY.width = PPT_BODY.right - PPT_BODY.left;

const pptLayoutConfig = {
  LAYOUT_16X9: {
    pptxLayout: "SNS_16X9",
    width: PPT_SLIDE_IN.w,
    height: PPT_SLIDE_IN.h,
    reportClass: "ppt-wide"
  }
};

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

// Escapes user-supplied text before inserting into innerHTML
function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Converts **text** to <strong>text</strong> and handles [MONTH] placeholder for HTML display.
 */
function formatRichTextHTML(text) {
  return formatReportRichTextHTML(text, getReportMonth, { italic: false });
}

function getKeyPointsSourceText() {
  const unified = getValue("kpContent");
  if (String(unified || "").trim()) return unified;

  const legacyBlocks = [
    { heading: "Device Decommissioning", field: "kpDecommissioning" },
    { heading: "Rule Implementation", field: "kpRuleImplementation" },
    { heading: "MSSP Contract & Log Retention Renewal", field: "kpContract" }
  ];

  const parts = [];
  legacyBlocks.forEach(({ heading, field }) => {
    const body = getValue(field);
    if (!String(body || "").trim()) return;
    parts.push(`**${heading}:**\n${body}`);
  });
  return parts.join("\n\n");
}

function parseKeyPointsSections(text) {
  const raw = String(text || "").replace(/\r\n/g, "\n").trim();
  if (!raw) return [];

  const headingRe = /(?:^|\n)\s*\*\*([^*\n]+?)\*\*:?\s*/g;
  const matches = [...raw.matchAll(headingRe)];
  if (!matches.length) {
    return [{ heading: null, body: raw }];
  }

  const sections = [];
  const lead = raw.slice(0, matches[0].index).trim();
  if (lead) sections.push({ heading: null, body: lead });

  matches.forEach((match, idx) => {
    const heading = match[1].trim().replace(/:$/, "");
    const bodyStart = match.index + match[0].length;
    const bodyEnd = idx + 1 < matches.length ? matches[idx + 1].index : raw.length;
    const body = raw.slice(bodyStart, bodyEnd).trim();
    sections.push({ heading, body });
  });

  return sections;
}

function renderKeyPointsHtml(text) {
  const sections = parseKeyPointsSections(text);
  if (!sections.length) return "";

  let headingNum = 0;
  return sections.map((section) => {
    if (section.heading) {
      headingNum += 1;
      const label = `${headingNum}. ${escapeHtml(section.heading)}:`;
      const bodyHtml = section.body
        ? `<div class="kp-body">${formatRichTextHTML(section.body)}</div>`
        : "";
      return `<div class="kp-block"><p class="kp-label">${label}</p>${bodyHtml}</div>`;
    }
    return `<div class="kp-block"><div class="kp-body">${formatRichTextHTML(section.body)}</div></div>`;
  }).join("");
}

function renderKeyPointsPreview() {
  const host = document.getElementById("keyPointsPreview");
  if (!host) return;
  host.innerHTML = renderKeyPointsHtml(getKeyPointsSourceText());
}

function estimateKeyPointsBodyHeight(body) {
  const lines = String(body || "").split("\n").length;
  return Math.min(2.8, Math.max(0.45, lines * 0.17));
}

function addKeyPointsToPptxSlide(slide, yStart = 0.8) {
  const sections = parseKeyPointsSections(getKeyPointsSourceText());
  let y = yStart;
  let headingNum = 0;

  sections.forEach((section) => {
    if (section.heading) {
      headingNum += 1;
      slide.addText(`${headingNum}. ${section.heading}:`, {
        x: 0.5,
        y,
        w: 9,
        h: 0.28,
        fontSize: CONFIG.REPORT_CONTENT_FONT_SIZE,
        bold: true,
        fontFace: "Calibri"
      });
      y += 0.26;
    }
    if (section.body) {
      const h = estimateKeyPointsBodyHeight(section.body);
      slide.addText(parseRichTextPptx(section.body), {
        x: 0.5,
        y,
        w: 9,
        h,
        fontSize: CONFIG.REPORT_CONTENT_FONT_SIZE,
        fontFace: "Calibri",
        valign: "top"
      });
      y += h + 0.12;
    }
  });
}

const EPS_BAR_COLORS = ["#4bc0c0", "#ffcd56", "#ff6384", "#ff9f40"];

const TOP_EPS_BAR_COLORS = [
  "#4bc0c0", "#ffcd56", "#ff6384", "#ff9f40", "#9966ff",
  "#c9cbcf", "#00b050", "#be2f2f", "#1e4f9a", "#8e5ea2"
];

const topEpsBarValuePlugin = {
  id: "topEpsBarValueLabels",
  afterDatasetsDraw(chart) {
    const { ctx, chartArea } = chart;
    if (!chartArea) return;
    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      meta.data.forEach((bar, index) => {
        const value = dataset.data[index];
        if (value == null) return;
        ctx.save();
        ctx.fillStyle = "#000";
        ctx.font = `bold ${CONFIG.REPORT_CONTENT_FONT_SIZE}px Calibri, Arial, sans-serif`;
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(String(Math.round(Number(value))), bar.x, Math.max(bar.y - 2, chartArea.top + 8));
        ctx.restore();
      });
    });
  }
};

function formatIndianNumber(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return "";
  const num = Number(raw.replace(/,/g, ""));
  if (!Number.isFinite(num)) return escapeHtml(raw);
  return num.toLocaleString("en-IN");
}

function formatAlertTableNumber(value) {
  const num = Number(String(value ?? "").replace(/,/g, ""));
  if (!Number.isFinite(num)) return "";
  if (num === 0) return "0";
  return num.toLocaleString("en-US");
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(String(e.target?.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function imageUrlToDataUrlViaCanvas(src) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      try {
        const c = document.createElement("canvas");
        c.width = img.naturalWidth;
        c.height = img.naturalHeight;
        const ctx = c.getContext("2d");
        if (!ctx) {
          resolve("");
          return;
        }
        ctx.drawImage(img, 0, 0);
        resolve(c.toDataURL("image/png"));
      } catch {
        resolve("");
      }
    };
    img.onerror = () => resolve("");
    img.src = src;
  });
}

async function ensureDefaultPuzzleDataUrl() {
  if (defaultPuzzleDataUrl) return defaultPuzzleDataUrl;
  const placeholderImg = document.querySelector("#puzzlePlaceholderLayout img");
  if (!placeholderImg) return "";
  const rawSrc = placeholderImg.getAttribute("src");
  if (!rawSrc) return "";
  const absUrl = new URL(rawSrc, window.location.href).toString();
  let dataUrl = "";
  try {
    const resp = await fetch(absUrl, { cache: "force-cache" });
    if (resp.ok) {
      const blob = await resp.blob();
      dataUrl = await blobToDataUrl(blob);
    }
  } catch (e) {
    console.warn("Puzzle fetch failed, trying canvas decode:", e);
  }
  if (!dataUrl) dataUrl = await imageUrlToDataUrlViaCanvas(absUrl);
  if (dataUrl) {
    defaultPuzzleDataUrl = dataUrl;
    placeholderImg.src = dataUrl;
  }
  return defaultPuzzleDataUrl;
}

function getEngagementGraphicDataUrl() {
  if (puzzleLogoDataUrl) return puzzleLogoDataUrl;
  if (defaultPuzzleDataUrl) return defaultPuzzleDataUrl;
  const placeholderImg = document.querySelector("#puzzlePlaceholderLayout img");
  return placeholderImg?.src || "";
}

/**
 * Parses **text** into an array of PptxGenJS text objects.
 */
function parseRichTextPptx(text, baseOptions = {}) {
  if (!text) return [];
  const parts = [];
  const monthReplaced = text.replace(/\[MONTH\]/g, getReportMonth());
  
  // Split by ** delimiters
  const segments = monthReplaced.split(/\*\*/);
  
  segments.forEach((segment, idx) => {
    if (segment === "") return;
    const isBold = idx % 2 !== 0; // Every second segment is bold
    parts.push({
      text: segment,
      options: { ...baseOptions, bold: isBold || baseOptions.bold }
    });
  });
  
  return parts;
}

// Internal field keys for month slots (epsJan, slaHighJanCount, etc.)
const MONTH_SLOT_KEYS = ["Jan", "Feb", "Mar", "Apr"];

function parseDateRangeEnd(str) {
  const endPart = String(str || "").split(/\s+to\s+/i).pop().trim();
  const match = endPart.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (!match) return null;
  const monthIdx = Number(match[2]) - 1;
  const year = Number(match[3]);
  if (monthIdx < 0 || monthIdx > 11) return null;
  return { monthIdx, year };
}

function getReportMonthColumnLabels() {
  const custom = getValue("monthColumns").trim();
  if (custom) {
    const labels = custom.split(",").map((part) => part.trim()).filter(Boolean);
    if (labels.length === 4) return labels;
  }

  const end = parseDateRangeEnd(getValue("dateRange"));
  if (end) {
    const labels = [];
    for (let offset = 3; offset >= 0; offset -= 1) {
      const date = new Date(end.year, end.monthIdx - offset, 1);
      labels.push(date.toLocaleString("en-US", { month: "short" }).replace(/\.$/, ""));
    }
    return labels;
  }

  return ["Jan", "Feb", "Mar", "Apr"];
}

function syncMonthColumnLabels() {
  const labels = getReportMonthColumnLabels();

  MONTH_SLOT_KEYS.forEach((slotKey, index) => {
    const label = labels[index] || slotKey;
    document.querySelectorAll(`[data-month-slot="${slotKey}"]`).forEach((el) => {
      el.textContent = label;
    });
    document.querySelectorAll(`[data-month-index="${index}"]`).forEach((el) => {
      const pctSuffix = el.hasAttribute("data-month-pct") ? " %" : "";
      el.textContent = `${label}${pctSuffix}`;
    });
  });

  const csvHint = document.getElementById("monthColumnsCsvHint");
  if (csvHint) {
    csvHint.textContent = labels.join(",");
  }
}

// Parses a DD-MM-YYYY date string and returns a locale month string, or null
function parseDateMatch(str) {
  const match = str.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (!match) return null;
  const monthIdx = Number(match[2]) - 1;
  const year = Number(match[3]);
  if (monthIdx < 0 || monthIdx > 11) return null;
  return new Date(year, monthIdx, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
}

// Removes characters unsafe in filenames
function sanitizeFilename(str) {
  return str.replace(/[^a-zA-Z0-9_\-]/g, "_").replace(/_+/g, "_");
}

function getApraavaSections() {
  return typeof TEMPLATE_REPORT_SECTIONS !== "undefined" ? TEMPLATE_REPORT_SECTIONS.apraava : null;
}

function isTemplateSectionActive(slideId) {
  const sections = getApraavaSections();
  if (sections === null) return true;
  return sections.includes(slideId);
}

function isExportSectionIncluded(slideId) {
  return isTemplateSectionActive(slideId);
}

function getExportSectionOrder() {
  const sections = getApraavaSections();
  if (sections !== null) return sections;
  return [
    "cover", "revision", "engagement", "alert-quad", "trend", "pot-incidents", "eps", "top-eps",
    "risks", "sla-monthly", "inventory", "key-points", "contact"
  ];
}
function getReportMonth() {
  const manual = ReportMonthUtils.resolveManualReportMonth(
    getValue("reportMonthLabel"),
    getValue("dateRange"),
    getValue("submittedOn")
  );
  if (manual) return manual;
  return parseDateMatch(getValue("dateRange"))
    || parseDateMatch(getValue("submittedOn"))
    || new Date().toLocaleString("en-US", { month: "long", year: "numeric" });
}

function getPptLayoutKey() {
  return PPT_LAYOUT_KEY;
}

function setDynamicPrintPageSize() {
  const { w, h } = PPT_SLIDE_IN;
  const styleId = "dynamicPrintPageSizeStyle";
  let styleEl = document.getElementById(styleId);
  if (!styleEl) {
    styleEl = document.createElement("style");
    styleEl.id = styleId;
    document.head.appendChild(styleEl);
  }
  styleEl.textContent = `@media print { @page { size: ${w}in ${h}in; margin: 0; } }`;
}

function getChartMonthLabel() {
  const labels = getReportMonthColumnLabels();
  return labels[labels.length - 1] || getReportMonth();
}

async function waitForChartsReady(delayMs = 800) {
  const charts = [
    trendChart,
    potIncidentsChart,
    epsTrendChart,
    topEpsChart,
    slaPctChart,
    ...apraavaQuadCharts
  ].filter(Boolean);
  renderSlaPctChart();
  window.dispatchEvent(new Event("resize"));
  charts.forEach((chart) => {
    try {
      chart.resize();
      chart.update("none");
    } catch (e) {
      // ignore chart resize errors during export prep
    }
  });
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  if (document.fonts && document.fonts.ready) {
    try {
      await document.fonts.ready;
    } catch (e) {
      // ignore font readiness errors
    }
  }
  await new Promise((r) => setTimeout(r, delayMs));
}

function beginExportVisibility() {
  const saved = [];
  document.querySelectorAll("[data-slide]").forEach((el) => {
    saved.push({ el, display: el.style.display });
    const slideId = el.getAttribute("data-slide");
    el.style.display = isExportSectionIncluded(slideId) ? "" : "none";
  });
  document.body.classList.add("is-exporting");
  return () => {
    document.body.classList.remove("is-exporting");
    saved.forEach(({ el, display }) => {
      el.style.display = display;
    });
  };
}

async function prepareReportForExport() {
  await ensureDefaultPuzzleDataUrl();
  const restoreVisibility = beginExportVisibility();
  applyData();
  await waitForChartsReady();
  return restoreVisibility;
}

function finishExportRestore() {
  if (exportRestoreCallback) {
    exportRestoreCallback();
    exportRestoreCallback = null;
    applyData();
  }
}

/** Replace Chart.js canvases with PNG images so browser print/PDF includes charts. */
function bakeChartsForPrint() {
  const baked = [];
  document.querySelectorAll("#reportRoot canvas").forEach((canvas) => {
    try {
      const rect = canvas.getBoundingClientRect();
      if (!rect.width || !rect.height) return;
      const img = document.createElement("img");
      img.src = canvas.toDataURL("image/png", 1.0);
      img.alt = "";
      img.className = "chart-print-bake";
      img.style.cssText = "display:block;width:100%;height:100%;object-fit:contain;";
      canvas.style.display = "none";
      canvas.insertAdjacentElement("afterend", img);
      baked.push({ canvas, img });
    } catch (e) {
      console.warn("Chart bake skipped:", e);
    }
  });
  return () => {
    baked.forEach(({ canvas, img }) => {
      canvas.style.display = "";
      img.remove();
    });
  };
}

function parseCsvCells(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === "," && !inQuotes) {
      out.push(cur.trim());
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur.trim());
  return out;
}

function normalizeKey(text) {
  return String(text || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function buildRiskTableBodyHtml(chunk) {
  if (!chunk.length) {
    return `<tr><td colspan="6">No risk rows found. Upload Risk CSV.</td></tr>`;
  }
  return chunk
    .map(
      (r) => `<tr>
  <td class="risk-col-sno">${escapeHtml(r.sno)}</td>
  <td class="risk-col-attack">${escapeHtml(r.attackType)}</td>
  <td class="risk-col-scenario">${escapeHtml(r.riskScenario)}</td>
  <td class="risk-col-cia">${escapeHtml(r.ciaTriad)}</td>
  <td class="risk-col-impact">${escapeHtml(r.businessImpact)}</td>
  <td class="risk-col-rating">${escapeHtml(r.riskRating)}</td>
</tr>`
    )
    .join("");
}

function getRiskSlideElements() {
  const host = document.getElementById("riskSlidesContainer");
  if (!host) return [];
  return [...host.querySelectorAll(":scope > .risk-slide")];
}

function appendRiskSlide(host, options) {
  const wrap = document.createElement("div");
  wrap.innerHTML = buildRiskSlideHtml(options).trim();
  const section = wrap.firstElementChild;
  if (section) host.appendChild(section);
  return section;
}

function buildRiskSlideHtml({ title, chunk, narrativeHtml = "", includeNarrative = false }) {
  return `
    <section class="page page-with-footer risk-slide risk-slide-compact" data-slide="risks" data-risk-rows="${chunk.length}">
      <div class="page-header-right">
        <div class="report-logo-slot">
          <img class="nav-sns-logo" alt="" />
        </div>
      </div>
      <h2 class="slide-title revision-title risk-slide-title">${escapeHtml(title)}</h2>
      ${includeNarrative ? `<div class="narrative slide-narrative risk-slide-narrative" id="riskNarrativePreview">${narrativeHtml}</div>` : ""}
      <div class="risk-table-wrap">
        <table class="risk-table">
          <thead>
            <tr>
              <th class="risk-col-sno">S.No</th>
              <th class="risk-col-attack">Attack Type</th>
              <th class="risk-col-scenario">Risk Scenario</th>
              <th class="risk-col-cia">Type of Risk(s)<br>CIA Triad</th>
              <th class="risk-col-impact">Potential Business Impact(s)</th>
              <th class="risk-col-rating">Risk Rating</th>
            </tr>
          </thead>
          <tbody>${buildRiskTableBodyHtml(chunk)}</tbody>
        </table>
      </div>
      <div class="page-footer-bar"></div>
    </section>
  `;
}

function getRiskMeasureStage() {
  let stage = document.getElementById("riskMeasureStage");
  if (stage) return stage;
  stage = document.createElement("div");
  stage.id = "riskMeasureStage";
  stage.className = "report ppt-wide";
  stage.setAttribute("aria-hidden", "true");
  stage.style.cssText =
    "position:absolute;left:-10000px;top:0;width:960px;visibility:hidden;pointer-events:none;z-index:-1;";
  document.body.appendChild(stage);
  return stage;
}

function riskSlideTableOverflows(slide, tolerancePx = 2) {
  const table = slide.querySelector(".risk-table");
  const footer = slide.querySelector(".page-footer-bar");
  if (!table || !footer) return false;
  const budget = footer.offsetTop - table.offsetTop - tolerancePx;
  if (budget <= 0) return true;
  const need = Math.max(table.offsetHeight, table.scrollHeight);
  return need > budget;
}

function measureRiskRowsPerSlide(rows, startIndex, isFirstSlide, narrativeHtml) {
  const stage = getRiskMeasureStage();
  let fitCount = 0;
  const remaining = rows.length - startIndex;
  for (let tryCount = 1; tryCount <= remaining; tryCount += 1) {
    const chunk = rows.slice(startIndex, startIndex + tryCount);
    stage.innerHTML = buildRiskSlideHtml({
      title: isFirstSlide ? "Potential Incidents – Risks Mitigated" : "Contn.,",
      chunk,
      narrativeHtml,
      includeNarrative: isFirstSlide && Boolean(String(narrativeHtml || "").trim())
    });
    const slide = stage.querySelector(".risk-slide");
    if (!slide) continue;
    if (riskSlideTableOverflows(slide)) break;
    fitCount = tryCount;
  }
  return Math.max(1, fitCount || 1);
}

function getRiskLayoutOptions() {
  const maxRows = parseInt(getValue("riskMaxRowsPerSlide"), 10);
  const fontPx = parseFloat(getValue("riskTableFontSize"));
  return {
    maxRowsPerSlide: Number.isFinite(maxRows) && maxRows > 0 ? maxRows : 0,
    tableFontSizePx: Number.isFinite(fontPx) && fontPx > 0 ? fontPx : 0
  };
}

function applyRiskTableFontSizeToSlide(slide, layoutOptions) {
  const fontPx = layoutOptions?.tableFontSizePx;
  if (!fontPx || !slide) return;
  slide.querySelectorAll(".risk-table").forEach((table) => {
    table.style.fontSize = `${fontPx}px`;
  });
}

function riskChunkFitsOnSlide(chunk, isFirstSlide, narrativeHtml, includeNarrative, layoutOptions) {
  const stage = getRiskMeasureStage();
  stage.innerHTML = `<div id="reportRoot" class="report ppt-wide">${buildRiskSlideHtml({
    title: isFirstSlide ? "Measure" : "Contn.,",
    chunk,
    narrativeHtml,
    includeNarrative
  })}</div>`;
  const slide = stage.querySelector(".risk-slide");
  if (slide) applyRiskTableFontSizeToSlide(slide, layoutOptions);
  return Boolean(slide) && !riskSlideTableOverflows(slide, 3);
}

const DEFAULT_RISK_ROWS_PER_SLIDE = 3;

function chunkRiskRowsForSlides(rows, narrativeText = "", layoutOptions = getRiskLayoutOptions()) {
  if (typeof RiskPagination !== "undefined" && RiskPagination.chunkRiskRowsForSlides) {
    return RiskPagination.chunkRiskRowsForSlides(
      rows,
      narrativeText,
      formatRichTextHTML,
      escapeHtml,
      "Contn.,",
      layoutOptions
    );
  }
  if (!rows.length) return [[]];
  const rowsPerSlideTarget =
    layoutOptions.maxRowsPerSlide > 0 ? layoutOptions.maxRowsPerSlide : DEFAULT_RISK_ROWS_PER_SLIDE;
  const hasNarrative = Boolean(String(narrativeText || "").trim());
  const narrativeHtmlFull = formatRichTextHTML(narrativeText);
  const chunks = [];
  let index = 0;
  while (index < rows.length) {
    const isFirst = chunks.length === 0;
    const includeNarrative = isFirst && hasNarrative;
    const narrativeHtml = isFirst ? narrativeHtmlFull : "";
    let best = 1;
    for (let tryCount = 1; tryCount <= rows.length - index; tryCount += 1) {
      const chunk = rows.slice(index, index + tryCount);
      if (riskChunkFitsOnSlide(chunk, isFirst, narrativeHtml, includeNarrative, layoutOptions)) {
        best = tryCount;
      } else {
        break;
      }
    }
    const take = Math.max(1, Math.min(best, rowsPerSlideTarget, rows.length - index));
    chunks.push(rows.slice(index, index + take));
    index += take;
  }
  if (chunks.length >= 2 && chunks[chunks.length - 1].length === 1) {
    const combined = [...chunks[chunks.length - 2], ...chunks[chunks.length - 1]];
    if (riskChunkFitsOnSlide(combined, false, "", false, layoutOptions)) {
      return [...chunks.slice(0, -2), combined];
    }
  }
  return chunks;
}

function initRiskTableUserControls() {
  const reflow = () => {
    renderRiskSlides();
  };
  ["riskMaxRowsPerSlide", "riskTableFontSize"].forEach((id) => {
    const el = document.getElementById(id);
    if (!el || el.dataset.riskLayoutBound === "1") return;
    el.dataset.riskLayoutBound = "1";
    el.addEventListener("change", reflow);
    el.addEventListener("input", reflow);
  });
  const btn = document.getElementById("riskReflowBtn");
  if (btn && btn.dataset.riskLayoutBound !== "1") {
    btn.dataset.riskLayoutBound = "1";
    btn.addEventListener("click", reflow);
  }
}

function resetRiskTableRowHeights(slide) {
  const wrap = slide?.querySelector(".risk-table-wrap");
  const table = slide?.querySelector(".risk-table");
  if (wrap) {
    wrap.classList.remove("risk-table-wrap--fill");
    wrap.style.height = "";
  }
  if (table) table.style.height = "";
  slide?.querySelectorAll(".risk-table tbody tr").forEach((tr) => {
    tr.style.height = "";
  });
}

function ensureRiskSlidesNoOverflow() {
  getRiskSlideElements().forEach(resetRiskTableRowHeights);
  let guard = 0;
  while (guard++ < 100 && rebalanceRiskSlidesTrimOverflow()) {
    /* move overflowing rows to next slide */
  }
  removeEmptyRiskSlides();
}

function riskSlidesOverflow() {
  return getRiskSlideElements().some(riskSlideTableOverflows);
}

function removeEmptyRiskSlides() {
  getRiskSlideElements().forEach((slide) => {
    const body = slide.querySelector(".risk-table tbody");
    if (!body?.querySelector("tr")) slide.remove();
  });
}

function rebalanceRiskSlidesFillSpace() {
  const host = document.getElementById("riskSlidesContainer");
  if (!host) return false;

  const slides = getRiskSlideElements();
  let changed = false;

  for (let i = 0; i < slides.length - 1; i += 1) {
    const slide = slides[i];
    const nextSlide = slides[i + 1];
    if (!nextSlide) break;

    while (true) {
      const nextBody = nextSlide.querySelector(".risk-table tbody");
      const nextRow = nextBody?.querySelector("tr");
      if (!nextRow) break;

      const currentBody = slide.querySelector(".risk-table tbody");
      if (!currentBody) break;

      currentBody.appendChild(nextRow);
      if (riskSlideTableOverflows(slide)) {
        nextBody.appendChild(nextRow);
        break;
      }
      changed = true;

      if (!nextBody.querySelector("tr")) {
        nextSlide.remove();
        slides.splice(i + 1, 1);
      }
    }
  }

  return changed;
}

function rebalanceRiskSlidesTrimOverflow() {
  const host = document.getElementById("riskSlidesContainer");
  if (!host) return false;

  const slides = getRiskSlideElements();
  let changed = false;

  for (let i = 0; i < slides.length; i += 1) {
    const slide = slides[i];
    while (riskSlideTableOverflows(slide)) {
      const body = slide.querySelector(".risk-table tbody");
      const lastRow = body?.lastElementChild;
      if (!lastRow) break;

      let nextSlide = slides[i + 1];
      if (!nextSlide) {
        nextSlide = document.createElement("section");
        nextSlide.className = "page page-with-footer risk-slide risk-slide-compact";
        nextSlide.setAttribute("data-slide", "risks");
        nextSlide.innerHTML = `
          <div class="page-header-right"><div class="report-logo-slot"><img class="nav-sns-logo" alt="SNS Logo" /></div></div>
          <h2 class="slide-title revision-title risk-slide-title">Contn.,</h2>
          <div class="risk-table-wrap">
            <table class="risk-table">
              <thead>
                <tr>
                  <th class="risk-col-sno">S.No</th>
                  <th class="risk-col-attack">Attack Type</th>
                  <th class="risk-col-scenario">Risk Scenario</th>
                  <th class="risk-col-cia">Type of Risk(s)<br>CIA Triad</th>
                  <th class="risk-col-impact">Potential Business Impact(s)</th>
                  <th class="risk-col-rating">Risk Rating</th>
                </tr>
              </thead>
              <tbody></tbody>
            </table>
          </div>
          <div class="page-footer-bar"></div>
        `;
        slide.after(nextSlide);
        slides.splice(i + 1, 0, nextSlide);
      }

      nextSlide.querySelector(".risk-table tbody").prepend(lastRow);
      changed = true;
    }
  }

  return changed;
}

function parseRiskCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];

  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some(
    (k) =>
      k.includes("attack") ||
      k.includes("scenario") ||
      k.includes("businessimpact") ||
      k.includes("riskrating") ||
      k.includes("ciatriad") ||
      k.includes("typeofrisk")
  );

  let dataRows = rows;
  let idx = {
    sno: -1,
    attack: 0,
    scenario: 1,
    typeOfRisk: -1,
    cia: 2,
    impact: 3,
    rating: 4
  };

  if (hasHeader) {
    dataRows = rows.slice(1);
    const findCol = (matchers, fallback) => {
      const pos = headerNorm.findIndex((k) => matchers.some((m) => k.includes(m)));
      return pos >= 0 ? pos : fallback;
    };
    idx = {
      sno: findCol(["sno", "serial"], -1),
      attack: findCol(["attacktype", "attack", "alerttype", "eventname"], 0),
      scenario: findCol(["riskscenario", "scenario", "description"], 1),
      typeOfRisk: findCol(["typeofrisks", "typeofrisk", "risktype"], -1),
      cia: findCol(["ciatriad", "cia"], 2),
      impact: findCol(["potentialbusinessimpact", "businessimpact", "impact", "consequence"], 3),
      rating: findCol(["riskrating", "rating", "severity"], 4)
    };
  } else if (rows[0].length >= 7) {
    idx = { sno: 0, attack: 1, scenario: 2, typeOfRisk: 3, cia: 4, impact: 5, rating: 6 };
  } else if (rows[0].length === 6) {
    idx = { sno: 0, attack: 1, scenario: 2, typeOfRisk: -1, cia: 3, impact: 4, rating: 5 };
  }

  return dataRows
    .map((r, i) => {
      const typeOfRisk = idx.typeOfRisk >= 0 ? (r[idx.typeOfRisk] || "") : "";
      const ciaOnly = idx.cia >= 0 ? (r[idx.cia] || "") : "";
      const ciaCombined = typeOfRisk && ciaOnly && typeOfRisk !== ciaOnly
        ? `${typeOfRisk}\n${ciaOnly}`
        : (typeOfRisk || ciaOnly);
      const snoRaw = idx.sno >= 0 ? (r[idx.sno] || "") : "";
      return {
        sno: snoRaw || String(i + 1),
        attackType: r[idx.attack] || "",
        riskScenario: r[idx.scenario] || "",
        ciaTriad: ciaCombined,
        businessImpact: r[idx.impact] || "",
        riskRating: r[idx.rating] || ""
      };
    })
    .filter(
      (r) =>
        r.attackType || r.riskScenario || r.ciaTriad || r.businessImpact || r.riskRating
    );
}

function makeColumnFinder(headerNorm, hasHeader) {
  return (matchers, fallback) => {
    if (!hasHeader) return fallback;
    const pos = headerNorm.findIndex((k) => matchers.some((m) => k.includes(m)));
    return pos >= 0 ? pos : fallback;
  };
}

function parseTrendCsvRows(csvText) {
  return ReportCharts.parseTrendCsvRows(csvText);
}

function parsePotIncidentsCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some(
    (k) =>
      k.includes("rule") ||
      k.includes("attack") ||
      k.includes("event") ||
      k.includes("high") ||
      k.includes("medium") ||
      k.includes("low")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = makeColumnFinder(headerNorm, hasHeader);
  const iRule = findCol(["rulename", "rule", "attacktype", "eventname", "incident"], 0);
  const iHigh = findCol(["high"], 1);
  const iMed = findCol(["medium", "med"], 2);
  const iLow = findCol(["low"], 3);

  return dataRows
    .map((r) => ({
      rule: r[iRule] || "",
      high: Math.max(0, parseFloat(r[iHigh]) || 0),
      medium: Math.max(0, parseFloat(r[iMed]) || 0),
      low: Math.max(0, parseFloat(r[iLow]) || 0)
    }))
    .filter((r) => r.rule);
}

function parseEpsTableCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some(
    (k) =>
      k.includes("reportingdevice") ||
      k.includes("eventname") ||
      k.includes("matchedevents") ||
      k.includes("jan")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = makeColumnFinder(headerNorm, hasHeader);
  const iSno = findCol(["sno", "serial"], -1);
  const iDevice = findCol(["reportingdevice", "device"], 0);
  const iEvent = findCol(["eventname", "event"], 1);
  const iJan = findCol(["jan"], 2);
  const iFeb = findCol(["feb"], 3);
  const iMar = findCol(["mar"], 4);
  const iApr = findCol(["apr"], 5);

  return dataRows
    .map((r, idx) => ({
      sno: (iSno >= 0 ? r[iSno] : null) || String(idx + 1),
      reportingDevice: r[iDevice] || "",
      eventName: r[iEvent] || "",
      jan: r[iJan] || "",
      feb: r[iFeb] || "",
      mar: r[iMar] || "",
      apr: r[iApr] || ""
    }))
    .filter((r) => r.reportingDevice || r.eventName || r.jan || r.feb || r.mar || r.apr);
}

function parseDeviceCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some(
    (k) => k.includes("device") || k.includes("host") || k.includes("eps") || k.includes("reporting")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = makeColumnFinder(headerNorm, hasHeader);
  const iSno = findCol(["sno", "serial"], -1);
  const iDevice = findCol(["reportingdevice", "devicename", "device", "host", "name"], 0);
  const iEps = findCol(["avgeps", "eps", "count", "value"], 1);

  return dataRows
    .map((r, idx) => ({
      sno: (iSno >= 0 ? r[iSno] : null) || String(idx + 1),
      device: r[iDevice] || "",
      eps: r[iEps] || ""
    }))
    .filter((r) => r.device)
    .slice(0, 10);
}

function parseInventoryCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some((k) => k.includes("device") || k.includes("count"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = makeColumnFinder(headerNorm, hasHeader);
  const iSno = findCol(["sno", "serial"], -1);
  const iDevice = findCol(["devicename", "device", "name"], 0);
  const iCount = findCol(["count", "total"], 1);

  return dataRows
    .map((r, idx) => ({
      sno: (iSno >= 0 ? r[iSno] : null) || String(idx + 1),
      deviceName: r[iDevice] || "",
      count: String(r[iCount] || "").replace(/,/g, "")
    }))
    .filter((r) => r.deviceName);
}

function parseCsvRows(csvText, columns) {
  return String(csvText || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, idx) => {
      const parts = parseCsvCells(line);
      const row = { sno: idx + 1 };
      columns.forEach((col, colIdx) => {
        row[col] = parts[colIdx] || "";
      });
      return row;
    });
}

function renderRiskSlides() {
  if (typeof RiskPagination === "undefined") {
    return Promise.resolve();
  }
  const copy = APRAAVA_COPY;
  const baseTitle = copy.riskTitle || "Potential Incidents – Risks Mitigated";
  return new Promise((resolve) => {
    RiskPagination.renderRiskSlides({
      rows: parseRiskCsvRows(getValue("riskCsv")),
      baseTitle,
      contdTitle: "Contn.,",
      narrativeText: getValue("riskNarrative"),
      escapeHtml,
      formatRichTextHTML,
      layoutOptions: getRiskLayoutOptions(),
      onComplete: (host) => {
        updateLogos();
        if (typeof initRiskTableColumnResize === "function") {
          initRiskTableColumnResize(host || document.getElementById("riskSlidesContainer"));
        }
        resolve();
      }
    });
  });
}


function renderTableRows(targetId, rows, renderRow) {
  const tableBody = document.getElementById(targetId);
  if (!tableBody) return;
  tableBody.innerHTML = "";
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = renderRow(row);
    tableBody.appendChild(tr);
  });
}

function updateLogos() {
  const snsEl = document.getElementById("snsLogoPreview");
  const clientEl = document.getElementById("clientLogoPreview");
  if (!snsEl || !clientEl) return;
  snsEl.style.display = snsLogoDataUrl ? "block" : "none";
  clientEl.style.display = clientLogoDataUrl ? "block" : "none";
  snsEl.src = snsLogoDataUrl;
  clientEl.src = clientLogoDataUrl;
  
  document.querySelectorAll(".nav-sns-logo").forEach(img => {
      img.src = snsLogoDataUrl;
      img.style.display = snsLogoDataUrl ? "block" : "none";
  });
  const lastSnsEl = document.getElementById("lastSnsLogoPreview");
  if (lastSnsEl) {
    lastSnsEl.src = snsLogoDataUrl;
    lastSnsEl.style.display = snsLogoDataUrl ? "block" : "none";
  }
  const puzzleGraphicEl = document.getElementById("puzzleGraphicPreview");
  const puzzleLayoutEl = document.getElementById("puzzlePlaceholderLayout");
  if (puzzleGraphicEl && puzzleLayoutEl) {
    if (puzzleLogoDataUrl) {
      puzzleGraphicEl.style.display = "block";
      puzzleGraphicEl.src = puzzleLogoDataUrl;
      puzzleLayoutEl.style.display = "none";
    } else {
      puzzleGraphicEl.style.display = "none";
      puzzleLayoutEl.style.display = "flex";
      const placeholderImg = puzzleLayoutEl.querySelector("img");
      if (placeholderImg && defaultPuzzleDataUrl) {
        placeholderImg.src = defaultPuzzleDataUrl;
      }
    }
  }
}

function parseChartPairs(csvText) {
  return csvText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [label, value] = line.split(",");
      return { label: (label || "").trim(), value: Number((value || "0").trim()) || 0 };
    });
}

function applyPremiumLayout() {
  document.querySelectorAll(".page").forEach(page => {
    // Skip cover, engagement, revision, and contact (contact has its own large logo)
    if (
      page.classList.contains("cover-hero") ||
      page.querySelector(".engagement-container") ||
      page.classList.contains("revision-page") ||
      page.classList.contains("slide-contact-summary") ||
      page.getAttribute("data-slide") === "contact"
    ) {
      page.querySelector(".page-header-right")?.remove();
      return;
    }
    
    // Upgrade to page container with footer padding
    if (!page.classList.contains("page-with-footer")) {
      page.classList.add("page-with-footer");
    }

    // Inject Logo Header
    if (!page.querySelector(".page-header-right")) {
      const header = document.createElement("div");
      header.className = "page-header-right";
      header.innerHTML = '<div class="report-logo-slot"><img class="nav-sns-logo" alt="SNS Logo" /></div>';
      page.insertBefore(header, page.firstChild);
    }

    // Inject Footer Bar
    if (!page.querySelector(".page-footer-bar")) {
      const footer = document.createElement("div");
      footer.className = "page-footer-bar";
      page.appendChild(footer);
    }

    // Standardize H2 Titles (skip centered slide titles)
    const h2 = page.querySelector("h2");
    if (h2 && !h2.classList.contains("revision-title") && !h2.classList.contains("slide-title-lg")) {
      h2.classList.add("revision-title");
    }
  });

  // Re-run updateLogos so the newly injected images get the logo URL
  updateLogos();
}

function applyDynamicTitles() {
  const copy = APRAAVA_COPY;
  document.querySelectorAll('[data-dynamic="summaryTitle"]').forEach((el) => {
    el.textContent = copy.summaryTitle;
  });
  document.querySelectorAll('[data-dynamic="summarySubtitle"]').forEach((el) => {
    el.textContent = copy.summarySubtitle;
  });
  document.querySelectorAll('[data-dynamic="trendTitle"]').forEach((el) => {
    el.textContent = copy.trendTitle;
  });
  document.querySelectorAll('[data-dynamic="riskTitle"]').forEach((el) => {
    el.textContent = copy.riskTitle || "Potential Incidents – Risks Mitigated";
  });
  document.querySelectorAll('[data-dynamic="potIncTitle"]').forEach((el) => {
    el.textContent = copy.potIncTitle || "Potential Incidents";
  });
}

function applyApraavaLayout() {
  configureSlideNavLabels();

  const root = document.getElementById("reportRoot");
  if (root) {
    const layoutClass = pptLayoutConfig.LAYOUT_16X9.reportClass;
    root.className = `report ${layoutClass}`;
  }

  const allowedSections = getApraavaSections();
  document.querySelectorAll("[data-slide]").forEach((el) => {
    if (allowedSections === null) {
      el.style.display = "";
      return;
    }
    const slideId = el.getAttribute("data-slide");
    el.style.display = allowedSections.includes(slideId) ? "" : "none";
  });

  if (!formSectionsOrdered) {
    formSectionsOrdered = true;
    reorderFormSectionsForTemplate();
  }
  if (slideNavController) slideNavController.refresh();
}

function reorderFormSectionsForTemplate() {
  const order = getApraavaSections();
  const nav = document.getElementById("slideNav");
  if (!order || !nav) return;

  const blocks = new Map();
  document.querySelectorAll(".panel [data-form-slide]").forEach((el) => {
    blocks.set(el.getAttribute("data-form-slide"), el);
  });

  let anchor = nav;
  order.forEach((slideId) => {
    if (slideId === "revision" || slideId === "cover" || slideId === "contact") return;
    const block = blocks.get(slideId);
    if (!block) return;
    anchor.insertAdjacentElement("afterend", block);
    anchor = block;
  });
}

function renderCharts() {
  const trendRows = parseTrendCsvRows(getValue("trendCsv"));
  const trendCtx = document.getElementById("trendChart");
  if (trendCtx) {
    if (trendChart) trendChart.destroy();
    const trendTitle = `Incident Summary – ${getReportMonth()}`;
    if (trendRows.length) {
      trendChart = new Chart(trendCtx, {
        type: "bar",
        data: {
          labels: trendRows.map((r) => r.date),
          datasets: ReportCharts.buildSeverityStackedDatasets(trendRows)
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          animation: false,
          plugins: ReportCharts.buildSeverityChartPlugins(trendTitle, {
            legendPosition: "bottom",
            fontSize: CONFIG.REPORT_CONTENT_FONT_SIZE
          }),
          scales: {
            x: ReportCharts.buildTrendChartXScaleOptions({ rowCount: trendRows.length }),
            y: ReportCharts.buildTrendChartYScaleOptions(trendRows)
          }
        }
      });
    }
  }

  const potIncRows = parsePotIncidentsCsvRows(getValue("potIncidentsCsv"));
  const potIncCtx = document.getElementById("potIncidentsChart");
  if (potIncCtx) {
    if (potIncidentsChart) potIncidentsChart.destroy();

    const potIncLabels = potIncRows.map((r) => r.rule);
    potIncidentsChart = new Chart(potIncCtx, {
      type: "bar",
      data: {
        labels: potIncLabels,
        datasets: ReportCharts.buildSeverityStackedDatasets(potIncRows)
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: ReportCharts.buildSeverityChartPlugins(`Incident Summary – ${getReportMonth()}`, {
          legendPosition: "bottom",
          fontSize: CONFIG.REPORT_CONTENT_FONT_SIZE
        }),
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { maxRotation: 45, minRotation: 45, font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE } } },
          y: { stacked: true, beginAtZero: true, suggestedMax: ReportCharts.getSeverityStackedMax(potIncRows) }
        }
      }
    });
  }

  const epsJan = parseFloat(getValue("epsJan")) || 0;
  const epsFeb = parseFloat(getValue("epsFeb")) || 0;
  const epsMar = parseFloat(getValue("epsMar")) || 0;
  const epsApr = parseFloat(getValue("epsApr")) || 0;
  const epsValues = [epsJan, epsFeb, epsMar, epsApr];

  const epsJanTable = document.getElementById("epsJanTable");
  const epsFebTable = document.getElementById("epsFebTable");
  const epsMarTable = document.getElementById("epsMarTable");
  const epsAprTable = document.getElementById("epsAprTable");
  [epsJanTable, epsFebTable, epsMarTable, epsAprTable].forEach((cell, idx) => {
    if (!cell) return;
    const raw = getValue(["epsJan", "epsFeb", "epsMar", "epsApr"][idx]);
    cell.textContent = isNaN(parseFloat(raw)) ? "" : String(epsValues[idx]);
  });

  const epsCtx = document.getElementById("epsTrendChart");
  if (epsCtx) {
    if (epsTrendChart) epsTrendChart.destroy();
    const epsMonthLabels = getReportMonthColumnLabels();
    const epsChartPlugins = { legend: { display: false } };
    ReportCharts.ensureDataLabelsPlugin();
    if (ReportCharts.isDataLabelsRegistered()) {
      epsChartPlugins.datalabels = ReportCharts.buildEpsTrendTotalValueDataLabelOptions({
        font: { weight: "bold", size: CONFIG.REPORT_CONTENT_FONT_SIZE }
      });
    }
    epsTrendChart = new Chart(epsCtx, {
      type: "bar",
      data: {
        labels: epsMonthLabels,
        datasets: [{
          label: "AVG EPS",
          data: epsValues,
          backgroundColor: EPS_BAR_COLORS,
          borderWidth: 0,
          barPercentage: 0.78,
          categoryPercentage: 0.86
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: epsChartPlugins,
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE } }
          },
          y: {
            beginAtZero: true,
            suggestedMax: 700,
            ticks: {
              font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE },
              stepSize: 100,
              maxTicksLimit: 8
            },
            grid: { color: "#d9d9d9", drawBorder: false }
          }
        }
      }
    });
  }

  const epsTableRows = parseEpsTableCsvRows(getValue("epsTableCsv"));
  renderTableRows(
    "epsTrendTableBody",
    epsTableRows,
    (row) => `
      <td>${row.sno}</td>
      <td class="eps-events-left">${escapeHtml(row.reportingDevice)}</td>
      <td class="eps-events-left">${escapeHtml(row.eventName)}</td>
      <td>${formatIndianNumber(row.jan)}</td>
      <td>${formatIndianNumber(row.feb)}</td>
      <td>${formatIndianNumber(row.mar)}</td>
      <td>${formatIndianNumber(row.apr || "")}</td>
    `
  );

  renderApraavaAlertQuad();

  const deviceRows = parseDeviceCsvRows(getValue("deviceCsv"));
  renderTopEpsChart(deviceRows);
}

function niceAxisMax(value) {
  const val = Math.max(Number(value) || 0, 1);
  if (val <= 10) return 10;
  const magnitude = Math.pow(10, Math.floor(Math.log10(val)));
  return Math.ceil(val / magnitude) * magnitude;
}

function parseAlertQuadCsv(csvText) {
  const dsHigh = [0, 0, 0, 0];
  const dsMed = [0, 0, 0, 0];
  const dsLow = [0, 0, 0, 0];

  String(csvText || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const parts = parseCsvCells(line);
      const sev = normalizeKey(parts[0] || "");
      if (!sev || sev === "severity" || sev === "jan" || sev.includes("month")) {
        return;
      }
      const vals = [
        Number(String(parts[1] || "").replace(/,/g, "")) || 0,
        Number(String(parts[2] || "").replace(/,/g, "")) || 0,
        Number(String(parts[3] || "").replace(/,/g, "")) || 0,
        Number(String(parts[4] || "").replace(/,/g, "")) || 0
      ];
      if (sev === "high" || sev.includes("high")) dsHigh.splice(0, 4, ...vals);
      if (sev === "medium" || sev.includes("medium") || sev === "med") dsMed.splice(0, 4, ...vals);
      if (sev === "low" || sev.includes("low")) dsLow.splice(0, 4, ...vals);
    });

  return { dsHigh, dsMed, dsLow };
}

function renderApraavaAlertQuad() {
  const aqKeys = [
    { csv: "overallAlertsCsv", chartId: "aqChart1", tableId: "aqTable1", narr: "overallAlertsNarrative", narrId: "aqNarrative1" },
    { csv: "truePositiveCsv", chartId: "aqChart2", tableId: "aqTable2", narr: "truePositiveNarrative", narrId: "aqNarrative2" },
    { csv: "falsePositiveCsv", chartId: "aqChart3", tableId: "aqTable3", narr: "falsePositiveNarrative", narrId: "aqNarrative3" },
    { csv: "apraavaIncidentsCsv", chartId: "aqChart4", tableId: "aqTable4", narr: "apraavaIncidentsNarrative", narrId: "aqNarrative4" }
  ];

  apraavaQuadCharts.forEach((chart) => chart && chart.destroy());
  apraavaQuadCharts = [];

  aqKeys.forEach((config) => {
    const narrEl = document.getElementById(config.narrId);
    if (narrEl) narrEl.innerHTML = formatRichTextHTML(getValue(config.narr));

    const { dsHigh, dsMed, dsLow } = parseAlertQuadCsv(getValue(config.csv));
    const axisMax = niceAxisMax(Math.max(...dsHigh, ...dsMed, ...dsLow));

    const tb = document.getElementById(config.tableId);
    if (tb) {
      tb.innerHTML = `
        <thead>
          <tr><th></th>${getReportMonthColumnLabels().map((label) => `<th>${escapeHtml(label)}</th>`).join("")}</tr>
        </thead>
        <tbody>
          <tr><td><span class="aq-sq high"></span> High</td><td>${formatAlertTableNumber(dsHigh[0])}</td><td>${formatAlertTableNumber(dsHigh[1])}</td><td>${formatAlertTableNumber(dsHigh[2])}</td><td>${formatAlertTableNumber(dsHigh[3])}</td></tr>
          <tr><td><span class="aq-sq med"></span> Medium</td><td>${formatAlertTableNumber(dsMed[0])}</td><td>${formatAlertTableNumber(dsMed[1])}</td><td>${formatAlertTableNumber(dsMed[2])}</td><td>${formatAlertTableNumber(dsMed[3])}</td></tr>
          <tr><td><span class="aq-sq low"></span> Low</td><td>${formatAlertTableNumber(dsLow[0])}</td><td>${formatAlertTableNumber(dsLow[1])}</td><td>${formatAlertTableNumber(dsLow[2])}</td><td>${formatAlertTableNumber(dsLow[3])}</td></tr>
        </tbody>
      `;
    }

    const cv = document.getElementById(config.chartId);
    if (!cv) return;

    apraavaQuadCharts.push(new Chart(cv, {
      type: "bar",
      data: {
        labels: getReportMonthColumnLabels(),
        datasets: [
          { label: "High", data: dsHigh, backgroundColor: "#ff0000", borderWidth: 0, barPercentage: 0.92, categoryPercentage: 0.82 },
          { label: "Medium", data: dsMed, backgroundColor: "#ffff00", borderWidth: 0, barPercentage: 0.92, categoryPercentage: 0.82 },
          { label: "Low", data: dsLow, backgroundColor: "#00b050", borderWidth: 0, barPercentage: 0.92, categoryPercentage: 0.82 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: false,
        plugins: {
          legend: { display: false },
          datalabels: { display: false }
        },
        scales: {
          x: {
            stacked: false,
            display: false,
            grid: { display: false }
          },
          y: {
            stacked: false,
            beginAtZero: true,
            suggestedMax: axisMax,
            ticks: {
              font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE },
              maxTicksLimit: 6
            },
            grid: { color: "#d9d9d9", drawBorder: false }
          }
        }
      }
    }));
  });
}

function ensureApraavaAlertQuadDefaults() {
  if (typeof APRAAVA_ALERT_QUAD_DEFAULTS === "undefined") return;
  Object.entries(APRAAVA_ALERT_QUAD_DEFAULTS).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
}

function ensureApraavaEpsDefaults() {
  if (typeof APRAAVA_EPS_DEFAULTS === "undefined") return;
  Object.entries(APRAAVA_EPS_DEFAULTS).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
}

function ensureApraavaTopEpsDefaults() {
  if (typeof APRAAVA_TOP_EPS_DEFAULTS === "undefined") return;
  Object.entries(APRAAVA_TOP_EPS_DEFAULTS).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
}

function ensureApraavaSlaDefaults() {
  if (typeof APRAAVA_SLA_DEFAULTS === "undefined") return;
  Object.entries(APRAAVA_SLA_DEFAULTS).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
}

function ensureApraavaInventoryDefaults() {
  if (typeof APRAAVA_INVENTORY_DEFAULTS === "undefined") return;
  Object.entries(APRAAVA_INVENTORY_DEFAULTS).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
}

function ensureApraavaSlideDefaults() {
  ensureApraavaAlertQuadDefaults();
  ensureApraavaEpsDefaults();
  ensureApraavaTopEpsDefaults();
  ensureApraavaSlaDefaults();
  ensureApraavaInventoryDefaults();
}

function formatInventoryNoteHtml(noteText) {
  const trimmed = String(noteText || "").trim();
  if (!trimmed) return "";
  const body = trimmed.replace(/^Note:\s*/i, "");
  return `<strong>Note:</strong> ${escapeHtml(body)}`;
}

function getSlaPctForChart(pctField, countField) {
  const count = parseInt(getValue(countField), 10) || 0;
  if (count <= 0) return null;
  const raw = String(getValue(pctField)).trim();
  if (!raw) return null;
  const pct = parseFloat(raw);
  return Number.isFinite(pct) ? pct : null;
}

function renderSlaPctChart() {
  const slaPctCtx = document.getElementById("slaPctChart");
  if (!slaPctCtx) return;

  if (slaPctChart) slaPctChart.destroy();

  const months = getReportMonthColumnLabels();
  const monthKeys = MONTH_SLOT_KEYS;
  const compact = document.body.classList.contains("is-exporting");
  const legendFontSize = compact ? 8 : CONFIG.REPORT_CONTENT_FONT_SIZE;

  slaPctChart = new Chart(slaPctCtx, {
    type: "bar",
    data: {
      labels: months,
      datasets: [
        {
          label: "S-1 (High) Within SLA (4 Hrs)",
          data: monthKeys.map((m) => getSlaPctForChart(`slaHigh${m}Pct`, `slaHigh${m}Count`)),
          backgroundColor: "#ff0000",
          borderWidth: 0,
          barPercentage: 0.82,
          categoryPercentage: 0.72
        },
        {
          label: "S-2 (Medium) Within SLA (8 Hrs)",
          data: monthKeys.map((m) => getSlaPctForChart(`slaMed${m}Pct`, `slaMed${m}Count`)),
          backgroundColor: "#ffff00",
          borderColor: "#d4c400",
          borderWidth: 1,
          barPercentage: 0.82,
          categoryPercentage: 0.72
        },
        {
          label: "S-3 (Low) Within SLA (24 Hrs)",
          data: monthKeys.map((m) => getSlaPctForChart(`slaLow${m}Pct`, `slaLow${m}Count`)),
          backgroundColor: "#00b050",
          borderWidth: 0,
          barPercentage: 0.82,
          categoryPercentage: 0.72
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: { padding: { top: 4, bottom: compact ? 10 : 6, left: 4, right: 4 } },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            font: { size: legendFontSize, weight: "bold" },
            padding: 6
          }
        },
        y: {
          beginAtZero: true,
          max: 120,
          ticks: {
            stepSize: 20,
            color: "#666666",
            font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE },
            callback(value) {
              return `${value}%`;
            }
          },
          grid: { color: "#bdbdbd", drawBorder: false }
        }
      },
      plugins: {
        legend: {
          position: "bottom",
          align: "center",
          labels: {
            boxWidth: compact ? 8 : 12,
            boxHeight: compact ? 8 : 12,
            padding: compact ? 4 : 10,
            color: "#666666",
            font: { size: legendFontSize }
          }
        },
        tooltip: { enabled: false },
        datalabels: { display: false }
      }
    }
  });

  requestAnimationFrame(() => {
    if (slaPctChart) {
      slaPctChart.resize();
      slaPctChart.update("none");
    }
  });
}

function renderTopEpsChart(deviceRows) {
  const topEpsCtx = document.getElementById("topEpsChart");
  if (!topEpsCtx) return;

  if (topEpsChart) topEpsChart.destroy();
  const labels = deviceRows.map((r) => r.device);
  const values = deviceRows.map((r) => parseFloat(r.eps) || 0);
  const colors = TOP_EPS_BAR_COLORS.slice(0, deviceRows.length);
  const dataMax = Math.max(...values, 0);
  const axisMax = dataMax > 140
    ? Math.ceil(dataMax / 20) * 20
    : 140;
  const compact = document.body.classList.contains("is-exporting");
  const legendFontSize = compact ? 8 : CONFIG.REPORT_CONTENT_FONT_SIZE;

  topEpsChart = new Chart(topEpsCtx, {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: "AVG EPS",
        data: values,
        backgroundColor: colors,
        borderWidth: 0,
        barPercentage: 0.78,
        categoryPercentage: 0.82
      }]
    },
    plugins: [topEpsBarValuePlugin],
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: { padding: { top: compact ? 8 : 14, bottom: compact ? 0 : 4, left: 2, right: 2 } },
      plugins: {
        legend: {
          display: deviceRows.length > 0,
          position: "bottom",
          labels: {
            boxWidth: compact ? 7 : 10,
            boxHeight: compact ? 7 : 10,
            padding: compact ? 2 : 5,
            font: { size: legendFontSize },
            generateLabels(chart) {
              const data = chart.data;
              return data.labels.map((label, i) => ({
                text: label,
                fillStyle: data.datasets[0].backgroundColor[i],
                strokeStyle: data.datasets[0].backgroundColor[i],
                lineWidth: 0,
                hidden: false,
                index: i
              }));
            }
          }
        },
        tooltip: { enabled: false },
        datalabels: { display: false }
      },
      scales: {
        x: {
          display: true,
          title: {
            display: true,
            text: "HOSTS",
            font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE, weight: "bold" }
          },
          ticks: { display: false },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          max: axisMax,
          suggestedMax: axisMax,
          title: {
            display: true,
            text: "EPS",
            font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE }
          },
          ticks: {
            stepSize: 20,
            font: { size: CONFIG.REPORT_CONTENT_FONT_SIZE },
            maxTicksLimit: 8
          },
          grid: { color: "#bdbdbd", drawBorder: false }
        }
      }
    }
  });

  requestAnimationFrame(() => {
    if (topEpsChart) {
      topEpsChart.resize();
      topEpsChart.update("none");
    }
  });
}

function applyData() {
  applyApraavaLayout();
  applyDynamicTitles();
  syncMonthColumnLabels();

  Object.entries(fields).forEach(([key, inputId]) => {
    const value = getValue(inputId);
    document.querySelectorAll(`[data-field="${key}"]`).forEach((el) => {
      el.textContent = value;
    });
  });
  document.querySelectorAll('[data-field="reviewedBy"]').forEach((el) => {
    el.textContent = STATIC_REVIEWED_BY;
  });
  document.querySelectorAll('[data-field="approvedBy"]').forEach((el) => {
    el.textContent = STATIC_APPROVED_BY;
  });
  document.querySelectorAll('[data-field="month"]').forEach((el) => {
    el.textContent = getReportMonth();
  });

  const execEl = document.getElementById("executiveSummaryPreview");
  if (execEl) {
    const raw = getValue("executiveSummary");
    execEl.innerHTML = raw.split("\n")
      .filter(p => p.trim())
      .map(p => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }



  const trendNarrEl = document.getElementById("trendNarrativePreview");
  if (trendNarrEl) trendNarrEl.innerHTML = formatRichTextHTML(getValue("trendNarrative"));

  const potIncNarrEl = document.getElementById("potIncidentsNarrativePreview");
  if (potIncNarrEl) potIncNarrEl.innerHTML = formatRichTextHTML(getValue("potIncidentsNarrative"));

  const riskNarrEl = document.getElementById("riskNarrativePreview");
  if (riskNarrEl && !document.getElementById("riskSlidesContainer")) {
    riskNarrEl.innerHTML = formatRichTextHTML(getValue("riskNarrative"));
  }

  renderKeyPointsPreview();

  const deviceRows = parseDeviceCsvRows(getValue("deviceCsv"));
  renderTableRows(
    "deviceTableBody",
    deviceRows,
    (row) => `
      <td>${row.sno}</td>
      <td class="top-eps-device">${escapeHtml(row.device)}</td>
      <td>${escapeHtml(row.eps)}</td>
    `
  );

  if (isTemplateSectionActive("risks")) {
    renderRiskSlides();
  } else {
    const riskHost = document.getElementById("riskSlidesContainer");
    if (riskHost) riskHost.innerHTML = "";
  }

  const slaTb = document.getElementById("slaTableBody");
  if (slaTb) {
    slaTb.innerHTML = `
      <tr>
        <td>S-1 (High)</td>
        <td>≤15 Min</td>
        <td>4 Hours</td>
        <td>${escapeHtml(getValue("slaHighJanCount"))}</td>
        <td>${escapeHtml(getValue("slaHighFebCount"))}</td>
        <td>${escapeHtml(getValue("slaHighMarCount"))}</td>
        <td>${escapeHtml(getValue("slaHighAprCount"))}</td>
      </tr>
      <tr>
        <td>S-2 (Medium)</td>
        <td>≤15 Min</td>
        <td>8 Hours</td>
        <td>${escapeHtml(getValue("slaMedJanCount"))}</td>
        <td>${escapeHtml(getValue("slaMedFebCount"))}</td>
        <td>${escapeHtml(getValue("slaMedMarCount"))}</td>
        <td>${escapeHtml(getValue("slaMedAprCount"))}</td>
      </tr>
      <tr>
        <td>S-3 (Low)</td>
        <td>≤30 Min</td>
        <td>24 Hours</td>
        <td>${escapeHtml(getValue("slaLowJanCount"))}</td>
        <td>${escapeHtml(getValue("slaLowFebCount"))}</td>
        <td>${escapeHtml(getValue("slaLowMarCount"))}</td>
        <td>${escapeHtml(getValue("slaLowAprCount"))}</td>
      </tr>
    `;
    renderSlaPctChart();
  }

  const inventoryRows = parseInventoryCsvRows(getValue("inventoryCsv"));
  let totalCount = 0;

  const inventoryHtml = inventoryRows.map((row) => {
    const cnt = parseInt(row.count, 10) || 0;
    totalCount += cnt;
    return `<tr>
      <td>${escapeHtml(row.sno)}</td>
      <td>${escapeHtml(row.deviceName)}</td>
      <td>${cnt}</td>
    </tr>`;
  }).join("");

  const invTb = document.getElementById("inventoryTableBody");
  if (invTb) invTb.innerHTML = inventoryHtml;

  const invTf = document.getElementById("inventoryTableFooter");
  if (invTf) {
    invTf.innerHTML = `
      <tr>
        <td colspan="2">Total</td>
        <td>${totalCount}</td>
      </tr>
    `;
  }

  const invNotePreview = document.getElementById("inventoryNotePreview");
  if (invNotePreview) {
    invNotePreview.innerHTML = formatInventoryNoteHtml(getValue("inventoryNote"));
  }



  updateLogos();
  applyPremiumLayout();
  renderCharts();
  initTableResizing();
  apraavaQuadCharts.forEach((chart) => {
    try {
      chart.resize();
      chart.update("none");
    } catch (e) {
      // ignore resize errors
    }
  });
  if (epsTrendChart) {
    try {
      epsTrendChart.resize();
      epsTrendChart.update("none");
    } catch (e) {
      // ignore resize errors
    }
  }
  if (topEpsChart) {
    try {
      topEpsChart.resize();
      topEpsChart.update("none");
    } catch (e) {
      // ignore resize errors
    }
  }
  if (slaPctChart) {
    try {
      slaPctChart.resize();
      slaPctChart.update("none");
    } catch (e) {
      // ignore resize errors
    }
  }

  if (typeof reapplyTextStyleOverrides === "function") {
    reapplyTextStyleOverrides();
  }
}

async function exportPdf() {
  const printBtn = document.getElementById("printBtn");
  if (printBtn?.disabled) return;
  const originalLabel = printBtn ? printBtn.textContent : "";
  if (printBtn) {
    printBtn.disabled = true;
    printBtn.textContent = "Preparing…";
  }
  try {
    finishExportRestore();
    const restoreVisibility = await prepareReportForExport();
    const restoreCharts = bakeChartsForPrint();
    exportRestoreCallback = () => {
      restoreCharts();
      restoreVisibility();
    };
    setDynamicPrintPageSize();
    const reportRoot = document.getElementById("reportRoot");
    if (reportRoot) reportRoot.scrollIntoView({ block: "start" });
    await new Promise((r) => requestAnimationFrame(r));
    window.print();
  } catch (err) {
    console.error("PDF export failed:", err);
    finishExportRestore();
    if (printBtn) {
      printBtn.disabled = false;
      printBtn.textContent = originalLabel || "Export PDF";
    }
    alert(`PDF export failed: ${err?.message || err}`);
  }
}

function canvasToPptxData(canvas) {
  if (!canvas) return null;
  return canvas.toDataURL("image/png");
}

function exportPptxCover(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Monthly SOC Report");
  slide.addText("MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", {
    x: 0.5,
    y: 0.65,
    w: 9,
    h: 0.35,
    fontSize: 12,
    color: CONFIG.PPTX_BLUE,
    bold: true
  });
  slide.addText(`${getReportMonth()}`, { x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 24, bold: true, color: CONFIG.PPTX_NAVY });
  slide.addText(`for ${getValue("customerName")}`, { x: 0.5, y: 1.65, w: 9, h: 0.4, fontSize: 16 });
  slide.addText(
    [
      `Version 1.0`,
      `Dated ${getValue("dateRange")}`,
      `Prepared By ${getValue("preparedBy")}`,
      `Reviewed By ${STATIC_REVIEWED_BY}`,
      `Approved By ${STATIC_APPROVED_BY}`,
      `Submitted On ${getValue("submittedOn")}`
    ].join("\n"),
    { x: 0.5, y: 2.3, w: 5.5, h: 2, fontSize: 11 }
  );
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 6.8, y: 0.65, w: 1.2, h: 0.55 });
  }
  if (clientLogoDataUrl) {
    slide.addImage({ data: clientLogoDataUrl, x: 8.1, y: 0.65, w: 1.2, h: 0.55 });
  }
  addPptxFooter(slide);
}

function exportPptxRevision(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Document Revision History");
  slide.addText(
    getValue("documentRevisionHistory")
      .split("\n")
      .filter(Boolean)
      .map((l) => `• ${l}`)
      .join("\n"),
    { x: 0.5, y: 0.65, w: 9, h: 3, fontSize: 11 }
  );
  addPptxFooter(slide);
}

function exportPptxEngagement(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "The Engagement");
  const introText = `${getValue("customerName")} has engaged with SNS to monitor and review the entity's security.\n\n`;
  const execBody = getValue("executiveSummary");
  slide.addText([...parseRichTextPptx(introText), ...parseRichTextPptx(execBody)], {
    x: 0.5,
    y: 0.65,
    w: 5.6,
    h: 4.2,
    fontSize: 11,
    valign: "top"
  });
  const engagementGraphic = getEngagementGraphicDataUrl();
  if (engagementGraphic) {
    slide.addImage({ data: engagementGraphic, x: 6.15, y: 0.65, w: 3.35, h: 4.2 });
  }
  addPptxFooter(slide);
}

function exportPptxAlertQuad(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Potential Incidents and Alert Summary");
  const quadIds = ["aqChart1", "aqChart2", "aqChart3", "aqChart4"];
  apraavaQuadCharts.forEach((chart) => {
    if (!chart) return;
    if (!chart.options.plugins) chart.options.plugins = {};
    chart.options.plugins.datalabels = { display: false };
    chart.update("none");
  });
  quadIds.forEach((id, idx) => {
    const cImg = canvasToPptxData(document.getElementById(id));
    if (cImg) {
      const isRight = idx % 2 !== 0;
      const isBot = Math.floor(idx / 2) > 0;
      slide.addImage({ data: cImg, x: isRight ? 5.0 : 0.5, y: isBot ? 3.0 : 0.8, w: 4.4, h: 2.0 });
    }
  });
  addPptxFooter(slide);
}

function exportPptxTotPot(pptx, copy) {
  const highVal = Math.max(0, parseFloat(getValue("totPotIncHigh")) || 0);
  const medVal = Math.max(0, parseFloat(getValue("totPotIncMedium")) || 0);
  const monthLabel = getChartMonthLabel();
  const maxAlert = Math.max(highVal, medVal, 1);
  const yAxisMax = Math.max(20, Math.ceil((maxAlert * 1.1) / 2) * 2);
  const barChartType = pptx.ChartType ? pptx.ChartType.bar : "bar";
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.summaryTitle || "Total Potential Incidents");
  slide.addChart(
    barChartType,
    [
      { name: "High", labels: [monthLabel], values: [highVal] },
      { name: "Medium", labels: [monthLabel], values: [medVal] }
    ],
    {
      x: 0.8,
      y: 0.75,
      w: 8.4,
      h: 3.5,
      barDir: "col",
      barGrouping: "clustered",
      chartColors: ["FF0000", "FFFF00"],
      showLegend: true,
      legendPos: "b",
      valAxisMaxVal: yAxisMax,
      valAxisMinVal: 0,
      valAxisMajorUnit: 2,
      showDataTable: true,
      dataTableFontSize: 9,
      catAxisLabelFontSize: 10,
      valAxisLabelFontSize: 10
    }
  );
  addPptxFooter(slide);
}

function exportPptxTrend(pptx, copy) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.trendTitle);
  const trendImg = canvasToPptxData(document.getElementById("trendChart"));
  if (trendImg) {
    slide.addImage({ data: trendImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), { x: 0.5, y: 4.3, w: 9, h: 0.3, fontSize: 13, bold: true });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), { x: 0.5, y: 4.6, w: 9, h: 0.7, fontSize: 11 });
  addPptxFooter(slide);
}

function exportPptxRuleSeverity(pptx, copy) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.ruleSeverityTitle || "Rule-Based Severity Categories");
  const ruleImg = canvasToPptxData(document.getElementById("potIncidentsChart"));
  if (ruleImg) {
    slide.addImage({ data: ruleImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  addPptxFooter(slide);
}

function exportPptxPotIncidents(pptx, copy) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.potIncTitle || "Potential Incidents");
  const potIncImg = canvasToPptxData(document.getElementById("potIncidentsChart"));
  if (potIncImg) {
    slide.addImage({ data: potIncImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("potIncidentsNarrative")), { x: 0.5, y: 4.4, w: 9, h: 0.8, fontSize: 11 });
  addPptxFooter(slide);
}

function exportPptxEps(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "EPS Trend Plot");
  const epsImg = canvasToPptxData(document.getElementById("epsTrendChart"));
  if (epsImg) {
    slide.addImage({ data: epsImg, x: 0.5, y: 0.8, w: 3.5, h: 4.0 });
  }
  const epsTblRows = parseEpsTableCsvRows(getValue("epsTableCsv"));
  if (epsTblRows.length > 0) {
    const monthLabels = getReportMonthColumnLabels();
    const tableData = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Event Name", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: monthLabels[0], options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: monthLabels[1], options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: monthLabels[2], options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: monthLabels[3], options: { bold: true, fill: { color: "F3F6FB" } } }
      ],
      ...epsTblRows.map((r) => [String(r.sno), r.reportingDevice, r.eventName, r.jan, r.feb, r.mar, r.apr || ""])
    ];
    slide.addTable(tableData, { x: 4.0, y: 0.8, w: 5.6, fontSize: 7, colW: [0.4, 1.1, 1.2, 0.65, 0.65, 0.65, 0.65] });
  }
  addPptxFooter(slide);
}

function exportPptxTopEps(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Top 10 Devices Contributing To Highest EPS");
  const topEpsGap = 0.12;
  const topEpsChartW = PPT_BODY.width * 0.56;
  const topEpsTableW = PPT_BODY.width - topEpsChartW - topEpsGap;
  const topEpsChartX = PPT_BODY.left;
  const topEpsTableX = topEpsChartX + topEpsChartW + topEpsGap;
  const devImg = canvasToPptxData(document.getElementById("topEpsChart"));
  if (devImg) {
    slide.addImage({
      data: devImg,
      x: topEpsChartX,
      y: PPT_BODY.top,
      w: topEpsChartW,
      h: PPT_BODY.height
    });
  }
  const devices = parseDeviceCsvRows(getValue("deviceCsv"));
  if (devices.length > 0) {
    const devRows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "8EEDE4" }, align: "center" } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "8EEDE4" }, align: "center" } },
        { text: "AVG EPS", options: { bold: true, fill: { color: "8EEDE4" }, align: "center" } }
      ],
      ...devices.map((d) => [
        { text: String(d.sno), options: { align: "center" } },
        { text: d.device, options: { align: "left" } },
        { text: d.eps, options: { align: "center" } }
      ])
    ];
    const headerRowH = 0.28;
    slide.addTable(devRows, {
      x: topEpsTableX,
      y: PPT_BODY.top,
      w: topEpsTableW,
      h: PPT_BODY.height,
      rowH: [headerRowH, ...Array(devices.length).fill((PPT_BODY.height - headerRowH) / devices.length)],
      fontSize: 8,
      colW: [0.48, topEpsTableW - 1.18, 0.7],
      align: "center",
      valign: "middle",
      border: { type: "solid", color: "00B0BA", pt: 0.5 }
    });
  }
  addPptxFooter(slide);
}

function exportPptxRisks(pptx, copy) {
  const risks = parseRiskCsvRows(getValue("riskCsv"));
  const riskTitle = copy.riskTitle || "Potential Incidents – Risks Mitigated";
  if (risks.length === 0) {
    const slide = pptx.addSlide();
    addPptxTitleBar(slide, riskTitle);
    slide.addText("No risk rows in CSV. Add rows under Potential Incidents CSV.", {
      x: 0.5,
      y: 0.65,
      w: 9,
      h: 0.5,
      fontSize: 11
    });
    addPptxFooter(slide);
    return;
  }
  const riskChunks = chunkRiskRowsForSlides(risks, getValue("riskNarrative"), getRiskLayoutOptions());
  riskChunks.forEach((chunk, i) => {
    const slide = pptx.addSlide();
    addPptxTitleBar(slide, i === 0 ? riskTitle : "Contn.,");
    if (i === 0) {
      slide.addText(parseRichTextPptx(getValue("riskNarrative")), { x: 0.35, y: 0.65, w: 9.3, h: 0.8, fontSize: 10 });
    }
    const rows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Attack Type", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Scenario", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Type of Risk(s)\nCIA Triad", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Potential Business Impact(s)", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Rating", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } }
      ],
      ...chunk.map((r) => [
        { text: String(r.sno), options: { align: "center", fontFace: "Calibri" } },
        { text: r.attackType, options: { align: "center", fontFace: "Calibri", valign: "middle" } },
        { text: r.riskScenario, options: { align: "center", fontFace: "Calibri", valign: "middle" } },
        { text: r.ciaTriad, options: { align: "center", fontFace: "Calibri", valign: "middle" } },
        { text: r.businessImpact, options: { align: "center", fontFace: "Calibri", valign: "middle" } },
        { text: r.riskRating, options: { align: "center", fontFace: "Calibri", valign: "middle" } }
      ])
    ];
    const tableY = i === 0 && String(getValue("riskNarrative") || "").trim() ? 1.5 : 1.15;
    const riskFont =
      getRiskLayoutOptions().tableFontSizePx > 0
        ? Math.min(18, Math.max(8, getRiskLayoutOptions().tableFontSizePx * 0.64))
        : 9;
    slide.addTable(rows, {
      x: 0.35,
      y: tableY,
      w: 9.3,
      fontSize: riskFont,
      border: { type: "solid", color: "7F7F7F", pt: 0.5 },
      colW: [0.38, 1.35, 1.95, 1.15, 3.35, 0.92],
      valign: "middle"
    });
    addPptxFooter(slide);
  });
}

function exportPptxSlaMonthly(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Response Time SLA");
  const slaMonthLabels = getReportMonthColumnLabels();
  const slaRows = [
    [
      { text: "Severity Level", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "SLA Response", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "SLA Resolution", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: slaMonthLabels[0], options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: slaMonthLabels[1], options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: slaMonthLabels[2], options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: slaMonthLabels[3], options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }
    ],
    [
      "S-1 (High)", "≤15 Min", "4 Hours",
      getValue("slaHighJanCount"), getValue("slaHighFebCount"), getValue("slaHighMarCount"), getValue("slaHighAprCount")
    ],
    [
      "S-2 (Medium)", "≤15 Min", "8 Hours",
      getValue("slaMedJanCount"), getValue("slaMedFebCount"), getValue("slaMedMarCount"), getValue("slaMedAprCount")
    ],
    [
      "S-3 (Low)", "≤30 Min", "24 Hours",
      getValue("slaLowJanCount"), getValue("slaLowFebCount"), getValue("slaLowMarCount"), getValue("slaLowAprCount")
    ]
  ];
  const slaTableTop = PPT_BODY.top;
  const slaTableH = 1.02;
  slide.addTable(slaRows, {
    x: PPT_BODY.left,
    y: slaTableTop,
    w: PPT_BODY.width,
    h: slaTableH,
    rowH: slaTableH / slaRows.length,
    fontSize: 8,
    align: "center",
    valign: "middle",
    border: { pt: 0.75, color: "000000" }
  });
  const slaSubtitleY = slaTableTop + slaTableH + 0.1;
  slide.addText("Remediation Time SLA in %", {
    x: PPT_BODY.left,
    y: slaSubtitleY,
    w: PPT_BODY.width,
    h: 0.28,
    fontSize: 14,
    bold: true
  });
  const slaChartTop = slaSubtitleY + 0.28 + 0.06;
  const slaPctImg = canvasToPptxData(document.getElementById("slaPctChart"));
  if (slaPctImg) {
    slide.addImage({
      data: slaPctImg,
      x: PPT_BODY.left + 0.1,
      y: slaChartTop,
      w: PPT_BODY.width - 0.2,
      h: PPT_BODY.bottom - slaChartTop
    });
  }
  addPptxFooter(slide);
}

function exportPptxInventory(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Integrated Device Inventory");
  const inv = parseInventoryCsvRows(getValue("inventoryCsv"));
  let invTotal = 0;
  inv.forEach((r) => { invTotal += (parseInt(r.count, 10) || 0); });
  const invRows = [
    [
      { text: "S.NO", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } },
      { text: "DEVICE NAME", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } },
      { text: "COUNT", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } }
    ],
    ...inv.map((r) => [
      { text: String(r.sno), options: { align: "center" } },
      { text: r.deviceName, options: { align: "left" } },
      { text: String(r.count), options: { align: "center" } }
    ]),
    [
      { text: "Total", options: { bold: true, align: "center", colspan: 2 } },
      { text: String(invTotal), options: { bold: true, align: "center" } }
    ]
  ];
  const invTableTop = 0.8;
  const invRowH = 0.32;
  slide.addTable(invRows, {
    x: 0.5,
    y: invTableTop,
    w: 9.0,
    rowH: invRowH,
    fontSize: 12,
    border: { pt: 1, color: "4cc0c5" },
    valign: "middle"
  });
  const invNote = getValue("inventoryNote");
  if (invNote) {
    const invTableBottom = invTableTop + invRows.length * invRowH;
    slide.addText(parseRichTextPptx(invNote), {
      x: 0.5,
      y: Math.min(invTableBottom + 0.15, PPT_BODY.bottom - 0.9),
      w: 9.0,
      h: 0.9,
      fontSize: 11,
      color: "000000",
      fontFace: "Calibri",
      valign: "top"
    });
  }
  addPptxFooter(slide);
}

function exportPptxKeyPoints(pptx) {
  const slide = pptx.addSlide();
  addPptxTitleBar(slide, "Key Points – Overall Summary");
  addKeyPointsToPptxSlide(slide, 0.8);
  addPptxFooter(slide);
}

function exportPptxContact(pptx) {
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { type: "none" },
    line: { color: "BFBFBF", width: 1 }
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 1.6, y: 1.0, w: 2.7 });
  }
  slide.addText("Your Trusted Security Advisor", {
    x: 1.6, y: 2.5, w: 7.4, h: 0.7,
    fontSize: 40,
    bold: true,
    color: "022a5b",
    align: "left",
    valign: "middle",
    fontFace: "Calibri"
  });
  slide.addText(
    [
      { text: "FOR FURTHER DETAILS:", options: { bold: true, color: "022a5b", fontSize: 16 } },
      { text: "Diptesh Saha", options: { bold: true, color: "1e4f9a", fontSize: 20 } },
      { text: "CISO & Practice Head - Cyber Security & Managed Security", options: { color: "1e4f9a", fontSize: 16 } },
      { text: "Contact No. 7338882888", options: { color: "1e4f9a", fontSize: 16 } },
      { text: "diptesh.s@snsin.com", options: { color: "1e4f9a", fontSize: 16 } }
    ],
    {
      x: 1.6, y: 3.3, w: 7.4, h: 2.0,
      align: "left",
      fontFace: "Calibri",
      lineSpacingMultiple: 1.2,
      valign: "top"
    }
  );
  slide.addImage({ data: getFooterBarDataUrl(), x: 0, y: 5.30, w: 10, h: 0.325 });
}

function exportPptxSection(pptx, sectionId, copy) {
  switch (sectionId) {
    case "cover": exportPptxCover(pptx); break;
    case "revision": exportPptxRevision(pptx); break;
    case "engagement": exportPptxEngagement(pptx); break;
    case "alert-quad": exportPptxAlertQuad(pptx); break;
    case "tot-pot": exportPptxTotPot(pptx, copy); break;
    case "trend": exportPptxTrend(pptx, copy); break;
    case "rule-severity": exportPptxRuleSeverity(pptx, copy); break;
    case "pot-incidents": exportPptxPotIncidents(pptx, copy); break;
    case "eps": exportPptxEps(pptx); break;
    case "top-eps": exportPptxTopEps(pptx); break;
    case "risks": exportPptxRisks(pptx, copy); break;
    case "sla-monthly": exportPptxSlaMonthly(pptx); break;
    case "inventory": exportPptxInventory(pptx); break;
    case "key-points": exportPptxKeyPoints(pptx); break;
    case "contact": exportPptxContact(pptx); break;
    default: break;
  }
}

function getPptxCtor() {
  return window.PptxGenJS || window.pptxgen;
}

function addPptxTitleBar(slide, title) {
  slide.addText(title, {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.45,
    fontSize: 14,
    bold: true,
    color: "FFFFFF",
    fill: { color: CONFIG.PPTX_NAVY },
    align: "left",
    valign: "middle",
    margin: [0.1, 0.35, 0.1, 0.35]
  });
}

let _footerBarDataUrl = null;
function getFooterBarDataUrl() {
  if (_footerBarDataUrl) return _footerBarDataUrl;
  const canvas = document.createElement("canvas");
  canvas.width = 1000;
  canvas.height = 40;
  const ctx = canvas.getContext("2d");
  const grad = ctx.createLinearGradient(0, 0, canvas.width, 0);
  grad.addColorStop(0, "#021a3f");
  grad.addColorStop(1, "#1e4f9a");
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  _footerBarDataUrl = canvas.toDataURL("image/png");
  return _footerBarDataUrl;
}

function addPptxFooter(slide) {
  slide.addImage({ data: getFooterBarDataUrl(), x: 0, y: 5.30, w: 10, h: 0.325 });
  slide.addText(CONFIG.COMPANY_NAME, {
    x: 0.35,
    y: 5,
    w: 9,
    h: 0.35,
    fontSize: 9,
    color: "666666"
  });
}

async function exportPptx() {
  const Ctor = getPptxCtor();
  if (!Ctor) {
    alert("PptxGenJS failed to load. Check network and refresh.");
    return;
  }

  const pptxBtn = document.getElementById("pptxBtn");
  if (pptxBtn?.disabled) return;
  const originalLabel = pptxBtn ? pptxBtn.textContent : "";
  if (pptxBtn) {
    pptxBtn.disabled = true;
    pptxBtn.textContent = "Preparing…";
  }

  const restore = await prepareReportForExport();
  try {
    applyData();
    await waitForChartsReady(400);

    const pptx = new Ctor();
    const selectedLayout = pptLayoutConfig.LAYOUT_16X9;
    pptx.defineLayout({
      name: selectedLayout.pptxLayout,
      width: selectedLayout.width,
      height: selectedLayout.height
    });
    pptx.layout = selectedLayout.pptxLayout;
    pptx.author = "SNS SOC";
    pptx.title = `${getValue("customerName")} Monthly SOC Report`;

    const copy = APRAAVA_COPY;
    const sections = getExportSectionOrder();

    sections.forEach((sectionId) => {
      exportPptxSection(pptx, sectionId, copy);
    });

    const fname = `${sanitizeFilename(getValue("customerName"))}_${sanitizeFilename(getReportMonth())}_Monthly_SOC_Report.pptx`;
    await pptx.writeFile(fname);
  } finally {
    restore();
    applyData();
    if (pptxBtn) {
      pptxBtn.disabled = false;
      pptxBtn.textContent = originalLabel || "Export PPTX";
    }
  }
}

async function exportDocx() {
  if (!window.docx || !window.saveAs) {
    alert("DOCX libraries failed to load.");
    return;
  }

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } = window.docx;

  const customer = getValue("customerName");
  const month = getReportMonth();
  const dateRange = getValue("dateRange");
  const preparedBy = getValue("preparedBy");
  const reviewedBy = STATIC_REVIEWED_BY;
  const approvedBy = STATIC_APPROVED_BY;
  const submittedOn = getValue("submittedOn");
  const keyPointSections = parseKeyPointsSections(getKeyPointsSourceText());
  let keyPointHeadingNum = 0;
  const keyPointParagraphs = keyPointSections.flatMap((section) => {
    const paragraphs = [];
    if (section.heading) {
      keyPointHeadingNum += 1;
      paragraphs.push(new Paragraph({
        children: [new TextRun({ text: `${keyPointHeadingNum}. ${section.heading}:`, bold: true })]
      }));
    }
    if (section.body) {
      section.body.split("\n").forEach((line) => {
        const trimmed = line.trim();
        if (trimmed) paragraphs.push(new Paragraph(trimmed));
      });
    }
    return paragraphs;
  });
  const devices = parseDeviceCsvRows(getValue("deviceCsv"));
  const risks = parseRiskCsvRows(getValue("riskCsv"));
  const slaMonthLabels = getReportMonthColumnLabels();
  const slaRows = [
    {
      severity: "S-1 (High)",
      response: "≤15 Min",
      resolution: "4 Hours",
      counts: MONTH_SLOT_KEYS.map((m) => getValue(`slaHigh${m}Count`))
    },
    {
      severity: "S-2 (Medium)",
      response: "≤15 Min",
      resolution: "8 Hours",
      counts: MONTH_SLOT_KEYS.map((m) => getValue(`slaMed${m}Count`))
    },
    {
      severity: "S-3 (Low)",
      response: "≤30 Min",
      resolution: "24 Hours",
      counts: MONTH_SLOT_KEYS.map((m) => getValue(`slaLow${m}Count`))
    }
  ];
  const inventories = parseInventoryCsvRows(getValue("inventoryCsv"));
  const executiveSummary = getValue("executiveSummary");
  const trendNote = getValue("trendNote");
  const revision = getValue("documentRevisionHistory");
  const alertQuadNarratives = [
    { title: "Overall Alerts", fieldId: "overallAlertsNarrative" },
    { title: "True Positive Alerts", fieldId: "truePositiveNarrative" },
    { title: "False Positive Alerts", fieldId: "falsePositiveNarrative" },
    { title: "Potential Incidents", fieldId: "apraavaIncidentsNarrative" }
  ];

  const deviceRows = [
    new TableRow({
      children: ["S.No", "Device", "AVG EPS"].map(
        (col) =>
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: col, bold: true })] })]
          })
      )
    }),
    ...devices.map(
      (row) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(String(row.sno))] }),
            new TableCell({ children: [new Paragraph(row.device)] }),
            new TableCell({ children: [new Paragraph(row.eps)] })
          ]
        })
    )
  ];

  const copy = APRAAVA_COPY;

  const children = [
    new Paragraph({
      children: [new TextRun({ text: "MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", bold: true })]
    }),
    new Paragraph({
      children: [new TextRun({ text: `Monthly SOC Report | ${month}`, bold: true, size: 32 })]
    }),
    new Paragraph(`for ${customer}`),
    new Paragraph(""),
    new Paragraph(`Version 1.0`),
    new Paragraph(`Dated ${dateRange}`),
    new Paragraph(`Prepared By ${preparedBy}`),
    new Paragraph(`Reviewed By ${reviewedBy}`),
    new Paragraph(`Approved By ${approvedBy}`),
    new Paragraph(`Submitted On ${submittedOn}`),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Document Revision History", bold: true })] }),
    new Paragraph(revision),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "The Engagement", bold: true })] }),
    new Paragraph(`${customer} has engaged with SNS to monitor and review the entity's security.`),
    new Paragraph(executiveSummary.replace(/\[MONTH\]/gi, month)),
    new Paragraph("")
  ];

  if (key === "apraava") {
    children.push(
      new Paragraph({ children: [new TextRun({ text: "Potential Incidents and Alert Summary", bold: true })] })
    );
    alertQuadNarratives.forEach(({ title, fieldId }) => {
      children.push(
        new Paragraph({ children: [new TextRun({ text: title, bold: true })] }),
        new Paragraph(getValue(fieldId)),
        new Paragraph("")
      );
    });
  }

  children.push(
    new Paragraph({ children: [new TextRun({ text: copy.trendTitle, bold: true })] }),
    new Paragraph(trendNote),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Potential Incidents - Risks Mitigated", bold: true })] }),
    ...risks.map(
      (risk) =>
        new Paragraph(
          `${risk.sno}. ${risk.attackType} | ${risk.riskScenario} | ${risk.ciaTriad} | ${risk.businessImpact} | ${risk.riskRating}`
        )
    ),
    new Paragraph("")
  );

  children.push(
    new Paragraph({ children: [new TextRun({ text: "Top EPS Devices", bold: true })] }),
    new Table({
      rows: deviceRows,
      width: { size: 100, type: WidthType.PERCENTAGE }
    }),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Total Alerts / TP-FP Charts", bold: true })] }),
    new Paragraph("Charts appear in the on-screen report, PDF printout, and PPTX export."),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Response & Remediation SLA", bold: true })] }),
    ...slaRows.map(
      (sla) =>
        new Paragraph(
          `${sla.severity}: Response ${sla.response}, Resolution ${sla.resolution}, ` +
          slaMonthLabels.map((label, i) => `${label} ${sla.counts[i]}`).join(", ")
        )
    ),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Integrated Device Inventory", bold: true })] }),
    ...inventories.map((inv) => new Paragraph(`${inv.sno}. ${inv.deviceName}: ${inv.count}`)),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Key Points - Overall Summary", bold: true })] }),
    ...keyPointParagraphs,
    new Paragraph(""),
    new Paragraph("Your Trusted Security Advisor"),
    new Paragraph(CONFIG.CONTACT_FOOTER_TEXT),
    new Paragraph(CONFIG.COMPANY_NAME)
  );

  const doc = new Document({
    sections: [
      {
        children
      }
    ]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${sanitizeFilename(customer)}_${sanitizeFilename(month)}_Monthly_SOC_Report.docx`);
}

function applyApraavaInitialData() {
  Object.entries(APRAAVA_INITIAL).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el && !String(el.value || "").trim()) {
      el.value = value;
    }
  });
  ensureApraavaSlideDefaults();
}

function bindImageInput(inputId, setter) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) {
      setter("");
      updateLogos();
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      setter(String(e.target?.result || ""));
      updateLogos();
    };
    reader.readAsDataURL(file);
  });
}

function initTableResizing() {
  const tables = document.querySelectorAll("table:not(.risk-table)");
  tables.forEach(table => {
    const headerRow = table.querySelector("thead tr");
    if (!headerRow) return;
    const cols = headerRow.querySelectorAll("th");
    // Only interior borders are resizable (length - 1 skips the right edge of the last column)
    for (let i = 0; i < cols.length - 1; i++) {
      if (cols[i].querySelector(".resizer")) continue;
      const resizer = document.createElement("div");
      resizer.className = "resizer";
      cols[i].appendChild(resizer);
      
      let x = 0, w = 0, nw = 0;
      const onMouseDown = (e) => {
        x = e.clientX;
        w = cols[i].offsetWidth;
        nw = cols[i+1].offsetWidth;
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      };
      const onMouseMove = (e) => {
        const dx = e.clientX - x;
        cols[i].style.width = `${w + dx}px`;
        cols[i+1].style.width = `${nw - dx}px`;
      };
      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };
      resizer.addEventListener("mousedown", onMouseDown);
    }
  });
}

function bindCsvInput(fileInputId, textareaId) {
  const input = document.getElementById(fileInputId);
  const textarea = document.getElementById(textareaId);
  if (!input || !textarea) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      textarea.value = String(e.target?.result || "");
      applyData();
    };
    reader.readAsText(file);
  });
}

async function init() {
  await ensureDefaultPuzzleDataUrl();
  applyApraavaInitialData();
  document.querySelectorAll(".panel input, .panel textarea, .panel select").forEach(el => {
    el.addEventListener("input", applyData);
  });
  document.getElementById("applyBtn").addEventListener("click", applyData);
  document.getElementById("printBtn").addEventListener("click", (e) => {
    e.preventDefault();
    exportPdf();
  });
  document.getElementById("pptxBtn").addEventListener("click", (e) => {
    e.preventDefault();
    exportPptx();
  });
  document.getElementById("docxBtn").addEventListener("click", exportDocx);
  window.addEventListener("afterprint", () => {
    finishExportRestore();
    const printBtn = document.getElementById("printBtn");
    if (printBtn) {
      printBtn.disabled = false;
      printBtn.textContent = "Export PDF";
    }
  });

  ["snsLogoInput", "clientLogoInput", "puzzleLogoInput"].forEach(id => {
    bindImageInput(id, (v) => {
      if (id === "snsLogoInput") snsLogoDataUrl = v;
      if (id === "clientLogoInput") clientLogoDataUrl = v;
      if (id === "puzzleLogoInput") puzzleLogoDataUrl = v;
      applyData();
    });
  });

  const csvUploads = [
    ["trendCsvFile", "trendCsv"],
    ["potIncidentsCsvFile", "potIncidentsCsv"],
    ["epsTableCsvFile", "epsTableCsv"],
    ["deviceCsvFile", "deviceCsv"],
    ["riskCsvFile", "riskCsv"],
    ["inventoryCsvFile", "inventoryCsv"],
    ["overallAlertsCsvFile", "overallAlertsCsv"],
    ["truePositiveCsvFile", "truePositiveCsv"],
    ["falsePositiveCsvFile", "falsePositiveCsv"],
    ["apraavaIncidentsCsvFile", "apraavaIncidentsCsv"]
  ];
  csvUploads.forEach(([fileId, textId]) => bindCsvInput(fileId, textId));
  initRiskTableUserControls();

  configureSlideNavLabels();
  if (typeof initSlideNav === "function") {
    slideNavController = initSlideNav(() => "apraava");
  }

  if (typeof initTextStyleEditor === "function") {
    initTextStyleEditor();
  }

  applyData();
  initTableResizing();
}

window.onload = init;
