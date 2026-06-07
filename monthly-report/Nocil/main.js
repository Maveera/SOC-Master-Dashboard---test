// Shared constants
const CONFIG = {
  PPTX_NAVY: "12284A",
  PPTX_BLUE: "1E4F9A",
  PPTX_ROWS_PER_SLIDE: 6,
  CONTACT_SLIDE_TEXT: "FOR FURTHER DETAILS:\nDiptesh Saha\nCISO & Practice Head - Cyber Security & Managed Security\nContact No. 7338882888\ndiptesh.s@snsin.com",
  CONTACT_FOOTER_TEXT: "For further details: Diptesh Saha | diptesh.s@snsin.com",
  COMPANY_NAME: "Secure Network Solutions India Pvt Ltd"
};

const fields = {
  customer: "customerName",
  range: "dateRange",
  preparedBy: "preparedBy",
  submittedOn: "submittedOn",
  executiveSummary: "executiveSummary"
};

const STATIC_REVIEWED_BY = "Kishore Kumar";
const STATIC_APPROVED_BY = "Diptesh Saha";

let snsLogoDataUrl = "";
let clientLogoDataUrl = "";
let puzzleLogoDataUrl = "";
let totPotIncChart;
let fortiSiemAlertsChart;
let truePositiveAlertsChart;
let falsePositiveAlertsChart;
let epsTotalChart;
let epsTopHostsChart;
let trendChart;
let ruleSeverityChart;
let responseSlaChart;
let remediationSlaChart;
let totPotPlotFramePluginRegistered = false;
let trendDataLabelsRegistered = false;
const DEFAULT_NOCIL_RULE_SEVERITY_CSV = [
  "Rule,High,Medium,Low",
  "Unusual Login Failure,4,0,0",
  "Suspicious Login,0,18,0"
].join("\n");

/**
 * PptxGenJS layout names must match the library (lowercase x).
 * LAYOUT_WIDE is 13.33"×7.5" — do not use; export coords assume 10" wide (LAYOUT_16x9 / LAYOUT_4x3).
 * @see https://gitbrent.github.io/PptxGenJS/docs/usage-pres-options/
 */
const PPT_LAYOUT_KEY = "LAYOUT_16X9";

const pptLayoutConfig = {
  LAYOUT_16X9: { pptxLayout: "LAYOUT_16x9", reportClass: "ppt-wide" }
};

/** Widescreen 16:9 — standard PowerPoint slide size (10 × 5.625 in). */
const PPT_SLIDE_IN = { w: 10, h: 5.625 };

const PPT_EXPORT_REF = PPT_SLIDE_IN;

function getPptSlideInches() {
  return PPT_SLIDE_IN;
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

let exportRestoreCallback = null;

async function waitForChartsReady(delayMs = 900) {
  const charts = [
    totPotIncChart,
    trendChart,
    ruleSeverityChart
  ].filter(Boolean);
  window.dispatchEvent(new Event("resize"));
  charts.forEach((chart) => {
    try {
      chart.resize();
      chart.update("none");
    } catch (e) {
      // ignore
    }
  });
  renderSlaStatus();
  await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
  if (document.fonts && document.fonts.ready) {
    try {
      await document.fonts.ready;
    } catch (e) {
      // ignore
    }
  }
  await new Promise((r) => setTimeout(r, delayMs));
}

async function prepareReportForExport() {
  document.body.classList.add("is-exporting");
  applyData();
  await renderRiskSlides();
  await waitForChartsReady();
  updateSlaSlide();
  if (typeof initSlaTableColumnResize === "function") {
    initSlaTableColumnResize(document.getElementById("slideSlaStatus"));
  }
  return () => {
    document.body.classList.remove("is-exporting");
  };
}

let reportPdfLayoutRestore = null;

function syncEngagementLayout() {
  if (typeof ReportEngagementLayout !== "undefined") {
    ReportEngagementLayout.syncEngagementLayout();
  }
}

function prepareNocilPrintLayout() {
  if (typeof ReportEngagementLayout === "undefined") {
    return () => {};
  }
  return ReportEngagementLayout.prepareReportPdfLayout(() => updateSlaSlide());
}

function finishNocilPrintLayout() {
  if (reportPdfLayoutRestore) {
    reportPdfLayoutRestore();
    reportPdfLayoutRestore = null;
  }
}

function finishExportRestore() {
  finishNocilPrintLayout();
  if (typeof finishPrintSlideOrder === "function") {
    finishPrintSlideOrder();
  }
  if (exportRestoreCallback) {
    exportRestoreCallback();
    exportRestoreCallback = null;
    applyData();
  }
}

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

function toHighResCanvasDataUrl(canvas, scale = 2) {
  if (!canvas || typeof canvas.toDataURL !== "function") return "";
  const width = canvas.width || canvas.clientWidth;
  const height = canvas.height || canvas.clientHeight;
  if (!width || !height) return canvas.toDataURL("image/png");

  const temp = document.createElement("canvas");
  temp.width = Math.max(1, Math.floor(width * scale));
  temp.height = Math.max(1, Math.floor(height * scale));
  const ctx = temp.getContext("2d");
  if (!ctx) return canvas.toDataURL("image/png");
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.drawImage(canvas, 0, 0, width, height);
  return temp.toDataURL("image/png");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatRichTextHTML(text) {
  return formatReportRichTextHTML(text, getReportMonth, { italic: true });
}

function parseRichTextPptx(text, baseOptions = {}) {
  if (!text) return [];
  const parts = [];
  const monthReplaced = text.replace(/\[MONTH\]/g, getReportMonth());
  const segments = monthReplaced.split(/\*\*/);
  segments.forEach((segment, idx) => {
    if (segment === "") return;
    const isBold = idx % 2 !== 0;
    parts.push({
      text: segment,
      options: { ...baseOptions, bold: isBold || baseOptions.bold }
    });
  });
  return parts;
}

function parseDateMatch(str) {
  const match = str.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (!match) return null;
  const monthIdx = Number(match[2]) - 1;
  const year = Number(match[3]);
  if (monthIdx < 0 || monthIdx > 11) return null;
  return new Date(year, monthIdx, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
}

function sanitizeFilename(str) {
  return str.replace(/[^a-zA-Z0-9_\-]/g, "_").replace(/_+/g, "_");
}

function getReportMonth() {
  const manual = ReportMonthUtils.resolveManualReportMonth(
    getValue("tableMonthLabel"),
    getValue("dateRange"),
    getValue("submittedOn")
  );
  if (manual) return manual;
  return (
    parseDateMatch(getValue("dateRange")) ||
    parseDateMatch(getValue("submittedOn")) ||
    new Date().toLocaleString("en-US", { month: "long", year: "numeric" })
  );
}

function getChartMonthLabel() {
  const reportMonth = getReportMonth();
  return reportMonth.split(" ")[0] || reportMonth;
}

function getTableMonthLabel() {
  return ReportMonthUtils.tableMonthFromManual(getValue("tableMonthLabel"), getReportMonth());
}

function parseTrendCsvRows(csvText) {
  return ReportCharts.parseTrendCsvRows(csvText);
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

const RISK_SLIDE_TITLE = "Potential Incident - Risks Mitigated";
const RISK_CONTD_TITLE = "Contn.,";

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
      attack: findCol(["attacktype", "attack", "alerttype", "eventname", "rulename"], 0),
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

function parseRuleSeverityCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells);
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some(
    (k) => k.includes("rule") || k.includes("high") || k.includes("medium") || k.includes("low")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iRule = findCol(["rulename", "rule", "attacktype", "eventname"], 0);
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

function parseEpsEventsRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some(
    (k) => k.includes("reportingdevice") || k.includes("eventtype") || k.includes("matchedevents")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iSno = findCol(["sno", "sno."], 0);
  const iDevice = findCol(["reportingdevice", "device"], 1);
  const iType = findCol(["eventtype", "type"], 2);
  const iName = findCol(["eventname", "name"], 3);
  const iCount = findCol(["matchedevents", "count", "events"], 4);
  return dataRows
    .map((r, idx) => ({
      sno: r[iSno] || String(idx + 1),
      device: r[iDevice] || "",
      eventType: r[iType] || "",
      eventName: r[iName] || "",
      matchedEvents: r[iCount] || ""
    }))
    .filter((r) => r.device || r.eventType || r.eventName || r.matchedEvents);
}

function parseSupportMajorRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("created") || k.includes("closed") || k.includes("inprocess"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iBy = findCol(["createdby", "created", "owner"], 0);
  const iClosed = findCol(["closed"], 1);
  const iInProcess = findCol(["inprocess", "in-progress", "inprogress"], 2);
  return dataRows
    .map((r) => ({
      by: r[iBy] || "",
      closed: Math.max(0, Number(r[iClosed]) || 0),
      inProcess: Math.max(0, Number(r[iInProcess]) || 0)
    }))
    .filter((r) => r.by)
    .slice(0, 20);
}

function parseSupportMinorRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("created") || k.includes("closed"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iBy = findCol(["createdby", "created", "owner"], 0);
  const iClosed = findCol(["closed"], 1);
  return dataRows
    .map((r) => ({
      by: r[iBy] || "",
      closed: Math.max(0, Number(r[iClosed]) || 0)
    }))
    .filter((r) => r.by)
    .slice(0, 20);
}

function parseInventoryRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("device") || k.includes("count"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iDevice = findCol(["device", "devicename", "name"], 0);
  const iCount = findCol(["count", "total"], 1);
  return dataRows
    .map((r, i) => ({
      sno: i + 1,
      device: r[iDevice] || "",
      count: Math.max(0, Number(String(r[iCount]).replace(/,/g, "")) || 0)
    }))
    .filter((r) => r.device);
}

function parseInventoryRowsFromDom() {
  const body = document.getElementById("inventoryTableBody");
  if (!body) return [];
  const rows = [];
  body.querySelectorAll("tr").forEach((tr) => {
    const cells = tr.querySelectorAll("td");
    if (cells.length < 3) return;
    const device = (cells[1].textContent || "").trim();
    const countText = (cells[2].textContent || "").trim();
    if (!device || /no rows found/i.test(device)) return;
    rows.push({
      sno: rows.length + 1,
      device,
      count: Math.max(0, parseInt(countText.replace(/,/g, ""), 10) || 0)
    });
  });
  return rows;
}

/** Inventory rows for preview + PPT — CSV first, then rendered table fallback. */
function getInventoryRows() {
  const fromCsv = parseInventoryRows(getValue("inventoryCsv"));
  if (fromCsv.length) return fromCsv;
  return parseInventoryRowsFromDom();
}

/** Stack table + note tightly; shrink rows only when content exceeds the slide. */
function getInventoryPptLayout(dataRowCount, noteText) {
  const tableRows = Math.max(1, dataRowCount) + 2;
  const titleBottom = 0.92;
  const footerReserve = 0.42;
  const noteGap = noteText.trim() ? 0.04 : 0;
  const maxBottom = PPT_SLIDE_IN.h - footerReserve;
  const noteLines = noteText.trim() ? Math.max(1, Math.ceil(noteText.trim().length / 88)) : 0;
  const noteH = noteLines ? Math.min(1.1, 0.18 + noteLines * 0.15) : 0;

  let rowH = 0.34;
  let fontSize = 12;
  let noteFontSize = 11;
  if (dataRowCount > 12) {
    rowH = 0.18;
    fontSize = 8.5;
    noteFontSize = 9;
  } else if (dataRowCount > 9) {
    rowH = 0.22;
    fontSize = 9.5;
    noteFontSize = 10;
  } else if (dataRowCount > 6) {
    rowH = 0.27;
    fontSize = 10.5;
    noteFontSize = 10;
  }

  const tableTop = titleBottom;
  let tableHeight = rowH * tableRows;
  let noteY = tableTop + tableHeight + noteGap;
  let totalBottom = noteY + noteH;

  if (totalBottom > maxBottom) {
    const maxTableH = Math.max(0.8, maxBottom - titleBottom - noteGap - noteH);
    rowH = maxTableH / tableRows;
    tableHeight = rowH * tableRows;
    noteY = tableTop + tableHeight + noteGap;
  }

  return {
    tableX: 1.4,
    tableW: 7.2,
    tableTop,
    tableBottom: tableTop + tableHeight,
    rowH,
    fontSize,
    noteY,
    noteH,
    noteFontSize,
    colW: [1.15, 4.05, 2.0]
  };
}

function buildInventoryPptTable(invRows) {
  const invTotal = invRows.reduce((a, r) => a + r.count, 0);
  const rowFill = (idx) => ({ color: idx % 2 === 0 ? "FFFFFF" : "F2F9FA" });
  const headerCell = (text) => ({
    text,
    options: { bold: true, fill: { color: "4CC0C5" }, color: "FFFFFF", align: "center" }
  });
  const dataCell = (text, idx, align = "center") => ({
    text,
    options: { align, fill: rowFill(idx), color: "111111" }
  });

  return [
    [headerCell("S.NO"), headerCell("DEVICE NAME"), headerCell("COUNT")],
    ...(invRows.length
      ? invRows.map((r, idx) => [
          dataCell(String(idx + 1), idx),
          dataCell(r.device, idx, "center"),
          dataCell(r.count.toLocaleString("en-US"), idx)
        ])
      : [[dataCell("-", 0), dataCell("No rows found. Upload Inventory CSV.", 0, "center"), dataCell("-", 0)]]),
    [
      {
        text: "Total",
        options: {
          bold: true,
          fill: { color: "4CC0C5" },
          color: "FFFFFF",
          align: "center",
          colspan: 2
        }
      },
      {
        text: invTotal.toLocaleString("en-US"),
        options: { bold: true, fill: { color: "4CC0C5" }, color: "000000", align: "center" }
      }
    ]
  ];
}

function renderIntegratedInventorySlide() {
  const body = document.getElementById("inventoryTableBody");
  const foot = document.getElementById("inventoryTableFoot");
  const note = document.getElementById("inventoryNotePreview");
  const wrap = document.querySelector(".slide-integrated-inventory .inventory-table-wrap");
  if (!body || !foot || !note) return;
  const rows = getInventoryRows();
  if (wrap) {
    wrap.classList.remove("inventory-table-wrap--compact", "inventory-table-wrap--dense");
    if (rows.length > 10) wrap.classList.add("inventory-table-wrap--dense");
    else if (rows.length > 6) wrap.classList.add("inventory-table-wrap--compact");
  }
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3">No rows found. Upload Inventory CSV.</td></tr>';
    foot.innerHTML = "";
  } else {
    body.innerHTML = rows
      .map(
        (r, i) =>
          `<tr><td>${i + 1}</td><td>${escapeHtml(r.device)}</td><td>${r.count.toLocaleString("en-US")}</td></tr>`
      )
      .join("");
    const total = rows.reduce((a, r) => a + r.count, 0);
    foot.innerHTML = `<tr><td colspan="2">Total</td><td>${total.toLocaleString("en-US")}</td></tr>`;
  }
  note.innerHTML = formatRichTextHTML(getValue("inventoryNote"));
}

function renderSupportTicketsSlide() {
  const majorBody = document.getElementById("supportMajorBody");
  const minorBody = document.getElementById("supportMinorBody");
  if (!majorBody || !minorBody) return;

  const majorRows = parseSupportMajorRows(getValue("supportMajorCsv"));
  const minorRows = parseSupportMinorRows(getValue("supportMinorCsv"));

  if (!majorRows.length) {
    majorBody.innerHTML = '<tr><td colspan="4">No rows found.</td></tr>';
  } else {
    const majorHtml = majorRows
      .map((r) => {
        const total = r.closed + r.inProcess;
        return `<tr><td>${escapeHtml(r.by)}</td><td>${r.closed}</td><td>${r.inProcess}</td><td>${total}</td></tr>`;
      })
      .join("");
    const sumClosed = majorRows.reduce((a, r) => a + r.closed, 0);
    const sumInProcess = majorRows.reduce((a, r) => a + r.inProcess, 0);
    majorBody.innerHTML =
      majorHtml +
      `<tr><td>Grand Total</td><td>${sumClosed}</td><td>${sumInProcess}</td><td>${sumClosed + sumInProcess}</td></tr>`;
  }

  if (!minorRows.length) {
    minorBody.innerHTML = '<tr><td colspan="2">No rows found.</td></tr>';
  } else {
    const minorHtml = minorRows
      .map((r) => `<tr><td>${escapeHtml(r.by)}</td><td>${r.closed}</td></tr>`)
      .join("");
    const sumMinor = minorRows.reduce((a, r) => a + r.closed, 0);
    minorBody.innerHTML = minorHtml + `<tr><td>Grand Total</td><td>${sumMinor}</td></tr>`;
  }
}

function renderEpsEventsTable() {
  const body = document.getElementById("epsEventsTableBody");
  if (!body) return;
  const rows = parseEpsEventsRows(getValue("epsEventsCsv"));
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="5" class="eps-num">No rows found. Upload EPS Events CSV.</td></tr>';
    return;
  }
  body.innerHTML = rows
    .map((r) => {
      const countNum = Number(String(r.matchedEvents).replace(/,/g, ""));
      const countText = Number.isFinite(countNum) ? countNum.toLocaleString("en-US") : r.matchedEvents;
      return `<tr>
        <td class="eps-num">${escapeHtml(String(r.sno))}</td>
        <td>${escapeHtml(r.device)}</td>
        <td>${escapeHtml(r.eventType)}</td>
        <td>${escapeHtml(r.eventName)}</td>
        <td class="eps-count">${escapeHtml(String(countText))}</td>
      </tr>`;
    })
    .join("");
}

function getRiskLayoutOptions() {
  const maxRows = parseInt(getValue("riskMaxRowsPerSlide"), 10);
  const fontPx = parseFloat(getValue("riskTableFontSize"));
  return {
    maxRowsPerSlide: Number.isFinite(maxRows) && maxRows > 0 ? maxRows : 0,
    tableFontSizePx: Number.isFinite(fontPx) && fontPx > 0 ? fontPx : 0
  };
}

function initRiskTableUserControls() {
  const reflow = () => {
    void renderRiskSlides();
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

function renderRiskSlides() {
  if (typeof RiskPagination === "undefined") {
    return Promise.resolve();
  }
  return new Promise((resolve) => {
    RiskPagination.renderRiskSlides({
      rows: parseRiskCsvRows(getValue("riskCsv")),
      baseTitle: RISK_SLIDE_TITLE,
      contdTitle: RISK_CONTD_TITLE,
      narrativeText: getValue("riskNarrative"),
      escapeHtml,
      formatRichTextHTML,
      layoutOptions: getRiskLayoutOptions(),
      onComplete: (host) => {
        applyPremiumLayout();
        updateLogos();
        if (typeof initRiskTableColumnResize === "function") {
          initRiskTableColumnResize(host || document.getElementById("riskSlidesContainer"));
        }
        resolve();
      }
    });
  });
}

function chunkRiskRowsForPptx(rows) {
  if (typeof RiskPagination !== "undefined" && RiskPagination.chunkRiskRowsForSlides) {
    return RiskPagination.chunkRiskRowsForSlides(
      rows,
      getValue("riskNarrative"),
      formatRichTextHTML,
      escapeHtml,
      RISK_CONTD_TITLE,
      getRiskLayoutOptions()
    );
  }
  const chunkSize = 3;
  return rows.length
    ? Array.from({ length: Math.ceil(rows.length / chunkSize) }, (_, i) =>
        rows.slice(i * chunkSize, i * chunkSize + chunkSize)
      )
    : [[]];
}

/** PPT risk slide — table fills from title/narrative down to footer (no empty band below rows). */
function getPptRiskSlideLayout(slideIndex, dataRowCount, sy) {
  const footerTop = 4.95 * sy;
  const tableGap = 0.03 * sy;
  const titleY = 0.55 * sy;
  const titleH = 0.38 * sy;
  const gapIn = 0.04 * sy;
  const hasNarrative = slideIndex === 0 && Boolean(String(getValue("riskNarrative") || "").trim());
  const rowCount = Math.max(1, dataRowCount) + 1; // +1 header row
  const minTableH = 0.35 * sy; // keep table visible; shrink narrative first

  let tableY;
  let narrativeY = 0;
  let narrativeH = 0;

  if (hasNarrative) {
    const raw = String(getValue("riskNarrative") || "").trim();
    const lines = Math.max(1, raw.split(/\r?\n/).length, Math.ceil(raw.length / 92));
    narrativeH = Math.min(1.15, 0.1 * lines + 0.18) * sy;
    narrativeY = 0.68 * sy;

    // Ensure header + narrative + table don't collide with the footer.
    // If narrative is too tall, shrink it so the table can fit (important for 3 rows).
    const maxNarrativeH = footerTop - (narrativeY + gapIn) - tableGap - minTableH;
    narrativeH = Math.max(0, Math.min(narrativeH, maxNarrativeH));
    tableY = narrativeY + narrativeH + gapIn;
  } else {
    tableY = titleY + titleH + gapIn;
  }

  const tableH = Math.max(minTableH, footerTop - tableY - tableGap);
  const rowH = tableH / rowCount;

  return { titleY, titleH, narrativeY, narrativeH, tableY, tableH, rowH, footerTop };
}

function renderRuleSeverityChart() {
  const canvas = document.getElementById("ruleSeverityChart");
  if (!canvas || typeof Chart === "undefined") return;
  const rows = parseRuleSeverityCsvRows(getValue("ruleSeverityCsv"));
  if (ruleSeverityChart) {
    ruleSeverityChart.destroy();
    ruleSeverityChart = undefined;
  }
  if (!rows.length) return;

  const maxVal = Math.max(...rows.map((r) => Math.max(r.high, r.medium, r.low)), 1);
  // Add headroom so right-edge labels (ex: Medium=18) stay visible.
  const axisMax = Math.max(6, Math.ceil(maxVal * 1.15));
  ensureTrendDataLabelsPlugin();
  const plugins = {
    title: {
      display: true,
      text: "Potential Incident - Severity Categories",
      font: { size: 16, weight: "bold" }
    },
    legend: {
      display: true,
      position: "bottom",
      labels: ReportCharts.buildSquareLegendLabels("#000")
    }
  };
  if (ReportCharts.isDataLabelsRegistered()) {
    plugins.datalabels = {
      display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
      color: "#000",
      anchor: "end",
      align: "right",
      offset: 3,
      clamp: true,
      clip: false,
      formatter: (v) => Math.round(Number(v)),
      font: { size: 10 }
    };
  }

  ruleSeverityChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: rows.map((r) => r.rule),
      datasets: ReportCharts.buildRuleSeverityDatasets(rows).map((ds) => ({
        ...ds,
        stack: "severity"
      }))
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins,
      scales: {
        x: {
          stacked: true,
          beginAtZero: true,
          max: axisMax,
          grid: { color: "#d4d4d4", drawBorder: false },
          ticks: { stepSize: 1, color: "#000" }
        },
        y: {
          stacked: true,
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000", font: { size: 11 } }
        }
      }
    }
  });
}

/** Light grey used outside the pie (box margin → circle edge); unfilled slice is white inside the circle. */
const SLA_HASH_GREY = "#ececec";

function classifySlaScore(rawPct) {
  const pct = Math.max(0, Math.min(100, Number(rawPct) || 0));
  if (pct >= 100) return { pct, label: "Within SLA", color: "#5b9bd5" };
  return { pct, label: "SLA Breached", color: "#ff0000" };
}

/** Derive response/remediation SLA % from Total and Closed incidents. */
function calculateSlaPercentages() {
  const total = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const closed = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const clampedClosed = Math.min(closed, total || closed);

  if (total === 0) {
    return { responsePct: 100, remediationPct: 100 };
  }

  return {
    responsePct: 100,
    remediationPct: Math.round((clampedClosed / total) * 100)
  };
}

function getCalculatedSlaInfo() {
  const { responsePct, remediationPct } = calculateSlaPercentages();
  return {
    response: classifySlaScore(responsePct),
    remediation: classifySlaScore(remediationPct)
  };
}

function renderSlaPieChart(canvasId, info) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  const dpr = window.devicePixelRatio || 1;
  const cssW = Math.max(10, canvas.clientWidth || 240);
  const cssH = Math.max(10, canvas.clientHeight || 140);
  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);

  const cx = cssW * 0.40;
  const cy = cssH * 0.5;
  const r = Math.min(cssW * 0.68, cssH) * 0.40;
  const start = -Math.PI / 2;
  const pct = Math.max(0, Math.min(100, Number(info.pct) || 0));
  const sweep = (pct / 100) * Math.PI * 2;
  const isFull = pct >= 99.999;

  const frameEl = canvas.parentElement;
  if (frameEl && frameEl.classList.contains("sla-pie-frame")) {
    frameEl.style.background = SLA_HASH_GREY;
  }

  ctx.fillStyle = SLA_HASH_GREY;
  ctx.fillRect(0, 0, cssW, cssH);

  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fillStyle = "#ffffff";
  ctx.fill();

  if (pct > 0) {
    ctx.beginPath();
    if (isFull) {
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
    } else {
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, start, start + sweep, false);
      ctx.closePath();
    }
    ctx.fillStyle = info.color;
    ctx.fill();
  }

  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.strokeStyle = "#4c4c4c";
  ctx.lineWidth = 1;
  ctx.stroke();

  ctx.fillStyle = pct >= 50 ? "#fff" : "#222";
  ctx.font = "bold 12px Calibri, Segoe UI, Arial";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(`${Math.round(pct)}%`, cx, cy);

  return null;
}

const SLA_INPUT_IDS = new Set(["slaIncidentCount", "slaClosedIncidentCount"]);

/** Updates SLA pie charts and summary text only; reference tables stay static in index.html. */
function renderSlaStatus() {
  const incidentCount = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const { response: responseInfo, remediation: remediationInfo } = getCalculatedSlaInfo();
  const closedIncidentCount = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const customer = getValue("customerName") || "customer";

  responseSlaChart = renderSlaPieChart("responseSlaChart", responseInfo);
  remediationSlaChart = renderSlaPieChart("remediationSlaChart", remediationInfo);

  const responseLegend = document.getElementById("responseSlaLegend");
  if (responseLegend) responseLegend.textContent = responseInfo.label;
  const remediationLegend = document.getElementById("remediationSlaLegend");
  if (remediationLegend) remediationLegend.textContent = remediationInfo.label;

  const responseSummary = document.getElementById("responseSlaSummary");
  if (responseSummary) {
    responseSummary.textContent = formatResponseSummaryText(
      incidentCount,
      customer,
      responseInfo.label
    );
    responseSummary.style.color = "#111";
    responseSummary.style.fontWeight = "bold";
    responseSummary.hidden = false;
  }
  const remediationSummary = document.getElementById("remediationSlaSummary");
  if (remediationSummary) {
    remediationSummary.textContent = formatRemediationSummaryText(
      closedIncidentCount,
      remediationInfo.label
    );
    remediationSummary.style.color = "#111";
    remediationSummary.style.fontWeight = "bold";
    remediationSummary.hidden = false;
  }
}

function slaCustomerLabel(customerName) {
  const name = (customerName || "Customer").trim();
  if (/nocil/i.test(name)) return name;
  return name;
}

function formatResponseSummaryText(incidentCount, customerName, responseLabel) {
  const customer = slaCustomerLabel(customerName);
  if (responseLabel === "Within SLA") {
    return `Overall, ${incidentCount} incidents were reported to ${customer}, all tickets are reported within SLA.`;
  }
  return `Overall, ${incidentCount} incidents were reported to ${customer}, response status: SLA Breached.`;
}

function formatRemediationSummaryText(closedCount, remediationLabel) {
  const phrase =
    remediationLabel === "Within SLA"
      ? "were remediated within SLA"
      : "were not remediated within SLA";
  return `Remediation Time for ${closedCount} closed tickets ${phrase}.`;
}

function updateSlaSlide() {
  renderSlaStatus();
  requestAnimationFrame(() => {
    renderSlaStatus();
  });
}

function handleSlaInputChange() {
  updateSlaSlide();
}

function getHiResChartDataUrl(canvas) {
  if (!canvas) return "";
  const sourceW = canvas.width || Math.max(1, Math.round(canvas.clientWidth || 1));
  const sourceH = canvas.height || Math.max(1, Math.round(canvas.clientHeight || 1));
  const upscale = 2;
  const out = document.createElement("canvas");
  out.width = Math.max(1, Math.round(sourceW * upscale));
  out.height = Math.max(1, Math.round(sourceH * upscale));
  const outCtx = out.getContext("2d");
  if (!outCtx) return canvas.toDataURL("image/png");
  outCtx.imageSmoothingEnabled = true;
  outCtx.imageSmoothingQuality = "high";
  outCtx.drawImage(canvas, 0, 0, out.width, out.height);
  return out.toDataURL("image/png");
}

/** Keeps #reportRoot in sync with widescreen 16:9 slide size (see styles.css .ppt-wide). */
function syncReportRootLayoutClass() {
  const root = document.getElementById("reportRoot");
  if (!root) return;
  const { reportClass } = pptLayoutConfig.LAYOUT_16X9;
  root.className = `report ${reportClass}`;
}

const TOT_POT_INC_SEVERITIES = [
  {
    label: "High",
    fieldId: "totPotIncHigh",
    color: "#ff0000",
    pptColor: "FF0000",
    tableRowId: "totPotIncTableHighRow",
    tableCellId: "totPotIncTableHigh"
  },
  {
    label: "Medium",
    fieldId: "totPotIncMedium",
    color: "#ffff00",
    pptColor: "FFFF00",
    tableRowId: "totPotIncTableMedRow",
    tableCellId: "totPotIncTableMed"
  },
  {
    label: "Low",
    fieldId: "totPotIncLow",
    color: "#00b050",
    pptColor: "00B050",
    tableRowId: "totPotIncTableLowRow",
    tableCellId: "totPotIncTableLow"
  }
];

function getTotPotIncSeverityEntries() {
  return TOT_POT_INC_SEVERITIES.map((severity) => {
    const raw = String(getValue(severity.fieldId) ?? "").trim();
    if (raw === "") return null;
    return {
      ...severity,
      raw,
      value: Math.max(0, parseFloat(raw) || 0)
    };
  }).filter(Boolean);
}

function syncTotPotIncSlideLayout() {
  const slide =
    document.getElementById("slideTotPotInc") ||
    document.querySelector('.slide-total-potential[data-slide="tot-pot"]');
  if (!slide) return;
  const count = getTotPotIncSeverityEntries().length;
  slide.classList.remove("totpot-rows-1", "totpot-rows-2", "totpot-rows-3");
  if (count >= 1 && count <= 3) {
    slide.classList.add(`totpot-rows-${count}`);
  }
}

function updateTotPotIncTable(monthLbl) {
  const hdr = document.getElementById("totPotIncTableColHdr");
  if (hdr) hdr.textContent = monthLbl;
  TOT_POT_INC_SEVERITIES.forEach((severity) => {
    const row = document.getElementById(severity.tableRowId);
    const cell = document.getElementById(severity.tableCellId);
    const raw = String(getValue(severity.fieldId) ?? "").trim();
    const visible = raw !== "";
    if (row) row.style.display = visible ? "" : "none";
    if (cell) cell.textContent = visible ? raw : "";
  });
  syncTotPotIncSlideLayout();
}

function ensureTotPotPlotFramePlugin() {
  if (typeof Chart === "undefined" || totPotPlotFramePluginRegistered) return;
  totPotPlotFramePluginRegistered = true;
  Chart.register({
    id: "totPotPlotFrame",
    afterDraw(chart) {
      if (
        !chart.canvas ||
        ![
          "totPotIncChart",
          "fortiSiemAlertsChart",
          "truePositiveAlertsChart",
          "falsePositiveAlertsChart",
          "epsTotalChart",
          "epsTopHostsChart"
        ].includes(
          chart.canvas.id
        )
      )
        return;
      const { ctx, chartArea } = chart;
      if (!chartArea || chartArea.width <= 0 || chartArea.height <= 0) return;
      ctx.save();
      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 1;
      ctx.strokeRect(
        chartArea.left + 0.5,
        chartArea.top + 0.5,
        chartArea.width - 1,
        chartArea.height - 1
      );
      ctx.restore();
    }
  });
}

function renderTotPotIncChart() {
  const canvas = document.getElementById("totPotIncChart");
  if (!canvas || typeof Chart === "undefined") return;

  const entries = getTotPotIncSeverityEntries();
  const labels = entries.map((entry) => entry.label);
  const data = entries.map((entry) => entry.value);
  const colors = entries.map((entry) => entry.color);
  const peak = Math.max(...data, 1);
  const yMax = Math.max(10, Math.ceil((peak * 1.1) / 5) * 5);

  if (totPotIncChart) totPotIncChart.destroy();
  ensureTotPotPlotFramePlugin();
  ensureTrendDataLabelsPlugin();

  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 11 };
  const plugins = {
    totPotPlotFrame: true,
    legend: { display: false },
    title: { display: false },
    tooltip: {
      enabled: true,
      callbacks: {
        label: (ctx) => `${ctx.label}: ${ctx.parsed.y}`
      }
    }
  };
  if (ReportCharts.isDataLabelsRegistered()) {
    plugins.datalabels = {
      display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
      color: "#000000",
      font: { weight: "bold", size: 11, family: axisFont.family },
      formatter: (v) => {
        const n = Number(v);
        return Number.isFinite(n) ? String(Math.round(n)) : "";
      },
      anchor: "end",
      align: "top",
      offset: 2,
      clamp: true,
      clip: false
    };
  }

  totPotIncChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "Incident Count",
          data,
          backgroundColor: colors,
          borderWidth: 0
        }
      ]
    },
    options: {
      // Default Chart.js bar animation is 1000ms; bars looked "wrong" (e.g. ~3 vs 14) until it finished.
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { left: 0, right: 2, top: 16, bottom: 2 } },
      elements: { bar: { borderWidth: 0 } },
      plugins,
      datasets: {
        bar: { categoryPercentage: 0.72, barPercentage: 0.72 }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { display: false }
        },
        y: {
          beginAtZero: true,
          max: yMax,
          min: 0,
          ticks: {
            stepSize: 5,
            font: axisFont,
            color: "#000000",
            padding: 8,
            mirror: false
          },
          grid: {
            display: true,
            color: "#e0e0e0",
            lineWidth: 1,
            drawBorder: false,
            drawTicks: false
          }
        }
      }
    }
  });

  syncTotPotIncSlideLayout();
  requestAnimationFrame(() => {
    if (totPotIncChart) totPotIncChart.resize();
  });
}

function renderFortiSiemAlertsChart() {
  const canvas = document.getElementById("fortiSiemAlertsChart");
  if (!canvas || typeof Chart === "undefined") return;

  const high = Math.max(0, parseFloat(getValue("fortiAlertHigh")) || 0);
  const medium = Math.max(0, parseFloat(getValue("fortiAlertMedium")) || 0);
  const low = Math.max(0, parseFloat(getValue("fortiAlertLow")) || 0);
  const peak = Math.max(high, medium, low, 1);
  const yMax = Math.max(2000, Math.ceil((peak * 1.1) / 2000) * 2000);
  const monthLabel = getTableMonthLabel();

  if (fortiSiemAlertsChart) fortiSiemAlertsChart.destroy();
  ensureTotPotPlotFramePlugin();

  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 11 };

  fortiSiemAlertsChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: [monthLabel],
      datasets: [
        {
          label: "High",
          data: [high],
          backgroundColor: "#ff0000",
          borderWidth: 0
        },
        { label: "Medium", data: [medium], backgroundColor: "#ffff00", borderWidth: 0 },
        { label: "Low", data: [low], backgroundColor: "#00b050", borderWidth: 0 }
      ]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { left: 0, right: 4, top: 8, bottom: 4 } },
      elements: { bar: { borderWidth: 0 } },
      plugins: {
        totPotPlotFrame: true,
        legend: {
          display: true,
          position: "bottom",
          labels: { boxWidth: 12, boxHeight: 8, usePointStyle: false, color: "#000", padding: 12 }
        },
        title: { display: false },
        tooltip: {
          enabled: true,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}`
          }
        }
      },
      datasets: {
        bar: { categoryPercentage: 0.65, barPercentage: 0.85 }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { font: axisFont, color: "#000000" }
        },
        y: {
          beginAtZero: true,
          max: yMax,
          min: 0,
          ticks: {
            stepSize: 2000,
            font: axisFont,
            color: "#000000",
            padding: 8,
            callback: (v) => v.toLocaleString("en-US")
          },
          grid: {
            display: true,
            color: "#e0e0e0",
            lineWidth: 1,
            drawBorder: false,
            drawTicks: false
          }
        }
      }
    }
  });

  requestAnimationFrame(() => {
    if (fortiSiemAlertsChart) fortiSiemAlertsChart.resize();
  });
}

function tfPosPickScale(peak) {
  const p = Math.max(peak * 1.05, 1);
  let step = 200;
  if (p > 6000) step = 2000;
  else if (p > 3000) step = 1000;
  else if (p > 1500) step = 500;
  else step = 200;
  const max = Math.ceil(p / step) * step;
  return { max, step };
}

function fillTfPosPlaceholders(template) {
  if (!template) return "";
  const tpH = Math.max(0, parseFloat(getValue("tfPosTrueHigh")) || 0);
  const tpM = Math.max(0, parseFloat(getValue("tfPosTrueMedium")) || 0);
  const tpL = Math.max(0, parseFloat(getValue("tfPosTrueLow")) || 0);
  const fpH = Math.max(0, parseFloat(getValue("tfPosFalseHigh")) || 0);
  const fpM = Math.max(0, parseFloat(getValue("tfPosFalseMedium")) || 0);
  const fpL = Math.max(0, parseFloat(getValue("tfPosFalseLow")) || 0);
  const fpTotal = fpH + fpM + fpL;
  return template
    .replace(/\[MONTH\]/g, getReportMonth())
    .replace(/\[TP_H\]/g, String(tpH))
    .replace(/\[TP_M\]/g, String(tpM))
    .replace(/\[TP_L\]/g, String(tpL))
    .replace(/\[FP_H\]/g, String(fpH))
    .replace(/\[FP_M\]/g, String(fpM))
    .replace(/\[FP_L\]/g, String(fpL))
    .replace(/\[FP_TOTAL\]/g, String(fpTotal));
}

function parseTfPosNoteLines(filled) {
  const lines = filled.split(/\r?\n/).map((l) => l.trim());
  const rawLines = lines.filter((l) => l.length > 0);
  if (!rawLines.length) return { items: [] };
  let items = [];
  if (/^Note:\s*$/i.test(rawLines[0])) {
    items = rawLines.slice(1);
  } else if (/^Note:\s+/i.test(rawLines[0])) {
    const after = rawLines[0].replace(/^Note:\s+/i, "").trim();
    items = after ? [after, ...rawLines.slice(1)] : rawLines.slice(1);
  } else {
    items = rawLines;
  }
  return { items };
}

function renderTfPosNotePreview(previewId, textareaId) {
  const el = document.getElementById(previewId);
  if (!el) return;
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  const heading = '<p class="tfpos-note-heading"><strong>Note:</strong></p>';
  if (!items.length) {
    el.innerHTML = heading;
    return;
  }
  const lis = items.map((line) => `<li>${formatRichTextHTML(line)}</li>`).join("");
  el.innerHTML = `${heading}<ul class="tfpos-note-list">${lis}</ul>`;
}

/** Plain text for PPTX: "Note:" then bullet lines. */
function formatTfPosNoteForPptx(textareaId) {
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  if (!items.length) return "Note:";
  return `Note:\n\n${items.map((line) => `• ${line}`).join("\n")}`;
}

function renderTfPosCharts() {
  const monthLabel = getTableMonthLabel();
  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 10 };

  const build = (canvasId, oldChart, high, medium, low) => {
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === "undefined") return undefined;
    const peak = Math.max(high, medium, low, 1);
    const { max: yMax, step: yStep } = tfPosPickScale(peak);
    ensureTotPotPlotFramePlugin();

    if (oldChart) oldChart.destroy();
    return new Chart(canvas, {
      type: "bar",
      data: {
        labels: [monthLabel],
        datasets: [
          { label: "High", data: [high], backgroundColor: "#ff0000", borderWidth: 0 },
          { label: "Medium", data: [medium], backgroundColor: "#ffff00", borderWidth: 0 },
          { label: "Low", data: [low], backgroundColor: "#00b050", borderWidth: 0 }
        ]
      },
      options: {
        animation: false,
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { left: 0, right: 2, top: 4, bottom: 2 } },
        elements: { bar: { borderWidth: 0 } },
        plugins: {
          totPotPlotFrame: true,
          legend: {
            display: true,
            position: "bottom",
            labels: { boxWidth: 10, boxHeight: 7, usePointStyle: false, color: "#000", padding: 8, font: { size: 9 } }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString("en-US")}`
            }
          }
        },
        datasets: {
          bar: { categoryPercentage: 0.65, barPercentage: 0.85 }
        },
        scales: {
          x: {
            grid: { display: false, drawBorder: false },
            ticks: { font: axisFont, color: "#000000", maxRotation: 0 }
          },
          y: {
            beginAtZero: true,
            max: yMax,
            min: 0,
            ticks: {
              stepSize: yStep,
              font: axisFont,
              color: "#000000",
              padding: 6,
              callback: (v) => v.toLocaleString("en-US")
            },
            grid: {
              display: true,
              color: "#e0e0e0",
              lineWidth: 1,
              drawBorder: false,
              drawTicks: false
            }
          }
        }
      }
    });
  };

  const tpH = Math.max(0, parseFloat(getValue("tfPosTrueHigh")) || 0);
  const tpM = Math.max(0, parseFloat(getValue("tfPosTrueMedium")) || 0);
  const tpL = Math.max(0, parseFloat(getValue("tfPosTrueLow")) || 0);
  const fpH = Math.max(0, parseFloat(getValue("tfPosFalseHigh")) || 0);
  const fpM = Math.max(0, parseFloat(getValue("tfPosFalseMedium")) || 0);
  const fpL = Math.max(0, parseFloat(getValue("tfPosFalseLow")) || 0);

  truePositiveAlertsChart = build("truePositiveAlertsChart", truePositiveAlertsChart, tpH, tpM, tpL);
  falsePositiveAlertsChart = build("falsePositiveAlertsChart", falsePositiveAlertsChart, fpH, fpM, fpL);

  requestAnimationFrame(() => {
    if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
    if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
  });
}

function parseEpsTopHostsRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("host") || k.includes("device") || k.includes("eps"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iHost = findCol(["host", "device", "reportingdevice", "name"], 0);
  const iEps = findCol(["eps", "avgeps", "value"], 1);
  return dataRows
    .map((r) => ({
      host: r[iHost] || "",
      eps: Math.max(0, Number(r[iEps]) || 0)
    }))
    .filter((r) => r.host)
    .slice(0, 10);
}

function renderEpsTrendPlot() {
  const totalCanvas = document.getElementById("epsTotalChart");
  const hostsCanvas = document.getElementById("epsTopHostsChart");
  if (!totalCanvas || !hostsCanvas || typeof Chart === "undefined") return;

  const epsChartXLabel = ReportMonthUtils.resolveEpsChartAxisLabel(getValue("epsMonthLabel"));
  const totalEps = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const hostRows = parseEpsTopHostsRows(getValue("epsTopHostsCsv"));
  const rows = hostRows.length ? hostRows : [{ host: "N/A", eps: 0 }];
  const palette = [
    "#4bc0c0",
    "#ffcd56",
    "#ff6384",
    "#ff9f40",
    "#9966ff",
    "#c45891",
    "#4db6ac",
    "#d4a017",
    "#d32f2f",
    "#e66a2c"
  ];

  ensureTrendDataLabelsPlugin();
  ensureTotPotPlotFramePlugin();
  if (epsTotalChart) epsTotalChart.destroy();
  if (epsTopHostsChart) epsTopHostsChart.destroy();

  const totalAxisMax = Math.max(10, Math.ceil((totalEps * 1.08) / 50) * 50);
  const totalPlugins = { legend: { display: false }, totPotPlotFrame: true };
  if (ReportCharts.isDataLabelsRegistered()) {
    totalPlugins.datalabels = ReportCharts.buildEpsTrendTotalValueDataLabelOptions({
      font: { weight: "bold", size: 12 }
    });
  }

  epsTotalChart = new Chart(totalCanvas, {
    type: "bar",
    data: {
      labels: [epsChartXLabel],
      datasets: [{ label: "EPS", data: [totalEps], backgroundColor: "#4bc0c0", borderWidth: 0 }]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 18, bottom: 2, left: 4, right: 6 } },
      plugins: totalPlugins,
      scales: {
        x: { grid: { display: false, drawBorder: false }, ticks: { color: "#000", font: { weight: "bold" } } },
        y: {
          beginAtZero: true,
          min: 0,
          max: totalAxisMax,
          ticks: { color: "#000", stepSize: Math.max(10, Math.round(totalAxisMax / 8)) },
          grid: { color: "#d9d9d9", drawBorder: false }
        }
      }
    }
  });

  const hostValues = rows.map((r) => r.eps);
  const hostAxisMax = Math.max(10, Math.ceil((Math.max(...hostValues, 1) * 1.15) / 10) * 10);
  epsTopHostsChart = new Chart(hostsCanvas, {
    type: "bar",
    data: {
      labels: rows.map((_, i) => String(i + 1)),
      datasets: [
        {
          label: "EPS",
          data: hostValues,
          backgroundColor: rows.map((_, i) => palette[i % palette.length]),
          borderWidth: 0
        }
      ]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          color: "#111",
          font: { weight: "bold", size: 10 },
          formatter: (v) => {
            const n = Number(v) || 0;
            return Number.isInteger(n) ? String(n) : n.toFixed(2).replace(/\.?0+$/, "");
          },
          anchor: "end",
          align: "top",
          offset: 2,
          clamp: true,
          clip: false
        },
        tooltip: {
          callbacks: {
            label: (ctx) => `${rows[ctx.dataIndex].host}: ${ctx.parsed.y}`
          }
        },
        totPotPlotFrame: true
      },
      scales: {
        x: {
          title: { display: true, text: "Hosts", color: "#000", font: { weight: "bold" } },
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000", font: { weight: "bold" } }
        },
        y: {
          title: { display: true, text: "EPS", color: "#000", font: { weight: "bold" } },
          beginAtZero: true,
          min: 0,
          max: hostAxisMax,
          ticks: { color: "#000" },
          grid: { color: "#c9c9c9", drawBorder: false }
        }
      }
    }
  });

  const legendHost = document.getElementById("epsHostsLegend");
  if (legendHost) {
    legendHost.innerHTML = rows
      .map(
        (r, i) =>
          `<div class="eps-host-item"><span class="eps-host-color" style="background:${palette[i % palette.length]}"></span><span class="eps-host-text">${i + 1} ${escapeHtml(r.host)}</span></div>`
      )
      .join("");
  }
}

function getTrendChartTitle() {
  const custom = getValue("trendChartTitle");
  if (custom) return custom.replace(/\[MONTH\]/g, getReportMonth());
  return `Potential Incident Summary - ${getReportMonth()}`;
}

function getTrendRowsForChart() {
  return ReportCharts.filterTrendRowsWithIncidentData(parseTrendCsvRows(getValue("trendCsv")));
}

function syncTrendSlideLayout() {
  const slide = document.getElementById("slideTrend");
  const wrap = slide?.querySelector(".trend-canvas-wrap");
  if (!wrap) return;
  const rowCount = trendChart?.data?.labels?.length || getTrendRowsForChart().length || 0;
  wrap.style.minHeight = rowCount > 10 ? "240px" : "180px";
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (!trendChart) return;
      const cp = rowCount > 12 ? 0.52 : rowCount > 8 ? 0.62 : 0.72;
      if (!trendChart.options.datasets) trendChart.options.datasets = {};
      if (!trendChart.options.datasets.bar) trendChart.options.datasets.bar = {};
      trendChart.options.datasets.bar.categoryPercentage = cp;
      trendChart.options.datasets.bar.barPercentage = 0.85;
      const ticks = trendChart.options.scales?.x?.ticks;
      if (ticks) {
        ticks.font = {
          size: rowCount > 12 ? 8 : rowCount > 8 ? 9 : 10,
          weight: "bold"
        };
        ticks.maxRotation = rowCount > 10 ? 60 : 45;
        ticks.minRotation = rowCount > 10 ? 45 : 0;
        ticks.autoSkip = false;
        ticks.maxTicksLimit = rowCount;
      }
      trendChart.update("none");
      trendChart.resize();
    });
  });
}

function ensureTrendDataLabelsPlugin() {
  ReportCharts.ensureDataLabelsPlugin();
  trendDataLabelsRegistered = ReportCharts.isDataLabelsRegistered();
}

function renderTrendChart() {
  const canvas = document.getElementById("trendChart");
  if (!canvas || typeof Chart === "undefined") return;
  const rows = getTrendRowsForChart();
  if (!rows.length) {
    if (trendChart) {
      trendChart.destroy();
      trendChart = undefined;
    }
    return;
  }

  ensureTrendDataLabelsPlugin();

  if (trendChart) trendChart.destroy();
  const rowCount = rows.length;
  const categoryPercentage = rowCount > 12 ? 0.52 : rowCount > 8 ? 0.62 : 0.72;
  const tickFontSize = rowCount > 12 ? 8 : rowCount > 8 ? 9 : 10;
  const plugins = ReportCharts.buildSeverityChartPlugins(getTrendChartTitle(), {
    legendPosition: "bottom",
    fontSize: 16,
    legendColor: "#000"
  });
  if (ReportCharts.isDataLabelsRegistered() && plugins.datalabels) {
    plugins.datalabels = {
      ...plugins.datalabels,
      color: "#000000",
      font: { weight: "bold", size: 10, family: "Calibri, Segoe UI, Arial, sans-serif" },
      backgroundColor: "rgba(255,255,255,0.92)",
      borderRadius: 2,
      padding: { top: 1, bottom: 1, left: 3, right: 3 }
    };
  }

  trendChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: rows.map((r) => r.date),
      datasets: ReportCharts.buildSeverityStackedDatasets(rows)
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      datasets: {
        bar: { categoryPercentage, barPercentage: 0.85 }
      },
      plugins,
      scales: {
        x: ReportCharts.buildTrendChartXScaleOptions({
          title: "Date",
          rowCount,
          font: { size: tickFontSize, weight: "bold" },
          maxRotation: rowCount > 10 ? 60 : 45,
          minRotation: rowCount > 10 ? 45 : 0
        }),
        y: ReportCharts.buildTrendChartYScaleOptions(rows, {
          title: "Count",
          font: { weight: "bold" }
        })
      }
    }
  });

  syncTrendSlideLayout();
}

function updateLogos() {
  const snsEl = document.getElementById("snsLogoPreview");
  const clientEl = document.getElementById("clientLogoPreview");
  if (snsEl) {
    snsEl.style.display = snsLogoDataUrl ? "block" : "none";
    snsEl.src = snsLogoDataUrl;
  }
  if (clientEl) {
    clientEl.style.display = clientLogoDataUrl ? "block" : "none";
    clientEl.src = clientLogoDataUrl;
  }
  document.querySelectorAll(".nav-sns-logo").forEach((img) => {
    img.src = snsLogoDataUrl;
    img.style.display = snsLogoDataUrl ? "block" : "none";
  });
  const lastSnsEl = document.getElementById("lastSnsLogoPreview");
  if (lastSnsEl) {
    lastSnsEl.src = snsLogoDataUrl;
    lastSnsEl.style.display = snsLogoDataUrl ? "block" : "none";
  }
}

function upgradeHeaderLogoSlot(page) {
  const header = page.querySelector(".page-header-right");
  if (header && !header.querySelector(".report-logo-slot")) {
    header.innerHTML = '<div class="report-logo-slot"><img class="nav-sns-logo" alt="SNS Logo" /></div>';
  }
}

function applyPremiumLayout() {
  document.querySelectorAll(".page").forEach((page) => {
    if (
      page.classList.contains("cover-hero") ||
      page.classList.contains("revision-page") ||
      page.classList.contains("slide-contact-summary") ||
      page.getAttribute("data-slide") === "contact"
    ) {
      page.querySelector(".page-header-right")?.remove();
      return;
    }
    if (!page.classList.contains("page-with-footer")) {
      page.classList.add("page-with-footer");
    }
    if (!page.querySelector(".page-header-right")) {
      const header = document.createElement("div");
      header.className = "page-header-right";
      header.innerHTML = '<div class="report-logo-slot"><img class="nav-sns-logo" alt="SNS Logo" /></div>';
      page.insertBefore(header, page.firstChild);
    } else {
      upgradeHeaderLogoSlot(page);
    }
    if (!page.querySelector(".page-footer-bar")) {
      const footer = document.createElement("div");
      footer.className = "page-footer-bar";
      page.appendChild(footer);
    }
    const h2 = page.querySelector("h2");
    if (h2 && !h2.classList.contains("revision-title") && !h2.classList.contains("slide-title-lg")) {
      h2.classList.add("revision-title");
    }
  });
  updateLogos();
}

function applyData() {
  setDynamicPrintPageSize();
  syncReportRootLayoutClass();
  const reportRoot = document.getElementById("reportRoot");
  if (reportRoot) void reportRoot.offsetHeight;

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
    execEl.innerHTML = getValue("executiveSummary")
      .split("\n")
      .filter((p) => p.trim())
      .map((p) => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }

  const puzzleEl = document.getElementById("puzzleGraphicPreview");
  if (puzzleEl && puzzleLogoDataUrl) {
    puzzleEl.src = puzzleLogoDataUrl;
    puzzleEl.style.display = "block";
    const ph = document.getElementById("puzzlePlaceholderLayout");
    if (ph) ph.style.display = "none";
  } else if (puzzleEl && !puzzleLogoDataUrl) {
    puzzleEl.style.display = "none";
    const ph = document.getElementById("puzzlePlaceholderLayout");
    if (ph) ph.style.display = "flex";
  }

  const monthLbl = getTableMonthLabel();
  updateTotPotIncTable(monthLbl);
  const trendNote = document.getElementById("trendNotePreview");
  if (trendNote) {
    trendNote.innerHTML = formatRichTextHTML(getValue("trendNote"));
  }
  const trendNarr = document.getElementById("trendNarrativePreview");
  if (trendNarr) {
    trendNarr.innerHTML = formatRichTextHTML(getValue("trendNarrative"));
  }
  updateLogos();
  void renderRiskSlides();
  applyPremiumLayout();
  renderTotPotIncChart();
  renderTrendChart();
  renderRuleSeverityChart();
  updateSlaSlide();
  renderIntegratedInventorySlide();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      syncEngagementLayout();
      if (totPotIncChart) totPotIncChart.resize();
      syncTrendSlideLayout();
      window.dispatchEvent(new Event("resize"));
    });
  });
  if (typeof reapplyTextStyleOverrides === "function") {
    reapplyTextStyleOverrides();
  }
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

async function exportPptx() {
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
  await renderRiskSlides();
  await waitForChartsReady(400);

  const layoutKey = getPptLayoutKey();
  const { w: slideW, h: slideH } = getPptSlideInches();
  const sx = slideW / PPT_EXPORT_REF.w;
  const sy = slideH / PPT_EXPORT_REF.h;

  const pptx = new PptxGenJS();
  pptx.layout = pptLayoutConfig.LAYOUT_16X9.pptxLayout;
  pptx.title = `${getValue("customerName") || "Customer"} Monthly SOC Report`;

  const margin = 0.5 * sx;
  const contentW = slideW - 2 * margin;
  const logoW = 1.2;
  const logoX = slideW - logoW - 0.06 * sx;

  const titleStyle = { fontFace: "Calibri" };
  const barChartType = pptx.ChartType ? pptx.ChartType.bar : "bar";

  let slide = pptx.addSlide();
  // Cover background to match web preview tone
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: slideW,
    h: slideH,
    line: { color: "FAFAFC", pt: 0 },
    fill: { color: "FAFAFC" }
  });
  // Client logo card (same concept as web card)
  slide.addShape(pptx.ShapeType.rect, {
    x: 2.7 * sx,
    y: 1.0 * sy,
    w: 4.6 * sx,
    h: 1.55 * sy,
    line: { color: "EFEFF2", pt: 1 },
    fill: { color: "FFFFFF" }
  });
  slide.addText("MANAGED INCIDENT RESPONSE &\nREMEDIATION SERVICE", {
    x: margin,
    y: 3.15 * sy,
    w: contentW,
    h: 1.0 * sy,
    fontSize: 24,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  slide.addText(`Monthly SOC Report | ${getReportMonth()}`, {
    x: 1.4 * sx,
    y: 4.6 * sy,
    w: 7.2 * sx,
    h: 0.4 * sy,
    fontSize: 16,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (clientLogoDataUrl) {
    slide.addImage({ data: clientLogoDataUrl, x: 3.15 * sx, y: 1.2 * sy, w: 3.7 * sx, h: 1.15 * sy });
  }
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 5.05 * sy,
    w: slideW,
    h: 0.575 * sy,
    line: { color: "042E5F", pt: 0 },
    fill: { color: "042E5F" }
  });

  slide = pptx.addSlide();
  slide.addText("Document Revision History", {
    x: 0.7 * sx,
    y: 0.85 * sy,
    w: 8.6 * sx,
    h: 0.55 * sy,
    fontSize: 24,
    bold: true,
    color: "000000",
    align: "left",
    ...titleStyle
  });
  const revTable = [
    [{ text: "Version", options: { bold: true, fill: { color: "FFFFFF" } } }, { text: "1.0", options: { bold: true, fill: { color: "FFFFFF" } } }],
    [{ text: "Dated", options: { fill: { color: "E1F4F4" } } }, { text: getValue("dateRange"), options: { fill: { color: "E1F4F4" } } }],
    [{ text: "Prepared By", options: { fill: { color: "FFFFFF" } } }, { text: getValue("preparedBy"), options: { fill: { color: "FFFFFF" } } }],
    [{ text: "Reviewed By", options: { fill: { color: "E1F4F4" } } }, { text: STATIC_REVIEWED_BY, options: { fill: { color: "E1F4F4" } } }],
    [{ text: "Approved By", options: { fill: { color: "FFFFFF" } } }, { text: STATIC_APPROVED_BY, options: { fill: { color: "FFFFFF" } } }],
    [{ text: "Submitted On", options: { fill: { color: "E1F4F4" } } }, { text: getValue("submittedOn"), options: { fill: { color: "E1F4F4" } } }]
  ];
  slide.addTable(revTable, {
    x: 0.7 * sx,
    y: 1.65 * sy,
    w: 8.6 * sx,
    colW: [2.8 * sx, 5.8 * sx],
    fontSize: 12,
    border: { pt: 0, color: "FFFFFF" },
    color: "111111",
    valign: "mid"
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 4.95 * sy,
    w: slideW,
    h: 0.675 * sy,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });

  slide = pptx.addSlide();
  slide.addText("The Engagement", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  slide.addText(
    [{ text: `${getValue("customerName") || "Customer"} has engaged with SNS to monitor and review the entity's security.`, options: { bold: true } }],
    {
      x: 0.65 * sx,
      y: 1.2 * sy,
      w: 5.2 * sx,
      h: 0.5 * sy,
      fontSize: 11,
      color: "111111",
      ...titleStyle
    }
  );
  slide.addText(parseRichTextPptx(getValue("executiveSummary")), {
    x: margin,
    y: 1.8 * sy,
    w: 5.3 * sx,
    h: 3.45 * sy,
    fontSize: 10.5,
    color: "000000",
    valign: "top",
    ...titleStyle
  });
  if (puzzleLogoDataUrl) {
    slide.addImage({ data: puzzleLogoDataUrl, x: 6.2 * sx, y: 1.6 * sy, w: 3.1 * sx, h: 3.0 * sy });
  }
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 4.95 * sy,
    w: slideW,
    h: 0.675 * sy,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });

  const monthLabel = getTableMonthLabel();
  const totPotEntries = getTotPotIncSeverityEntries();
  const maxAlert = Math.max(...totPotEntries.map((entry) => entry.value), 1);
  const yAxisMax = Math.max(10, Math.ceil((maxAlert * 1.1) / 5) * 5);
  const yAxisMajorUnit = yAxisMax <= 10 ? 2 : 5;

  // Slide 4: Total Potential Incident
  slide = pptx.addSlide();
  slide.addText("Total Potential Incident", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  if (totPotEntries.length) {
    slide.addChart(
      barChartType,
      totPotEntries.map((entry) => ({
        name: entry.label,
        labels: [monthLabel],
        values: [entry.value]
      })),
      {
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.95 * sy,
        barDir: "col",
        barGrouping: "clustered",
        chartColors: totPotEntries.map((entry) => entry.pptColor),
        showLegend: true,
        legendPos: "b",
        valAxisMaxVal: yAxisMax,
        valAxisMinVal: 0,
        valAxisMajorUnit: yAxisMajorUnit,
        showValue: true,
        dataLabelPosition: "outEnd",
        dataLabelFontSize: 11,
        dataLabelColor: "000000",
        dataLabelFormatCode: "#,##0",
        showDataTable: true,
        dataTableFontSize: 9,
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 10
      }
    );
  }

  // Slide 5: Potential Incident Tickets Trend
  slide = pptx.addSlide();
  slide.addText("Potential Incident Tickets Trend", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const trendCanvasForPpt = document.getElementById("trendChart");
  if (trendCanvasForPpt) {
    slide.addImage({ data: toHighResCanvasDataUrl(trendCanvasForPpt, 2), x: 0.45 * sx, y: 0.9 * sy, w: 9.1 * sx, h: 3.3 * sy });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), {
    x: 0.5 * sx,
    y: 4.35 * sy,
    w: 9.0 * sx,
    h: 0.3 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), {
    x: 0.5 * sx,
    y: 4.7 * sy,
    w: 9.0 * sx,
    h: 0.7 * sy,
    fontSize: 10.5,
    color: "000000",
    fontFace: "Calibri"
  });


  // Slide 6+: Risk Mitigation (one or more slides — matches web preview pagination)
  const riskRowsPptx = parseRiskCsvRows(getValue("riskCsv"));
  const riskChunksPptx = chunkRiskRowsForPptx(riskRowsPptx);
  const narrative = parseRichTextPptx(getValue("riskNarrative"));
  const riskCellOpts = { fontFace: "Calibri", valign: "middle" };

  riskChunksPptx.forEach((chunk, i) => {
    slide = pptx.addSlide();
    const riskLayout = getPptRiskSlideLayout(i, chunk.length, sy);

    slide.addText(i === 0 ? RISK_SLIDE_TITLE : RISK_CONTD_TITLE, {
      x: margin,
      y: riskLayout.titleY,
      w: contentW,
      h: riskLayout.titleH,
      ...titleStyle,
      align: "center",
      fontSize: 24,
      bold: true,
      color: "000000"
    });
    if (snsLogoDataUrl) {
      slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
    }

    if (riskLayout.narrativeH > 0) {
      slide.addText(narrative, {
        x: margin,
        y: riskLayout.narrativeY,
        w: contentW,
        h: riskLayout.narrativeH,
        fontSize: 10.5,
        color: "111111",
        fontFace: "Calibri",
        align: "justify",
        valign: "top"
      });
    }

    const riskTableData = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } },
        { text: "Attack Type", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } },
        { text: "Risk Scenario", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } },
        { text: "Type of Risk(s)\nCIA Triad", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } },
        { text: "Potential Business Impact(s)", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } },
        { text: "Risk Rating", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", ...riskCellOpts } }
      ],
      ...chunk.map((r) => [
        { text: String(r.sno), options: { align: "center", ...riskCellOpts } },
        { text: r.attackType, options: { align: "center", ...riskCellOpts } },
        { text: r.riskScenario, options: { align: "center", ...riskCellOpts } },
        { text: r.ciaTriad, options: { align: "center", ...riskCellOpts } },
        { text: r.businessImpact, options: { align: "center", ...riskCellOpts } },
        { text: r.riskRating, options: { align: "center", bold: true, ...riskCellOpts } }
      ])
    ];

    const riskFont =
      getRiskLayoutOptions().tableFontSizePx > 0
        ? Math.min(18, Math.max(8, getRiskLayoutOptions().tableFontSizePx * 0.64))
        : 9;
    slide.addTable(riskTableData, {
      x: margin,
      y: riskLayout.tableY,
      w: contentW,
      h: riskLayout.tableH,
      rowH: riskLayout.rowH,
      fontSize: riskFont,
      colW: [0.38 * sx, 1.35 * sx, 1.95 * sx, 1.15 * sx, 3.35 * sx, 0.92 * sx],
      border: { type: "solid", color: "7F7F7F", pt: 0.5 },
      autoPage: false,
      valign: "middle"
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0,
      y: riskLayout.footerTop,
      w: slideW,
      h: 0.675 * sy,
      line: { color: "1E4F9A", pt: 0 },
      fill: { color: "1E4F9A" }
    });
  });

  // Slide 7: Rule-Based Severity Categories
  const ruleRows = parseRuleSeverityCsvRows(getValue("ruleSeverityCsv"));
  if (ruleRows.length) {
    slide = pptx.addSlide();
    slide.addText("Rule-Based Severity Categories For Potential Incident", {
      x: margin,
      y: 0.6 * sy,
      w: contentW,
      h: 0.55 * sy,
      fontSize: 28,
      bold: true,
      color: "000000",
      align: "center",
      ...titleStyle
    });
    if (snsLogoDataUrl) {
      slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
    }

    const ruleCanvas = document.getElementById("ruleSeverityChart");
    if (ruleCanvas) {
      slide.addImage({
        data: ruleCanvas.toDataURL("image/png"),
        x: 0.5 * sx,
        y: 1.4 * sy,
        w: 9.0 * sx,
        h: 3.8 * sy
      });
    }
  }

  // Slide 8: SLA Performance
  const incidentCount = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const closedIncidentCount = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const { response: responseInfo, remediation: remediationInfo } = getCalculatedSlaInfo();
  const customer = getValue("customerName") || "customer";

  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("Response Time SLA", {
    x: 0.5 * sx,
    y: 0.6 * sy,
    w: 2.8 * sx,
    h: 0.4 * sy,
    fontSize: 22,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("Remediation Time SLA", {
    x: 0.5 * sx,
    y: 3.0 * sy,
    w: 3.2 * sx,
    h: 0.4 * sy,
    fontSize: 22,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });

  const responseCanvas = document.getElementById("responseSlaChart");
  if (responseCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(responseCanvas, 2),
      x: 0.35 * sx,
      y: 1.0 * sy,
      w: 2.85 * sx,
      h: 1.75 * sy
    });
  }
  const remediationCanvas = document.getElementById("remediationSlaChart");
  if (remediationCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(remediationCanvas, 2),
      x: 0.35 * sx,
      y: 3.45 * sy,
      w: 2.85 * sx,
      h: 1.75 * sy
    });
  }

  const slaHeaderFill = { color: "FFFFFF" };
  const slaHeaderOpts = { bold: true, fill: slaHeaderFill, align: "center", fontFace: "Calibri" };
  const slaTableW = 6.15 * sx;
  const responseColW = [0.48, 0.12, 0.14, 0.26].map((p) => p * slaTableW);

  const responseTbl = [
    [
      { text: "Severity Level Description", options: slaHeaderOpts },
      { text: "Severity Level", options: slaHeaderOpts },
      { text: "Response", options: slaHeaderOpts },
      { text: "Resolution/Remediation", options: slaHeaderOpts }
    ],
    [
      { text: "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community", options: { fill: { color: "EEF6F9" } } },
      { text: "S-1", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "≤15 Minutes", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "4 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      { text: "S-2", options: { align: "center" } },
      { text: "≤15 Minutes", options: { align: "center" } },
      { text: "8 Hours", options: { align: "center" } }
    ],
    [
      { text: "Low Severity A Request/Query/Service that does not change existing service structure and no cost implication", options: { fill: { color: "EEF6F9" } } },
      { text: "S-3", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "≤30 Minutes", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "24 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ]
  ];
  slide.addTable(responseTbl, {
    x: 3.4 * sx,
    y: 0.75 * sy,
    w: slaTableW,
    fontSize: 8.5,
    colW: responseColW,
    border: { pt: 0.6, color: "C8D8DD" },
    valign: "middle",
    fontFace: "Calibri"
  });
  slide.addText(formatResponseSummaryText(incidentCount, customer, responseInfo.label), {
    x: 3.45 * sx,
    y: 2.95 * sy,
    w: 6.2 * sx,
    h: 0.35 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri"
  });

  const remediationColW = [0.58, 0.14, 0.28].map((p) => p * slaTableW);

  const remediationTbl = [
    [
      { text: "Severity Level Description", options: slaHeaderOpts },
      { text: "Severity Level", options: slaHeaderOpts },
      { text: "Resolution/Remediation", options: slaHeaderOpts }
    ],
    [
      { text: "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community", options: { fill: { color: "EEF6F9" } } },
      { text: "High", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "4 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      { text: "Medium", options: { align: "center" } },
      { text: "8 Hours", options: { align: "center" } }
    ],
    [
      { text: "Low A Request/Query/Service that does not change existing service structure and no cost implication", options: { fill: { color: "EEF6F9" } } },
      { text: "Low", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "24 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ]
  ];
  slide.addTable(remediationTbl, {
    x: 3.4 * sx,
    y: 3.2 * sy,
    w: slaTableW,
    fontSize: 8.5,
    colW: remediationColW,
    border: { pt: 0.6, color: "C8D8DD" },
    valign: "middle",
    fontFace: "Calibri"
  });
  slide.addText(formatRemediationSummaryText(closedIncidentCount, remediationInfo.label), {
    x: 3.45 * sx,
    y: 4.88 * sy,
    w: 6.2 * sx,
    h: 0.45 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri",
    valign: "top"
  });

  // Slide 9: Integrated Device Inventory
  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("Integrated Device Inventory", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 32,
    bold: true,
    color: "000000",
    align: "center",
    fontFace: "Calibri"
  });

  const invRows = getInventoryRows();
  const invNoteRaw = getValue("inventoryNote");
  const invLayout = getInventoryPptLayout(invRows.length, invNoteRaw);
  const invTbl = buildInventoryPptTable(invRows);

  slide.addTable(invTbl, {
    x: invLayout.tableX * sx,
    y: invLayout.tableTop * sy,
    w: invLayout.tableW * sx,
    fontSize: invLayout.fontSize,
    rowH: invLayout.rowH * sy,
    colW: invLayout.colW.map((w) => w * sx),
    border: { pt: 0.5, color: "FFFFFF" },
    valign: "middle"
  });

  if (invNoteRaw.trim()) {
    slide.addText(invNoteRaw, {
      x: invLayout.tableX * sx,
      y: invLayout.noteY * sy,
      w: invLayout.tableW * sx,
      h: invLayout.noteH * sy,
      fontSize: invLayout.noteFontSize,
      color: "000000",
      align: "left",
      fontFace: "Calibri",
      valign: "top",
      margin: 0
    });
  }

  exportPptxContact(pptx);

  await pptx.writeFile(`Nocil_Monthly_SOC_Report.pptx`);
  } finally {
    restore();
    applyData();
    if (pptxBtn) {
      pptxBtn.disabled = false;
      pptxBtn.textContent = originalLabel || "Export PPTX";
    }
  }
}

function bindImageInput(inputId, setter) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        setter(String(e.target.result));
        applyData();
      };
      reader.readAsDataURL(file);
    }
  });
}

function bindCsvInput(fileInputId, textareaId) {
  const input = document.getElementById(fileInputId);
  const target = document.getElementById(textareaId);
  if (!input || !target) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      target.value = String(e.target.result || "");
      applyData();
    };
    reader.readAsText(file);
  });
}

async function init() {
  let panelApplyScheduled = false;
  function schedulePanelApply(event) {
    if (panelApplyScheduled) return;
    panelApplyScheduled = true;
    const inputId = event?.target?.id;
    requestAnimationFrame(() => {
      panelApplyScheduled = false;
      if (inputId && SLA_INPUT_IDS.has(inputId)) {
        handleSlaInputChange(inputId);
        return;
      }
      applyData();
    });
  }
  document.querySelectorAll(".sidebar input, .sidebar textarea, .sidebar select").forEach((el) => {
    el.addEventListener("input", schedulePanelApply);
    el.addEventListener("change", schedulePanelApply);
  });
  document.getElementById("printBtn").addEventListener("click", async (e) => {
    e.preventDefault();
    const printBtn = document.getElementById("printBtn");
    if (printBtn?.disabled) return;
    const originalLabel = printBtn ? printBtn.textContent : "";
    if (printBtn) {
      printBtn.disabled = true;
      printBtn.textContent = "Preparing…";
    }
    try {
      finishExportRestore();
      exportRestoreCallback = await prepareReportForExport();
      setDynamicPrintPageSize();
      if (typeof preparePrintSlideOrder === "function") {
        preparePrintSlideOrder("nocil");
      }
      finishNocilPrintLayout();
      reportPdfLayoutRestore = prepareNocilPrintLayout();
      const engagementSlide = document.querySelector(
        '#reportRoot .page[data-slide="engagement"]'
      );
      if (engagementSlide) {
        engagementSlide.scrollIntoView({ block: "center" });
      }
      await new Promise((resolve) => requestAnimationFrame(resolve));
      updateSlaSlide();
      await new Promise((resolve) => requestAnimationFrame(resolve));
      updateSlaSlide();
      syncEngagementLayout();
      await new Promise((resolve) => requestAnimationFrame(resolve));
      syncEngagementLayout();
      await new Promise((resolve) => setTimeout(resolve, 150));
      syncEngagementLayout();
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
  });
  window.addEventListener("afterprint", () => {
    finishExportRestore();
    const printBtn = document.getElementById("printBtn");
    if (printBtn) {
      printBtn.disabled = false;
      printBtn.textContent = "Export PDF";
    }
  });
  document.getElementById("pptxBtn").addEventListener("click", (e) => {
    e.preventDefault();
    exportPptx();
  });
  
  const menuToggle = document.getElementById("menuToggle");
  const appShell = document.getElementById("appShell");
  if (menuToggle && appShell) {
    menuToggle.addEventListener("click", () => {
      appShell.classList.toggle("sidebar-collapsed");
      setTimeout(updateSlaSlide, 300);
    });
  }

  let slaResizeTimer;
  window.addEventListener("resize", () => {
    clearTimeout(slaResizeTimer);
    slaResizeTimer = setTimeout(updateSlaSlide, 150);
  });

  ["snsLogoInput", "clientLogoInput", "puzzleLogoInput"].forEach((id) => {
    bindImageInput(id, (v) => {
      if (id === "snsLogoInput") snsLogoDataUrl = v;
      if (id === "clientLogoInput") clientLogoDataUrl = v;
      if (id === "puzzleLogoInput") puzzleLogoDataUrl = v;
    });
  });
  bindCsvInput("trendCsvFile", "trendCsv");
  bindCsvInput("riskCsvFile", "riskCsv");
  initRiskTableUserControls();
  bindCsvInput("ruleSeverityCsvFile", "ruleSeverityCsv");
  bindCsvInput("inventoryCsvFile", "inventoryCsv");

  // Keep Nocil rule-severity chart aligned with the approved reference slide
  // whenever no CSV has been uploaded yet.
  const ruleSeverityEl = document.getElementById("ruleSeverityCsv");
  if (ruleSeverityEl && !String(ruleSeverityEl.value || "").trim()) {
    ruleSeverityEl.value = DEFAULT_NOCIL_RULE_SEVERITY_CSV;
  }

  applyData();
  if (typeof initSlideNav === "function") {
    initSlideNav(() => "nocil");
  }
  if (typeof initTextStyleEditor === "function") {
    initTextStyleEditor();
  }
}

window.onload = init;
