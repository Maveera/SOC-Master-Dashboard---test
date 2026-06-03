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
let defaultPuzzleDataUrl = "";
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
let slaPieCenterPluginRegistered = false;

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
let reportPdfLayoutRestore = null;

const RISK_SLIDE_TITLE = "Potential Alerts - Risks Mitigated";
const RISK_CONTD_TITLE = "Contn.,";

function syncEngagementLayout() {
  if (typeof ReportEngagementLayout !== "undefined") {
    ReportEngagementLayout.syncEngagementLayout();
  }
}

function finishReportPdfLayout() {
  finishBluPinePrintLayout();
}

function finishBluPinePrintLayout() {
  if (reportPdfLayoutRestore) {
    reportPdfLayoutRestore();
    reportPdfLayoutRestore = null;
  }
}

async function waitForChartsReady(delayMs = 900) {
  const charts = [
    totPotIncChart,
    fortiSiemAlertsChart,
    truePositiveAlertsChart,
    falsePositiveAlertsChart,
    epsTotalChart,
    epsTopHostsChart,
    trendChart,
    ruleSeverityChart,
    responseSlaChart,
    remediationSlaChart
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
  await ensureDefaultPuzzleDataUrl();
  document.body.classList.add("is-exporting");
  applyData();
  await renderRiskSlides();
  await waitForChartsReady();
  syncAllSlideLayoutsForExport();
  if (typeof initSlaTableColumnResize === "function") {
    initSlaTableColumnResize(document.getElementById("slideSlaStatus"));
  }
  return () => {
    document.body.classList.remove("is-exporting");
  };
}

function syncAllSlideLayoutsForExport() {
  updateSlaSlide();
  syncEngagementPuzzleImage();
  syncEngagementLayout();
  if (totPotIncChart) {
    totPotIncChart.resize();
    syncTotpotTableLayout(totPotIncChart, "#slideTotPotInc .totpot-table-legend-wrap");
  }
  if (fortiSiemAlertsChart) {
    fortiSiemAlertsChart.resize();
    syncFortiSiemSlideLayout();
  }
  if (truePositiveAlertsChart || falsePositiveAlertsChart) {
    syncTfPosSlideLayout();
    syncTfPosTableLayouts();
  }
  syncEpsSlideLayout();
  syncEpsEventsSlideLayout();
  syncSupportTicketsSlideLayout();
  if (trendChart) trendChart.resize();
  if (ruleSeverityChart) ruleSeverityChart.resize();
  renderKeyPointsPreview();
  window.dispatchEvent(new Event("resize"));
}

function prepareBluPinePrintLayout() {
  if (typeof ReportEngagementLayout === "undefined") {
    return () => {};
  }
  return ReportEngagementLayout.prepareReportPdfLayout(() => {
    syncEngagementPuzzleImage();
    syncEngagementLayout();
    updateSlaSlide();
  });
}

function finishExportRestore() {
  finishReportPdfLayout();
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
  const monthReplaced = text.replace(/\[MONTH\]/g, `**${getReportMonth()}**`);
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

function getKeyPointsSourceText() {
  return getValue("keyPointsSummary");
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
  return sections
    .map((section) => {
      if (section.heading) {
        headingNum += 1;
        const label = `${headingNum}. ${escapeHtml(section.heading)}:`;
        const bodyHtml = section.body
          ? `<div class="kp-body">${formatRichTextHTML(section.body)}</div>`
          : "";
        return `<div class="kp-block"><p class="kp-label">${label}</p>${bodyHtml}</div>`;
      }
      return `<div class="kp-block"><div class="kp-body">${formatRichTextHTML(section.body)}</div></div>`;
    })
    .join("");
}

function renderKeyPointsPreview() {
  const host = document.getElementById("keyPointsSummaryPreview");
  if (!host) return;
  if (shouldSkipApplyForNode(host)) return;
  host.innerHTML = renderKeyPointsHtml(getKeyPointsSourceText());
}

function estimateKeyPointsBodyHeight(body) {
  const lines = String(body || "").split("\n").length;
  return Math.min(2.8, Math.max(0.45, lines * 0.17));
}

function addKeyPointsToPptxSlide(slide, sx, sy, yStart = 1.35) {
  const sections = parseKeyPointsSections(getKeyPointsSourceText());
  let y = yStart * sy;
  let headingNum = 0;
  const fontSize = 11;
  const txtBase = { fontSize, fontFace: "Calibri" };

  sections.forEach((section) => {
    if (section.heading) {
      headingNum += 1;
      slide.addText(`${headingNum}. ${section.heading}:`, {
        x: 0.65 * sx,
        y,
        w: 8.7 * sx,
        h: 0.28 * sy,
        bold: true,
        ...txtBase
      });
      y += 0.26 * sy;
    }
    if (section.body) {
      const h = estimateKeyPointsBodyHeight(section.body) * sy;
      slide.addText(parseRichTextPptx(section.body, txtBase), {
        x: 0.65 * sx,
        y,
        w: 8.7 * sx,
        h,
        valign: "top",
        ...txtBase
      });
      y += h + 0.12 * sy;
    }
  });
}

const INLINE_EDITABLE_NARRATIVES = {
  executiveSummaryPreview: "executiveSummary",
  trendNotePreview: "trendNote",
  trendNarrativePreview: "trendNarrative",
  keyPointsSummaryPreview: "keyPointsSummary",
  inventoryNotePreview: "inventoryNote",
  tfPosNoteTruePreview: "tfPosNoteTrue",
  tfPosNoteFalsePreview: "tfPosNoteFalse"
};

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeMonthPlaceholderFromMarkdown(source) {
  const m = getReportMonth();
  return String(source || "").replace(new RegExp(`\\*\\*${escapeRegExp(m)}\\*\\*`, "g"), "[MONTH]");
}

function shouldSkipApplyForNode(node) {
  if (!node || typeof node.contains !== "function") return false;
  const ae = document.activeElement;
  if (!ae || !node.contains(ae)) return false;
  return ae.isContentEditable === true;
}

function inlineHtmlToMarkdown(el) {
  let out = "";
  el.childNodes.forEach((node) => {
    if (node.nodeType === 3) {
      out += node.textContent;
    } else if (node.nodeType === 1) {
      const tag = node.tagName;
      if (tag === "STRONG" || tag === "B" || node.classList.contains("fw-bold")) {
        out += `**${node.textContent}**`;
      } else if (tag === "EM" || tag === "I") {
        out += `*${node.textContent}*`;
      } else if (tag === "BR") {
        out += "\n";
      } else if (tag === "P") {
        const inner = inlineHtmlToMarkdown(node);
        out += (out ? "\n\n" : "") + inner;
      } else {
        out += inlineHtmlToMarkdown(node);
      }
    }
  });
  return out;
}

function narrativeHtmlToMarkdownSource(html) {
  const wrap = document.createElement("div");
  wrap.innerHTML = String(html || "").trim();
  const blocks = wrap.querySelectorAll(":scope > .kp-block");
  if (blocks.length) {
    const parts = [];
    blocks.forEach((block) => {
      const label = block.querySelector(".kp-label");
      const body = block.querySelector(".kp-body");
      if (label) {
        const heading = label.textContent.replace(/^\d+\.\s*/, "").replace(/:$/, "").trim();
        const bodyText = body ? inlineHtmlToMarkdown(body).trim() : "";
        parts.push(bodyText ? `**${heading}:**\n${bodyText}` : `**${heading}:**`);
      } else if (body) {
        parts.push(inlineHtmlToMarkdown(body).trim());
      }
    });
    return parts.filter(Boolean).join("\n\n");
  }
  const ps = wrap.querySelectorAll(":scope > p");
  if (ps.length) {
    return Array.from(ps)
      .map((p) => inlineHtmlToMarkdown(p))
      .join("\n\n");
  }
  return inlineHtmlToMarkdown(wrap);
}

function addEngagementNativeSlide(pptx, slideW, slideH) {
  const slide = pptx.addSlide();
  const customer = getValue("customerName") || "Customer";
  const summaryLines = getValue("executiveSummary")
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);
  const puzzleData = puzzleLogoDataUrl || defaultPuzzleDataUrl || "";
  const baseTxt = { fontFace: "Times New Roman", fontSize: 11, color: "000000" };
  const textParts = [
    { text: customer, options: { ...baseTxt, bold: true } },
    {
      text: " has engaged with SNS to monitor and review the entity's security.",
      options: { ...baseTxt }
    }
  ];
  summaryLines.forEach((line, idx) => {
    textParts.push({ text: idx === 0 ? "\n\n" : "\n\n", options: { ...baseTxt } });
    textParts.push(...parseRichTextPptx(line, baseTxt));
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: slideW - 1.35, y: 0.18, w: 1.2, h: 0.5 });
  }
  slide.addText("The Engagement", {
    x: 0.5,
    y: 0.35,
    w: 4.2,
    h: 0.5,
    fontSize: 24,
    bold: true,
    color: "000000",
    fontFace: "Times New Roman"
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.48,
    y: 0.86,
    w: 4.58,
    h: 3.92,
    line: { color: "FFFFFF", pt: 0 },
    fill: { color: "F6FBFB" }
  });
  slide.addText(textParts, {
    x: 0.58,
    y: 0.94,
    w: 4.38,
    h: 3.76,
    valign: "top",
    fontFace: "Times New Roman",
    color: "000000"
  });
  if (puzzleData) {
    slide.addImage({ data: puzzleData, x: 5.2, y: 0.9, w: 4.25, h: 3.85 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: slideH - 0.42,
    w: slideW,
    h: 0.42,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });
  return slide;
}

function initInlineEditableReport() {
  const report = document.getElementById("reportRoot");
  if (!report) return;
  if (report.dataset.inlineEditableBound === "1") return;
  report.dataset.inlineEditableBound = "1";

  const skipField = new Set(["month", "reviewedBy", "approvedBy"]);
  report.querySelectorAll("[data-field]").forEach((el) => {
    const key = el.getAttribute("data-field");
    if (!key || skipField.has(key)) return;
    if (!fields[key]) return;
    el.setAttribute("contenteditable", "true");
    el.setAttribute("spellcheck", "true");
    el.classList.add("report-editable-field");
  });

  Object.keys(INLINE_EDITABLE_NARRATIVES).forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.setAttribute("contenteditable", "true");
    el.setAttribute("spellcheck", "true");
    el.classList.add("report-editable-narrative");
    el.dataset.syncField = INLINE_EDITABLE_NARRATIVES[id];
  });

  const numericCells = [
    ["totPotIncTableHigh", "totPotIncHigh"],
    ["totPotIncTableMed", "totPotIncMedium"],
    ["totPotIncTableLow", "totPotIncLow"],
    ["fortiSiemTableHigh", "fortiAlertHigh"],
    ["fortiSiemTableMed", "fortiAlertMedium"],
    ["fortiSiemTableLow", "fortiAlertLow"],
    ["tfPosTrueTableHigh", "tfPosTrueHigh"],
    ["tfPosTrueTableMed", "tfPosTrueMedium"],
    ["tfPosTrueTableLow", "tfPosTrueLow"],
    ["tfPosFalseTableHigh", "tfPosFalseHigh"],
    ["tfPosFalseTableMed", "tfPosFalseMedium"],
    ["tfPosFalseTableLow", "tfPosFalseLow"]
  ];
  numericCells.forEach(([cellId, inputId]) => {
    const el = document.getElementById(cellId);
    if (!el) return;
    el.setAttribute("contenteditable", "true");
    el.classList.add("report-editable-field");
    el.dataset.syncInput = inputId;
  });

  report.addEventListener("focusout", (e) => {
    const t = e.target;
    if (!t.closest || !t.closest("#reportRoot")) return;
    const rel = e.relatedTarget;

    const syncCell = t.closest("[data-sync-input]");
    if (syncCell && syncCell.dataset.syncInput) {
      if (rel && syncCell.contains(rel)) return;
      const inp = document.getElementById(syncCell.dataset.syncInput);
      if (inp) inp.value = syncCell.textContent.replace(/\s+/g, " ").trim();
      applyData();
      return;
    }

    const narrative = t.closest(".report-editable-narrative");
    if (narrative && narrative.dataset.syncField) {
      if (rel && narrative.contains(rel)) return;
      const inp = document.getElementById(narrative.dataset.syncField);
      if (inp) {
        let src = narrativeHtmlToMarkdownSource(narrative.innerHTML);
        src = normalizeMonthPlaceholderFromMarkdown(src);
        inp.value = src;
      }
      applyData();
      return;
    }

    const fieldEl = t.closest("[data-field].report-editable-field");
    const key = fieldEl && fieldEl.getAttribute("data-field");
    if (fieldEl && key && fields[key]) {
      if (rel && fieldEl.contains(rel)) return;
      const inp = document.getElementById(fields[key]);
      if (inp) inp.value = fieldEl.textContent.trim();
      applyData();
    }
  });
}

const DRAWING_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function pptxIsInTable(pEl) {
  let x = pEl.parentElement;
  while (x) {
    if (x.localName === "tbl" && x.namespaceURI === DRAWING_A_NS) return true;
    x = x.parentElement;
  }
  return false;
}

function extractParagraphsFromSlideXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const paras = doc.getElementsByTagNameNS(DRAWING_A_NS, "p");
  const lines = [];
  for (let i = 0; i < paras.length; i++) {
    const pEl = paras[i];
    if (pptxIsInTable(pEl)) continue;
    const ts = pEl.getElementsByTagNameNS(DRAWING_A_NS, "t");
    let line = "";
    for (let j = 0; j < ts.length; j++) line += ts[j].textContent || "";
    const t = line.replace(/\r/g, "").trim();
    if (t) lines.push(t);
  }
  return lines;
}

function extractTablesFromSlideXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const tbls = doc.getElementsByTagNameNS(DRAWING_A_NS, "tbl");
  const tables = [];
  for (let ti = 0; ti < tbls.length; ti++) {
    const rows = [];
    const trs = tbls[ti].getElementsByTagNameNS(DRAWING_A_NS, "tr");
    for (let ri = 0; ri < trs.length; ri++) {
      const cells = [];
      const tcs = trs[ri].getElementsByTagNameNS(DRAWING_A_NS, "tc");
      for (let ci = 0; ci < tcs.length; ci++) {
        const txBody = tcs[ci].getElementsByTagNameNS(DRAWING_A_NS, "txBody")[0];
        let cellText = "";
        if (txBody) {
          const ps = txBody.getElementsByTagNameNS(DRAWING_A_NS, "p");
          for (let pi = 0; pi < ps.length; pi++) {
            const ts = ps[pi].getElementsByTagNameNS(DRAWING_A_NS, "t");
            for (let tj = 0; tj < ts.length; tj++) cellText += ts[tj].textContent || "";
            if (pi < ps.length - 1) cellText += "\n";
          }
        }
        cells.push(cellText.trim());
      }
      rows.push(cells);
    }
    tables.push(rows);
  }
  return tables;
}

function csvEscapeCell(v) {
  const s = String(v ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function csvEscapeRow(cells) {
  return cells.map(csvEscapeCell).join(",");
}

function classifySlideKind(paragraphs) {
  const head = (paragraphs[0] || "").trim();
  if (head.includes("Document Revision History")) return "revision";
  if (head.includes("The Engagement")) return "engagement";
  if (head.includes("Total Potential Incident")) return "totPot";
  if (head.includes("Potential Incident Tickets Trend")) return "trend";
  if (head.includes("Potential Alerts - Risks Mitigated")) return "risk";
  if (head.includes("Rule-Based Severity")) return "ruleSeverity";
  if (head.includes("Response Time SLA") || head.includes("Remediation Time SLA")) return "sla";
  if (head.includes("Total Alerts Triggered in FortiSIEM")) return "forti";
  if (head.includes("Total Number Of True")) return "tfPos";
  if (head.includes("EPS Trend Plot")) return "epsTrend";
  if (head.includes("Highest EPS Consuming Events")) return "epsEvents";
  if (head.includes("Overall Support Ticket")) return "support";
  if (head.includes("Integrated Device Inventory")) return "inventory";
  if (head.includes("Key Points")) return "keyPoints";
  if (head.includes("Your Trusted Security Advisor")) return "contact";
  return "unknown";
}

function normalizeMonthForImportText(text) {
  const m = getReportMonth();
  return String(text || "").replace(new RegExp(escapeRegExp(m), "g"), "[MONTH]");
}

function importFromRevisionParagraphs(paragraphs, updates) {
  paragraphs.forEach((line) => {
    if (/^Dated\s+/i.test(line)) updates.dateRange = line.replace(/^Dated\s+/i, "").trim();
    if (/^Prepared By\s+/i.test(line)) updates.preparedBy = line.replace(/^Prepared By\s+/i, "").trim();
    if (/^Submitted On\s+/i.test(line)) updates.submittedOn = line.replace(/^Submitted On\s+/i, "").trim();
  });
}

function importFromEngagementParagraphs(paragraphs, updates) {
  const leadIdx = paragraphs.findIndex((p) => p.includes("has engaged with SNS"));
  if (leadIdx < 0) return;
  const lead = paragraphs[leadIdx];
  const m = lead.match(/^(.+?)\s+has engaged with SNS/i);
  if (m) updates.customerName = m[1].replace(/\*\*/g, "").trim();
  const rest = paragraphs
    .slice(leadIdx + 1)
    .filter((p) => p !== "The Engagement" && !/^Version\s+\d/i.test(p));
  if (rest.length) {
    updates.executiveSummary = normalizeMonthForImportText(rest.join("\n\n"));
  }
}

function importFromTotPotParagraphs(paragraphs, updates) {
  const tidx = paragraphs.findIndex(
    (p) => p.includes("Total Potential Incident") || p.includes("Total Potential Alerts")
  );
  if (tidx < 0) return;
  const slice = paragraphs.slice(tidx + 1);
  const monthWords = /^(January|February|March|April|May|June|July|August|September|October|November|December)$/i;
  const nums = [];
  for (const p of slice) {
    const t = p.trim();
    if (!t || monthWords.test(t) || /^Incident Count$/i.test(t)) continue;
    if (/^(High|Medium|Low)$/i.test(t)) continue;
    if (/^Potential Incident/i.test(t)) break;
    if (/^\d+(\.\d+)?$/.test(t)) nums.push(t);
  }
  if (nums[0] !== undefined) updates.totPotIncHigh = nums[0];
  if (nums[1] !== undefined) updates.totPotIncMedium = nums[1];
  if (nums[2] !== undefined) updates.totPotIncLow = nums[2];
}

function getTrendChartTitle() {
  const custom = getValue("trendChartTitle");
  if (custom) return custom.replace(/\[MONTH\]/g, getReportMonth());
  return `Potential Incident Summary - ${getReportMonth()}`;
}

function getTrendXAxisLabel() {
  const v = getValue("trendXAxisLabel");
  return v || "Date";
}

function getTrendYAxisLabel() {
  const v = getValue("trendYAxisLabel");
  return v || "Count";
}

function importFromTrendParagraphs(paragraphs, updates) {
  const tidx = paragraphs.findIndex(
    (p) =>
      p.includes("Potential Incident Tickets Trend") || p.includes("Potential Alert Tickets Trend")
  );
  if (tidx < 0) return;
  const slice = paragraphs.slice(tidx + 1);
  const stopIdx = slice.findIndex(
    (p) =>
      p.includes("Potential Alerts") ||
      p.includes("Risks Mitigated") ||
      p.includes("Rule-Based Severity")
  );
  const body = stopIdx >= 0 ? slice.slice(0, stopIdx) : slice;
  const summaryIdx = body.findIndex((p) => /summary/i.test(p) && p.length < 80);
  if (summaryIdx >= 0) {
    updates.trendChartTitle = normalizeMonthForImportText(body[summaryIdx]);
  }
  const noteIdx = body.findIndex((p) => /^note:/i.test(p.trim()));
  const narrStart = noteIdx >= 0 ? noteIdx + 1 : summaryIdx >= 0 ? summaryIdx + 1 : 0;
  if (noteIdx >= 0) updates.trendNote = normalizeMonthForImportText(body[noteIdx]);
  const narrParts = body
    .slice(narrStart)
    .filter((p) => p.length > 3 && !/^note:/i.test(p) && !/incident details are provided/i.test(p));
  if (narrParts.length) updates.trendNarrative = normalizeMonthForImportText(narrParts.join("\n\n"));
}

function importRiskTableToCsv(rows) {
  if (!rows.length) return "";
  const start = rows[0][0] && /S\.?No|Attack/i.test(rows[0][0]) ? 1 : 0;
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    if (r.length < 6) continue;
    if (/No risk rows/i.test(String(r[1] || ""))) continue;
    out.push(csvEscapeRow([r[1], r[2], r[3], r[4], r[5]]));
  }
  return out.join("\n");
}

function importFromRiskSlide(paragraphs, tables, updates, riskAccum) {
  const narIdx = paragraphs.findIndex((p) => p.includes("Potential Alerts - Risks Mitigated"));
  const slice = narIdx >= 0 ? paragraphs.slice(narIdx + 1) : paragraphs;
  const narrative = slice.filter((p) => p.length > 40).join("\n\n");
  if (narrative && !riskAccum.narrativeDone) {
    updates.riskNarrative = normalizeMonthForImportText(narrative);
    riskAccum.narrativeDone = true;
  }
  tables.forEach((tbl) => {
    const csv = importRiskTableToCsv(tbl);
    if (csv) riskAccum.csvParts.push(csv);
  });
}

function importFromSlaParagraphs(paragraphs, updates) {
  const t = paragraphs.join("\n");
  const m1 = t.match(/Overall,\s*(\d+)\s*incidents/i);
  if (m1) updates.slaIncidentCount = m1[1];
  const m2 = t.match(/Remediation Time for\s*(\d+)\s*closed/i);
  if (m2) updates.slaClosedIncidentCount = m2[1];
}

function wrapTfImportedNote(body) {
  const b = String(body || "").trim();
  if (!b) return "Note:";
  if (/^Note:/i.test(b)) return b;
  return `Note:\n${b}`;
}

function importTfNotesFromParagraphs(paragraphs, updates) {
  const blob = paragraphs.join("\n");
  const segs = `\n${blob}`.split(/\n\s*Note:\s*/i);
  if (segs.length < 2) return;
  const bodies = segs
    .slice(1)
    .map((s) => s.trim().replace(/^•\s*/gm, "").trim())
    .filter(Boolean);
  if (bodies[0]) updates.tfPosNoteTrue = wrapTfImportedNote(bodies[0]);
  if (bodies[1]) updates.tfPosNoteFalse = wrapTfImportedNote(bodies[1]);
}

function importFromEpsTrend(paragraphs, tables, updates) {
  if (tables.length && tables[0].length) {
    const r0 = tables[0][0];
    if (r0 && r0.length >= 2) {
      updates.epsMonthLabel = String(r0[0]).trim();
      const n = parseFloat(String(r0[1]).replace(/,/g, ""));
      if (Number.isFinite(n)) updates.epsTotalValue = String(n);
    }
  }
}

function importEpsEventsTableToCsv(rows) {
  if (!rows.length) return "";
  const start = rows[0][0] && /S\.?No|Reporting/i.test(rows[0][0]) ? 1 : 0;
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    if (r.length < 5) continue;
    if (/No rows found/i.test(String(r[1] || ""))) continue;
    out.push(csvEscapeRow([r[1], r[2], r[3], r[4], r[5]]));
  }
  return out.join("\n");
}

function importSupportFromTables(tables, updates) {
  if (tables.length < 1) return;
  const major = tables[0];
  const majorRows = [];
  for (let i = 0; i < major.length; i++) {
    const r = major[i];
    if (r.length < 4) continue;
    if (/Created by/i.test(String(r[0] || ""))) continue;
    if (/Grand Total/i.test(String(r[0] || ""))) continue;
    if (/^Major tickets$/i.test(String(r[0] || ""))) continue;
    if (/^Status$/i.test(String(r[0] || ""))) continue;
    if (/^Closed$/i.test(String(r[1] || "")) && i < 4) continue;
    const by = String(r[0] || "").trim();
    if (!by || by === "Grand Total") continue;
    majorRows.push(csvEscapeRow([by, r[1], r[2]]));
  }
  if (majorRows.length) updates.supportMajorCsv = majorRows.join("\n");

  if (tables.length >= 2) {
    const minor = tables[1];
    const minorRows = [];
    for (let i = 0; i < minor.length; i++) {
      const r = minor[i];
      if (r.length < 2) continue;
      if (/Created by/i.test(String(r[0] || ""))) continue;
      if (/Minor tickets$/i.test(String(r[0] || ""))) continue;
      if (/Grand Total/i.test(String(r[0] || ""))) continue;
      const by = String(r[0] || "").trim();
      if (!by) continue;
      minorRows.push(csvEscapeRow([by, r[1]]));
    }
    if (minorRows.length) updates.supportMinorCsv = minorRows.join("\n");
  }
}

function importInventoryFromSlide(paragraphs, tables, updates) {
  if (tables.length && tables[0].length) {
    const rows = tables[0];
    const start = rows[0][0] && /S\.?No|Device/i.test(rows[0][0]) ? 1 : 0;
    const out = [];
    for (let i = start; i < rows.length; i++) {
      const r = rows[i];
      if (r.length < 2) continue;
      if (/^Total$/i.test(String(r[1] || "").trim())) continue;
      out.push(csvEscapeRow([r[1], r[2]]));
    }
    if (out.length) updates.inventoryCsv = out.join("\n");
  }
  const tidx = paragraphs.findIndex((p) => p.includes("Integrated Device Inventory"));
  const tail = paragraphs.slice(tidx + 1).filter((p) => p.length > 20);
  if (tail.length) updates.inventoryNote = normalizeMonthForImportText(tail[tail.length - 1]);
}

function keyPointsFromParagraphs(paragraphs, updates) {
  const kidx = paragraphs.findIndex((p) => p.includes("Key Points"));
  const slice = kidx >= 0 ? paragraphs.slice(kidx + 1) : [];
  if (!slice.length) return;

  const out = [];
  let i = 0;
  while (i < slice.length) {
    const p = slice[i].trim();
    if (!p) {
      i += 1;
      continue;
    }
    const headingMatch = p.match(/^(\d+)\.\s+(.+?):?\s*$/);
    if (headingMatch && p.length < 100 && !p.startsWith("•")) {
      out.push(`**${headingMatch[2].trim()}:**`);
      i += 1;
      const bodyLines = [];
      while (i < slice.length) {
        const next = slice[i].trim();
        if (!next) {
          i += 1;
          continue;
        }
        if (/^\d+\.\s+/.test(next) && next.length < 100 && !next.startsWith("•")) break;
        bodyLines.push(next);
        i += 1;
      }
      if (bodyLines.length) out.push(bodyLines.join("\n"));
    } else {
      out.push(p);
      i += 1;
    }
  }
  updates.keyPointsSummary = normalizeMonthForImportText(out.join("\n"));
}

const PPTX_IMPORT_FIELD_LABELS = {
  customerName: "Customer name",
  executiveSummary: "Executive summary (Engagement)",
  dateRange: "Report date range",
  preparedBy: "Prepared by",
  submittedOn: "Submitted on",
  trendChartTitle: "Trend — chart title",
  trendXAxisLabel: "Trend — X-axis label",
  trendYAxisLabel: "Trend — Y-axis label",
  trendNote: "Trend — note",
  trendNarrative: "Trend — narrative",
  riskNarrative: "Risk slide — narrative",
  riskCsv: "Risk CSV",
  slaIncidentCount: "SLA — incident count",
  slaClosedIncidentCount: "SLA — closed incident count",
  tfPosNoteTrue: "True positive — note",
  tfPosNoteFalse: "False positive — note",
  epsMonthLabel: "EPS — month label",
  epsTotalValue: "EPS — total value",
  epsTopHostsCsv: "EPS — top hosts CSV",
  epsEventsCsv: "EPS events CSV",
  supportMajorCsv: "Support — major tickets CSV",
  supportMinorCsv: "Support — minor tickets CSV",
  inventoryCsv: "Inventory CSV",
  inventoryNote: "Inventory note",
  keyPointsSummary: "Key points summary"
};

const LAST_PPTX_IMPORT_STORAGE_KEY = "bluPineLastPptxImport";

function renderLastImportLog(fileName, appliedIds) {
  const meta = document.getElementById("lastImportLogMeta");
  const list = document.getElementById("lastImportLogList");
  const box = document.getElementById("lastImportLog");
  if (!meta || !list || !box) return;

  const when = new Date().toLocaleString();
  meta.textContent = `${when} — ${fileName || "presentation.pptx"} — ${appliedIds.length} field(s)`;
  list.innerHTML = "";
  appliedIds.forEach((id) => {
    const li = document.createElement("li");
    li.textContent = PPTX_IMPORT_FIELD_LABELS[id] || id;
    list.appendChild(li);
  });
  box.hidden = false;

  try {
    sessionStorage.setItem(
      LAST_PPTX_IMPORT_STORAGE_KEY,
      JSON.stringify({ t: Date.now(), file: fileName || "", ids: appliedIds })
    );
  } catch (e) {
    // Storage may be unavailable (private mode).
  }
}

function restoreLastImportLogFromStorage() {
  try {
    const raw = sessionStorage.getItem(LAST_PPTX_IMPORT_STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (!data || !Array.isArray(data.ids) || !data.ids.length) return;
    const meta = document.getElementById("lastImportLogMeta");
    const list = document.getElementById("lastImportLogList");
    const box = document.getElementById("lastImportLog");
    if (!meta || !list || !box) return;
    const when = data.t ? new Date(data.t).toLocaleString() : "";
    meta.textContent = `${when}${when ? " — " : ""}${data.file || "presentation.pptx"} — ${data.ids.length} field(s) (restored from this session)`;
    list.innerHTML = "";
    data.ids.forEach((id) => {
      const li = document.createElement("li");
      li.textContent = PPTX_IMPORT_FIELD_LABELS[id] || id;
      list.appendChild(li);
    });
    box.hidden = false;
  } catch (e) {
    // ignore
  }
}

function clearLastImportLog() {
  const box = document.getElementById("lastImportLog");
  if (box) box.hidden = true;
  try {
    sessionStorage.removeItem(LAST_PPTX_IMPORT_STORAGE_KEY);
  } catch (e) {
    // ignore
  }
}

async function importEditablePptxIntoForm(file) {
  if (typeof JSZip === "undefined") {
    alert("JSZip did not load. Check your network and refresh the page.");
    return;
  }
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const slidePaths = Object.keys(zip.files)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/i)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)/i)[1], 10);
      return na - nb;
    });

  if (!slidePaths.length) {
    alert("No slides found in this file.");
    return;
  }

  const updates = {};
  const riskAccum = { narrativeDone: false, csvParts: [] };

  for (const path of slidePaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    const xml = await entry.async("string");
    const paragraphs = extractParagraphsFromSlideXml(xml);
    const tables = extractTablesFromSlideXml(xml);
    const kind = classifySlideKind(paragraphs);

    switch (kind) {
      case "revision":
        importFromRevisionParagraphs(paragraphs, updates);
        break;
      case "engagement":
        importFromEngagementParagraphs(paragraphs, updates);
        break;
      case "totPot":
        importFromTotPotParagraphs(paragraphs, updates);
        break;
      case "trend":
        importFromTrendParagraphs(paragraphs, updates);
        break;
      case "risk":
        importFromRiskSlide(paragraphs, tables, updates, riskAccum);
        break;
      case "sla":
        importFromSlaParagraphs(paragraphs, updates);
        break;
      case "tfPos":
        importTfNotesFromParagraphs(paragraphs, updates);
        break;
      case "epsTrend":
        importFromEpsTrend(paragraphs, tables, updates);
        break;
      case "epsEvents":
        if (tables[0]) updates.epsEventsCsv = importEpsEventsTableToCsv(tables[0]);
        break;
      case "support":
        importSupportFromTables(tables, updates);
        break;
      case "inventory":
        importInventoryFromSlide(paragraphs, tables, updates);
        break;
      case "keyPoints":
        keyPointsFromParagraphs(paragraphs, updates);
        break;
      default:
        break;
    }
  }

  if (riskAccum.csvParts.length) {
    updates.riskCsv = riskAccum.csvParts.join("\n");
  }

  const applied = [];
  Object.entries(updates).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el && typeof val === "string") {
      el.value = val;
      applied.push(id);
    }
  });

  if (!applied.length) {
    alert(
      "No editable text fields were found. Use an PPTX exported with “Export PPTX (editable)” from this app, or edit text in PowerPoint (not only pictures)."
    );
    return;
  }

  applyData();
  renderLastImportLog(file && file.name ? file.name : "presentation.pptx", applied);
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

function syncEngagementPuzzleImage() {
  const preview = document.getElementById("puzzleGraphicPreview");
  const ph = document.getElementById("puzzlePlaceholderLayout");
  const phImg = ph?.querySelector("img");
  if (!preview || !ph) return;

  if (puzzleLogoDataUrl) {
    preview.src = puzzleLogoDataUrl;
    preview.style.display = "block";
    ph.style.display = "none";
    return;
  }

  preview.removeAttribute("src");
  preview.style.display = "none";
  ph.style.display = "flex";
  if (phImg) {
    phImg.style.display = "block";
    if (defaultPuzzleDataUrl) {
      phImg.src = defaultPuzzleDataUrl;
    } else if (!phImg.getAttribute("src")) {
      phImg.src = "./final_puzzle.png";
    }
  }
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
    syncEngagementPuzzleImage();
  } else {
    console.warn("Unable to prepare default puzzle data URL for:", absUrl);
  }
  return defaultPuzzleDataUrl;
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
      const ciaCombined =
        typeOfRisk && ciaOnly && typeOfRisk !== ciaOnly
          ? `${typeOfRisk}\n${ciaOnly}`
          : typeOfRisk || ciaOnly;
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
    .map((r) => ({
      device: r[iDevice] || "",
      count: Math.max(0, Number(String(r[iCount]).replace(/,/g, "")) || 0)
    }))
    .filter((r) => r.device);
}

function renderIntegratedInventorySlide() {
  const body = document.getElementById("inventoryTableBody");
  const foot = document.getElementById("inventoryTableFoot");
  const note = document.getElementById("inventoryNotePreview");
  if (!body || !foot || !note) return;
  const rows = parseInventoryRows(getValue("inventoryCsv"));
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
  if (!shouldSkipApplyForNode(note)) {
    note.innerHTML = formatRichTextHTML(getValue("inventoryNote"));
  }
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
  syncSupportTicketsSlideLayout();
  requestAnimationFrame(() => {
    requestAnimationFrame(syncSupportTicketsSlideLayout);
  });
}

function getSupportTicketsGridBudget(slide, grid) {
  const footer = slide.querySelector(".page-footer-bar");
  const gap = 14;
  if (grid && footer) {
    const budget = footer.offsetTop - grid.offsetTop - gap;
    if (Number.isFinite(budget) && budget > 80) return budget;
  }
  const title = slide.querySelector(".support-main-title");
  const header = slide.querySelector(".page-header-right");
  const reservedTop =
    (header?.offsetHeight || 0) + (title?.offsetHeight || 0) + slidePaddingTop(slide) + 8;
  const footerH = footer?.offsetHeight || 42;
  return Math.max(120, slide.clientHeight - reservedTop - footerH - 20);
}

function syncSupportTicketsSlideLayout() {
  const slide = document.getElementById("slideSupportTickets");
  if (!slide) return;
  const grid = slide.querySelector(".support-grid");
  if (!grid) return;
  const footer = slide.querySelector(".page-footer-bar");
  const gap = 14;

  grid.style.height = "";
  grid.style.maxHeight = "";
  slide.querySelectorAll(".support-table").forEach((table) => {
    table.style.height = "";
    table.style.maxHeight = "";
    table.style.fontSize = "";
    table.querySelectorAll("tbody tr").forEach((tr) => {
      tr.style.height = "";
    });
    table.querySelectorAll("th, td").forEach((cell) => {
      cell.style.padding = "";
      cell.style.fontSize = "";
    });
  });

  void slide.offsetHeight;
  const offsetBudget = getSupportTicketsGridBudget(slide, grid);
  const flexBudget = grid.clientHeight;
  const cssMax = parseFloat(window.getComputedStyle(grid).maxHeight) || offsetBudget;
  const budget = Math.min(
    offsetBudget,
    flexBudget > 80 ? flexBudget : offsetBudget,
    Number.isFinite(cssMax) && cssMax > 80 ? cssMax : offsetBudget
  );

  grid.style.height = `${budget}px`;
  grid.style.maxHeight = `${budget}px`;

  let scale = 1;
  const baseFontRem = 0.82;
  const baseBandRem = 1.34;
  const baseHeadRem = 1.12;

  const applyScale = () => {
    const fontRem = baseFontRem * scale;
    const pad = Math.max(2, Math.round(6 * scale));
    slide.querySelectorAll(".support-table").forEach((table) => {
      table.style.fontSize = `${fontRem}rem`;
      table.style.height = "100%";
      table.style.maxHeight = "100%";
      const band = table.querySelector(".support-band");
      if (band) band.style.fontSize = `${baseBandRem * scale}rem`;
      table.querySelectorAll("thead th:not(.support-band)").forEach((th) => {
        th.style.fontSize = `${baseHeadRem * scale}rem`;
      });
      table.querySelectorAll("th, td").forEach((cell) => {
        cell.style.padding = `${pad}px 5px`;
      });
    });
  };

  const overflows = () => {
    if (footer && grid.offsetTop + grid.offsetHeight > footer.offsetTop - gap) return true;
    const cardH = grid.clientHeight || budget;
    return [...slide.querySelectorAll(".support-table")].some(
      (table) => table.scrollHeight > cardH + 1
    );
  };

  applyScale();
  for (let pass = 0; pass < 48; pass++) {
    if (!overflows()) break;
    if (scale <= 0.76) break;
    scale -= 0.04;
    applyScale();
  }
}

function initSupportTicketsSlideLayoutWatcher() {
  const slide = document.getElementById("slideSupportTickets");
  if (!slide || slide.dataset.supportLayoutWatch === "1") return;
  slide.dataset.supportLayoutWatch = "1";
  const run = () => syncSupportTicketsSlideLayout();
  if ("IntersectionObserver" in window) {
    const obs = new IntersectionObserver(
      (entries) => {
        if (entries.some((e) => e.isIntersecting)) run();
      },
      { threshold: 0.12 }
    );
    obs.observe(slide);
  }
  const preview = document.querySelector(".preview-wrap");
  if (preview) {
    let timer;
    preview.addEventListener(
      "scroll",
      () => {
        clearTimeout(timer);
        timer = setTimeout(run, 100);
      },
      { passive: true }
    );
  }
}

function renderEpsEventsTable() {
  const body = document.getElementById("epsEventsTableBody");
  if (!body) return;
  const rows = parseEpsEventsRows(getValue("epsEventsCsv")).slice(0, 12);
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="5" class="eps-num">No rows found. Upload EPS Events CSV.</td></tr>';
    syncEpsEventsSlideLayout();
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
  syncEpsEventsSlideLayout();
  requestAnimationFrame(() => {
    requestAnimationFrame(syncEpsEventsSlideLayout);
  });
}

function getEpsEventsWrapBudget(slide, wrap) {
  const footer = slide.querySelector(".page-footer-bar");
  const gap = 6;
  if (wrap && footer) {
    const budget = footer.offsetTop - wrap.offsetTop - gap;
    if (Number.isFinite(budget) && budget > 80) return budget;
  }
  const title = slide.querySelector(".eps-events-title");
  const header = slide.querySelector(".page-header-right");
  const reservedTop =
    (header?.offsetHeight || 0) + (title?.offsetHeight || 0) + slidePaddingTop(slide) + 8;
  const footerH = footer?.offsetHeight || 42;
  const reservedBottom = footerH + 16;
  return Math.max(120, slide.clientHeight - reservedTop - reservedBottom);
}

function initEpsEventsSlideLayoutWatcher() {
  const slide = document.getElementById("slideEpsEvents");
  if (!slide || slide.dataset.epsLayoutWatch === "1") return;
  slide.dataset.epsLayoutWatch = "1";
  const run = () => syncEpsEventsSlideLayout();
  if ("IntersectionObserver" in window) {
    const obs = new IntersectionObserver(
      (entries) => {
        if (entries.some((e) => e.isIntersecting)) run();
      },
      { threshold: 0.12 }
    );
    obs.observe(slide);
  }
  const preview = document.querySelector(".preview-wrap");
  if (preview) {
    let timer;
    preview.addEventListener(
      "scroll",
      () => {
        clearTimeout(timer);
        timer = setTimeout(run, 100);
      },
      { passive: true }
    );
  }
}

function syncEpsEventsSlideLayout() {
  const slide = document.getElementById("slideEpsEvents");
  if (!slide) return;
  const wrap = slide.querySelector(".eps-events-wrap");
  const table = slide.querySelector(".eps-events-table");
  if (!wrap || !table) return;

  const tbody = table.querySelector("tbody");
  const thead = table.querySelector("thead");
  const tbodyRows = tbody ? [...tbody.querySelectorAll("tr")] : [];
  const footer = slide.querySelector(".page-footer-bar");
  const gap = 6;

  wrap.style.height = "";
  wrap.style.maxHeight = "";
  table.style.height = "";
  table.style.maxHeight = "";
  table.style.fontSize = "";
  tbodyRows.forEach((tr) => {
    tr.style.height = "";
  });
  table.querySelectorAll("th, td").forEach((cell) => {
    cell.style.padding = "";
  });

  void slide.offsetHeight;
  const offsetBudget = getEpsEventsWrapBudget(slide, wrap);
  const flexBudget = wrap.clientHeight;
  const budget =
    flexBudget > 80 ? Math.min(offsetBudget, flexBudget) : offsetBudget;

  table.style.height = `${budget}px`;
  table.style.maxHeight = `${budget}px`;

  let fontRem = 0.98;
  let padH = 5;
  let bodyH = Math.max(40, budget - (thead?.offsetHeight || 0) - 2);

  const applyBase = () => {
    table.style.fontSize = `${fontRem}rem`;
    table.querySelectorAll("th, td").forEach((cell) => {
      cell.style.paddingTop = "";
      cell.style.paddingBottom = "";
      cell.style.paddingLeft = `${padH}px`;
      cell.style.paddingRight = `${padH}px`;
    });
  };

  const distributeRowHeights = () => {
    const rowCount = Math.max(tbodyRows.length, 1);
    const heights = Array(rowCount).fill(Math.floor(bodyH / rowCount));
    let remainder = bodyH - heights.reduce((sum, h) => sum + h, 0);
    for (let i = 0; i < remainder; i++) {
      heights[i % rowCount] += 1;
    }
    tbodyRows.forEach((tr, idx) => {
      tr.style.height = `${heights[idx] || heights[0]}px`;
    });
    return heights[0] || Math.floor(bodyH / rowCount);
  };

  const applyFill = () => {
    applyBase();
    const theadH = thead ? thead.offsetHeight : 0;
    bodyH = Math.max(40, budget - theadH - 2);
    const sampleRowH = distributeRowHeights();
    const padV = Math.max(2, Math.min(8, Math.floor((sampleRowH - fontRem * 14) / 2)));
    table.querySelectorAll("thead th").forEach((cell) => {
      cell.style.paddingTop = `${Math.max(3, padV)}px`;
      cell.style.paddingBottom = `${Math.max(3, padV)}px`;
    });
    table.querySelectorAll("tbody td").forEach((cell) => {
      cell.style.paddingTop = `${padV}px`;
      cell.style.paddingBottom = `${padV}px`;
    });
  };

  const overflowsFooter = () => {
    if (!footer) return false;
    const bottom = wrap.offsetTop + Math.max(wrap.offsetHeight, table.offsetHeight);
    return bottom > footer.offsetTop - gap;
  };

  applyFill();
  for (let pass = 0; pass < 48; pass++) {
    const over = table.scrollHeight - budget;
    if (over <= 1 && !overflowsFooter()) break;
    if (over > 1 && bodyH > tbodyRows.length * 14) {
      bodyH -= Math.max(1, Math.ceil(over / 2));
      applyFill();
      continue;
    }
    if (fontRem <= 0.66) break;
    fontRem -= 0.04;
    padH = Math.max(3, padH - 1);
    applyFill();
  }
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

function getPptRiskSlideLayout(slideIndex, dataRowCount, sy) {
  const footerTop = 4.95 * sy;
  const tableGap = 0.03 * sy;
  const titleY = 0.55 * sy;
  const titleH = 0.38 * sy;
  const gapIn = 0.04 * sy;
  const hasNarrative = slideIndex === 0 && Boolean(String(getValue("riskNarrative") || "").trim());
  const rowCount = Math.max(1, dataRowCount) + 1;

  let tableY;
  let narrativeY = 0;
  let narrativeH = 0;

  if (hasNarrative) {
    const raw = String(getValue("riskNarrative") || "").trim();
    const lines = Math.max(1, raw.split(/\r?\n/).length, Math.ceil(raw.length / 92));
    narrativeH = Math.min(1.15, 0.1 * lines + 0.18) * sy;
    narrativeY = 0.68 * sy;
    tableY = narrativeY + narrativeH + gapIn;
  } else {
    tableY = titleY + titleH + gapIn;
  }

  const tableH = Math.max(0.55 * sy, footerTop - tableY - tableGap);
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

function ensureSlaPieCenterPlugin() {
  if (slaPieCenterPluginRegistered || typeof Chart === "undefined") return;
  slaPieCenterPluginRegistered = true;
  Chart.register({
    id: "slaPieCenter",
    afterDraw(chart, args, opts) {
      const txt = opts && opts.text ? String(opts.text) : "";
      if (!txt) return;
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data || !meta.data[0]) return;
      const arc = meta.data[0];
      const { ctx } = chart;
      ctx.save();
      ctx.fillStyle = "#000";
      ctx.font = "bold 12px Calibri, Segoe UI, Arial";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(txt, arc.x, arc.y);
      ctx.restore();
    }
  });
}

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

/** Light grey used outside the pie (box margin → circle edge); unfilled slice is white inside the circle. */
const SLA_HASH_GREY = "#ececec";

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

function formatResponseSummaryText(incidentCount, customerName, responseLabel) {
  const customer = (customerName || "Customer").trim();
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

/** Keeps #reportRoot in sync with widescreen 16:9 slide size (see styles.css .ppt-wide). */
function syncReportRootLayoutClass() {
  const root = document.getElementById("reportRoot");
  if (!root) return;
  const { reportClass } = pptLayoutConfig.LAYOUT_16X9;
  root.className = `report ${reportClass}`;
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

function parseSeverityFieldValue(fieldId) {
  const raw = String(getValue(fieldId) ?? "").trim();
  if (raw === "") return null;
  const value = Math.max(0, parseFloat(raw) || 0);
  if (value <= 0) return null;
  return { raw, value };
}

function severityFieldIsVisible(fieldId) {
  return parseSeverityFieldValue(fieldId) !== null;
}

function getTotPotIncSeverityEntries() {
  return TOT_POT_INC_SEVERITIES.map((severity) => {
    const parsed = parseSeverityFieldValue(severity.fieldId);
    if (!parsed) return null;
    return {
      ...severity,
      raw: parsed.raw,
      value: parsed.value
    };
  }).filter(Boolean);
}

const FORTI_SIEM_SEVERITIES = [
  { label: "High", fieldId: "fortiAlertHigh", tableCellId: "fortiSiemTableHigh" },
  { label: "Medium", fieldId: "fortiAlertMedium", tableCellId: "fortiSiemTableMed" },
  { label: "Low", fieldId: "fortiAlertLow", tableCellId: "fortiSiemTableLow" }
];

const TF_POS_TRUE_SEVERITIES = [
  { label: "High", fieldId: "tfPosTrueHigh", tableCellId: "tfPosTrueTableHigh" },
  { label: "Medium", fieldId: "tfPosTrueMedium", tableCellId: "tfPosTrueTableMed" },
  { label: "Low", fieldId: "tfPosTrueLow", tableCellId: "tfPosTrueTableLow" }
];

const TF_POS_FALSE_SEVERITIES = [
  { label: "High", fieldId: "tfPosFalseHigh", tableCellId: "tfPosFalseTableHigh" },
  { label: "Medium", fieldId: "tfPosFalseMedium", tableCellId: "tfPosFalseTableMed" },
  { label: "Low", fieldId: "tfPosFalseLow", tableCellId: "tfPosFalseTableLow" }
];

function updateSeverityTableRows(severities, monthLbl, hdrId) {
  if (monthLbl && hdrId) {
    const hdr = document.getElementById(hdrId);
    if (hdr) hdr.textContent = monthLbl;
  }
  severities.forEach((severity) => {
    const cell = document.getElementById(severity.tableCellId);
    const row = cell?.closest("tr");
    const parsed = parseSeverityFieldValue(severity.fieldId);
    const visible = parsed !== null;
    if (row) row.style.display = visible ? "" : "none";
    if (cell) cell.textContent = visible ? formatTotpotTableVal(parsed.raw) : "";
  });
}

function updateFortiSiemTable(monthLbl) {
  updateSeverityTableRows(FORTI_SIEM_SEVERITIES, monthLbl, "fortiSiemTableColHdr");
  syncFortiSiemSlideLayout();
}

function updateTfPosTables(monthLbl) {
  updateSeverityTableRows(TF_POS_TRUE_SEVERITIES, monthLbl, "tfPosTrueTableColHdr");
  updateSeverityTableRows(TF_POS_FALSE_SEVERITIES, monthLbl, "tfPosFalseTableColHdr");
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
    const visible = severityFieldIsVisible(severity.fieldId);
    if (row) row.style.display = visible ? "" : "none";
    if (cell) cell.textContent = visible ? raw : "";
  });
  syncTotPotIncSlideLayout();
}

function renderTotPotIncChart() {
  const canvas = document.getElementById("totPotIncChart");
  if (!canvas || typeof Chart === "undefined") return;

  const entries = getTotPotIncSeverityEntries();
  if (!entries.length) {
    if (totPotIncChart) totPotIncChart.destroy();
    totPotIncChart = undefined;
    syncTotPotIncSlideLayout();
    return;
  }
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

function formatTotpotTableVal(raw) {
  const s = String(raw ?? "").trim();
  if (s === "") return "0";
  const n = Number(s);
  if (!Number.isFinite(n)) return s;
  return String(Math.round(n));
}

function syncTotpotTableLayout(chart, wrapSelector) {
  const wrap = document.querySelector(wrapSelector);
  const canvas = chart?.canvas;
  if (!wrap || !canvas || !chart?.chartArea) return;
  const scale = canvas.clientWidth / (canvas.width || 1);
  const left = chart.chartArea.left * scale;
  const width = chart.chartArea.width * scale;
  wrap.style.marginLeft = `${Math.round(left)}px`;
  wrap.style.width = `${Math.round(width)}px`;
  wrap.style.maxWidth = "none";
}

function syncFortiSiemTableLayout() {
  const wrap = document.querySelector("#slideFortiSiemAlerts .totpot-table-legend-wrap");
  if (!wrap) return;
  wrap.style.marginLeft = "";
  wrap.style.width = "";
  wrap.style.maxWidth = "";
}

function syncFortiSiemSlideLayout() {
  syncFortiSiemTableLayout();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (fortiSiemAlertsChart) {
        fortiSiemAlertsChart.resize();
      }
    });
  });
}

function syncTfPosTableLayouts() {
  document.querySelectorAll("#slideTfPosAlerts .totpot-table-legend-wrap").forEach((wrap) => {
    wrap.style.marginLeft = "";
    wrap.style.width = "";
    wrap.style.maxWidth = "";
  });
}

function scheduleSyncTfPosSlideLayout() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      syncTfPosSlideLayout();
    });
  });
}

function slidePaddingTop(slide) {
  const style = window.getComputedStyle(slide);
  return parseFloat(style.paddingTop) || 0;
}

function getTotPotStyleYScale(peak) {
  return peak <= 100
    ? { max: Math.max(10, Math.ceil((peak * 1.1) / 5) * 5), step: 5, mirror: false }
    : (() => {
        const picked = tfPosPickScale(peak);
        return { max: picked.max, step: picked.step, mirror: false };
      })();
}

function buildTotPotStyleSeverityChart(canvas, oldChart, high, medium, low) {
  if (!canvas || typeof Chart === "undefined") return undefined;
  const entries = [
    { label: "High", value: high, color: "#ff0000" },
    { label: "Medium", value: medium, color: "#ffff00" },
    { label: "Low", value: low, color: "#00b050" }
  ].filter((entry) => entry.value > 0);
  if (!entries.length) {
    if (oldChart) oldChart.destroy();
    return undefined;
  }
  const labels = entries.map((entry) => entry.label);
  const data = entries.map((entry) => entry.value);
  const colors = entries.map((entry) => entry.color);
  const peak = Math.max(...data, 1);
  const yScale = getTotPotStyleYScale(peak);

  if (oldChart) oldChart.destroy();
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

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [{ label: "Alert Count", data, backgroundColor: colors, borderWidth: 0 }]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { left: 0, right: 2, top: 16, bottom: 2 } },
      elements: { bar: { borderWidth: 0 } },
      plugins,
      datasets: { bar: { categoryPercentage: 0.72, barPercentage: 0.72 } },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { display: false }
        },
        y: {
          beginAtZero: true,
          max: yScale.max,
          min: 0,
          ticks: {
            stepSize: yScale.step,
            font: axisFont,
            color: "#000000",
            padding: 8,
            mirror: yScale.mirror,
            ...(yScale.step >= 200 ? { callback: (v) => v.toLocaleString("en-US") } : {})
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
}

function renderFortiSiemAlertsChart() {
  const canvas = document.getElementById("fortiSiemAlertsChart");
  if (!canvas) return;

  const high = Math.max(0, parseFloat(getValue("fortiAlertHigh")) || 0);
  const medium = Math.max(0, parseFloat(getValue("fortiAlertMedium")) || 0);
  const low = Math.max(0, parseFloat(getValue("fortiAlertLow")) || 0);

  fortiSiemAlertsChart = buildTotPotStyleSeverityChart(
    canvas,
    fortiSiemAlertsChart,
    high,
    medium,
    low
  );

  syncFortiSiemSlideLayout();
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
  if (shouldSkipApplyForNode(el)) return;
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  const heading = '<p class="tfpos-note-heading"><strong>Note:</strong></p>';
  if (!items.length) {
    el.innerHTML = heading;
    scheduleSyncTfPosSlideLayout();
    return;
  }
  if (items.length === 1) {
    el.innerHTML = `${heading}<p class="tfpos-note-text">${formatRichTextHTML(items[0])}</p>`;
    scheduleSyncTfPosSlideLayout();
    return;
  }
  const lis = items.map((line) => `<li>${formatRichTextHTML(line)}</li>`).join("");
  el.innerHTML = `${heading}<ul class="tfpos-note-list">${lis}</ul>`;
  scheduleSyncTfPosSlideLayout();
}

/** Plain text for PPTX: "Note:" then bullet lines. */
function formatTfPosNoteForPptx(textareaId) {
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  if (!items.length) return "Note:";
  return `Note:\n\n${items.map((line) => `• ${line}`).join("\n")}`;
}

function renderTfPosCharts() {
  const tpH = Math.max(0, parseFloat(getValue("tfPosTrueHigh")) || 0);
  const tpM = Math.max(0, parseFloat(getValue("tfPosTrueMedium")) || 0);
  const tpL = Math.max(0, parseFloat(getValue("tfPosTrueLow")) || 0);
  const fpH = Math.max(0, parseFloat(getValue("tfPosFalseHigh")) || 0);
  const fpM = Math.max(0, parseFloat(getValue("tfPosFalseMedium")) || 0);
  const fpL = Math.max(0, parseFloat(getValue("tfPosFalseLow")) || 0);

  truePositiveAlertsChart = buildTotPotStyleSeverityChart(
    document.getElementById("truePositiveAlertsChart"),
    truePositiveAlertsChart,
    tpH,
    tpM,
    tpL
  );
  falsePositiveAlertsChart = buildTotPotStyleSeverityChart(
    document.getElementById("falsePositiveAlertsChart"),
    falsePositiveAlertsChart,
    fpH,
    fpM,
    fpL
  );

  requestAnimationFrame(() => {
    syncTfPosSlideLayout();
  });
}

function syncTfPosSlideLayout() {
  const slide = document.getElementById("slideTfPosAlerts");
  if (!slide) return;
  const title = slide.querySelector(".tfpos-main-title");
  const header = slide.querySelector(".page-header-right");
  const reservedTop =
    (header?.offsetHeight || 0) + (title?.offsetHeight || 0) + slidePaddingTop(slide) + 8;
  const reservedBottom = 52;
  const grid = slide.querySelector(".tfpos-grid");
  const slideBudget = Math.max(
    160,
    grid?.clientHeight || slide.clientHeight - reservedTop - reservedBottom
  );
  const noteBaseRem = 0.82;
  const cols = [...slide.querySelectorAll(".tfpos-col")];

  slide.querySelectorAll(".tfpos-note-block").forEach((note) => {
    note.style.fontSize = "";
  });
  slide.querySelectorAll(".tfpos-chart-wrap").forEach((wrap) => {
    wrap.style.height = "";
    wrap.style.minHeight = "";
  });

  syncTfPosTableLayouts();

  const resizeCharts = () => {
    if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
    if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
  };

  resizeCharts();

  const colContentHeight = (col) => {
    let h = 0;
    col.querySelectorAll(".tfpos-subtitle, .tfpos-frame-inner, .tfpos-note-block").forEach((el) => {
      h += Math.max(el.offsetHeight, el.scrollHeight);
    });
    return h;
  };

  let tallest = Math.max(0, ...cols.map(colContentHeight));
  if (tallest > slideBudget) {
    let noteScale = 1;
    for (let i = 0; i < 12 && tallest > slideBudget && noteScale > 0.68; i++) {
      noteScale -= 0.05;
      slide.querySelectorAll(".tfpos-note-block").forEach((note) => {
        note.style.fontSize = `${noteScale * noteBaseRem}rem`;
      });
      resizeCharts();
      tallest = Math.max(0, ...cols.map(colContentHeight));
    }
  }
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

function pickEpsTotalYScale(peak) {
  const p = Math.max(Number(peak) || 0, 1);
  const padded = p * 1.1;
  const steps = [10, 20, 25, 50, 100, 200, 250, 500, 1000];
  let step = 10;
  for (const s of steps) {
    if (padded / s <= 6) {
      step = s;
      break;
    }
  }
  const max = Math.max(step, Math.ceil(padded / step) * step);
  return { max, step };
}

function syncEpsSlideLayout() {
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (epsTotalChart) epsTotalChart.resize();
      if (epsTopHostsChart) epsTopHostsChart.resize();
    });
  });
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

  const totalY = pickEpsTotalYScale(totalEps);
  const totalAxisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 11, weight: "bold" };
  const totalPlugins = {
    legend: { display: false },
    totPotPlotFrame: true
  };
  if (ReportCharts.isDataLabelsRegistered()) {
    totalPlugins.datalabels = ReportCharts.buildEpsTrendTotalValueDataLabelOptions({
      font: { weight: "bold", size: 12, family: totalAxisFont.family }
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
      elements: { bar: { borderWidth: 0 } },
      datasets: { bar: { categoryPercentage: 0.62, barPercentage: 0.88, maxBarThickness: 140 } },
      plugins: totalPlugins,
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000000", font: totalAxisFont, maxRotation: 0, minRotation: 0 }
        },
        y: {
          beginAtZero: true,
          min: 0,
          max: totalY.max,
          grace: 0,
          ticks: {
            color: "#000000",
            stepSize: totalY.step,
            font: totalAxisFont,
            callback: (v) => {
              const n = Number(v);
              return Number.isFinite(n) && Number.isInteger(n) ? String(n) : "";
            }
          },
          grid: { color: "#d9d9d9", drawBorder: false, drawTicks: false }
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
  syncEpsSlideLayout();
}

function ensureTrendDataLabelsPlugin() {
  ReportCharts.ensureDataLabelsPlugin();
  trendDataLabelsRegistered = ReportCharts.isDataLabelsRegistered();
}

function renderTrendChart() {
  const canvas = document.getElementById("trendChart");
  if (!canvas || typeof Chart === "undefined") return;
  const rows = ReportCharts.filterTrendRowsWithIncidentData(parseTrendCsvRows(getValue("trendCsv")));
  if (!rows.length) {
    if (trendChart) {
      trendChart.destroy();
      trendChart = undefined;
    }
    return;
  }

  ensureTrendDataLabelsPlugin();

  if (trendChart) trendChart.destroy();
  const plugins = ReportCharts.buildSeverityChartPlugins(getTrendChartTitle(), {
    legendPosition: "bottom",
    fontSize: 16,
    legendColor: "#000"
  });

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
      plugins,
      scales: {
        x: ReportCharts.buildTrendChartXScaleOptions({ title: getTrendXAxisLabel(), rowCount: rows.length }),
        y: ReportCharts.buildTrendChartYScaleOptions(rows, { title: getTrendYAxisLabel() })
      }
    }
  });

  requestAnimationFrame(() => {
    if (trendChart) trendChart.resize();
  });
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
    if (pageSkipsHeaderLogo(page)) {
      page.querySelector(".page-header-right")?.remove();
      if (!page.querySelector(".page-footer-bar")) {
        const footer = document.createElement("div");
        footer.className = "page-footer-bar";
        page.appendChild(footer);
      }
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

function pageSkipsHeaderLogo(page) {
  return (
    page.classList.contains("cover-hero") ||
    page.classList.contains("revision-page") ||
    page.getAttribute("data-slide") === "engagement" ||
    page.classList.contains("slide-contact-summary") ||
    page.getAttribute("data-slide") === "contact"
  );
}

function applyData() {
  syncReportRootLayoutClass();
  const reportRoot = document.getElementById("reportRoot");
  if (reportRoot) void reportRoot.offsetHeight;

  Object.entries(fields).forEach(([key, inputId]) => {
    const value = getValue(inputId);
    document.querySelectorAll(`[data-field="${key}"]`).forEach((el) => {
      if (shouldSkipApplyForNode(el)) return;
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
  if (execEl && !shouldSkipApplyForNode(execEl)) {
    execEl.innerHTML = getValue("executiveSummary")
      .split("\n")
      .filter((p) => p.trim())
      .map((p) => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }

  syncEngagementPuzzleImage();

  const monthLbl = getTableMonthLabel();
  updateTotPotIncTable(monthLbl);
  updateFortiSiemTable(monthLbl);
  updateTfPosTables(monthLbl);
  renderEpsEventsTable();
  renderSupportTicketsSlide();
  renderIntegratedInventorySlide();
  renderKeyPointsPreview();
  const epsTotalVal = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const epsMonthEl = document.getElementById("epsTotalMonthLabelCell");
  if (epsMonthEl) epsMonthEl.textContent = monthLbl;
  const epsValEl = document.getElementById("epsTotalValueCell");
  if (epsValEl) epsValEl.textContent = String(epsTotalVal);

  const trendNote = document.getElementById("trendNotePreview");
  if (trendNote && !shouldSkipApplyForNode(trendNote)) {
    trendNote.innerHTML = formatRichTextHTML(getValue("trendNote"));
  }
  const trendNarr = document.getElementById("trendNarrativePreview");
  if (trendNarr && !shouldSkipApplyForNode(trendNarr)) {
    trendNarr.innerHTML = formatRichTextHTML(getValue("trendNarrative"));
  }

  void renderRiskSlides();
  updateLogos();
  applyPremiumLayout();
  renderTotPotIncChart();
  renderFortiSiemAlertsChart();
  renderTfPosCharts();
  scheduleSyncTfPosSlideLayout();
  const tfPosNoteTruePreview = document.getElementById("tfPosNoteTruePreview");
  if (tfPosNoteTruePreview && !shouldSkipApplyForNode(tfPosNoteTruePreview)) {
    renderTfPosNotePreview("tfPosNoteTruePreview", "tfPosNoteTrue");
  }
  const tfPosNoteFalsePreview = document.getElementById("tfPosNoteFalsePreview");
  if (tfPosNoteFalsePreview && !shouldSkipApplyForNode(tfPosNoteFalsePreview)) {
    renderTfPosNotePreview("tfPosNoteFalsePreview", "tfPosNoteFalse");
  }
  renderEpsTrendPlot();
  renderTrendChart();
  renderRuleSeverityChart();
  renderSlaStatus();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      syncEngagementLayout();
      if (totPotIncChart) totPotIncChart.resize();
      if (fortiSiemAlertsChart) {
        syncFortiSiemSlideLayout();
      }
      if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
      if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
      syncEpsSlideLayout();
      syncEpsEventsSlideLayout();
      syncSupportTicketsSlideLayout();
      if (trendChart) trendChart.resize();
      if (ruleSeverityChart) ruleSeverityChart.resize();
      updateSlaSlide();
      syncTfPosSlideLayout();
      window.dispatchEvent(new Event("resize"));
    });
  });
  if (typeof reapplyTextStyleOverrides === "function") {
    reapplyTextStyleOverrides();
  }
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

function engagementSlideOnClone(doc) {
  const engage = doc.getElementById("slideEngagement");
  if (!engage) return;
  const textBox = engage.querySelector(".engagement-text");
  if (textBox) {
    textBox.style.setProperty("background-color", "#f6fbfb", "important");
    textBox.style.setProperty("background-image", "none", "important");
    textBox.style.setProperty("border", "none", "important");
    textBox.style.setProperty("box-sizing", "border-box", "important");
  }
  engage.querySelectorAll("strong, .fw-bold").forEach((node) => {
    node.style.setProperty("font-weight", "700", "important");
    node.style.setProperty("color", "#000000", "important");
  });
  const preview = engage.querySelector("#puzzleGraphicPreview");
  const previewSrc = preview && preview.getAttribute("src");
  if (preview && previewSrc && previewSrc.length > 8) {
    preview.style.setProperty("display", "block", "important");
    preview.style.setProperty("object-fit", "fill", "important");
  }
  engage.querySelectorAll(".puzzle-placeholder img").forEach((img) => {
    img.style.setProperty("display", "block", "important");
    img.style.setProperty("visibility", "visible", "important");
    img.style.setProperty("opacity", "1", "important");
    img.style.setProperty("width", "100%", "important");
    img.style.setProperty("height", "100%", "important");
    img.style.setProperty("object-fit", "fill", "important");
  });
}

async function waitForImagesInElement(root) {
  if (!root) return;
  const imgs = Array.from(root.querySelectorAll("img"));
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise((resolve) => {
          if (img.complete && img.naturalWidth > 0) {
            resolve();
            return;
          }
          const done = () => resolve();
          img.addEventListener("load", done, { once: true });
          img.addEventListener("error", done, { once: true });
        })
    )
  );
}

async function captureElToPng(el, scale = 2) {
  if (!el) return "";
  if (typeof html2canvas !== "function") return "";
  try {
    const rect = el.getBoundingClientRect();
    const width = Math.max(
      1,
      Math.round(Math.max(el.scrollWidth || 0, rect.width, el.clientWidth || 0))
    );
    const height = Math.max(
      1,
      Math.round(Math.max(el.scrollHeight || 0, rect.height, el.clientHeight || 0))
    );
    const hasEngagement = Boolean(
      el.id === "slideEngagement" ||
        (el.classList && el.classList.contains("slide-engagement")) ||
        (el.querySelector && el.querySelector(".engagement-text"))
    );
    const capturePromise = html2canvas(el, {
      scale,
      backgroundColor: "#ffffff",
      useCORS: true,
      allowTaint: false,
      width,
      height,
      windowWidth: width,
      windowHeight: height,
      logging: false,
      onclone: hasEngagement ? (doc) => engagementSlideOnClone(doc) : undefined
    }).catch((err) => {
      console.warn("captureElToPng html2canvas failed:", err);
      return null;
    });

    const timeoutMs = 12000;
    const timeoutPromise = new Promise((resolve) =>
      setTimeout(() => resolve("__timeout__"), timeoutMs)
    );

    const result = await Promise.race([capturePromise, timeoutPromise]);
    if (!result || result === "__timeout__") return "";
    return result.toDataURL("image/png");
  } catch (e) {
    console.warn("captureElToPng failed:", e);
    return "";
  }
}

function canvasHasRenderedPixels(canvas) {
  try {
    if (!canvas || typeof canvas.getContext !== "function") return false;
    const w = canvas.width;
    const h = canvas.height;
    if (!w || !h) return false;
    if (w < 5 || h < 5) return false;
    const ctx = canvas.getContext("2d");
    if (!ctx || typeof ctx.getImageData !== "function") return false;
    const x = Math.floor(w / 2);
    const y = Math.floor(h / 2);
    const d = ctx.getImageData(x, y, 1, 1).data;
    // If the canvas hasn't been drawn, alpha is often 0 (transparent).
    return d && d.length >= 4 && d[3] > 0;
  } catch (e) {
    return true; // If we can't inspect, don't block export.
  }
}

async function waitForPageCanvases(page, timeoutMs = 4000) {
  const canvases = Array.from(page.querySelectorAll("canvas"));
  if (!canvases.length) return;

  const start = performance.now();
  while (performance.now() - start < timeoutMs) {
    const ok = canvases.every((c) => canvasHasRenderedPixels(c));
    if (ok) return;
    await new Promise((r) => setTimeout(r, 200));
  }
}

async function exportPptxFromDom() {
  await ensureDefaultPuzzleDataUrl();
  applyData();
  document.body.classList.add("pptx-export-capture");
  try {
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    if (totPotIncChart) totPotIncChart.resize();
    if (fortiSiemAlertsChart) {
      syncFortiSiemSlideLayout();
    }
    if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
    if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
    syncEpsSlideLayout();
    syncEpsEventsSlideLayout();
    if (trendChart) trendChart.resize();
    if (ruleSeverityChart) ruleSeverityChart.resize();
    if (responseSlaChart) responseSlaChart.resize();
    if (remediationSlaChart) remediationSlaChart.resize();
    window.dispatchEvent(new Event("resize"));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    await new Promise((r) => setTimeout(r, 750));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    if (document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (e) {
        // Ignore font readiness errors; continue export.
      }
    }

    const { w: slideW, h: slideH } = getPptSlideInches();
    const pptx = new PptxGenJS();
    pptx.layout = pptLayoutConfig.LAYOUT_16X9.pptxLayout;

    pptx.title = `${getValue("customerName") || "Customer"} Monthly SOC Report`;

    const addEngagementFallbackSlide = () => {
      addEngagementNativeSlide(pptx, slideW, slideH);
    };

    const pages = Array.from(document.querySelectorAll("#reportRoot .page"));
    for (const page of pages) {
      // Ensure the element is rendered with correct layout before html2canvas capture.
      page.scrollIntoView({ block: "center", inline: "nearest" });
      await new Promise((r) => setTimeout(r, 400));
      await waitForPageCanvases(page, 5000);
      await waitForImagesInElement(page);
      let img = await captureElToPng(page, 2);
      if (!img) img = await captureElToPng(page, 1.4);
      if (!img) {
        page.scrollIntoView({ block: "center", inline: "nearest" });
        await new Promise((r) => setTimeout(r, 400));
        img = await captureElToPng(page, 1);
      }
      if (!img) {
        console.warn("Page capture still blank, using low-scale fallback for:", page.id || page.className);
        img = await captureElToPng(page, 0.8);
      }
      if (img) {
        const slide = pptx.addSlide();
        slide.addImage({ data: img, x: 0, y: 0, w: slideW, h: slideH });
      } else if (page.id === "slideEngagement") {
        console.warn("Using engagement fallback slide generation.");
        addEngagementFallbackSlide();
      } else {
        const slide = pptx.addSlide();
        slide.addText("Slide capture failed for this page. Please re-export after refresh.", {
          x: 0.6,
          y: 2.6,
          w: slideW - 1.2,
          h: 0.5,
          fontSize: 16,
          bold: true,
          color: "AA0000",
          align: "center"
        });
      }
    }

    await pptx.writeFile({ fileName: `${sanitizeFilename(getValue("customerName") || "Customer")}_Monthly_SOC_Report.pptx`, compression: true });
  } finally {
    document.body.classList.remove("pptx-export-capture");
  }
}

async function exportPptx() {
  return exportPptxLegacy();
}

/** Slide images matching the browser preview (not editable in PowerPoint). */
async function exportPptxPreviewImages() {
  return exportPptxFromDom();
}

/** Native PptxGenJS slides (editable text); layout differs from the web preview. */
async function exportPptxEditableNative() {
  return exportPptxLegacy();
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

function exportPptxContact(pptx, sx, sy) {
  const slide = pptx.addSlide();
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 10 * sx,
    h: 5.625 * sy,
    fill: { type: "none" },
    line: { color: "BFBFBF", width: 1 }
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 1.6 * sx, y: 1.0 * sy, w: 2.7 * sx });
  }
  slide.addText("Your Trusted Security Advisor", {
    x: 1.6 * sx,
    y: 2.5 * sy,
    w: 7.4 * sx,
    h: 0.7 * sy,
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
      {
        text: "CISO & Practice Head - Cyber Security & Managed Security",
        options: { color: "1e4f9a", fontSize: 16 }
      },
      { text: "Contact No. 7338882888", options: { color: "1e4f9a", fontSize: 16 } },
      { text: "diptesh.s@snsin.com", options: { color: "1e4f9a", fontSize: 16 } }
    ],
    {
      x: 1.6 * sx,
      y: 3.3 * sy,
      w: 7.4 * sx,
      h: 2.0 * sy,
      align: "left",
      fontFace: "Calibri",
      lineSpacingMultiple: 1.2,
      valign: "top"
    }
  );
  slide.addImage({
    data: getFooterBarDataUrl(),
    x: 0,
    y: 5.3 * sy,
    w: 10 * sx,
    h: 0.325 * sy
  });
}

async function exportPptxLegacy() {
  const pptxBtn = document.getElementById("pptxBtn");
  if (pptxBtn?.disabled) return;
  const originalLabel = pptxBtn ? pptxBtn.textContent : "";
  if (pptxBtn) {
    pptxBtn.disabled = true;
    pptxBtn.textContent = "Preparing…";
  }

  const restore = await prepareReportForExport();
  try {
  await ensureDefaultPuzzleDataUrl();
  applyData();
  await waitForChartsReady(400);

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
  slide.addText(`${getValue("customerName") || "Customer"} Monthly SOC Report`, {
    x: margin,
    y: 1.1 * sy,
    w: contentW,
    fontSize: 24,
    bold: true,
    color: "12284A",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }

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

  addEngagementNativeSlide(pptx, slideW, slideH);

  const monthLabel = getTableMonthLabel();
  const totPotEntries = getTotPotIncSeverityEntries();
  const maxAlert = Math.max(...totPotEntries.map((entry) => entry.value), 1);
  const yAxisMax = Math.max(10, Math.ceil((maxAlert * 1.1) / 5) * 5);
  const yAxisMajorUnit = yAxisMax <= 10 ? 2 : 5;
  const fortiHigh = Math.max(0, parseFloat(getValue("fortiAlertHigh")) || 0);
  const fortiMed = Math.max(0, parseFloat(getValue("fortiAlertMedium")) || 0);
  const fortiLow = Math.max(0, parseFloat(getValue("fortiAlertLow")) || 0);
  const fortiPeak = Math.max(fortiHigh, fortiMed, fortiLow, 1);
  const fortiYMax = Math.max(2000, Math.ceil((fortiPeak * 1.1) / 2000) * 2000);

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
    slide.addImage({
      data: getHiResChartDataUrl(trendCanvasForPpt, trendChart),
      x: 0.45 * sx,
      y: 0.9 * sy,
      w: 9.1 * sx,
      h: 3.3 * sy
    });
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
  const riskRows = parseRiskCsvRows(getValue("riskCsv"));
  const riskChunks = chunkRiskRowsForPptx(riskRows);
  const riskNarrativePptx = parseRichTextPptx(getValue("riskNarrative"));
  const riskCellOpts = { fontFace: "Calibri", valign: "middle" };

  riskChunks.forEach((chunk, i) => {
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
      slide.addText(riskNarrativePptx, {
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

  slide = pptx.addSlide();
  slide.addText("Rule-Based Severity Categories For Potential Incidents", {
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
  const ruleSeverityCanvas = document.getElementById("ruleSeverityChart");
  if (ruleSeverityCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(ruleSeverityCanvas, ruleSeverityChart),
      x: 0.35 * sx,
      y: 0.9 * sy,
      w: 9.3 * sx,
      h: 4.15 * sy
    });
  }

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
      data: getHiResChartDataUrl(responseCanvas, responseSlaChart),
      x: 0.35 * sx,
      y: 1.0 * sy,
      w: 2.9 * sx,
      h: 1.8 * sy
    });
  }
  const remediationCanvas = document.getElementById("remediationSlaChart");
  if (remediationCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(remediationCanvas, remediationSlaChart),
      x: 0.35 * sx,
      y: 3.45 * sy,
      w: 2.9 * sx,
      h: 1.8 * sy
    });
  }
  const responseTbl = [
    [
      { text: "Severity Level Description", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Severity Level", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Response", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Resolution/Remediation", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } }
    ],
    [
      "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community",
      "S-1",
      "≤15 Minutes",
      "4 Hours"
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      "S-2",
      "≤15 Minutes",
      "8 Hours"
    ],
    [
      "Low Severity A Request/Query/Service that does not change existing service structure and no cost implication",
      "S-3",
      "≤30 Minutes",
      "24 Hours"
    ]
  ];
  slide.addTable(responseTbl, {
    x: 3.45 * sx,
    y: 0.75 * sy,
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [3.8 * sx, 1.1 * sx, 1.3 * sx, 1.35 * sx],
    border: { pt: 0.6, color: "C8D8DD" }
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
  const remediationTbl = [
    [
      { text: "Severity Level Description", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Severity Level", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Resolution/Remediation", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } }
    ],
    [
      "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community",
      "High",
      "4 Hours"
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      "Medium",
      "8 Hours"
    ],
    [
      "Low A Request/Query/Service that does not change existing service structure and no cost implication",
      "Low",
      "24 Hours"
    ]
  ];
  slide.addTable(remediationTbl, {
    x: 3.45 * sx,
    y: 3.2 * sy,
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [4.55 * sx, 1.0 * sx, 1.45 * sx],
    border: { pt: 0.6, color: "C8D8DD" }
  });
  slide.addText(formatRemediationSummaryText(closedIncidentCount, remediationInfo.label), {
    x: 3.45 * sx,
    y: 5.15 * sy,
    w: 6.2 * sx,
    h: 0.35 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri"
  });

  slide = pptx.addSlide();
  slide.addText("Total Alerts Triggered in FortiSIEM", {
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
  const fortiBandEl = document.getElementById("fortiSiemBandedBlock");
  const fortiBandImg = await captureElToPng(fortiBandEl, 2);
  if (fortiBandImg) {
    try {
      slide.addImage({
        data: fortiBandImg,
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.55 * sy
      });
    } catch (e) {
      console.warn("fortiBandImg addImage failed, fallback:", e);
      slide.addChart(
        barChartType,
        [
          { name: "High", labels: [monthLabel], values: [fortiHigh] },
          { name: "Medium", labels: [monthLabel], values: [fortiMed] },
          { name: "Low", labels: [monthLabel], values: [fortiLow] }
        ],
        {
          x: 0.7 * sx,
          y: 0.95 * sy,
          w: 8.6 * sx,
          h: 2.55 * sy,
          barDir: "col",
          barGrouping: "clustered",
          chartColors: ["FF0000", "FFFF00", "00B050"],
          showLegend: true,
          legendPos: "b",
          valAxisMaxVal: fortiYMax,
          valAxisMinVal: 0,
          valAxisMajorUnit: 2000,
          showDataTable: true,
          dataTableFontSize: 9,
          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 10
        }
      );
    }
  } else {
    slide.addChart(
      barChartType,
      [
        { name: "High", labels: [monthLabel], values: [fortiHigh] },
        { name: "Medium", labels: [monthLabel], values: [fortiMed] },
        { name: "Low", labels: [monthLabel], values: [fortiLow] }
      ],
      {
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.55 * sy,
        barDir: "col",
        barGrouping: "clustered",
        chartColors: ["FF0000", "FFFF00", "00B050"],
        showLegend: true,
        legendPos: "b",
        valAxisMaxVal: fortiYMax,
        valAxisMinVal: 0,
        valAxisMajorUnit: 2000,
        showDataTable: true,
        dataTableFontSize: 9,
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 10
      }
    );
  }

  slide = pptx.addSlide();
  slide.addText("Total Number Of True & False Positive Alerts", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 26,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("True Positive Alerts based on Severity", {
    x: 0.35 * sx,
    y: 0.88 * sy,
    w: 4.6 * sx,
    h: 0.35 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("False Positive Alerts based on Severity", {
    x: 5.05 * sx,
    y: 0.88 * sy,
    w: 4.6 * sx,
    h: 0.35 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  const tpCanvasEl = document.getElementById("truePositiveAlertsChart");
  const fpCanvasEl = document.getElementById("falsePositiveAlertsChart");
  if (tpCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(tpCanvasEl, truePositiveAlertsChart),
      x: 0.35 * sx,
      y: 1.15 * sy,
      w: 4.6 * sx,
      h: 2.65 * sy
    });
  }
  if (fpCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(fpCanvasEl, falsePositiveAlertsChart),
      x: 5.05 * sx,
      y: 1.15 * sy,
      w: 4.6 * sx,
      h: 2.65 * sy
    });
  }
  slide.addText(formatTfPosNoteForPptx("tfPosNoteTrue"), {
    x: 0.35 * sx,
    y: 3.95 * sy,
    w: 4.6 * sx,
    h: 1.85 * sy,
    fontSize: 9,
    color: "000000",
    valign: "top",
    fontFace: "Calibri"
  });
  slide.addText(formatTfPosNoteForPptx("tfPosNoteFalse"), {
    x: 5.05 * sx,
    y: 3.95 * sy,
    w: 4.6 * sx,
    h: 1.85 * sy,
    fontSize: 9,
    color: "000000",
    valign: "top",
    fontFace: "Calibri"
  });

  const epsTableMonthLabel = monthLabel;
  const epsTotalValue = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const epsRows = parseEpsTopHostsRows(getValue("epsTopHostsCsv"));
  const epsLegendRows = epsRows.length ? epsRows : [{ host: "N/A", eps: 0 }];

  slide = pptx.addSlide();
  slide.addText("EPS Trend Plot", {
    x: margin,
    y: 0.32 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 30,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const epsTotalCanvasEl = document.getElementById("epsTotalChart");
  if (epsTotalCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(epsTotalCanvasEl, epsTotalChart),
      x: 0.12 * sx,
      y: 0.82 * sy,
      w: 3.9 * sx,
      h: 3.95 * sy
    });
  }
  const epsTopCanvasEl = document.getElementById("epsTopHostsChart");
  if (epsTopCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(epsTopCanvasEl, epsTopHostsChart),
      x: 4.52 * sx,
      y: 0.82 * sy,
      w: 5.25 * sx,
      h: 2.45 * sy
    });
  }
  slide.addTable(
    [
      [{ text: epsTableMonthLabel, options: { align: "left" } }, { text: String(epsTotalValue), options: { align: "center", bold: true } }]
    ],
    {
      x: 0.12 * sx,
      y: 4.68 * sy,
      w: 3.9 * sx,
      fontSize: 11,
      colW: [1.9 * sx, 2.0 * sx],
      border: { pt: 0.6, color: "D0D0D0" }
    }
  );
  const legendTableRows = [];
  for (let i = 0; i < 5; i += 1) {
    const left = epsLegendRows[i];
    const right = epsLegendRows[i + 5];
    legendTableRows.push([
      left ? `${i + 1} ${left.host}` : "",
      right ? `${i + 6} ${right.host}` : ""
    ]);
  }
  slide.addTable(legendTableRows, {
    x: 4.52 * sx,
    y: 3.35 * sy,
    w: 5.25 * sx,
    h: 1.5 * sy,
    fontSize: 8.5,
    colW: [2.55 * sx, 2.55 * sx],
    border: { pt: 0.5, color: "D0D0D0" }
  });

  const epsEventsRows = parseEpsEventsRows(getValue("epsEventsCsv"));
  slide = pptx.addSlide();
  slide.addText("Highest EPS Consuming Events For Mentioned Firewalls", {
    x: margin,
    y: 0.42 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 20,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const epsEventsTableData = [
    [
      { text: "S.No", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Reporting Device", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Event Type", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Event Name", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Matched Events", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } }
    ]
  ];
  if (epsEventsRows.length) {
    epsEventsRows.slice(0, 12).forEach((r, idx) => {
      const n = Number(String(r.matchedEvents).replace(/,/g, ""));
      epsEventsTableData.push([
        String(r.sno || idx + 1),
        r.device,
        r.eventType,
        r.eventName,
        Number.isFinite(n) ? n.toLocaleString("en-US") : String(r.matchedEvents || "")
      ]);
    });
  } else {
    epsEventsTableData.push(["-", "No rows found. Upload EPS Events CSV.", "", "", ""]);
  }
  const epsEventsTableH = 4.2 * sy;
  const epsEventsRowH = epsEventsTableData.length
    ? epsEventsTableH / epsEventsTableData.length
    : epsEventsTableH;
  slide.addTable(epsEventsTableData, {
    x: 0.65 * sx,
    y: 1.05 * sy,
    w: 8.7 * sx,
    h: epsEventsTableH,
    fontSize: 8.6,
    rowH: epsEventsTableData.map(() => epsEventsRowH),
    valign: "middle",
    colW: [0.65 * sx, 2.05 * sx, 1.95 * sx, 2.15 * sx, 1.9 * sx],
    border: { pt: 0.6, color: "6CC0C6" }
  });

  const majorRows = parseSupportMajorRows(getValue("supportMajorCsv"));
  const minorRows = parseSupportMinorRows(getValue("supportMinorCsv"));
  const majorDataRows = majorRows.length
    ? majorRows.map((r) => [r.by, String(r.closed), String(r.inProcess), String(r.closed + r.inProcess)])
    : [["-", "0", "0", "0"]];
  const majorClosedTotal = majorRows.reduce((a, r) => a + r.closed, 0);
  const majorInProcessTotal = majorRows.reduce((a, r) => a + r.inProcess, 0);
  majorDataRows.push(["Grand Total", String(majorClosedTotal), String(majorInProcessTotal), String(majorClosedTotal + majorInProcessTotal)]);

  const minorDataRows = minorRows.length ? minorRows.map((r) => [r.by, String(r.closed)]) : [["-", "0"]];
  const minorTotal = minorRows.reduce((a, r) => a + r.closed, 0);
  minorDataRows.push(["Grand Total", String(minorTotal)]);

  slide = pptx.addSlide();
  slide.addText("Overall Support Ticket Handled By SNS (Firewall Support)", {
    x: margin,
    y: 0.45 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 23,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addTable(
    [
      [{ text: "Major tickets", options: { bold: true, align: "center", fill: { color: "082A6F" }, color: "FFFFFF" } }, "", "", ""],
      [
        { text: "Created by", options: { bold: true, align: "center" } },
        { text: "Status", options: { bold: true, align: "center" } },
        "",
        { text: "Grand Total", options: { bold: true, align: "center" } }
      ],
      [
        "",
        { text: "Closed", options: { bold: true, align: "center" } },
        { text: "In-process", options: { bold: true, align: "center" } },
        ""
      ],
      ...majorDataRows
    ],
    {
      x: 0.1 * sx,
      y: 1.2 * sy,
      w: 4.75 * sx,
      fontSize: 10,
      colW: [1.35 * sx, 1.1 * sx, 1.1 * sx, 1.2 * sx],
      border: { pt: 0.6, color: "888888" }
    }
  );
  slide.addTable(
    [
      [{ text: "Minor tickets", options: { bold: true, align: "center", fill: { color: "082A6F" }, color: "FFFFFF" } }, ""],
      [
        { text: "Created by", options: { bold: true, align: "center" } },
        { text: "Status", options: { bold: true, align: "center" } }
      ],
      [
        "",
        { text: "Closed", options: { bold: true, align: "center" } }
      ],
      ...minorDataRows
    ],
    {
      x: 5.15 * sx,
      y: 1.2 * sy,
      w: 4.75 * sx,
      fontSize: 10,
      colW: [2.35 * sx, 2.35 * sx],
      border: { pt: 0.6, color: "888888" }
    }
  );

  const inventoryRows = parseInventoryRows(getValue("inventoryCsv"));
  const inventoryTotal = inventoryRows.reduce((a, r) => a + r.count, 0);
  const inventoryTableData = [
    [
      { text: "S.NO", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } },
      { text: "DEVICE NAME", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } },
      { text: "COUNT", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } }
    ],
    ...(inventoryRows.length
      ? inventoryRows.map((r, i) => [String(i + 1), r.device, r.count.toLocaleString("en-US")])
      : [["-", "No rows found. Upload Inventory CSV.", ""]]),
    ["", "Total", inventoryTotal.toLocaleString("en-US")]
  ];
  slide = pptx.addSlide();
  slide.addText("Integrated Device Inventory", {
    x: margin,
    y: 0.35 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 30,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addTable(inventoryTableData, {
    x: 1.4 * sx,
    y: 1.25 * sy,
    w: 7.2 * sx,
    fontSize: 10.5,
    colW: [1.15 * sx, 4.05 * sx, 2.0 * sx],
    border: { pt: 0.6, color: "E2EDF0" }
  });
  slide.addText(getValue("inventoryNote"), {
    x: 0.65 * sx,
    y: 4.9 * sy,
    w: 8.7 * sx,
    h: 0.55 * sy,
    fontSize: 11,
    color: "111111",
    fontFace: "Calibri"
  });

  slide = pptx.addSlide();
  slide.addText("Key Points - Overall Summary", {
    x: margin,
    y: 0.35 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  addKeyPointsToPptxSlide(slide, sx, sy);

  exportPptxContact(pptx, sx, sy);

  await pptx.writeFile({ fileName: `${sanitizeFilename(getValue("customerName") || "Customer")}_Monthly_SOC_Report.pptx`, compression: true });
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
        handleSlaInputChange();
        return;
      }
      applyData();
    });
  }
  document.querySelectorAll(".panel input, .panel textarea, .panel select").forEach((el) => {
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
        preparePrintSlideOrder("blupine");
      }
      finishBluPinePrintLayout();
      reportPdfLayoutRestore = prepareBluPinePrintLayout();
      const engagementSlide = document.querySelector(
        '#reportRoot .page[data-slide="engagement"]'
      );
      if (engagementSlide) {
        engagementSlide.scrollIntoView({ block: "center" });
      }
      await ensureDefaultPuzzleDataUrl();
      syncEngagementPuzzleImage();
      await new Promise((resolve) => requestAnimationFrame(resolve));
      syncEngagementLayout();
      await new Promise((resolve) => requestAnimationFrame(resolve));
      syncEngagementLayout();
      await new Promise((resolve) => setTimeout(resolve, 150));
      syncEngagementLayout();
      syncAllSlideLayoutsForExport();
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
    exportPptx().catch((err) => console.error("PPTX export failed:", err));
  });

  const panelToggleBtn = document.getElementById("panelToggleBtn");
  const panelBackdrop = document.getElementById("panelBackdrop");
  const panelMobileMq = window.matchMedia("(max-width: 900px)");
  const closePanel = () => document.body.classList.remove("panel-open");
  const openPanel = () => document.body.classList.add("panel-open");

  function bindMobilePanelToggle() {
    if (!panelToggleBtn || !panelMobileMq.matches) {
      return;
    }

    if (panelToggleBtn.dataset.mobileToggleBound === "1") {
      return;
    }
    panelToggleBtn.dataset.mobileToggleBound = "1";

    panelToggleBtn.addEventListener("click", () => {
      if (document.body.classList.contains("panel-open")) {
        closePanel();
      } else {
        openPanel();
      }
    });

    if (panelBackdrop) {
      panelBackdrop.addEventListener("click", closePanel);
    }
  }

  bindMobilePanelToggle();
  panelMobileMq.addEventListener("change", () => {
    if (!panelMobileMq.matches) {
      closePanel();
    } else {
      bindMobilePanelToggle();
    }
  });

  let slaResizeTimer;
  window.addEventListener("resize", () => {
    clearTimeout(slaResizeTimer);
    slaResizeTimer = setTimeout(() => {
      updateSlaSlide();
      syncTfPosSlideLayout();
      syncEpsEventsSlideLayout();
      syncSupportTicketsSlideLayout();
    }, 150);
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
  bindCsvInput("epsTopHostsCsvFile", "epsTopHostsCsv");
  bindCsvInput("epsEventsCsvFile", "epsEventsCsv");
  bindCsvInput("supportMajorCsvFile", "supportMajorCsv");
  bindCsvInput("supportMinorCsvFile", "supportMinorCsv");
  bindCsvInput("inventoryCsvFile", "inventoryCsv");

  await ensureDefaultPuzzleDataUrl();
  applyData();
  let engagementResizeTimer;
  window.addEventListener("resize", () => {
    clearTimeout(engagementResizeTimer);
    engagementResizeTimer = setTimeout(syncEngagementLayout, 150);
  });
  initInlineEditableReport();
  if (typeof initSlideNav === "function") {
    initSlideNav(() => "blupine");
  }
  initEpsEventsSlideLayoutWatcher();
  initSupportTicketsSlideLayoutWatcher();
  if (typeof initTextStyleEditor === "function") {
    initTextStyleEditor();
  }
}

window.onload = init;
