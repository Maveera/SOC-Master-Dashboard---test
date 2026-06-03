/**
 * Shared severity stacked-bar chart helpers (High / Low / Medium).
 */
(function (global) {
  "use strict";

  const SEVERITY_CHART_COLORS = {
    high: "#ff0000",
    low: "#00b050",
    medium: "#ffff00"
  };

  let dataLabelsRegistered = false;

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

  function makeColumnFinder(headerNorm, hasHeader) {
    return (matchers, fallback) => {
      if (!hasHeader) return fallback;
      const pos = headerNorm.findIndex((k) => matchers.some((m) => k.includes(m)));
      return pos >= 0 ? pos : fallback;
    };
  }

  function parseTrendCsvRows(csvText) {
    const lines = String(csvText || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);
    if (!lines.length) return [];

    const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
    const headerNorm = rows[0].map(normalizeKey);
    const hasHeader = headerNorm.some(
      (k) => k.includes("date") || k.includes("day") || k.includes("high") || k.includes("medium") || k.includes("low")
    );
    const dataRows = hasHeader ? rows.slice(1) : rows;
    const findCol = makeColumnFinder(headerNorm, hasHeader);
    const iDate = findCol(["date", "day"], 0);
    const iHigh = findCol(["high"], 1);
    const iMed = findCol(["medium", "med"], 2);
    const iLow = findCol(["low"], 3);

    return dataRows
      .map((r) => ({
        date: r[iDate] || "",
        high: Math.max(0, parseFloat(r[iHigh]) || 0),
        medium: Math.max(0, parseFloat(r[iMed]) || 0),
        low: Math.max(0, parseFloat(r[iLow]) || 0)
      }))
      .filter((r) => r.date && r.high + r.medium + r.low > 0);
  }

  function ensureDataLabelsPlugin() {
    if (dataLabelsRegistered || typeof Chart === "undefined") return false;
    if (typeof ChartDataLabels === "undefined") return false;
    Chart.register(ChartDataLabels);
    if (!Chart.defaults.plugins) Chart.defaults.plugins = {};
    Chart.defaults.plugins.datalabels = { display: false };
    dataLabelsRegistered = true;
    return true;
  }

  function buildSeverityStackedDatasets(rows) {
    return [
      { label: "High", data: rows.map((r) => r.high), backgroundColor: SEVERITY_CHART_COLORS.high, borderWidth: 0, stack: "s" },
      { label: "Low", data: rows.map((r) => r.low), backgroundColor: SEVERITY_CHART_COLORS.low, borderWidth: 0, stack: "s" },
      { label: "Medium", data: rows.map((r) => r.medium), backgroundColor: SEVERITY_CHART_COLORS.medium, borderWidth: 0, stack: "s" }
    ];
  }

  function buildRuleSeverityDatasets(rows) {
    return [
      { label: "High", data: rows.map((r) => r.high), backgroundColor: SEVERITY_CHART_COLORS.high, borderWidth: 0 },
      { label: "Low", data: rows.map((r) => r.low), backgroundColor: SEVERITY_CHART_COLORS.low, borderWidth: 0 },
      { label: "Medium", data: rows.map((r) => r.medium), backgroundColor: SEVERITY_CHART_COLORS.medium, borderWidth: 0 }
    ];
  }

  function buildSquareLegendLabels(color) {
    return {
      boxWidth: 12,
      boxHeight: 12,
      usePointStyle: false,
      color: color || "#595959"
    };
  }

  function buildSeverityChartPlugins(titleText, options) {
    const opts = options || {};
    const legendPosition = opts.legendPosition || "bottom";
    const fontSize = opts.fontSize || 14;
    const legendColor = opts.legendColor || "#595959";
    const legendTitle = opts.legendTitle;

    ensureDataLabelsPlugin();
    const plugins = {
      title: {
        display: true,
        text: titleText,
        font: { size: fontSize, weight: "bold" }
      },
      legend: {
        display: true,
        position: legendPosition,
        labels: buildSquareLegendLabels(legendColor)
      }
    };
    if (legendTitle) {
      plugins.legend.title = {
        display: true,
        text: legendTitle,
        color: legendColor,
        font: { weight: "bold" }
      };
    }
    if (dataLabelsRegistered) {
      plugins.datalabels = {
        display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
        color: "#000",
        font: { weight: "bold", size: 10 },
        formatter: (v) => Math.round(Number(v)),
        anchor: "center",
        align: "center",
        clamp: true
      };
    }
    return plugins;
  }

  function getSeverityStackedMax(rows) {
    return Math.max(...rows.map((r) => r.high + r.low + r.medium), 1);
  }

  function rowHasTrendIncidentData(row) {
    if (!row) return false;
    return Number(row.high) + Number(row.medium) + Number(row.low) > 0;
  }

  /** Keep only dates with at least one High/Medium/Low count. */
  function filterTrendRowsWithIncidentData(rows) {
    return (rows || []).filter(rowHasTrendIncidentData);
  }

  function buildTrendChartXScaleOptions(options) {
    const opts = options || {};
    const rowCount = Math.max(Number(opts.rowCount) || 0, 1);
    const scale = {
      stacked: true,
      grid: { display: false, drawBorder: false },
      ticks: {
        color: opts.tickColor || "#000",
        maxRotation: opts.maxRotation != null ? opts.maxRotation : 45,
        minRotation: opts.minRotation != null ? opts.minRotation : 45,
        autoSkip: false,
        maxTicksLimit: rowCount,
        font: opts.font
      }
    };
    if (opts.title) {
      scale.title = { display: true, text: opts.title, color: opts.tickColor || "#000" };
    }
    return scale;
  }

  function buildTrendChartYScaleOptions(rows, options) {
    const opts = options || {};
    const stackedMax = getSeverityStackedMax(rows);
    const yMax = Math.max(6, Math.ceil(stackedMax));
    const scale = {
      stacked: true,
      beginAtZero: true,
      min: 0,
      max: yMax,
      ticks: { stepSize: 1, color: opts.tickColor || "#000", font: opts.font },
      grid: { color: "#e0e0e0", drawBorder: false }
    };
    if (opts.title) {
      scale.title = { display: true, text: opts.title, color: opts.tickColor || "#000" };
    }
    return scale;
  }

  /** EPS Trend Plot left chart: datalabel shows EPS count above the bar. */
  function buildEpsTrendTotalValueDataLabelOptions(options) {
    const opts = options || {};
    const font = opts.font || { weight: "bold", size: 12, family: "Calibri, Segoe UI, Arial, sans-serif" };
    return {
      display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
      color: opts.color || "#000000",
      font,
      formatter: (v) => {
        const n = Number(v);
        return Number.isFinite(n) ? String(Math.round(n)) : "";
      },
      anchor: "end",
      align: "top",
      offset: opts.offset != null ? opts.offset : 4,
      clamp: true,
      clip: false
    };
  }

  global.ReportCharts = {
    SEVERITY_CHART_COLORS,
    parseTrendCsvRows,
    ensureDataLabelsPlugin,
    isDataLabelsRegistered: () => dataLabelsRegistered,
    buildSeverityStackedDatasets,
    buildRuleSeverityDatasets,
    buildSquareLegendLabels,
    buildSeverityChartPlugins,
    getSeverityStackedMax,
    rowHasTrendIncidentData,
    filterTrendRowsWithIncidentData,
    buildTrendChartXScaleOptions,
    buildTrendChartYScaleOptions,
    buildEpsTrendTotalValueDataLabelOptions
  };

  ensureDataLabelsPlugin();
})(typeof window !== "undefined" ? window : globalThis);
