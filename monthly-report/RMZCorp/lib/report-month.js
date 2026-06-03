/**
 * Manual report month + table month helpers for monthly report templates.
 */
(function (global) {
  function parseDateRangeEnd(str) {
    const endPart = String(str || "").split(/\s+to\s+/i).pop().trim();
    const match = endPart.match(/(\d{2})-(\d{2})-(\d{4})/);
    if (!match) return null;
    const monthIdx = Number(match[2]) - 1;
    const year = Number(match[3]);
    if (monthIdx < 0 || monthIdx > 11) return null;
    return { monthIdx, year };
  }

  function resolveManualReportMonth(manual, dateRange, submittedOn) {
    const custom = String(manual || "").trim();
    if (!custom) return null;
    if (/\b(19|20)\d{2}\b/.test(custom)) return custom;

    let year = parseDateRangeEnd(dateRange)?.year;
    if (!year) {
      const sm = String(submittedOn || "").match(/(\d{2})-(\d{2})-(\d{4})/);
      if (sm) year = Number(sm[3]);
    }
    if (!year) year = new Date().getFullYear();
    return `${custom} ${year}`;
  }

  function tableMonthFromManual(manual, reportMonth) {
    const custom = String(manual || "").trim();
    if (custom) {
      const short = custom.split(/\s+/)[0];
      return short || custom;
    }
    const full = String(reportMonth || "").trim();
    return full.split(/\s+/)[0] || full;
  }

  /** X-axis category under the EPS Trend Plot bar — EPS Month Label field only (not report month). */
  function resolveEpsChartAxisLabel(epsMonthLabelField) {
    const custom = String(epsMonthLabelField || "").trim();
    return custom || "AVG EPS";
  }

  global.ReportMonthUtils = {
    parseDateRangeEnd,
    resolveManualReportMonth,
    tableMonthFromManual,
    resolveEpsChartAxisLabel
  };
})(typeof window !== "undefined" ? window : globalThis);
