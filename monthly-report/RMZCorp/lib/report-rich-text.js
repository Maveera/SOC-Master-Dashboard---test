/**
 * Shared rich-text formatting for monthly report narrative fields.
 */
(function (global) {
  "use strict";

  function escapeHtml(str) {
    return String(str || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  /**
   * @param {string} text
   * @param {() => string} getMonthLabel
   * @param {{ italic?: boolean }} [options]
   */
  function formatReportRichTextHTML(text, getMonthLabel, options) {
    const opts = options || {};
    if (!text) return "";
    let formatted = escapeHtml(text);
    formatted = formatted.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
    if (opts.italic) {
      formatted = formatted.replace(/\*(.*?)\*/g, "<em>$1</em>");
    }
    if (typeof getMonthLabel === "function") {
      formatted = formatted.replace(/\[MONTH\]/g, `<span class="fw-bold">${getMonthLabel()}</span>`);
    }
    formatted = formatted.replace(/\r\n|\r|\n/g, "<br>");
    return formatted;
  }

  global.formatReportRichTextHTML = formatReportRichTextHTML;
})(typeof window !== "undefined" ? window : globalThis);
