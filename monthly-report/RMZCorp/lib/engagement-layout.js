/**
 * Engagement slide layout (equal-height columns, text fit, PDF export prep).
 * Used by Nocil, RMZCorp, BluPine and other templates with a matching engagement slide.
 */
(function (global) {
  const PDF_EXPORT_CLASSES = ["report-pdf-export", "nocil-pdf-export"];
  const ENGAGEMENT_ABOVE_FOOTER_GAP = 28;
  const ENGAGEMENT_MIN_FOOTER_GAP = 10;

  function isReportPdfExport() {
    return PDF_EXPORT_CLASSES.some((c) => document.body.classList.contains(c));
  }

  function setEngagementStyle(el, prop, value) {
    if (!el) return;
    if (isReportPdfExport()) {
      el.style.setProperty(prop, value, "important");
    } else {
      el.style.setProperty(prop, value);
    }
  }

  function clearEngagementStyle(el, prop) {
    if (!el) return;
    el.style.removeProperty(prop);
  }

  function fitEngagementTextToPanel() {
    const slide = document.querySelector('#reportRoot .page[data-slide="engagement"]');
    const textEl = slide?.querySelector(".engagement-text");
    if (!textEl) return false;

    const narrative = textEl.querySelector(".narrative");
    const paragraphs = [...textEl.querySelectorAll("p")];

    clearEngagementTextFit();

    let sizePx = parseFloat(getComputedStyle(textEl).fontSize) || 15.68;
    let lineHeight = 1.42;
    let paraGap = 12;
    let padV = 16;

    const apply = () => {
      const fontVal = `${sizePx}px`;
      setEngagementStyle(textEl, "font-size", fontVal);
      setEngagementStyle(textEl, "line-height", String(lineHeight));
      setEngagementStyle(textEl, "padding", `${padV}px 18px`);
      paragraphs.forEach((p) => {
        setEngagementStyle(p, "margin-bottom", `${paraGap}px`);
      });
      if (narrative) {
        setEngagementStyle(narrative, "font-size", fontVal);
        setEngagementStyle(narrative, "line-height", String(lineHeight));
      }
    };

    const overflows = () => {
      if (textEl.scrollHeight > textEl.clientHeight + 2) return true;
      const lastP = paragraphs[paragraphs.length - 1];
      if (!lastP) return false;
      const box = textEl.getBoundingClientRect();
      const last = lastP.getBoundingClientRect();
      return last.bottom > box.bottom - 1;
    };

    apply();
    let guard = 0;
    while (guard++ < 140 && overflows()) {
      if (paraGap > 3) {
        paraGap -= 1;
      } else if (lineHeight > 1.14) {
        lineHeight = Math.max(1.14, lineHeight - 0.02);
      } else if (sizePx > 7.5) {
        sizePx -= 0.5;
      } else if (padV > 4) {
        padV -= 1;
      } else {
        return false;
      }
      apply();
    }
    return !overflows();
  }

  function clearEngagementTextFit() {
    const textEl = document.querySelector(
      '#reportRoot .page[data-slide="engagement"] .engagement-text'
    );
    if (!textEl) return;
    ["font-size", "line-height", "padding"].forEach((prop) => clearEngagementStyle(textEl, prop));
    textEl.querySelectorAll("p").forEach((p) => clearEngagementStyle(p, "margin-bottom"));
    const narrative = textEl.querySelector(".narrative");
    if (narrative) {
      ["font-size", "line-height"].forEach((prop) => clearEngagementStyle(narrative, prop));
    }
  }

  function syncEngagementColumnHeights(footerGap = ENGAGEMENT_ABOVE_FOOTER_GAP) {
    const slide = document.querySelector('#reportRoot .page[data-slide="engagement"]');
    const container = slide?.querySelector(".engagement-container");
    const footer = slide?.querySelector(".page-footer-bar");
    const title = slide?.querySelector(".engagement-slide-title, .revision-title");
    if (!slide || !container || !footer || !title) return;

    const top = title.offsetTop + title.offsetHeight + 8;
    const footerTop = footer.offsetTop;
    const height = Math.max(220, Math.floor(footerTop - top - footerGap));

    container.style.flex = "0 0 auto";
    container.style.height = `${height}px`;
    container.style.minHeight = `${height}px`;
    container.style.maxHeight = `${height}px`;
    container.style.marginBottom = "0";
    void container.offsetHeight;
  }

  function syncEngagementLayout() {
    clearEngagementTextFit();
    let gap = ENGAGEMENT_ABOVE_FOOTER_GAP;
    syncEngagementColumnHeights(gap);
    let fits = fitEngagementTextToPanel();
    while (!fits && gap > ENGAGEMENT_MIN_FOOTER_GAP) {
      gap -= 4;
      syncEngagementColumnHeights(gap);
      clearEngagementTextFit();
      fits = fitEngagementTextToPanel();
    }
  }

  function clearEngagementColumnHeights() {
    const container = document.querySelector(
      '#reportRoot .page[data-slide="engagement"] .engagement-container'
    );
    if (!container) return;
    container.style.removeProperty("height");
    container.style.removeProperty("min-height");
    container.style.removeProperty("max-height");
    container.style.removeProperty("margin-bottom");
    container.style.removeProperty("flex");
  }

  function syncPrintCanvasToPreview() {
    const root = document.getElementById("reportRoot");
    if (!root) return;
    root.style.width = "960px";
    root.style.maxWidth = "960px";
  }

  function clearPrintCanvasInline() {
    const root = document.getElementById("reportRoot");
    if (!root) return;
    root.style.removeProperty("width");
    root.style.removeProperty("max-width");
  }

  function prepareReportPdfLayout(onRefresh) {
    document.body.classList.add("report-pdf-export");
    syncPrintCanvasToPreview();
    syncEngagementLayout();
    if (typeof onRefresh === "function") onRefresh();
    return function cleanup() {
      document.body.classList.remove("report-pdf-export");
      clearPrintCanvasInline();
      clearEngagementColumnHeights();
      clearEngagementTextFit();
    };
  }

  global.ReportEngagementLayout = {
    ENGAGEMENT_ABOVE_FOOTER_GAP,
    syncEngagementLayout,
    clearEngagementColumnHeights,
    clearEngagementTextFit,
    fitEngagementNarrativeForPrint: syncEngagementLayout,
    syncPrintCanvasToPreview,
    clearPrintCanvasInline,
    prepareReportPdfLayout
  };
})(typeof window !== "undefined" ? window : globalThis);
