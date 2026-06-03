/**
 * Ensures browser print/PDF page order matches TEMPLATE_REPORT_SECTIONS (same as PPTX).
 */
(function (global) {
  "use strict";

  let domOrderSnapshot = null;

  function snapshotReportDomOrder() {
    const reportRoot = document.getElementById("reportRoot");
    if (!reportRoot) return null;
    return [...reportRoot.childNodes];
  }

  function restoreReportDomOrder() {
    const reportRoot = document.getElementById("reportRoot");
    if (!reportRoot || !domOrderSnapshot) return;
    domOrderSnapshot.forEach((node) => {
      if (node.parentNode !== reportRoot) {
        reportRoot.appendChild(node);
      }
    });
    domOrderSnapshot = null;
    const riskHost = document.getElementById("riskSlidesContainer");
    if (riskHost) riskHost.style.removeProperty("display");
  }

  function normalizeReportSlideOrderForPrint(templateKey) {
    const reportRoot = document.getElementById("reportRoot");
    const order =
      global.TEMPLATE_REPORT_SECTIONS && global.TEMPLATE_REPORT_SECTIONS[templateKey];
    if (!reportRoot || !order || !order.length) return;

    const sequence = [];
    order.forEach((sectionId) => {
      if (sectionId === "risks") {
        const host = document.getElementById("riskSlidesContainer");
        if (host) {
          sequence.push(...host.querySelectorAll(":scope > .risk-slide"));
        }
        return;
      }
      sequence.push(...reportRoot.querySelectorAll(`:scope > section[data-slide="${sectionId}"]`));
    });

    sequence.forEach((slide) => reportRoot.appendChild(slide));

    const riskHost = document.getElementById("riskSlidesContainer");
    if (riskHost) {
      if (!riskHost.querySelector(".risk-slide")) {
        riskHost.style.display = "none";
      }
      reportRoot.appendChild(riskHost);
    }
  }

  function preparePrintSlideOrder(templateKey) {
    domOrderSnapshot = snapshotReportDomOrder();
    normalizeReportSlideOrderForPrint(templateKey);
  }

  function finishPrintSlideOrder() {
    restoreReportDomOrder();
  }

  global.preparePrintSlideOrder = preparePrintSlideOrder;
  global.finishPrintSlideOrder = finishPrintSlideOrder;
  global.normalizeReportSlideOrderForPrint = normalizeReportSlideOrderForPrint;
})(typeof window !== "undefined" ? window : globalThis);
