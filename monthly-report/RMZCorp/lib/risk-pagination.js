/**
 * Shared overflow-aware risk table pagination (Apraava pattern).
 */
(function (global) {
  "use strict";

  function buildRiskTableBodyHtml(chunk, escapeHtml) {
    if (!chunk.length) {
      return `<tr><td colspan="6">No risk rows found. Upload Risk CSV.</td></tr>`;
    }
    return chunk
      .map(
        (r) => `<tr>
  <td class="risk-col-sno">${escapeHtml(String(r.sno))}</td>
  <td class="risk-col-attack">${escapeHtml(r.attackType)}</td>
  <td class="risk-col-scenario">${escapeHtml(r.riskScenario)}</td>
  <td class="risk-col-cia">${escapeHtml(r.ciaTriad)}</td>
  <td class="risk-col-impact">${escapeHtml(r.businessImpact)}</td>
  <td class="risk-col-rating">${escapeHtml(r.riskRating)}</td>
</tr>`
      )
      .join("");
  }

  function buildRiskSlideHtml({
    title,
    chunk,
    narrativeHtml = "",
    includeNarrative = false,
    isContinuation = false,
    escapeHtml
  }) {
    return `
    <section class="page page-with-footer risk-slide risk-slide-compact" data-slide="risks" data-risk-rows="${chunk.length}" data-risk-contd="${isContinuation ? "true" : "false"}">
      <div class="page-header-right">
        <div class="report-logo-slot">
          <img class="nav-sns-logo" alt="SNS Logo" />
        </div>
      </div>
      <h2 class="slide-title revision-title risk-slide-title">${escapeHtml(title)}</h2>
      ${includeNarrative ? `<div class="narrative slide-narrative risk-slide-narrative">${narrativeHtml}</div>` : ""}
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
          <tbody>${buildRiskTableBodyHtml(chunk, escapeHtml)}</tbody>
        </table>
      </div>
      <div class="page-footer-bar"></div>
    </section>
  `;
  }

  function getRiskSlideElements(hostId) {
    const host = document.getElementById(hostId || "riskSlidesContainer");
    if (!host) return [];
    return [...host.querySelectorAll(":scope > .risk-slide")];
  }

  function getRiskMeasureStage() {
    let stage = document.getElementById("riskMeasureStage");
    if (stage) {
      if (stage.parentElement && stage.parentElement.id === "reportRoot") {
        document.body.appendChild(stage);
      }
      return stage;
    }
    stage = document.createElement("div");
    stage.id = "riskMeasureStage";
    stage.setAttribute("aria-hidden", "true");
    stage.style.cssText =
      "position:fixed;left:0;top:0;width:960px;opacity:0;pointer-events:none;z-index:-1;";
    document.body.appendChild(stage);
    return stage;
  }

  function mountRiskMeasureSlide(html) {
    const stage = getRiskMeasureStage();
    stage.innerHTML = `<div id="reportRoot" class="report ppt-wide">${html}</div>`;
    return stage.querySelector(".risk-slide");
  }

  /** Space between table top and footer — table may extend past footer when clipped by overflow:hidden. */
  function getRiskTableBudgetPx(slide) {
    const table = slide.querySelector(".risk-table");
    const footer = slide.querySelector(".page-footer-bar");
    if (!table || !footer) return 0;
    return footer.offsetTop - table.offsetTop - 2;
  }

  function riskSlideTableOverflows(slide, tolerancePx = 6) {
    const wrap = slide.querySelector(".risk-table-wrap");
    const table = slide.querySelector(".risk-table");
    const footer = slide.querySelector(".page-footer-bar");
    if (!table || !footer) return false;

    if (wrap?.classList.contains("risk-table-wrap--fill")) {
      const need = Math.max(table.offsetHeight, table.scrollHeight);
      if (need > wrap.clientHeight - tolerancePx) return true;
    }

    const anchorTop = wrap?.offsetTop ?? table.offsetTop;
    const budget = footer.offsetTop - anchorTop - tolerancePx;
    if (budget <= 0) return true;
    const need = Math.max(table.offsetHeight, table.scrollHeight);
    return need > budget;
  }

  function riskChunkFitsOnSlide(
    chunk,
    isFirstSlide,
    narrativeHtml,
    includeNarrative,
    escapeHtml,
    contdTitle,
    layoutOptions
  ) {
    const slide = mountRiskMeasureSlide(
      buildRiskSlideHtml({
        title: isFirstSlide ? "Measure" : contdTitle,
        chunk,
        narrativeHtml,
        includeNarrative,
        isContinuation: !isFirstSlide,
        escapeHtml
      })
    );
    if (slide) applyRiskTableFontSizeToSlide(slide, layoutOptions);
    return Boolean(slide) && !riskSlideTableOverflows(slide, 3);
  }

  function applyRiskTableFontSizeToSlide(slide, layoutOptions) {
    const { tableFontSizePx } = parseLayoutOptions(layoutOptions);
    if (!tableFontSizePx || !slide) return;
    slide.querySelectorAll(".risk-table").forEach((table) => {
      table.style.fontSize = `${tableFontSizePx}px`;
    });
  }

  /** Pack as many rows as fit per slide (adjustable height — no empty gap below table). */
  const DEFAULT_ROWS_PER_SLIDE = 3;

  function parseLayoutOptions(layoutOptions) {
    const maxRows = parseInt(layoutOptions?.maxRowsPerSlide, 10);
    const fontPx = parseFloat(layoutOptions?.tableFontSizePx);
    const userMax = Number.isFinite(maxRows) && maxRows > 0 ? maxRows : 0;
    return {
      maxRowsPerSlide: userMax,
      rowsPerSlideTarget: userMax > 0 ? userMax : DEFAULT_ROWS_PER_SLIDE,
      tableFontSizePx: Number.isFinite(fontPx) && fontPx > 0 ? fontPx : 0,
      skipFitOnColumnResize: layoutOptions?.skipFitOnColumnResize === true
    };
  }

  function chunkRiskRowsSequential(
    rows,
    narrativeText,
    formatRichTextHTML,
    escapeHtml,
    contdTitle,
    layoutOptions
  ) {
    if (!rows.length) return [[]];
    const { maxRowsPerSlide, rowsPerSlideTarget } = parseLayoutOptions(layoutOptions);

    if (maxRowsPerSlide > 0) {
      const chunks = [];
      for (let i = 0; i < rows.length; i += maxRowsPerSlide) {
        chunks.push(rows.slice(i, i + maxRowsPerSlide));
      }
      return chunks;
    }

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
        if (
          riskChunkFitsOnSlide(
            chunk,
            isFirst,
            narrativeHtml,
            includeNarrative,
            escapeHtml,
            contdTitle,
            layoutOptions
          )
        ) {
          best = tryCount;
        } else {
          break;
        }
      }
      const take = Math.max(1, Math.min(best, rowsPerSlideTarget, rows.length - index));
      chunks.push(rows.slice(index, index + take));
      index += take;
    }
    return mergeOrphanLastRiskChunk(chunks, escapeHtml, contdTitle, layoutOptions);
  }

  function mergeOrphanLastRiskChunk(chunks, escapeHtml, contdTitle, layoutOptions) {
    if (chunks.length < 2 || chunks[chunks.length - 1].length !== 1) return chunks;
    const prev = chunks[chunks.length - 2];
    const combined = [...prev, ...chunks[chunks.length - 1]];
    if (riskChunkFitsOnSlide(combined, false, "", false, escapeHtml, contdTitle, layoutOptions)) {
      return [...chunks.slice(0, -2), combined];
    }
    return chunks;
  }

  /** Clear forced row heights (legacy); never compress rows below content. */
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

  function ensureRiskSlidesNoOverflow(hostId, escapeHtml, contdTitle, layoutOptions) {
    getRiskSlideElements(hostId).forEach(resetRiskTableRowHeights);
    const { maxRowsPerSlide } = parseLayoutOptions(layoutOptions || {});
    if (maxRowsPerSlide > 0) {
      removeEmptyRiskSlides(hostId);
      syncRiskContdAttributes(hostId);
      return;
    }
    let guard = 0;
    while (guard++ < 100 && rebalanceRiskSlidesTrimOverflow(hostId, escapeHtml, contdTitle)) {
      /* move overflowing rows to next slide until stable */
    }
    removeEmptyRiskSlides(hostId);
    syncRiskContdAttributes(hostId);
  }

  function countRiskRowsInSlides(hostId) {
    return getRiskSlideElements(hostId).reduce((total, slide) => {
      const body = slide.querySelector(".risk-table tbody");
      if (!body) return total;
      const rows = [...body.querySelectorAll("tr")].filter(
        (tr) => !tr.querySelector("td[colspan]")
      );
      return total + rows.length;
    }, 0);
  }

  function measureRiskRowsPerSlide(rows, startIndex, isFirstSlide, narrativeHtml, escapeHtml, contdTitle) {
    const stage = getRiskMeasureStage();
    let fitCount = 0;
    const remaining = rows.length - startIndex;
    for (let tryCount = 1; tryCount <= remaining; tryCount += 1) {
      const chunk = rows.slice(startIndex, startIndex + tryCount);
      const slide = mountRiskMeasureSlide(
        buildRiskSlideHtml({
          title: isFirstSlide ? "Measure" : contdTitle,
          chunk,
          narrativeHtml,
          includeNarrative: isFirstSlide && Boolean(String(narrativeHtml || "").trim()),
          isContinuation: !isFirstSlide,
          escapeHtml
        })
      );
      if (!slide) continue;
      if (riskSlideTableOverflows(slide)) break;
      fitCount = tryCount;
    }
    return Math.max(1, fitCount || 1);
  }

  function chunkRiskRowsForSlides(
    rows,
    narrativeText,
    formatRichTextHTML,
    escapeHtml,
    contdTitle,
    layoutOptions
  ) {
    return chunkRiskRowsSequential(
      rows,
      narrativeText,
      formatRichTextHTML,
      escapeHtml,
      contdTitle,
      layoutOptions
    );
  }

  function applyRiskTableFontSize(hostId, layoutOptions) {
    const { tableFontSizePx } = parseLayoutOptions(layoutOptions);
    if (!tableFontSizePx) return;
    getRiskSlideElements(hostId).forEach((slide) => {
      slide.querySelectorAll(".risk-table").forEach((table) => {
        table.style.fontSize = `${tableFontSizePx}px`;
      });
    });
  }

  function appendRiskSlide(host, options) {
    const wrap = document.createElement("div");
    wrap.innerHTML = buildRiskSlideHtml(options).trim();
    const section = wrap.firstElementChild;
    if (section) host.appendChild(section);
    return section;
  }

  /** Insert continuation slide directly after `afterSlide` (never at end of host). */
  function insertRiskSlideAfter(afterSlide, nextSlide) {
    if (!afterSlide || !nextSlide) return;
    afterSlide.insertAdjacentElement("afterend", nextSlide);
  }

  function createEmptyRiskContdSlide(escapeHtml, contdTitle) {
    const wrap = document.createElement("div");
    wrap.innerHTML = buildRiskSlideHtml({
      title: contdTitle,
      chunk: [],
      narrativeHtml: "",
      includeNarrative: false,
      isContinuation: true,
      escapeHtml
    }).trim();
    return wrap.firstElementChild;
  }

  function removeEmptyRiskSlides(hostId) {
    getRiskSlideElements(hostId).forEach((slide) => {
      const body = slide.querySelector(".risk-table tbody");
      if (!body?.querySelector("tr")) slide.remove();
    });
  }

  function syncRiskContdAttributes(hostId) {
    getRiskSlideElements(hostId).forEach((slide, idx) => {
      slide.setAttribute("data-risk-contd", idx > 0 ? "true" : "false");
    });
  }

  function prepareRiskSlideTableWrap(slide) {
    const footer = slide.querySelector(".page-footer-bar");
    const wrap = slide.querySelector(".risk-table-wrap");
    const table = slide.querySelector(".risk-table");
    const tbody = table?.querySelector("tbody");
    if (!footer || !wrap || !table || !tbody) return null;

    const available = Math.max(72, footer.offsetTop - wrap.offsetTop - 10);
    wrap.classList.add("risk-table-wrap--fill");
    wrap.style.height = `${available}px`;

    const dataRows = [...tbody.querySelectorAll("tr")].filter(
      (tr) => !tr.querySelector("td[colspan]")
    );
    dataRows.forEach((tr) => {
      tr.style.height = "";
    });
    return { table, dataRows, available };
  }

  function fitRiskSlideTextToBox(slide, table, startFontPx) {
    let testFont = startFontPx;
    while (testFont >= 8 && riskSlideTableOverflows(slide, 6)) {
      table.style.fontSize = `${testFont}px`;
      testFont -= 0.5;
    }
    if (riskSlideTableOverflows(slide, 6)) {
      table.style.fontSize = "8px";
    }
  }

  /** Fit one risk slide table inside footer (equal row heights + font shrink). */
  function fitRiskSlideTable(slide, layoutOptions) {
    const prep = prepareRiskSlideTableWrap(slide);
    if (!prep) return;
    const { tableFontSizePx } = parseLayoutOptions(layoutOptions);
    const { table, dataRows, available } = prep;
    const baseFont =
      tableFontSizePx > 0 ? tableFontSizePx : parseFloat(getComputedStyle(table).fontSize) || 14;

    if (!dataRows.length) {
      table.style.height = "auto";
      table.style.fontSize = `${baseFont}px`;
      return;
    }

    const thead = table.querySelector("thead");
    table.style.height = `${available}px`;
    let fontPx = baseFont;

    const distributeRowHeights = () => {
      const theadH = thead ? thead.offsetHeight : 0;
      const bodyBudget = Math.max(42, available - theadH - 2);
      const rowCount = Math.max(dataRows.length, 1);
      const heights = Array(rowCount).fill(Math.floor(bodyBudget / rowCount));
      let remainder = bodyBudget - heights.reduce((sum, h) => sum + h, 0);
      for (let i = 0; i < remainder; i += 1) {
        heights[i % rowCount] += 1;
      }
      dataRows.forEach((tr, idx) => {
        tr.style.height = `${heights[idx] || heights[0]}px`;
      });
    };

    const applySizing = () => {
      table.style.fontSize = `${fontPx}px`;
      distributeRowHeights();
    };

    applySizing();
    while (fontPx >= 7.5 && riskSlideTableOverflows(slide, 4)) {
      fontPx -= 0.5;
      applySizing();
    }
    if (riskSlideTableOverflows(slide, 4)) {
      table.style.fontSize = "7.5px";
      distributeRowHeights();
    }
  }

  function applyRiskTableVerticalFill(hostId, layoutOptions) {
    getRiskSlideElements(hostId).forEach((slide) => fitRiskSlideTable(slide, layoutOptions));
  }

  function rebalanceRiskSlidesFillSpace(hostId) {
    const host = document.getElementById(hostId || "riskSlidesContainer");
    if (!host) return false;

    const slides = getRiskSlideElements(hostId);
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
          break;
        }
      }
    }

    return changed;
  }

  function rebalanceRiskSlidesTrimOverflow(hostId, escapeHtml, contdTitle) {
    const host = document.getElementById(hostId || "riskSlidesContainer");
    if (!host) return false;

    const slides = getRiskSlideElements(hostId);
    let changed = false;

    for (let i = 0; i < slides.length; i += 1) {
      const slide = slides[i];
      while (riskSlideTableOverflows(slide)) {
        const body = slide.querySelector(".risk-table tbody");
        const lastRow = body?.lastElementChild;
        if (!lastRow) break;

        let nextSlide = slides[i + 1];
        if (!nextSlide) {
          nextSlide = createEmptyRiskContdSlide(escapeHtml, contdTitle);
          insertRiskSlideAfter(slide, nextSlide);
          slides.splice(i + 1, 0, nextSlide);
        }

        nextSlide.querySelector(".risk-table tbody").prepend(lastRow);
        changed = true;
      }
    }

    return changed;
  }

  function renderRiskSlides(config) {
    const {
      hostId = "riskSlidesContainer",
      rows = [],
      baseTitle,
      contdTitle = "Contn.,",
      narrativeText = "",
      escapeHtml,
      formatRichTextHTML,
      layoutOptions = {},
      onComplete
    } = config;

    const host = document.getElementById(hostId);
    if (!host) return;
    host.innerHTML = "";

    const narrativeHtml = formatRichTextHTML(narrativeText);

    chunkRiskRowsForSlides(
      rows,
      narrativeText,
      formatRichTextHTML,
      escapeHtml,
      contdTitle,
      layoutOptions
    ).forEach((chunk, idx) => {
      appendRiskSlide(host, {
        title: idx === 0 ? baseTitle : contdTitle,
        chunk,
        narrativeHtml,
        includeNarrative: idx === 0 && Boolean(String(narrativeText || "").trim()),
        isContinuation: idx > 0,
        escapeHtml
      });
    });

    removeEmptyRiskSlides(hostId);
    applyRiskTableFontSize(hostId, layoutOptions);
    if (typeof onComplete === "function") {
      requestAnimationFrame(() => {
        applyRiskTableFontSize(hostId, layoutOptions);
        ensureRiskSlidesNoOverflow(hostId, escapeHtml, contdTitle, layoutOptions);
        applyRiskTableVerticalFill(hostId, layoutOptions);
        initRiskTableResizeOnHost(host, hostId, layoutOptions, escapeHtml, contdTitle);
        requestAnimationFrame(() => onComplete(host));
      });
    } else {
      requestAnimationFrame(() => {
        applyRiskTableFontSize(hostId, layoutOptions);
        ensureRiskSlidesNoOverflow(hostId, escapeHtml, contdTitle, layoutOptions);
        applyRiskTableVerticalFill(hostId, layoutOptions);
        initRiskTableResizeOnHost(host, hostId, layoutOptions, escapeHtml, contdTitle);
      });
    }
  }

  function initRiskTableResizeOnHost(host, hostId, layoutOptions, escapeHtml, contdTitle) {
    if (typeof global.initRiskTableColumnResize !== "function") return;
    const scope = host || document.getElementById(hostId || "riskSlidesContainer");
    if (!scope) return;
    global.initRiskTableColumnResize(scope, {
      onColumnResizeEnd(table) {
        const slide = table?.closest?.(".risk-slide");
        const { maxRowsPerSlide, skipFitOnColumnResize } = parseLayoutOptions(layoutOptions);
        if (slide && !skipFitOnColumnResize) fitRiskSlideTable(slide, layoutOptions);
        if (maxRowsPerSlide === 0 && !skipFitOnColumnResize) {
          ensureRiskSlidesNoOverflow(hostId, escapeHtml, contdTitle, layoutOptions);
          applyRiskTableVerticalFill(hostId, layoutOptions);
        }
        initRiskTableResizeOnHost(scope, hostId, layoutOptions, escapeHtml, contdTitle);
      }
    });
  }

  function readRiskChunksFromDom(hostId, sourceRows) {
    const slides = getRiskSlideElements(hostId);
    if (!slides.length) return null;
    const chunks = slides
      .map((slide) =>
        [...slide.querySelectorAll(".risk-table tbody tr")]
          .filter((tr) => !tr.querySelector("td[colspan]"))
          .map((tr) => ({
            sno: tr.cells[0]?.textContent?.trim() || "",
            attackType: tr.cells[1]?.textContent?.trim() || "",
            riskScenario: tr.cells[2]?.textContent?.trim() || "",
            ciaTriad: tr.cells[3]?.textContent?.trim() || "",
            businessImpact: tr.cells[4]?.textContent?.trim() || "",
            riskRating: tr.cells[5]?.textContent?.trim() || ""
          }))
      )
      .filter((chunk) => chunk.length > 0);
    if (!chunks.length) return null;

    if (sourceRows?.length) {
      const flat = chunks.flat();
      if (flat.length !== sourceRows.length) return null;
      for (let i = 0; i < sourceRows.length; i += 1) {
        if (String(flat[i].sno) !== String(sourceRows[i].sno)) return null;
      }
    }
    return chunks;
  }

  global.RiskPagination = {
    renderRiskSlides,
    chunkRiskRowsForSlides,
    chunkRiskRowsSequential,
    readRiskChunksFromDom,
    buildRiskSlideHtml,
    buildRiskTableBodyHtml,
    ensureRiskSlidesNoOverflow,
    applyRiskTableVerticalFill,
    fitRiskSlideTable,
    resetRiskTableRowHeights,
    applyRiskTableFontSize,
    parseLayoutOptions
  };
})(typeof window !== "undefined" ? window : globalThis);
