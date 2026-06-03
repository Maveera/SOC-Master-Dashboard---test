/**
 * Drag column borders on risk tables (.risk-table) in the slide preview.
 */
(function (global) {
  "use strict";

  const MIN_COL_PX = 48;
  const SELECTOR = ".risk-table";

  function injectPrintStyle() {
    const id = "risk-table-resize-print";
    if (document.getElementById(id)) return;
    const style = document.createElement("style");
    style.id = id;
    style.textContent =
      "@media print{" +
      SELECTOR +
      " .resizer{display:none!important}}" +
      SELECTOR +
      " thead th{position:relative;overflow:visible}" +
      SELECTOR +
      " .resizer{position:absolute;top:0;right:0;width:8px;height:100%;cursor:col-resize;z-index:4;user-select:none;touch-action:none}" +
      SELECTOR +
      " .resizer:hover{background:rgba(0,102,204,.35)}";
    document.head.appendChild(style);
  }

  function setColumnWidth(table, colIndex, widthPx) {
    const w = `${Math.max(MIN_COL_PX, Math.round(widthPx))}px`;
    const headerRow = table.querySelector("thead tr");
    if (headerRow && headerRow.children[colIndex]) {
      headerRow.children[colIndex].style.width = w;
    }
    table.querySelectorAll("tbody tr").forEach((row) => {
      if (row.children[colIndex]) row.children[colIndex].style.width = w;
    });
  }

  function bindTable(table, options) {
    if (table.dataset.riskResizeBound === "1") return;
    const onColumnResizeEnd = options?.onColumnResizeEnd;
    const headerRow = table.querySelector("thead tr");
    if (!headerRow) return;
    const cols = headerRow.querySelectorAll("th");
    if (cols.length < 2) return;

    table.dataset.riskResizeBound = "1";
    table.style.tableLayout = "fixed";

    for (let i = 0; i < cols.length - 1; i += 1) {
      const th = cols[i];
      if (th.querySelector(".resizer")) continue;

      th.style.position = "relative";

      const resizer = document.createElement("div");
      resizer.className = "resizer";
      resizer.setAttribute("role", "separator");
      resizer.setAttribute("aria-orientation", "vertical");
      resizer.title = "Drag to resize column";
      th.appendChild(resizer);

      let startX = 0;
      let startW = 0;
      let nextStartW = 0;

      const onMouseMove = (e) => {
        const dx = e.clientX - startX;
        const left = Math.max(MIN_COL_PX, startW + dx);
        const right = Math.max(MIN_COL_PX, nextStartW - dx);
        setColumnWidth(table, i, left);
        setColumnWidth(table, i + 1, right);
      };

      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
        document.body.classList.remove("risk-table-col-resizing");
        if (typeof onColumnResizeEnd === "function") onColumnResizeEnd(table);
      };

      const onMouseDown = (e) => {
        e.preventDefault();
        e.stopPropagation();
        startX = e.clientX;
        startW = th.offsetWidth;
        nextStartW = cols[i + 1].offsetWidth;
        document.body.classList.add("risk-table-col-resizing");
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      };

      resizer.addEventListener("mousedown", onMouseDown);
    }
  }

  function initRiskTableColumnResize(root, options) {
    injectPrintStyle();
    const scope = root && root.querySelectorAll ? root : document;
    scope.querySelectorAll(SELECTOR).forEach((table) => bindTable(table, options));
  }

  global.initRiskTableColumnResize = initRiskTableColumnResize;
})(
  typeof window !== "undefined" ? window : globalThis
);
