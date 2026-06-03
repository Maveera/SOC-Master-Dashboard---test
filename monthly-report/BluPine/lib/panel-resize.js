(function () {
  const STORAGE_KEY = "sns-monthly-report-panel-width";
  const MIN_WIDTH = 240;
  const MAX_RATIO = 0.55;
  const MOBILE_MQ = window.matchMedia("(max-width: 900px)");

  function isEmbeddedView() {
    try {
      return window.self !== window.top;
    } catch (error) {
      return true;
    }
  }

  function isSidebarLayout() {
    const shell = document.querySelector(".app-shell");
    return !!(shell && shell.querySelector(".sidebar") && !shell.querySelector(":scope > .panel, :scope > aside.panel"));
  }

  function isNavPanelCollapsed() {
    return document.body.classList.contains("nav-panel-collapsed");
  }

  function isNavHidden() {
    if (isNavPanelCollapsed()) {
      return true;
    }

    if (isSidebarLayout()) {
      const shell = document.getElementById("appShell") || document.querySelector(".app-shell");
      return !!(shell && shell.classList.contains("sidebar-collapsed"));
    }

    return false;
  }

  function syncNavHiddenState() {
    document.body.classList.toggle("nav-shell-hidden", !isMobileLayout() && isNavHidden());
    updateFloatingExportBarVisibility();
  }

  function isMobileLayout() {
    return MOBILE_MQ.matches;
  }

  function markEmbeddedView() {
    if (!isEmbeddedView() || isMobileLayout()) {
      return;
    }

    document.documentElement.classList.add("is-embedded");

    const shell = document.querySelector(".app-shell");
    if (shell && shell.querySelector(".sidebar") && !shell.querySelector(".panel")) {
      shell.classList.add("sidebar-collapsed");
    }
  }

  function setNavPanelCollapsed(collapsed) {
    document.body.classList.toggle("nav-panel-collapsed", collapsed);

    const toggleBtn = document.getElementById("panelToggleBtn");
    if (toggleBtn) {
      toggleBtn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      toggleBtn.setAttribute(
        "aria-label",
        collapsed ? "Show navigation panel" : "Hide navigation panel"
      );
    }

    const resizer = document.querySelector(".panel-resizer");
    if (resizer) {
      updateResizerVisibility(resizer);
    }

    syncNavHiddenState();
  }

  function syncFloatingExportButtons() {
    const printBtn = document.getElementById("printBtn");
    const pptxBtn = document.getElementById("pptxBtn");
    const floatingPrintBtn = document.getElementById("floatingPrintBtn");
    const floatingPptxBtn = document.getElementById("floatingPptxBtn");
    if (!printBtn || !pptxBtn || !floatingPrintBtn || !floatingPptxBtn) {
      return;
    }

    floatingPrintBtn.disabled = printBtn.disabled;
    floatingPrintBtn.textContent = printBtn.textContent;
    floatingPptxBtn.disabled = pptxBtn.disabled;
    floatingPptxBtn.textContent = pptxBtn.textContent;
  }

  function updateFloatingExportBarVisibility() {
    const bar = document.getElementById("floatingExportBar");
    if (!bar) {
      return;
    }

    const show = !isMobileLayout() && isNavHidden();
    bar.hidden = !show;
    if (show) {
      syncFloatingExportButtons();
    }
  }

  function initFloatingExportBar() {
    if (document.getElementById("floatingExportBar")) {
      return;
    }

    const printBtn = document.getElementById("printBtn");
    const pptxBtn = document.getElementById("pptxBtn");
    if (!printBtn || !pptxBtn) {
      return;
    }

    const bar = document.createElement("div");
    bar.id = "floatingExportBar";
    bar.className = "floating-export-bar";
    bar.hidden = true;
    bar.innerHTML =
      '<button type="button" id="floatingPrintBtn" class="floating-export-btn">Export PDF</button>' +
      '<button type="button" id="floatingPptxBtn" class="floating-export-btn">Export PPTX</button>';
    document.body.appendChild(bar);

    const floatingPrintBtn = document.getElementById("floatingPrintBtn");
    const floatingPptxBtn = document.getElementById("floatingPptxBtn");

    floatingPrintBtn.addEventListener("click", function () {
      printBtn.click();
    });
    floatingPptxBtn.addEventListener("click", function () {
      pptxBtn.click();
    });

    const observer = new MutationObserver(syncFloatingExportButtons);
    observer.observe(printBtn, { attributes: true, childList: true, subtree: true, characterData: true });
    observer.observe(pptxBtn, { attributes: true, childList: true, subtree: true, characterData: true });

    updateFloatingExportBarVisibility();
  }

  function syncNavPanelForViewport() {
    if (isMobileLayout()) {
      document.body.classList.remove("nav-panel-collapsed");
      document.body.classList.remove("panel-open");
      document.body.classList.remove("nav-shell-hidden");
      return;
    }

    const panelToggleBtn = document.getElementById("panelToggleBtn");
    if (panelToggleBtn && panelToggleBtn.dataset.navToggleBound !== "1") {
      setNavPanelCollapsed(true);
    }

    initSidebarLayout();
  }

  function initSidebarLayout() {
    if (isMobileLayout() || !isSidebarLayout()) {
      return;
    }

    const shell = document.getElementById("appShell") || document.querySelector(".app-shell");
    if (!shell) {
      return;
    }

    if (!shell.dataset.sidebarDesktopInit) {
      shell.classList.add("sidebar-collapsed");
      shell.dataset.sidebarDesktopInit = "1";
    }

    if (!shell.dataset.sidebarObserverBound) {
      shell.dataset.sidebarObserverBound = "1";
      new MutationObserver(syncNavHiddenState).observe(shell, {
        attributes: true,
        attributeFilter: ["class"]
      });
    }

    syncNavHiddenState();
  }

  function initPanelToggle() {
    if (isMobileLayout()) {
      return;
    }

    const toggleBtn = document.getElementById("panelToggleBtn");
    const shell = document.querySelector(".app-shell");
    const panel = shell ? shell.querySelector(".panel") : null;
    if (!toggleBtn || !panel) {
      return;
    }

    if (toggleBtn.dataset.navToggleBound === "1") {
      return;
    }
    toggleBtn.dataset.navToggleBound = "1";

    setNavPanelCollapsed(true);

    toggleBtn.addEventListener("click", function () {
      setNavPanelCollapsed(!isNavPanelCollapsed());
    });
  }

  function getPanelElement(shell) {
    return shell.querySelector(":scope > .panel, :scope > .sidebar, :scope > aside.panel, :scope > aside.sidebar");
  }

  function clampWidth(widthPx) {
    const max = Math.floor(window.innerWidth * MAX_RATIO);
    return Math.min(max, Math.max(MIN_WIDTH, widthPx));
  }

  function applyWidth(widthPx) {
    const clamped = clampWidth(widthPx);
    document.documentElement.style.setProperty("--form-panel-width", clamped + "px");
    localStorage.setItem(STORAGE_KEY, String(clamped));
  }

  function loadSavedWidth() {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved && !isMobileLayout()) {
      applyWidth(Number(saved));
    }
  }

  function ensureResizer(shell, panel) {
    let resizer = shell.querySelector(":scope > .panel-resizer");
    if (!resizer) {
      resizer = document.createElement("div");
      resizer.className = "panel-resizer";
      resizer.setAttribute("role", "separator");
      resizer.setAttribute("aria-orientation", "vertical");
      resizer.setAttribute("aria-label", "Resize form panel");
      resizer.title = "Drag to resize";
      panel.insertAdjacentElement("afterend", resizer);
    }
    return resizer;
  }

  function updateResizerVisibility(resizer) {
    const hide = isMobileLayout() || isNavPanelCollapsed();
    resizer.hidden = hide;
    resizer.classList.toggle("is-hidden", hide);
  }

  function bindResize(resizer, panel) {
    let dragging = false;

    function onPointerDown(event) {
      if (isMobileLayout() || isNavPanelCollapsed()) {
        return;
      }
      dragging = true;
      document.body.classList.add("panel-resizing");
      event.preventDefault();
    }

    function onPointerMove(event) {
      if (!dragging) {
        return;
      }
      const clientX = event.touches ? event.touches[0].clientX : event.clientX;
      applyWidth(clientX);
      event.preventDefault();
    }

    function onPointerUp() {
      if (!dragging) {
        return;
      }
      dragging = false;
      document.body.classList.remove("panel-resizing");
    }

    resizer.addEventListener("mousedown", onPointerDown);
    resizer.addEventListener("touchstart", onPointerDown, { passive: false });
    window.addEventListener("mousemove", onPointerMove);
    window.addEventListener("touchmove", onPointerMove, { passive: false });
    window.addEventListener("mouseup", onPointerUp);
    window.addEventListener("touchend", onPointerUp);
  }

  function initPanelResize() {
    const shell = document.querySelector(".app-shell");
    const panel = shell ? getPanelElement(shell) : null;
    if (!shell || !panel) {
      return;
    }

    loadSavedWidth();
    const resizer = ensureResizer(shell, panel);
    updateResizerVisibility(resizer);
    bindResize(resizer, panel);
  }

  function onViewportChange() {
    syncNavPanelForViewport();
    initPanelResize();
    initSidebarLayout();
    syncNavHiddenState();
    if (!isMobileLayout()) {
      initPanelToggle();
    }
  }

  MOBILE_MQ.addEventListener("change", onViewportChange);

  function boot() {
    syncNavPanelForViewport();
    markEmbeddedView();
    initPanelResize();
    initFloatingExportBar();
    initSidebarLayout();
    initPanelToggle();
    syncNavHiddenState();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();
