const TOOL_CONFIG = {
  "SOC Daily report": {
    url: "https://soc-daily-report.vercel.app/",
    mode: "iframe",
    mobileFit: true,
    about: "Generate editable SOC daily reports with EPS metrics, incident tables, severity breakdowns, and branded client documentation for end-of-day SOC handoffs."
  },
  "SOC Incident Analysis": {
    url: "https://soc-incident-analysis.vercel.app/",
    mode: "iframe",
    about: "SOC Intelligence Console for incident rule analysis. Configure AI keys, customer context, and investigation setup to support faster triage and response."
  },
  "SOC Monthly Report": {
    url: "./monthly-report/",
    mode: "iframe",
    mobileFit: true,
    about: "SOC monthly reporting console for Apraava, Nocil, RMZ, and Blue Pine. Edit, modify, and upload monthly SOC report data."
  },
  "SOC Master Consolidation": {
    url: "https://soc-master-consolidation-report.vercel.app/",
    mode: "iframe",
    mobileFit: true,
    about: "Central hub to consolidate SOC metrics, incidents, and reporting outputs into a single master view for leadership and operational review for 6 month of data and you can get the Consolitate report on this."
  },
  "FortiSIEM Full NoteBook": {
    url: "https://forti-siem-book.vercel.app/",
    mode: "iframe",
    about: "FortiSIEM notebook workspace for queries, parser references, and SIEM investigation notes used during active SOC monitoring."
  },
  "SOC MTTR and MTTA Calculator": {
    url: "https://itil-tracker.vercel.app/",
    mode: "iframe",
    about: "IncidentPulse analytics tool to import incident reports and calculate MTTR and MTTA with exportable results for SOC performance tracking."
  },
  "SOC Parser learning": {
    url: "https://fortisieimparser.free.nf/",
    mode: "external",
    message: "This parser requires browser cookies and cannot run inside an embedded panel. Open it in a new tab to use it normally.",
    about: "Interactive parser learning environment for FortiSIEM log formats. Opens in a new browser tab because embedded cookies are required."
  },
  "Viperintel Threat Intel": {
    url: "https://viper-intel.streamlit.app/#threat-map",
    mode: "external",
    message: "Streamlit blocks iframe embedding for security. Open Viperintel in a new tab to access the threat map.",
    about: "Threat intelligence dashboard with global threat mapping and IOC context. Opens externally because Streamlit blocks iframe embedding."
  }
};

const TOOL_INFO_STORAGE_KEY = "soc-dashboard-tool-info-dismissed";

const searchInput = document.getElementById("nav-search");
const navLinks = document.querySelectorAll(".nav-link");
const categories = document.querySelectorAll(".nav-category");
const noResults = document.getElementById("no-results");
const welcomeOverlay = document.getElementById("welcome-overlay");
const activeToolLabel = document.getElementById("active-tool");
const clockEl = document.getElementById("clock");
const workspacePanels = document.getElementById("workspace-panels");
const externalPanel = document.getElementById("external-panel");
const externalToolName = document.getElementById("external-tool-name");
const externalToolMessage = document.getElementById("external-tool-message");
const externalOpenBtn = document.getElementById("external-open-btn");
const appRoot = document.querySelector(".app");
const sidebarToggle = document.getElementById("sidebar-toggle");
const sidebarResizer = document.getElementById("sidebar-resizer");
const sidebarBackdrop = document.getElementById("sidebar-backdrop");
const MOBILE_MQ = window.matchMedia("(max-width: 900px)");
const SIDEBAR_WIDTH_KEY = "soc-dashboard-sidebar-width";
const SIDEBAR_MIN_WIDTH = 220;
const SIDEBAR_MAX_RATIO = 0.5;
const toolInfoModal = document.getElementById("tool-info-modal");
const toolInfoTitle = document.getElementById("tool-info-title");
const toolInfoBody = document.getElementById("tool-info-body");
const toolInfoDismiss = document.getElementById("tool-info-dismiss");
const toolInfoOk = document.getElementById("tool-info-ok");
const toolInfoBackdrop = document.getElementById("tool-info-backdrop");

const iframeByUrl = new Map();
const loadedUrls = new Set();
const MOBILE_FIT_WIDTH = 1024;
let activeTool = null;
let currentModalTool = null;

function isLoadableToolUrl(url) {
  if (!url || url === "#" || url.startsWith("#")) {
    return false;
  }

  try {
    const target = new URL(url, window.location.href);
    const current = new URL(window.location.href);
    const normalizePath = function (path) {
      return path.replace(/\/index\.html$/i, "/").replace(/\/$/, "") || "/";
    };

    if (
      target.origin === current.origin &&
      normalizePath(target.pathname) === normalizePath(current.pathname)
    ) {
      return false;
    }

    return true;
  } catch (error) {
    return false;
  }
}

function isToolOnline(toolName) {
  const config = TOOL_CONFIG[toolName];
  if (!config) {
    return false;
  }
  return isLoadableToolUrl(config.url);
}

function getDismissedToolInfoMap() {
  try {
    return JSON.parse(localStorage.getItem(TOOL_INFO_STORAGE_KEY) || "{}");
  } catch (error) {
    return {};
  }
}

function isToolInfoDismissed(toolName) {
  return !!getDismissedToolInfoMap()[toolName];
}

function dismissToolInfo(toolName) {
  const dismissed = getDismissedToolInfoMap();
  dismissed[toolName] = true;
  localStorage.setItem(TOOL_INFO_STORAGE_KEY, JSON.stringify(dismissed));
}

function openToolInfoModal(toolName) {
  const config = TOOL_CONFIG[toolName];
  if (!config || !config.about || isToolInfoDismissed(toolName)) {
    return;
  }

  currentModalTool = toolName;
  toolInfoTitle.textContent = toolName;
  toolInfoBody.textContent = config.about;
  toolInfoDismiss.checked = false;
  toolInfoModal.hidden = false;
}

function closeToolInfoModal() {
  toolInfoModal.hidden = true;
  currentModalTool = null;
}

function updateToolStatuses() {
  navLinks.forEach(function (link) {
    const toolName = link.dataset.tool;
    const dot = link.querySelector(".status-dot");
    const label = link.querySelector(".status-label");

    if (!dot || !label) {
      return;
    }

    const online = isToolOnline(toolName);
    dot.classList.toggle("online", online);
    dot.classList.toggle("offline", !online);
    label.classList.toggle("online", online);
    label.classList.toggle("offline", !online);
    label.textContent = online ? "Online" : "Offline";
    dot.title = online ? "Online" : "Offline";
    dot.setAttribute("aria-label", "Status: " + (online ? "online" : "offline"));
  });
}

function getPlaceholderDocument(toolName) {
  return [
    "<!DOCTYPE html>",
    "<html lang=\"en\">",
    "<head>",
    "<meta charset=\"UTF-8\" />",
    "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\" />",
    "<style>",
    "body { margin: 0; min-height: 100vh; display: grid; place-items: center; background: #0d1117; color: #8b949e; font-family: Inter, system-ui, sans-serif; }",
    ".panel { text-align: center; max-width: 420px; padding: 2rem; border: 1px solid rgba(255,74,74,0.2); border-radius: 12px; background: #11141a; }",
    "h1 { margin: 0 0 0.75rem; font-size: 1rem; color: #e6edf3; }",
    "p { margin: 0; line-height: 1.6; font-size: 0.9rem; }",
    "</style>",
    "</head>",
    "<body>",
    "<div class=\"panel\">",
    "<h1>" + toolName + "</h1>",
    "<p>This tool is currently offline and unavailable in the workspace.</p>",
    "</div>",
    "</body>",
    "</html>"
  ].join("");
}

function toolNeedsMobileFit(toolName) {
  const config = TOOL_CONFIG[toolName];
  return !!(config && config.mobileFit);
}

function hideAllPanels() {
  workspacePanels.querySelectorAll(".iframe-panel").forEach(function (panel) {
    panel.classList.remove("active");
    panel.hidden = true;
  });
  externalPanel.classList.remove("active");
  externalPanel.hidden = true;
}

function showPanel(panel) {
  hideAllPanels();
  panel.hidden = false;
  panel.classList.add("active");
  scheduleMobileIframeFit();
}

function applyMobileIframeFit(panel) {
  if (!panel || !panel.classList.contains("iframe-panel--mobile-fit")) {
    return;
  }

  const iframe = panel.querySelector(".workspace-iframe");
  if (!iframe) {
    return;
  }

  if (!isMobileLayout()) {
    panel.classList.remove("is-mobile-scaled");
    iframe.removeAttribute("style");
    return;
  }

  const containerWidth = panel.clientWidth;
  const containerHeight = panel.clientHeight;
  if (!containerWidth || !containerHeight) {
    return;
  }

  const scale = containerWidth / MOBILE_FIT_WIDTH;
  panel.classList.add("is-mobile-scaled");
  iframe.style.width = MOBILE_FIT_WIDTH + "px";
  iframe.style.height = Math.ceil(containerHeight / scale) + "px";
  iframe.style.transform = "scale(" + scale + ")";
  iframe.style.transformOrigin = "top left";
}

function applyAllMobileIframeFits() {
  if (!isMobileLayout()) {
    workspacePanels.querySelectorAll(".iframe-panel--mobile-fit").forEach(function (panel) {
      panel.classList.remove("is-mobile-scaled");
      const iframe = panel.querySelector(".workspace-iframe");
      if (iframe) {
        iframe.removeAttribute("style");
      }
    });
    return;
  }

  workspacePanels.querySelectorAll(".iframe-panel--mobile-fit.active").forEach(applyMobileIframeFit);
}

function scheduleMobileIframeFit() {
  requestAnimationFrame(function () {
    applyAllMobileIframeFits();
    requestAnimationFrame(applyAllMobileIframeFits);
  });
}

function getIframeForTool(toolName, url) {
  if (iframeByUrl.has(url)) {
    return iframeByUrl.get(url);
  }

  const panel = document.createElement("div");
  panel.className = "iframe-panel";
  panel.hidden = true;

  const iframe = document.createElement("iframe");
  iframe.className = "workspace-iframe";
  iframe.title = "SOC tool workspace";
  iframe.src = "about:blank";

  if (toolNeedsMobileFit(toolName)) {
    panel.classList.add("iframe-panel--mobile-fit");
  }

  panel.appendChild(iframe);
  workspacePanels.appendChild(panel);
  iframeByUrl.set(url, { panel: panel, iframe: iframe });
  return iframeByUrl.get(url);
}

function showExternalTool(toolName, config) {
  externalToolName.textContent = toolName;
  externalToolMessage.textContent = config.message || "This tool opens in a new browser tab.";
  externalOpenBtn.href = config.url;
  showPanel(externalPanel);
}

function showIframeTool(toolName, url) {
  if (!isLoadableToolUrl(url)) {
    const placeholderKey = "placeholder:" + toolName;
    const entry = getIframeForTool(toolName, placeholderKey);
    if (!loadedUrls.has(placeholderKey)) {
      entry.iframe.removeAttribute("src");
      entry.iframe.srcdoc = getPlaceholderDocument(toolName);
      loadedUrls.add(placeholderKey);
    }
    showPanel(entry.panel);
    return;
  }

  const entry = getIframeForTool(toolName, url);
  if (!loadedUrls.has(url)) {
    entry.iframe.removeAttribute("srcdoc");
    entry.iframe.src = url;
    loadedUrls.add(url);
  }
  showPanel(entry.panel);
}

function isMobileLayout() {
  return MOBILE_MQ.matches;
}

function clampSidebarWidth(widthPx) {
  const max = Math.floor(window.innerWidth * SIDEBAR_MAX_RATIO);
  return Math.min(max, Math.max(SIDEBAR_MIN_WIDTH, widthPx));
}

function applySidebarWidth(widthPx) {
  const clamped = clampSidebarWidth(widthPx);
  document.documentElement.style.setProperty("--sidebar-width", clamped + "px");
  localStorage.setItem(SIDEBAR_WIDTH_KEY, String(clamped));
}

function loadSavedSidebarWidth() {
  const saved = localStorage.getItem(SIDEBAR_WIDTH_KEY);
  if (saved && !isMobileLayout()) {
    applySidebarWidth(Number(saved));
  }
}

function updateSidebarResizerVisibility() {
  if (!sidebarResizer) {
    return;
  }
  const hide = isMobileLayout() || document.body.classList.contains("sidebar-collapsed");
  sidebarResizer.hidden = hide;
  sidebarResizer.classList.toggle("is-hidden", hide);
}

function initSidebarResize() {
  if (!sidebarResizer) {
    return;
  }

  loadSavedSidebarWidth();
  updateSidebarResizerVisibility();

  let dragging = false;

  function onPointerDown(event) {
    if (isMobileLayout() || document.body.classList.contains("sidebar-collapsed")) {
      return;
    }
    dragging = true;
    document.body.classList.add("sidebar-resizing");
    event.preventDefault();
  }

  function onPointerMove(event) {
    if (!dragging) {
      return;
    }
    const clientX = event.touches ? event.touches[0].clientX : event.clientX;
    applySidebarWidth(clientX);
    event.preventDefault();
  }

  function onPointerUp() {
    if (!dragging) {
      return;
    }
    dragging = false;
    document.body.classList.remove("sidebar-resizing");
  }

  sidebarResizer.addEventListener("mousedown", onPointerDown);
  sidebarResizer.addEventListener("touchstart", onPointerDown, { passive: false });
  window.addEventListener("mousemove", onPointerMove);
  window.addEventListener("touchmove", onPointerMove, { passive: false });
  window.addEventListener("mouseup", onPointerUp);
  window.addEventListener("touchend", onPointerUp);
}

function setSidebarOpen(open) {
  if (isMobileLayout()) {
    appRoot.classList.toggle("sidebar-open", open);
    appRoot.classList.remove("sidebar-collapsed");
    document.body.classList.toggle("sidebar-nav-open", open);
    document.body.classList.remove("sidebar-collapsed");
  } else {
    appRoot.classList.toggle("sidebar-collapsed", !open);
    appRoot.classList.remove("sidebar-open");
    document.body.classList.toggle("sidebar-collapsed", !open);
    document.body.classList.remove("sidebar-nav-open");
  }

  sidebarToggle.setAttribute("aria-expanded", open ? "true" : "false");
  sidebarToggle.setAttribute("aria-label", open ? "Hide navigation" : "Show navigation");

  if (sidebarBackdrop) {
    const showBackdrop = open && isMobileLayout();
    sidebarBackdrop.classList.toggle("visible", showBackdrop);
    sidebarBackdrop.hidden = !showBackdrop;
    sidebarBackdrop.setAttribute("aria-hidden", showBackdrop ? "false" : "true");
  }

  scheduleMobileIframeFit();
  updateSidebarResizerVisibility();
}

function closeSidebarOnMobile() {
  if (isMobileLayout()) {
    setSidebarOpen(false);
  }
}

function isSidebarOpen() {
  if (isMobileLayout()) {
    return document.body.classList.contains("sidebar-nav-open");
  }
  return !document.body.classList.contains("sidebar-collapsed");
}

function syncSidebarForViewport() {
  if (isMobileLayout()) {
    appRoot.classList.remove("sidebar-open", "sidebar-collapsed");
    document.body.classList.remove("sidebar-nav-open", "sidebar-collapsed");
    if (sidebarBackdrop) {
      sidebarBackdrop.classList.remove("visible");
      sidebarBackdrop.hidden = true;
      sidebarBackdrop.setAttribute("aria-hidden", "true");
    }
    sidebarToggle.setAttribute("aria-expanded", "false");
    sidebarToggle.setAttribute("aria-label", "Show navigation");
  } else {
    appRoot.classList.remove("sidebar-open");
    appRoot.classList.add("sidebar-collapsed");
    document.body.classList.remove("sidebar-nav-open");
    document.body.classList.add("sidebar-collapsed");
    if (sidebarBackdrop) {
      sidebarBackdrop.classList.remove("visible");
      sidebarBackdrop.hidden = true;
      sidebarBackdrop.setAttribute("aria-hidden", "true");
    }
    sidebarToggle.setAttribute("aria-expanded", "false");
    sidebarToggle.setAttribute("aria-label", "Show navigation");
  }
  updateSidebarResizerVisibility();
}

function loadTool(toolName) {
  if (activeTool === toolName) {
    return;
  }

  activeTool = toolName;
  welcomeOverlay.classList.add("hidden");
  activeToolLabel.textContent = toolName;

  const config = TOOL_CONFIG[toolName];
  if (!config) {
    hideAllPanels();
    return;
  }

  if (config.mode === "external" && isLoadableToolUrl(config.url)) {
    showExternalTool(toolName, config);
  } else {
    showIframeTool(toolName, config.url);
  }

  openToolInfoModal(toolName);
}

navLinks.forEach(function (link) {
  const toolName = link.dataset.tool;

  link.addEventListener("click", function (event) {
    event.preventDefault();

    if (link.classList.contains("active")) {
      return;
    }

    navLinks.forEach(function (l) {
      l.classList.remove("active");
    });
    link.classList.add("active");

    if (toolName) {
      loadTool(toolName);
      closeSidebarOnMobile();
    }
  });
});

toolInfoOk.addEventListener("click", function () {
  if (toolInfoDismiss.checked && currentModalTool) {
    dismissToolInfo(currentModalTool);
  }
  closeToolInfoModal();
});

toolInfoBackdrop.addEventListener("click", function () {
  closeToolInfoModal();
});

document.addEventListener("keydown", function (event) {
  if (event.key === "Escape") {
    if (!toolInfoModal.hidden) {
      closeToolInfoModal();
      return;
    }
    if (isMobileLayout() && isSidebarOpen()) {
      closeSidebarOnMobile();
    }
  }
});

searchInput.addEventListener("input", function () {
  const query = searchInput.value.trim().toLowerCase();
  let visibleCount = 0;

  categories.forEach(function (category) {
    let categoryVisible = 0;
    category.querySelectorAll(".nav-link").forEach(function (link) {
      const label = link.querySelector(".nav-link-label").textContent.toLowerCase();
      const match = !query || label.includes(query);
      link.classList.toggle("hidden", !match);
      if (match) categoryVisible++;
    });
    category.classList.toggle("hidden", categoryVisible === 0);
    visibleCount += categoryVisible;
  });

  noResults.classList.toggle("visible", visibleCount === 0 && query.length > 0);
});

sidebarToggle.addEventListener("click", function () {
  setSidebarOpen(!isSidebarOpen());
});

if (sidebarBackdrop) {
  sidebarBackdrop.addEventListener("click", function () {
    closeSidebarOnMobile();
  });
}

MOBILE_MQ.addEventListener("change", function () {
  syncSidebarForViewport();
  scheduleMobileIframeFit();
  updateSidebarResizerVisibility();
});

window.addEventListener("resize", scheduleMobileIframeFit);

function initInspectProtection() {
  document.addEventListener("contextmenu", function (event) {
    event.preventDefault();
  });

  document.addEventListener("keydown", function (event) {
    const key = event.key ? event.key.toLowerCase() : "";
    const blockedShortcut =
      key === "f12" ||
      (event.ctrlKey && event.shiftKey && (key === "i" || key === "j" || key === "c")) ||
      (event.ctrlKey && key === "u") ||
      (event.metaKey && event.altKey && key === "i");

    if (blockedShortcut) {
      event.preventDefault();
      event.stopPropagation();
    }
  });
}

function updateClock() {
  const now = new Date();
  clockEl.textContent = now.toLocaleString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
}

syncSidebarForViewport();
initSidebarResize();
initInspectProtection();
scheduleMobileIframeFit();
updateToolStatuses();
updateClock();
setInterval(updateClock, 1000);
