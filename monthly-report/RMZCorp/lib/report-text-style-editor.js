/**
 * Preview text style editor — click slide text to adjust font size & color.
 * Edits are preview-only and do not affect PPTX/PDF/DOCX export.
 */
(function (global) {
  "use strict";

  const textStyleOverrides = new Map();
  let textEditMode = false;
  let tseSelectedEl = null;
  let tseSelectedKey = null;
  let tseToolbar = null;
  let tseStyleIdCounter = 0;

  function tseGetKey(el) {
    if (el.id) return el.id;
    if (el.dataset.styleId) return el.dataset.styleId;
    const key = `tse-${Date.now().toString(36)}-${tseStyleIdCounter++}`;
    el.dataset.styleId = key;
    return key;
  }

  function tseFindByKey(key) {
    const escape = global.CSS && CSS.escape ? CSS.escape : (s) => s;
    return document.getElementById(key) || document.querySelector(`[data-style-id="${escape(key)}"]`);
  }

  function tseRgbToHex(color) {
    const m = color && color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i);
    if (!m) return "#000000";
    const toHex = (n) => Number(n).toString(16).padStart(2, "0");
    return `#${toHex(m[1])}${toHex(m[2])}${toHex(m[3])}`;
  }

  function tseIsTextElement(t) {
    if (!t || t.nodeType !== 1) return null;
    if (["CANVAS", "IMG", "SVG", "PATH", "INPUT", "TEXTAREA", "BUTTON", "SELECT"].includes(t.tagName)) return null;
    if (t.closest && t.closest("svg")) return null;
    if (!t.textContent || !t.textContent.trim()) return null;
    return t;
  }

  function tseStoreOverride() {
    if (!tseSelectedEl || !tseSelectedKey) return;
    const cs = getComputedStyle(tseSelectedEl);
    textStyleOverrides.set(tseSelectedKey, {
      fontSize: tseSelectedEl.style.fontSize || cs.fontSize,
      color: tseSelectedEl.style.color || cs.color
    });
  }

  function tseBuildToolbar() {
    if (tseToolbar) return tseToolbar;
    const bar = document.createElement("div");
    bar.className = "tse-toolbar";
    bar.hidden = true;
    bar.innerHTML = `
      <label class="tse-field">Size (px)
        <input type="number" min="1" max="400" step="1" class="tse-font-size" />
      </label>
      <label class="tse-field">Color
        <input type="color" class="tse-color" />
      </label>
      <button type="button" class="tse-close" title="Close">&times;</button>
    `;
    document.body.appendChild(bar);

    const sizeInput = bar.querySelector(".tse-font-size");
    const colorInput = bar.querySelector(".tse-color");

    const onSize = () => {
      if (!tseSelectedEl) return;
      const v = parseFloat(sizeInput.value);
      if (!Number.isFinite(v) || v <= 0) return;
      tseSelectedEl.style.fontSize = `${v}px`;
      tseStoreOverride();
    };
    const onColor = () => {
      if (!tseSelectedEl) return;
      tseSelectedEl.style.color = colorInput.value;
      tseStoreOverride();
    };

    sizeInput.addEventListener("input", onSize);
    sizeInput.addEventListener("change", onSize);
    colorInput.addEventListener("input", onColor);
    colorInput.addEventListener("change", onColor);
    bar.querySelector(".tse-close").addEventListener("click", () => tseDeselect());

    tseToolbar = bar;
    return bar;
  }

  function tsePositionToolbar() {
    if (!tseToolbar || !tseSelectedEl) return;
    const r = tseSelectedEl.getBoundingClientRect();
    const bar = tseToolbar;
    bar.hidden = false;
    const bw = bar.offsetWidth || 220;
    const bh = bar.offsetHeight || 46;
    let top = r.top - bh - 8;
    if (top < 8) top = r.bottom + 8;
    let left = r.left;
    left = Math.max(8, Math.min(left, global.innerWidth - bw - 8));
    top = Math.max(8, Math.min(top, global.innerHeight - bh - 8));
    bar.style.top = `${top}px`;
    bar.style.left = `${left}px`;
  }

  function tseSelect(el) {
    tseBuildToolbar();
    if (tseSelectedEl && tseSelectedEl !== el) tseSelectedEl.classList.remove("text-edit-selected");
    tseSelectedEl = el;
    tseSelectedKey = tseGetKey(el);
    el.classList.add("text-edit-selected");

    const cs = getComputedStyle(el);
    tseToolbar.querySelector(".tse-font-size").value = Math.round(parseFloat(cs.fontSize) || 14);
    tseToolbar.querySelector(".tse-color").value = tseRgbToHex(cs.color);
    tsePositionToolbar();
  }

  function tseDeselect() {
    if (tseSelectedEl) tseSelectedEl.classList.remove("text-edit-selected");
    tseSelectedEl = null;
    tseSelectedKey = null;
    if (tseToolbar) tseToolbar.hidden = true;
  }

  function reapplyTextStyleOverrides() {
    textStyleOverrides.forEach((style, key) => {
      const el = tseFindByKey(key);
      if (!el) return;
      if (style.fontSize) el.style.fontSize = style.fontSize;
      if (style.color) el.style.color = style.color;
    });

    if (textEditMode && tseSelectedKey) {
      const el = tseFindByKey(tseSelectedKey);
      if (el) {
        tseSelectedEl = el;
        el.classList.add("text-edit-selected");
        tsePositionToolbar();
      } else {
        tseDeselect();
      }
    }
  }

  function tseOnDocumentClick(e) {
    if (!textEditMode) return;
    if (tseToolbar && tseToolbar.contains(e.target)) return;

    const root = document.getElementById("reportRoot");
    if (root && root.contains(e.target)) {
      const el = tseIsTextElement(e.target);
      if (el) {
        e.preventDefault();
        e.stopPropagation();
        tseSelect(el);
      } else {
        tseDeselect();
      }
      return;
    }
    tseDeselect();
  }

  function tseSetMode(on) {
    textEditMode = !!on;
    const root = document.getElementById("reportRoot");
    if (root) root.classList.toggle("text-edit-active", textEditMode);
    const hint = document.querySelector(".tse-hint");
    if (hint) hint.hidden = !textEditMode;
    if (!textEditMode) tseDeselect();
  }

  function initTextStyleEditor() {
    const toggle = document.getElementById("textEditToggle");
    if (toggle) {
      toggle.checked = false;
      toggle.addEventListener("change", () => tseSetMode(toggle.checked));
    }
    document.addEventListener("click", tseOnDocumentClick, true);
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && textEditMode) tseDeselect();
    });
    global.addEventListener("resize", () => { if (tseSelectedEl) tsePositionToolbar(); });
    const preview = document.querySelector(".preview-wrap");
    if (preview) {
      preview.addEventListener("scroll", () => { if (tseSelectedEl) tsePositionToolbar(); }, { passive: true });
    }
  }

  global.initTextStyleEditor = initTextStyleEditor;
  global.reapplyTextStyleOverrides = reapplyTextStyleOverrides;
})(typeof window !== "undefined" ? window : globalThis);
