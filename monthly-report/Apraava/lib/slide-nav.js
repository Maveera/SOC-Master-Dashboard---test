/**
 * Slide navigation for Monthly SOC Report — order matches TEMPLATE_REPORT_SECTIONS (PPT export order).
 * Click a nav item to scroll the preview to that slide.
 */
(function (global) {
  const LABELS = global.SLIDE_NAV_LABELS || {};

  function getVisibleSlides() {
    return Array.from(document.querySelectorAll("#reportRoot [data-slide]")).filter((el) => {
      const style = el.style.display;
      return style !== "none" && !el.hidden;
    });
  }

  function scrollToSlide(slideId) {
    const preview = document.querySelector(".preview-wrap");
    const slides = getVisibleSlides().filter((el) => el.getAttribute("data-slide") === slideId);
    const slide = slides[0];
    if (!slide || !preview) return;

    const previewRect = preview.getBoundingClientRect();
    const slideRect = slide.getBoundingClientRect();
    const top = slideRect.top - previewRect.top + preview.scrollTop - 12;
    preview.scrollTo({ top: Math.max(0, top), behavior: "smooth" });

    setActiveNavItem(slideId);
  }

  function scrollToFormSection(slideId) {
    const section =
      document.querySelector(`.panel [data-form-slide="${slideId}"]`) ||
      document.querySelector(`.sidebar [data-form-slide="${slideId}"]`) ||
      document.querySelector(`[data-form-slide="${slideId}"]`);
    if (!section || section.offsetParent === null) return;
    section.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function setActiveNavItem(slideId) {
    document.querySelectorAll(".slide-nav-item").forEach((btn) => {
      btn.classList.toggle("is-active", btn.dataset.slide === slideId);
    });
  }

  function buildSlideNav(templateKey) {
    const nav = document.getElementById("slideNav");
    const list = document.getElementById("slideNavList") || nav;
    if (!nav) return;

    const sections =
      typeof global.TEMPLATE_REPORT_SECTIONS !== "undefined"
        ? global.TEMPLATE_REPORT_SECTIONS[templateKey]
        : null;

    if (!sections || !sections.length) {
      if (list !== nav) list.innerHTML = "";
      else nav.innerHTML = "";
      nav.hidden = true;
      return;
    }

    nav.hidden = false;
    const itemsHtml = sections
      .map((slideId, index) => {
        const label = LABELS[slideId] || slideId.replace(/-/g, " ");
        return `<button type="button" class="slide-nav-item" data-slide="${slideId}" title="Go to slide ${index + 1}">
          <span class="slide-nav-num">${index + 1}</span>
          <span class="slide-nav-label">${label}</span>
        </button>`;
      })
      .join("");

    if (list !== nav) {
      list.innerHTML = itemsHtml;
    } else {
      nav.innerHTML = itemsHtml;
    }

    (list !== nav ? list : nav).querySelectorAll(".slide-nav-item").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.dataset.slide;
        scrollToSlide(id);
        scrollToFormSection(id);
      });
    });
  }

  function bindPreviewScrollSync() {
    const preview = document.querySelector(".preview-wrap");
    if (!preview || preview.dataset.slideNavBound === "1") return;
    preview.dataset.slideNavBound = "1";

    let ticking = false;
    preview.addEventListener(
      "scroll",
      () => {
        if (ticking) return;
        ticking = true;
        requestAnimationFrame(() => {
          ticking = false;
          const slides = getVisibleSlides();
          if (!slides.length) return;

          const anchor = preview.scrollTop + 48;
          let active = slides[0].getAttribute("data-slide");
          const previewRect = preview.getBoundingClientRect();
          for (const slide of slides) {
            const slideTop = slide.getBoundingClientRect().top - previewRect.top + preview.scrollTop;
            if (slideTop <= anchor) {
              active = slide.getAttribute("data-slide");
            }
          }
          setActiveNavItem(active);
        });
      },
      { passive: true }
    );
  }

  function initSlideNav(getTemplateKeyFn) {
    const resolveKey =
      typeof getTemplateKeyFn === "function"
        ? getTemplateKeyFn
        : () => {
            const el = document.getElementById("templatePreset");
            return el && el.value ? el.value : "apraava";
          };

    const refresh = () => {
      buildSlideNav(resolveKey());
      bindPreviewScrollSync();
    };

    refresh();
    return { refresh, scrollToSlide };
  }

  global.initSlideNav = initSlideNav;
  global.scrollToReportSlide = scrollToSlide;
})(window);
