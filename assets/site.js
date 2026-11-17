/* Scroll behaviour for processchem.no: header state, progress bar, reveal-on-scroll,
   gentle parallax, mobile menu. Everything degrades to a normal static page without JS
   or when the visitor asks for reduced motion. */
(function () {
  var root = document.documentElement;
  root.classList.add("js");
  var reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  var head = document.querySelector(".site-head");
  var bar = document.querySelector(".progress");
  var drift = [].slice.call(document.querySelectorAll("[data-drift]"));
  var ticking = false;

  function frame() {
    ticking = false;
    var y = window.pageYOffset || 0;
    var max = Math.max(1, document.documentElement.scrollHeight - window.innerHeight);
    if (head) head.classList.toggle("scrolled", y > 24);
    if (bar) bar.style.transform = "scaleX(" + Math.min(1, y / max).toFixed(4) + ")";
    if (!reduce) {
      for (var i = 0; i < drift.length; i++) {
        var el = drift[i];
        var rect = el.parentNode.getBoundingClientRect();
        if (rect.bottom < -200 || rect.top > window.innerHeight + 200) continue;
        var f = parseFloat(el.getAttribute("data-drift")) || 0;
        var centre = rect.top + rect.height / 2 - window.innerHeight / 2;
        el.style.transform = "translate3d(0," + (centre * f).toFixed(1) + "px,0)";
      }
    }
  }
  function request() { if (!ticking) { ticking = true; window.requestAnimationFrame(frame); } }
  window.addEventListener("scroll", request, { passive: true });
  window.addEventListener("resize", request);
  frame();

  // Reveal on scroll. Children of a [data-stagger] container arrive one after another.
  var items = [].slice.call(document.querySelectorAll(".reveal"));
  [].forEach.call(document.querySelectorAll("[data-stagger]"), function (group) {
    [].forEach.call(group.querySelectorAll(".reveal"), function (el, i) {
      el.style.transitionDelay = (i * 90) + "ms";
    });
  });
  if (reduce || !("IntersectionObserver" in window)) {
    items.forEach(function (el) { el.classList.add("in"); });
  } else {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) { en.target.classList.add("in"); io.unobserve(en.target); }
      });
    }, { threshold: 0.14, rootMargin: "0px 0px -6% 0px" });
    items.forEach(function (el) { io.observe(el); });
  }

  // The header's height, for things that stick just below it (the docs Contents bar).
  if (head) {
    var setHead = function () { root.style.setProperty("--head-h", head.offsetHeight + "px"); };
    setHead();
    if ("ResizeObserver" in window) new ResizeObserver(setHead).observe(head);
    else window.addEventListener("resize", setHead);
  }

  // Fold-out panels: the header menu and, on docs pages, the Contents list. Only one is
  // open at a time; a link tap, Escape or a tap elsewhere closes it.
  var panels = [];
  function panel(button, list, onOpen) {
    if (!button || !list) return;
    var p = {
      close: function () { list.classList.remove("open"); button.setAttribute("aria-expanded", "false"); }
    };
    button.addEventListener("click", function () {
      var open = !list.classList.contains("open");
      panels.forEach(function (o) { o.close(); });
      if (open) {
        list.classList.add("open");
        button.setAttribute("aria-expanded", "true");
        if (onOpen) onOpen();
      }
    });
    [].forEach.call(list.querySelectorAll("a"), function (a) { a.addEventListener("click", p.close); });
    p.owns = function (el) { return button.contains(el) || list.contains(el); };
    panels.push(p);
  }
  var siteNav = document.querySelector(".site-head .nav");
  panel(document.querySelector(".menu-btn"), siteNav);
  // "Get the app" opens a small menu with both stores. It lives inside the phone-width
  // menu, so opening it must not close that menu: it is handled apart from `panels`.
  var getBtn = document.querySelector(".get-app .cta-link");
  var getMenu = document.querySelector(".get-menu");
  if (getBtn && getMenu) {
    var closeGet = function () { getMenu.classList.remove("open"); getBtn.setAttribute("aria-expanded", "false"); };
    getBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      var open = !getMenu.classList.contains("open");
      getMenu.classList.toggle("open", open);
      getBtn.setAttribute("aria-expanded", open ? "true" : "false");
      if (open) { var first = getMenu.querySelector("a"); if (first) first.focus(); }
    });
    document.addEventListener("click", function (e) { if (!getMenu.contains(e.target)) closeGet(); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeGet(); });
  }
  // Report form: the e-mail subject carries the reporter's subject, and a picture over
  // the form service's 10 MB limit is refused here rather than failing after sending.
  var report = document.getElementById("report-form");
  if (report) {
    report.addEventListener("submit", function (e) {
      var err = report.querySelector(".err");
      var files = report.querySelector('input[type="file"]').files || [];
      var total = 0;
      for (var i = 0; i < files.length; i++) total += files[i].size;
      if (total > 10 * 1024 * 1024) {
        e.preventDefault();
        err.textContent = "The pictures add up to more than 10 MB. Choose fewer or smaller ones.";
        return;
      }
      var subj = report.querySelector('[name="subject"]').value.trim();
      report.querySelector('[name="_subject"]').value = "ProcessChem bug report: " + subj;
      err.textContent = "";
    });
  }
  var docsList = document.querySelector(".docs-nav-list");
  panel(document.querySelector(".docs-toggle"), docsList, function () {
    var cur = docsList.querySelector('[aria-current="page"]');
    if (cur) docsList.scrollTop = cur.offsetTop - docsList.clientHeight / 2;
  });
  document.addEventListener("click", function (e) {
    panels.forEach(function (p) { if (!p.owns(e.target)) p.close(); });
  });
  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape") panels.forEach(function (p) { p.close(); });
  });
})();
