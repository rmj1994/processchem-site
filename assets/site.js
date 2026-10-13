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

  // Mobile menu
  var btn = document.querySelector(".menu-btn");
  var nav = document.querySelector(".site-head .nav");
  if (btn && nav) {
    btn.addEventListener("click", function () {
      var open = nav.classList.toggle("open");
      btn.setAttribute("aria-expanded", open ? "true" : "false");
    });
    [].forEach.call(nav.querySelectorAll("a"), function (a) {
      a.addEventListener("click", function () {
        nav.classList.remove("open");
        btn.setAttribute("aria-expanded", "false");
      });
    });
  }
})();
