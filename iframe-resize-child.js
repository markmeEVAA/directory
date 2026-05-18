// EVAA iframe auto-resize — CHILD script.
// Loaded by each per-sport roster page (baseball.html, basketball.html, …).
// Reports the page's true content height to the parent (SE) so the iframe
// can be sized exactly right and scrolling="no" can be applied.
(function () {
  function postHeight() {
    if (window.parent === window) return; // not in an iframe
    var doc = document.documentElement;
    var body = document.body;
    var h = Math.max(
      body.scrollHeight, body.offsetHeight,
      doc.scrollHeight, doc.offsetHeight, doc.clientHeight
    );
    window.parent.postMessage({ source: "evaa-roster", type: "height", height: h }, "*");
  }
  window.addEventListener("DOMContentLoaded", postHeight);
  window.addEventListener("load", postHeight);
  window.addEventListener("resize", postHeight);
  if (window.ResizeObserver) {
    new ResizeObserver(postHeight).observe(document.body);
  }
  // Also report after table sort clicks etc.
  document.addEventListener("click", function () { setTimeout(postHeight, 50); });
})();
