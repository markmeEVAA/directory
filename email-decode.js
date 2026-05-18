// EVAA email obfuscation — DECODER.
// Each public roster page generates email cells like:
//   <span class="evaa-email" data-e="dXNlckBldmFhc3BvcnRzLm9yZw=="></span>
// The data-e attribute is a Base64-encoded email address (NOT secret encryption,
// just a way to keep the address out of the raw HTML source so basic email-
// harvesting bots that don't execute JavaScript can't scrape addresses by
// grepping for "@" or "mailto:".)
//
// This script runs on page load, decodes each span, and replaces it with a
// real clickable mailto link. Modern browsers (including the iframe context
// inside SportsEngine) handle this transparently. Headless bots with full JS
// engines can still get the addresses; this stops the common case.
(function () {
  function decodeAll() {
    var spans = document.querySelectorAll(".evaa-email[data-e]");
    for (var i = 0; i < spans.length; i++) {
      var el = spans[i];
      if (el.dataset.decoded === "1") continue;
      try {
        var addr = atob(el.dataset.e);
        var a = document.createElement("a");
        a.href = "mailto:" + addr;
        a.textContent = addr;
        el.innerHTML = "";
        el.appendChild(a);
        el.dataset.decoded = "1";
      } catch (e) {
        // bad data; leave the span empty
      }
    }
  }
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", decodeAll);
  } else {
    decodeAll();
  }
})();
