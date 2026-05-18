// EVAA iframe auto-resize — HOST script.
// Paste a single <script src=".../iframe-resize-host.js"></script> into the
// SportsEngine CODE page element (after the iframe). The script listens for
// height reports from any EVAA roster iframe and resizes it to fit, so the
// iframe shows the full content with no internal scrollbar.
(function () {
  window.addEventListener("message", function (e) {
    var data = e.data;
    if (!data || data.source !== "evaa-roster" || data.type !== "height") return;
    if (typeof data.height !== "number" || data.height <= 0) return;

    // Find the iframe whose contentWindow sent this message.
    var iframes = document.querySelectorAll(
      'iframe[src*="markmeevaa.github.io/directory/"]'
    );
    for (var i = 0; i < iframes.length; i++) {
      if (iframes[i].contentWindow === e.source) {
        iframes[i].style.height = data.height + "px";
        iframes[i].setAttribute("scrolling", "no");
        return;
      }
    }
  });
})();
