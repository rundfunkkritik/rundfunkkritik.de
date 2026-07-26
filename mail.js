/* Setzt E-Mail-Adressen erst im Browser zusammen, damit sie nicht im
   Quelltext stehen. Ohne JavaScript zeigt das <noscript> die Adresse in der
   Form "kontakt (at) domain" - lesbar, aber kein gueltiges Sammler-Muster. */
(function () {
  function r(s) { return s.split("").reverse().join(""); }
  function bau(el) {
    var n = el.getAttribute("data-n"), d = el.getAttribute("data-d");
    if (!n || !d) { return; }
    var adr = r(n) + String.fromCharCode(64) + r(d);
    var a = document.createElement("a");
    a.href = "mailto:" + adr;
    var pre = el.getAttribute("data-pre") || "";
    var post = el.getAttribute("data-post") || "";
    a.textContent = el.getAttribute("data-t") || (pre + adr + post);
    var st = el.getAttribute("style");
    if (st) { a.setAttribute("style", st); }
    var cl = el.getAttribute("data-c");
    if (cl) { a.setAttribute("class", cl); }
    while (el.firstChild) { el.removeChild(el.firstChild); }
    el.appendChild(a);
  }
  function los() {
    var l = document.querySelectorAll("span.e-mail");
    for (var i = 0; i < l.length; i++) { bau(l[i]); }
  }
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", los);
  } else {
    los();
  }
})();
