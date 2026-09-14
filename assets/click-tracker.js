/**
 * 리포트 원문 링크(.source a) 클릭을 기사 단위로 집계합니다.
 * - GA4: outbound_article_click 이벤트
 * - 구글 시트: sheetWebhookUrl로 POST (Apps Script 웹앱)
 */
(function () {
  var cfg = window.AI_TREND_ANALYTICS || {};
  var gaId = (cfg.gaId || "").trim();
  var sheetUrl = (cfg.sheetWebhookUrl || "").trim();
  if (!gaId && !sheetUrl) return;

  function getReportDate() {
    var title = document.title || "";
    var m = title.match(/\((\d{4}-\d{2}-\d{2})\)/);
    if (m) return m[1];
    var meta = document.querySelector(".meta");
    if (meta) {
      m = (meta.textContent || "").match(/(\d{4}-\d{2}-\d{2})/);
      if (m) return m[1];
    }
    return "";
  }

  function getArticleTitle(anchor) {
    var item = anchor.closest(".item");
    if (!item) return "";
    var h3 = item.querySelector("h3");
    if (!h3) return "";
    var clone = h3.cloneNode(true);
    var badges = clone.querySelectorAll(".badge");
    for (var i = 0; i < badges.length; i++) badges[i].remove();
    return (clone.textContent || "").replace(/\s+/g, " ").trim();
  }

  // Asia/Seoul 로컬 시각: 2026-09-14 11:12:44
  function seoulTimestamp(date) {
    var d = date || new Date();
    try {
      var parts = new Intl.DateTimeFormat("en-CA", {
        timeZone: "Asia/Seoul",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false
      }).formatToParts(d);
      var map = {};
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].type !== "literal") map[parts[i].type] = parts[i].value;
      }
      return map.year + "-" + map.month + "-" + map.day + " " + map.hour + ":" + map.minute + ":" + map.second;
    } catch (err) {
      var kst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
      return kst.toISOString().replace("T", " ").replace(/\.\d{3}Z$/, "");
    }
  }

  function initGa() {
    if (!gaId || window.__aiTrendGaReady) return;
    window.__aiTrendGaReady = true;
    window.dataLayer = window.dataLayer || [];
    window.gtag = window.gtag || function () { window.dataLayer.push(arguments); };
    window.gtag("js", new Date());
    window.gtag("config", gaId, { send_page_view: true });

    var s = document.createElement("script");
    s.async = true;
    s.src = "https://www.googletagmanager.com/gtag/js?id=" + encodeURIComponent(gaId);
    document.head.appendChild(s);
  }

  function sendGa(payload) {
    if (!gaId || typeof window.gtag !== "function") return;
    window.gtag("event", "outbound_article_click", {
      article_url: payload.article_url,
      article_title: payload.article_title,
      report_date: payload.report_date,
      transport_type: "beacon"
    });
  }

  function sendSheet(payload) {
    if (!sheetUrl) return;
    try {
      if (navigator.sendBeacon) {
        var blob = new Blob([JSON.stringify(payload)], { type: "text/plain" });
        navigator.sendBeacon(sheetUrl, blob);
        return;
      }
    } catch (e) {}
    fetch(sheetUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(payload),
      keepalive: true
    }).catch(function () {});
  }

  initGa();

  document.addEventListener(
    "click",
    function (e) {
      var a = e.target && e.target.closest ? e.target.closest("p.source a[href]") : null;
      if (!a) return;
      var href = a.href;
      if (!href || href.indexOf("http") !== 0) return;

      var payload = {
        ts: seoulTimestamp(new Date()),
        report_date: getReportDate(),
        article_title: getArticleTitle(a).slice(0, 180),
        article_url: href
      };

      sendGa(payload);
      sendSheet(payload);
    },
    true
  );
})();
