/**
 * 구글 시트 원문 링크 클릭 로그 + 기사별 요약
 *
 * 시트 구성:
 * - 로그: 클릭 원본 (클릭시각, 리포트일, 기사제목, 원문URL)
 * - 요약: 한눈에 보는 집계 (순위, 리포트일, 기사제목, 클릭수, 주간클릭비중%, 원문URL, 최근클릭)
 *
 * 설정:
 * 1) 이 코드를 Apps Script에 붙여넣고 저장
 * 2) 배포 > 배포 관리 > 수정(연필) > 새 버전 > 배포
 * 3) (기존 로그가 있으면) 편집기에서 rebuildSummary 함수 실행 1회
 *
 * 참고: 지금 집계는 '클릭수/클릭비중'입니다.
 * 진짜 클릭률(CTR=클릭/페이지뷰)은 페이지뷰 수집이 추가로 필요합니다.
 */
var LOG_HEADERS = ["클릭시각", "리포트일", "기사제목", "원문URL"];
var SUMMARY_HEADERS = ["순위", "리포트일", "기사제목", "클릭수", "주간클릭비중(%)", "원문URL", "최근클릭"];

function toSeoulTimestamp(value) {
  var tz = "Asia/Seoul";
  var d = value ? new Date(value) : new Date();
  if (isNaN(d.getTime())) {
    if (typeof value === "string" && value.indexOf("-") > -1) return value;
    d = new Date();
  }
  return Utilities.formatDate(d, tz, "yyyy-MM-dd HH:mm:ss");
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("로그");
  if (sheet) return sheet;

  // 기존 첫 시트를 로그로 사용(이미 데이터가 있는 경우)
  sheet = ss.getSheets()[0];
  if (sheet.getName() !== "로그" && sheet.getName() !== "요약") {
    try { sheet.setName("로그"); } catch (e) {}
  }
  if (sheet.getName() !== "로그") {
    sheet = getOrCreateSheet_(ss, "로그");
  }
  return sheet;
}

function ensureLogHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LOG_HEADERS);
    return;
  }
  var first = String(sheet.getRange(1, 1).getValue() || "");
  // 영문 헤더로 시작한 기존 시트는 그대로 두고 데이터만 이어서 쌓음
  if (!first) sheet.appendRow(LOG_HEADERS);
}

function doPost(e) {
  var logSheet = getLogSheet_();
  ensureLogHeader_(logSheet);

  var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
  var data = {};
  try {
    data = JSON.parse(raw);
  } catch (err) {
    data = {};
  }

  logSheet.appendRow([
    toSeoulTimestamp(data.ts),
    data.report_date || "",
    data.article_title || "",
    data.article_url || ""
  ]);

  rebuildSummary_();

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput("ai-trend-weekly click logger ok")
    .setMimeType(ContentService.MimeType.TEXT);
}

/** 편집기에서 수동 실행: 기존 로그로 요약 시트 재생성 */
function rebuildSummary() {
  rebuildSummary_();
}

function rebuildSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = getLogSheet_();
  var summary = getOrCreateSheet_(ss, "요약");

  var lastRow = logSheet.getLastRow();
  var map = {}; // key = report_date + \t + url
  var reportTotals = {};

  if (lastRow >= 2) {
    var values = logSheet.getRange(2, 1, lastRow, 4).getValues();
    for (var i = 0; i < values.length; i++) {
      var ts = String(values[i][0] || "");
      var reportDate = String(values[i][1] || "");
      var title = String(values[i][2] || "");
      var url = String(values[i][3] || "");
      if (!url && !title) continue;

      var key = reportDate + "\t" + url;
      if (!map[key]) {
        map[key] = {
          report_date: reportDate,
          article_title: title,
          article_url: url,
          clicks: 0,
          last_ts: ts
        };
      }
      map[key].clicks += 1;
      if (title) map[key].article_title = title;
      if (ts > map[key].last_ts) map[key].last_ts = ts;
      reportTotals[reportDate] = (reportTotals[reportDate] || 0) + 1;
    }
  }

  var rows = [];
  for (var k in map) {
    if (!map.hasOwnProperty(k)) continue;
    var item = map[k];
    var total = reportTotals[item.report_date] || 0;
    var share = total ? Math.round((item.clicks / total) * 1000) / 10 : 0;
    rows.push([
      0, // rank placeholder
      item.report_date,
      item.article_title,
      item.clicks,
      share,
      item.article_url,
      item.last_ts
    ]);
  }

  // 리포트일 내림차순, 클릭수 내림차순
  rows.sort(function (a, b) {
    if (a[1] === b[1]) return b[3] - a[3];
    return a[1] < b[1] ? 1 : -1;
  });

  var rankByReport = {};
  for (var r = 0; r < rows.length; r++) {
    var rd = rows[r][1];
    rankByReport[rd] = (rankByReport[rd] || 0) + 1;
    rows[r][0] = rankByReport[rd];
  }

  summary.clear();
  summary.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]);
  summary.getRange(1, 1, 1, SUMMARY_HEADERS.length).setFontWeight("bold");

  if (rows.length) {
    summary.getRange(2, 1, rows.length, SUMMARY_HEADERS.length).setValues(rows);
  }

  summary.setFrozenRows(1);
  summary.autoResizeColumns(1, SUMMARY_HEADERS.length);
  // 요약 시트를 앞으로
  try { ss.setActiveSheet(summary); ss.moveActiveSheet(1); } catch (e2) {}
}
