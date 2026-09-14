/**
 * 구글 시트 원문 링크 클릭 로그용 Apps Script
 *
 * 설정 방법:
 * 1) 구글 시트 새 문서 생성
 * 2) 확장 프로그램 > Apps Script 에 이 코드 붙여넣기
 * 3) doPost 저장 후 배포 > 새 배포 > 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 웹 앱 URL을 assets/analytics-config.js 의 sheetWebhookUrl 에 넣기
 * 5) 시트 1행에 헤더 자동 생성됨: ts / report_date / article_title / article_url
 * 6) 피벗 테이블로 article_url 기준 COUNTA 하면 기사별 클릭 수
 */
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ts", "report_date", "article_title", "article_url"]);
  }

  var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
  var data = {};
  try {
    data = JSON.parse(raw);
  } catch (err) {
    data = { raw: raw };
  }

  sheet.appendRow([
    data.ts || new Date().toISOString(),
    data.report_date || "",
    data.article_title || "",
    data.article_url || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput("ai-trend-weekly click logger ok")
    .setMimeType(ContentService.MimeType.TEXT);
}
