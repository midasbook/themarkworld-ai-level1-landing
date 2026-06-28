/**
 * 원데이 전자책쓰기 7월 과정 신청 DB 저장용 Google Apps Script
 *
 * 사용 방법
 * 1. Google Sheets에서 새 스프레드시트를 만듭니다.
 * 2. 확장 프로그램 > Apps Script를 열고 이 코드를 붙여 넣습니다.
 * 3. setupApplicationSheet()를 한 번 실행해 헤더를 만듭니다.
 * 4. 배포 > 새 배포 > 웹 앱을 선택합니다.
 * 5. 실행 사용자: 나, 액세스 권한: 모든 사용자로 배포합니다.
 * 6. 발급된 웹 앱 URL을 index.html의 GOOGLE_SCRIPT_URL 값에 넣습니다.
 */

const SPREADSHEET_ID = "";
const SHEET_NAME = "신청 DB";

const HEADERS = [
  "접수시각",
  "과정명",
  "성함",
  "연락처",
  "희망 반",
  "이메일",
  "전자책 주제",
  "랜딩페이지 URL",
  "처리상태",
  "메모"
];

function setupApplicationSheet() {
  const sheet = getApplicationSheet_();
  ensureHeaders_(sheet);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
}

function doGet() {
  return json_({
    ok: true,
    message: "원데이 전자책쓰기 신청 저장 앱스크립트가 정상 동작 중입니다."
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getApplicationSheet_();
    ensureHeaders_(sheet);

    const data = e && e.parameter ? e.parameter : {};
    const submittedAt = data.submittedAt
      ? new Date(data.submittedAt)
      : new Date();

    sheet.appendRow([
      submittedAt,
      data.course || "원데이 전자책쓰기 7월 과정",
      data.name || "",
      data.phone || "",
      data.classType || "",
      data.email || "",
      data.topic || "",
      data.pageUrl || "",
      "신규",
      ""
    ]);

    return json_({ ok: true });
  } catch (error) {
    return json_({
      ok: false,
      error: error.message
    });
  } finally {
    lock.releaseLock();
  }
}

function getApplicationSheet_() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error("스프레드시트를 찾을 수 없습니다. SPREADSHEET_ID를 입력하거나 시트에 바인딩된 Apps Script에서 실행해 주세요.");
  }

  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = firstRow.some(Boolean);

  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground("#082b61")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
  }
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
