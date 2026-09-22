/**
 * Re:Nnocent Lab 캠페인 신청서 -> Google Sheet 연동용 Apps Script
 *
 * 연결 대상 스프레드시트:
 * https://docs.google.com/spreadsheets/d/1G9jSkHvkb1bbCF7WjW002bVNFDv-gORrHluQ_EnzHYY/edit
 *
 * [기존 코드 대비 수정 사항]
 * - 한글이 깨져 들어가는 문제 수정: e.postData.contents 대신
 *   e.postData.getDataAsString('UTF-8')로 명시적으로 UTF-8 디코딩하도록 변경
 * - setupHeaders() 함수 추가: 실행 한 번으로 시트1 1행에 헤더를 넣어줌
 * - doPost 실행 시 시트가 완전히 비어있으면 헤더를 자동으로 먼저 넣도록 처리
 *
 * 설치 / 재설치 방법
 * 1) 스프레드시트를 열고 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 2) 기존 Code.gs 내용을 전부 지우고 이 파일 내용 전체를 붙여넣기 후 저장(Ctrl+S)
 * 3) 상단 함수 선택 드롭다운에서 "setupHeaders" 선택 후 ▶ 실행
 *    - 처음 실행 시 권한 승인 창이 뜨면 본인 계정으로 허용
 *    - 실행 후 시트1 1행에 헤더가 채워짐
 * 4) [배포] > [배포 관리] 클릭 -> 기존 배포 옆 연필(수정) 아이콘 클릭
 *    -> 버전: "새 버전" 선택 -> [배포] 클릭
 *    (※ "새 배포"가 아니라 반드시 기존 배포를 "수정"해야 웹 앱 URL이 그대로 유지됩니다)
 * 5) URL은 그대로이므로 index.html 파일들의 APPLY_ENDPOINT는 다시 바꿀 필요 없음
 */

var SPREADSHEET_ID = "1G9jSkHvkb1bbCF7WjW002bVNFDv-gORrHluQ_EnzHYY";
var SHEET_NAME = "시트1";

var HEADERS = [
  "접수시각", "타입구분", "고료구분", "주차선택",
  "이름", "인스타그램", "휴대폰", "이메일",
  "우편번호", "주소", "상세주소", "요청사항"
];

// Apps Script 편집기에서 이 함수를 한 번 실행하면 시트1 1행에 헤더가 채워집니다.
function setupHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
}

function doPost(e) {
  try {
    // e.postData.contents 대신 getDataAsString('UTF-8')을 사용해야
    // 한글/이모지 등 멀티바이트 문자가 깨지지 않습니다.
    var rawBody = e.postData.getDataAsString("UTF-8");
    var data = JSON.parse(rawBody);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.guideType || "",   // 타입구분 (예: A Type)
      data.feeTier || "",     // 고료구분 (예: 5, 10, ..., 고료조정)
      data.week || "",        // 주차선택 (예: 10월 4주차(10/19~10/25))
      data.name || "",
      data.insta || "",
      data.phone || "",
      data.email || "",
      data.zip || "",
      data.addr || "",
      data.addr2 || "",
      data.note || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 배포 후 브라우저로 웹 앱 URL을 직접 열어보면 이 함수가 실행되어
// 정상 배포 여부를 간단히 확인할 수 있습니다.
function doGet(e) {
  return ContentService.createTextOutput("OK - Re:Nnocent Lab apply endpoint is running.");
}
