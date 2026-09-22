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
 * - 휴대폰(G열) / 우편번호(I열) 앞자리 0 소실 방지:
 *   appendRow는 숫자로만 된 문자열을 자동으로 숫자로 바꿔버려 앞자리 0이
 *   사라지는 경우가 있어(예: "06035" -> "6035"), 대상 행을 직접 계산해
 *   해당 셀에 먼저 "일반 텍스트" 서식(setNumberFormat("@"))을 지정하고
 *   flush()로 확정한 뒤에 값을 쓰도록 변경. 휴대폰은 index.html 쪽에서도
 *   항상 010-1234-5678 형태(하이픈 포함)로 정규화해서 보내므로 이중으로 방지됨
 * - 중복 제출 방지: 같은 사람(휴대폰+타입구분+주차선택)이 최근 10분 안에
 *   이미 접수된 상태로 다시 요청이 오면 새 행을 추가하지 않고 그대로
 *   성공 응답만 반환 (느린 응답 때문에 사용자가 제출 버튼을 여러 번 눌러도
 *   시트에 같은 신청이 중복으로 쌓이지 않도록 함)
 * - LockService로 동시 요청을 순차 처리해 경쟁 상태(race condition) 방지
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
 * 6) (한 번만) 함수 선택 드롭다운에서 "cleanupDuplicateRows" 선택 후 ▶ 실행
 *    -> 지금까지 쌓인 중복 신청 행이 정리됩니다. 실행 전 시트 사본을
 *       만들어두면 안전합니다 (파일 > 사본 만들기)
 */

var SPREADSHEET_ID = "1G9jSkHvkb1bbCF7WjW002bVNFDv-gORrHluQ_EnzHYY";
var SHEET_NAME = "시트1";

var HEADERS = [
  "접수시각", "타입구분", "고료구분", "주차선택",
  "이름", "인스타그램", "휴대폰", "이메일",
  "우편번호", "주소", "상세주소", "요청사항"
];

// 같은 사람의 중복 제출로 판단할 시간 범위 (밀리초)
var DUPLICATE_WINDOW_MS = 10 * 60 * 1000; // 10분
// 중복 검사 시 최근 몇 행까지 살펴볼지 (성능을 위해 상한)
var DUPLICATE_CHECK_MAX_ROWS = 100;

// 휴대폰(G열), 우편번호(I열)를 "일반 텍스트" 서식으로 고정해
// 앞자리 0이 숫자 변환으로 사라지는 것을 방지합니다.
function forceTextColumns(sheet) {
  sheet.getRange("G:G").setNumberFormat("@");
  sheet.getRange("I:I").setNumberFormat("@");
}

// Apps Script 편집기에서 이 함수를 한 번 실행하면
// 시트1 1행에 헤더가 채워지고, 휴대폰/우편번호 컬럼이 텍스트 서식으로 고정됩니다.
function setupHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  forceTextColumns(sheet);
}

// 같은 휴대폰 + 타입구분 + 주차선택 조합이 최근 DUPLICATE_WINDOW_MS 안에
// 이미 접수되어 있으면 true를 반환합니다.
function isDuplicateSubmission(sheet, data) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // 헤더뿐이거나 빈 시트

  var rowsToCheck = Math.min(DUPLICATE_CHECK_MAX_ROWS, lastRow - 1);
  var startRow = lastRow - rowsToCheck + 1;
  var values = sheet.getRange(startRow, 1, rowsToCheck, 7).getValues(); // A~G열
  var now = new Date().getTime();

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var ts = row[0] instanceof Date ? row[0].getTime() : null;
    if (ts !== null && (now - ts) > DUPLICATE_WINDOW_MS) {
      // 시간순으로 쌓이므로, 이보다 위(과거) 행들은 전부 더 오래된 데이터입니다.
      break;
    }
    var rowGuideType = row[1];
    var rowWeek = row[3];
    var rowPhone = row[6];
    if (rowGuideType === (data.guideType || "") &&
        rowWeek === (data.week || "") &&
        rowPhone === (data.phone || "")) {
      return true;
    }
  }
  return false;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: "server busy, try again" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

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

    if (isDuplicateSubmission(sheet, data)) {
      // 이미 접수된 신청입니다. 새 행을 추가하지 않고 성공으로 응답만 합니다.
      return ContentService
        .createTextOutput(JSON.stringify({ result: "success", duplicate: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // appendRow는 숫자로만 된 문자열을 자동으로 숫자로 바꿔버려 앞자리 0을
    // 지워버리는 경우가 있어, 대상 행을 직접 계산해 해당 셀에 먼저
    // "일반 텍스트" 서식을 지정하고 flush로 확정한 뒤 값을 씁니다.
    var targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 7).setNumberFormat("@"); // 휴대폰
    sheet.getRange(targetRow, 9).setNumberFormat("@"); // 우편번호
    SpreadsheetApp.flush();

    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([[
      new Date(),
      data.guideType || "",   // 타입구분 (예: A Type)
      data.feeTier || "",     // 고료구분 (예: 5, 10, ..., 고료조정)
      data.week || "",        // 주차선택 (예: 10월 4주차(10/19~10/25))
      data.name || "",
      data.insta || "",
      String(data.phone || ""),
      data.email || "",
      String(data.zip || ""),
      data.addr || "",
      data.addr2 || "",
      data.note || ""
    ]]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 배포 후 브라우저로 웹 앱 URL을 직접 열어보면 이 함수가 실행되어
// 정상 배포 여부를 간단히 확인할 수 있습니다.
function doGet(e) {
  return ContentService.createTextOutput("OK - Re:Nnocent Lab apply endpoint is running.");
}

// ---------------------------------------------------------------
// 지금까지 쌓인 중복 신청 행을 한 번에 정리하는 함수입니다.
// Apps Script 편집기에서 함수 드롭다운 -> "cleanupDuplicateRows" 선택 -> ▶ 실행
// (같은 타입구분 + 주차선택 + 휴대폰 조합이 여러 번 나오면, 가장 먼저
//  들어온 행 1개만 남기고 나머지는 전부 삭제합니다. 실행 전 시트를
//  사본으로 백업해두는 것을 권장합니다: 파일 > 사본 만들기)
// ---------------------------------------------------------------
function cleanupDuplicateRows() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    Logger.log("정리할 데이터가 없습니다.");
    return;
  }

  var numRows = lastRow - 1; // 헤더 제외
  var values = sheet.getRange(2, 1, numRows, HEADERS.length).getValues();

  var seen = {};
  var rowsToDelete = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var guideType = row[1];
    var week = row[3];
    var phone = row[6];
    var key = guideType + "||" + week + "||" + phone;

    if (seen[key]) {
      rowsToDelete.push(i + 2); // 실제 시트 행 번호 (헤더가 1행이므로 +2)
    } else {
      seen[key] = true;
    }
  }

  // 아래(나중 행)부터 위로 삭제해야 행 번호가 밀리지 않습니다.
  for (var j = rowsToDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(rowsToDelete[j]);
  }

  Logger.log("중복 " + rowsToDelete.length + "개 행을 삭제했습니다.");
}
