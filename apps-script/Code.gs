/**
 * Re:Nnocent Lab 캠페인 신청서 -> Google Sheet 연동용 Apps Script
 *
 * 연결 대상 스프레드시트:
 * https://docs.google.com/spreadsheets/d/1G9jSkHvkb1bbCF7WjW002bVNFDv-gORrHluQ_EnzHYY/edit
 *
 * 설치 방법
 * 1) 위 스프레드시트를 열고 상단 메뉴 [확장 프로그램] > [Apps Script] 클릭
 * 2) 기본으로 열린 Code.gs 내용을 전부 지우고 이 파일 내용을 붙여넣기
 * 3) 저장(Ctrl+S) 후 상단 [배포] > [새 배포] 클릭
 * 4) 유형 선택(톱니바퀴 아이콘)에서 "웹 앱" 선택
 *    - 설명: 아무 이름(예: rennocent-apply)
 *    - 다음 사용자로 실행: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 5) [배포] 클릭 -> 권한 승인(본인 계정으로 허용) -> 발급된 "웹 앱 URL"(...../exec) 복사
 * 6) 각 index.html 파일의 <script> 안에 있는
 *      var APPLY_ENDPOINT = "";
 *    이 부분에 복사한 URL을 붙여넣기 (type-a ~ type-g 총 7개 파일)
 *
 * 스프레드시트 "시트1"의 헤더(1행)를 아래 순서로 미리 넣어두는 것을 권장합니다.
 * 접수시각 | 타입구분 | 고료구분 | 주차선택 | 이름 | 인스타그램 | 휴대폰 | 이메일 | 우편번호 | 주소 | 상세주소 | 요청사항
 *
 * 코드를 수정한 뒤에는 반드시 [배포] > [배포 관리] > 연필 아이콘 > [새 버전]으로
 * 다시 배포해야 실제 웹 앱에 반영됩니다. (저장만으로는 반영되지 않습니다)
 */

var SPREADSHEET_ID = "1G9jSkHvkb1bbCF7WjW002bVNFDv-gORrHluQ_EnzHYY";
var SHEET_NAME = "시트1";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

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
