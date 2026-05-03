/**
 * 벽돌깨기 게임 - Google Apps Script 백엔드
 *
 * [설정 방법]
 * 1. Google Drive에서 "새 Google 스프레드시트" 생성
 * 2. 스프레드시트 URL에서 ID 복사 (https://docs.google.com/spreadsheets/d/[이 부분]/edit)
 * 3. 아래 SPREADSHEET_ID에 붙여넣기
 * 4. 확장 프로그램 > Apps Script 클릭
 * 5. 이 코드 전체를 붙여넣기
 * 6. 배포 > 새 배포 > 유형: 웹 앱
 *    - 실행 주체: 본인
 *    - 액세스 권한: 누구나
 * 7. 승인 > 배포
 * 8. 배포된 URL을 게임 페이지에 입력
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const action = (e.parameter && e.parameter.action) || body.action;
    const table = (e.parameter && e.parameter.table) || body.table;

    let result;
    switch(action) {
      case 'init':   result = initSheet(); break;
      case 'select': result = selectRows(table, body); break;
      case 'insert': result = insertRow(table, body); break;
      default:       result = { data: null, error: 'Unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ data: null, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function initSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('scores');
  if (!sheet) {
    sheet = ss.insertSheet('scores');
    sheet.getRange(1, 1, 1, 5).setValues([['id', 'name', 'score', 'photo', 'played_at']]);
    sheet.setFrozenRows(1);
  }
  return { data: 'scores 시트 초기화 완료!', error: null };
}

function selectRows(tableName, body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tableName);
  if (!sheet) return { data: [], error: null };

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { data: [], error: null };

  const headers = data[0];
  let rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] === '' ? null : row[i]; });
    return obj;
  });

  if (body.order) {
    const col = body.order.column;
    const asc = body.order.ascending;
    rows.sort((a, b) => {
      const va = a[col] || 0;
      const vb = b[col] || 0;
      return asc ? (va > vb ? 1 : -1) : (va < vb ? 1 : -1);
    });
  }

  if (body.limit) rows = rows.slice(0, body.limit);

  return { data: rows, error: null };
}

function insertRow(tableName, body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(tableName);
  if (!sheet) {
    sheet = ss.insertSheet(tableName);
    sheet.getRange(1, 1, 1, 5).setValues([['id', 'name', 'score', 'photo', 'played_at']]);
    sheet.setFrozenRows(1);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const record = body.records || {};
  if (!record.id) record.id = Utilities.getUuid();

  const row = headers.map(h => record[h] !== undefined ? record[h] : '');
  sheet.appendRow(row);

  return { data: record, error: null };
}
