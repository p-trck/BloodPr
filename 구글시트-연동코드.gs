/**
 * 혈압 관리 앱 - 구글 시트 연동용 Google Apps Script
 * 설치 방법은 README.md 참고
 */
const SHEET_NAME = '혈압기록';
const HEADER = ['ID','날짜','시간','시간대','수축기','이완기','맥박','복약','메모','단계'];
const SLOT_KO = {morning:'아침', day:'낮', evening:'저녁'};
const SLOT_EN = {'아침':'morning', '낮':'day', '저녁':'evening'};

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

/** 앱에서 데이터를 보낼 때 (전체 교체 저장) */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action !== 'replace') throw new Error('unknown action');
    const sh = getSheet_();
    sh.clearContents();
    sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    const rows = (data.records || []).map(function(r) {
      return [
        r.id,
        r.dt.slice(0, 10),
        r.dt.slice(11, 16),
        SLOT_KO[r.slot] || r.slot,
        r.sys, r.dia, r.pulse,
        r.med ? 'O' : '',
        r.memo || '',
        r.stage || ''
      ];
    });
    if (rows.length) {
      // 날짜/시간 열을 텍스트 서식으로 (자동 변환 방지)
      sh.getRange(2, 1, rows.length, 3).setNumberFormat('@');
      sh.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
    }
    return json_({ ok: true, count: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 앱에서 데이터를 가져갈 때 */
function doGet(e) {
  try {
    const sh = getSheet_();
    const values = sh.getDataRange().getValues();
    const records = [];
    for (let i = 1; i < values.length; i++) {
      const v = values[i];
      if (!v[0]) continue;
      records.push({
        id: String(v[0]),
        dt: fmtDate_(v[1]) + 'T' + fmtTime_(v[2]),
        sys: Number(v[4]), dia: Number(v[5]), pulse: Number(v[6]),
        slot: SLOT_EN[String(v[3])] || 'day',
        med: String(v[7]) === 'O' ? 1 : 0,
        memo: String(v[8] || '')
      });
    }
    return json_({ ok: true, records: records });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v);
}
function fmtTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v).slice(0, 5);
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
